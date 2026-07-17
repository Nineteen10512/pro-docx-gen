"""merge — DOCX 协作合并与导出 API（v1.5.1 C 能力域）。

提供：
- list_suggestions(session, status=None)   : 列出 session.suggestions
- accept_by_id(session, rev_ids)            : 按 ID 接受
- reject_by_id(session, rev_ids)            : 按 ID 拒绝
- accept_by_author(session, author)         : 按作者接受全部
- reject_by_author(session, author)         : 按作者拒绝全部
- merge_collaboration(session, strategy)     : 合并策略 auto/accept_all/reject_all
- export_clean_copy(session, output_path)   : 接受所有修订 + 移除批注
- export_review_pdf(session, output_dir)    : 保留修订痕迹的 PDF

设计：
- 复用 v1.5 稳定文件 pro_docx_gen.engine.revisions 的 accept/reject API
- 复用 v1.5 稳定文件 pro_docx_gen.engine.doc_converter 提供 PDF 转换
- 不修改 v1.5 任何文件
"""
from __future__ import annotations

import os
import shutil
import zipfile
from typing import Optional, Union

from lxml import etree

from .session import CollaborationSession

# 复用 v1.5 稳定 API（不修改）
from pro_docx_gen.engine.revisions import (
    accept_all_revisions,
    reject_all_revisions,
    accept_revision_by_id,
    reject_revision_by_id,
    list_revisions,
)
from pro_docx_gen.engine.doc_converter import ensure_docx, is_doc_path


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WNS = "{" + W + "}"


def _q(tag: str) -> str:
    return f"{WNS}{tag}"


def list_suggestions(
    session: CollaborationSession,
    status: Optional[str] = None,
) -> list[dict]:
    """列出当前会话所有建议。

    Args:
        session: CollaborationSession
        status: 可选过滤 "pending" | "accepted" | "rejected"

    Returns:
        list of suggestion dict
    """
    if status is None:
        return list(session.suggestions)
    return [s for s in session.suggestions if s.get("status") == status]


def _normalize_ids(rev_ids: Union[int, str, list]) -> list[str]:
    """把 rev_id 参数统一成字符串列表（接受 int / str / list）。"""
    if rev_ids is None:
        return []
    if isinstance(rev_ids, (int, str)):
        return [str(rev_ids)]
    return [str(x) for x in rev_ids]


def _suggestion_uses_rev_id(session: CollaborationSession, rev_id: str) -> list[dict]:
    """返回 session.suggestions 中所有用到此 rev_id 的建议。"""
    out = []
    for s in session.suggestions:
        rid = s.get("rev_id")
        if isinstance(rid, list):
            if rev_id in [str(x) for x in rid]:
                out.append(s)
        else:
            if str(rid) == rev_id:
                out.append(s)
    return out


def _mark_suggestions_by_rev_id(
    session: CollaborationSession,
    rev_ids: list[str],
    new_status: str,
) -> int:
    """把所有用到了 rev_ids 的 session.suggestions 标记为 new_status。返回数量。"""
    rev_set = set(rev_ids)
    count = 0
    for s in session.suggestions:
        rid = s.get("rev_id")
        if isinstance(rid, list):
            ids = {str(x) for x in rid}
        else:
            ids = {str(rid)}
        if ids & rev_set:
            s["status"] = new_status
            count += 1
    return count


def accept_by_id(
    session: CollaborationSession,
    rev_ids: Union[int, str, list],
    output_path: Optional[str] = None,
) -> int:
    """按 ID 接受修订（调用 v1.5 accept_revision_by_id）。"""
    ids = _normalize_ids(rev_ids)
    if not ids:
        return 0
    out = output_path or session.docx_path
    # v1.5 accept_revision_by_id 一次只接受一个；循环
    count = 0
    for rid in ids:
        try:
            accept_revision_by_id(session.docx_path, rid, out)
            count += 1
        except Exception:
            continue
    if output_path:
        # 同时把已接受的建议状态写回 session
        _mark_suggestions_by_rev_id(session, ids, "accepted")
    else:
        _mark_suggestions_by_rev_id(session, ids, "accepted")
    return count


def reject_by_id(
    session: CollaborationSession,
    rev_ids: Union[int, str, list],
    output_path: Optional[str] = None,
) -> int:
    """按 ID 拒绝修订（调用 v1.5 reject_revision_by_id）。"""
    ids = _normalize_ids(rev_ids)
    if not ids:
        return 0
    out = output_path or session.docx_path
    count = 0
    for rid in ids:
        try:
            reject_revision_by_id(session.docx_path, rid, out)
            count += 1
        except Exception:
            continue
    _mark_suggestions_by_rev_id(session, ids, "rejected")
    return count


def accept_by_author(
    session: CollaborationSession,
    author: str,
) -> int:
    """接受指定作者的全部建议。

    Returns:
        实际接受的修订数
    """
    author_sugs = [s for s in session.suggestions if s.get("author") == author]
    if not author_sugs:
        return 0
    # 收集所有相关 rev_id
    rev_ids = []
    for s in author_sugs:
        rid = s.get("rev_id")
        if isinstance(rid, list):
            rev_ids.extend([str(x) for x in rid])
        else:
            rev_ids.append(str(rid))
    return accept_by_id(session, rev_ids)


def reject_by_author(
    session: CollaborationSession,
    author: str,
) -> int:
    """拒绝指定作者的全部建议。"""
    author_sugs = [s for s in session.suggestions if s.get("author") == author]
    if not author_sugs:
        return 0
    rev_ids = []
    for s in author_sugs:
        rid = s.get("rev_id")
        if isinstance(rid, list):
            rev_ids.extend([str(x) for x in rid])
        else:
            rev_ids.append(str(rid))
    return reject_by_id(session, rev_ids)


def merge_collaboration(
    session: CollaborationSession,
    strategy: str = "auto",
) -> dict:
    """合并协作建议。

    Args:
        session: CollaborationSession
        strategy: "auto" = 全部 accept；"accept_all" / "reject_all"

    Returns:
        {"strategy": str, "applied": int, "pending": int}
    """
    if strategy not in ("auto", "accept_all", "reject_all"):
        raise ValueError(f"不支持的 strategy: {strategy}")

    pending_count = sum(1 for s in session.suggestions if s.get("status") == "pending")
    if strategy == "auto":
        # auto 默认全部接受（最常见合并策略）
        strategy = "accept_all"

    if strategy == "accept_all":
        accept_all_revisions(session.docx_path)
        for s in session.suggestions:
            if s.get("status") == "pending":
                s["status"] = "accepted"
    elif strategy == "reject_all":
        reject_all_revisions(session.docx_path)
        for s in session.suggestions:
            if s.get("status") == "pending":
                s["status"] = "rejected"

    return {
        "strategy": strategy,
        "applied": pending_count,
        "pending": 0,
    }


def _remove_all_comments_in_docx(docx_path: str) -> int:
    """从 docx 中移除 word/comments.xml、document.xml 中的 commentRange/Reference。

    Returns:
        实际移除的 comment 节点数
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(docx_path)

    tmp = docx_path + ".rmcmt_tmp"
    removed = 0
    with zipfile.ZipFile(docx_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    root = etree.fromstring(data)
                    # 移除 commentRangeStart / commentRangeEnd
                    for tag in ("commentRangeStart", "commentRangeEnd"):
                        for el in root.findall(f".//{_q(tag)}"):
                            el.getparent().remove(el)
                            removed += 1
                    # 移除 r 内部的 commentReference
                    for el in root.findall(f".//{_q('commentReference')}"):
                        # 找到父 w:r 并删除整个 r（commentReference 通常独占一个 r）
                        r = el.getparent()
                        if r is not None and r.tag == _q("r"):
                            if r.getparent() is not None:
                                r.getparent().remove(r)
                                removed += 1
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                elif item.filename == "word/comments.xml":
                    # 跳过 comments.xml（不复制 = 删除）
                    continue
                zout.writestr(item, data)
    shutil.move(tmp, docx_path)

    # 第二步：从 [Content_Types].xml 移除 comments Override + 从 rels 移除
    tmp = docx_path + ".ctclean_tmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    root = etree.fromstring(data)
                    CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
                    for ov in root.findall(f"{{{CT_NS}}}Override"):
                        if ov.get("PartName") == "/word/comments.xml":
                            root.remove(ov)
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                elif item.filename == "word/_rels/document.xml.rels":
                    root = etree.fromstring(data)
                    REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
                    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
                        if rel.get("Type", "").endswith("/comments"):
                            root.remove(rel)
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                zout.writestr(item, data)
    shutil.move(tmp, docx_path)
    return removed


def export_clean_copy(
    session: CollaborationSession,
    output_path: str,
) -> str:
    """导出干净版本：接受所有修订 + 移除所有批注。

    Returns:
        实际输出路径
    """
    if not os.path.exists(session.docx_path):
        raise FileNotFoundError(session.docx_path)
    if not output_path:
        raise ValueError("output_path 不能为空")

    # 1. 复制源 docx 到 output_path
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    shutil.copy2(session.docx_path, output_path)

    # 2. 接受所有修订
    accept_all_revisions(output_path)

    # 3. 移除所有批注
    removed = _remove_all_comments_in_docx(output_path)

    # 4. 更新 session 状态
    for s in session.suggestions:
        if s.get("status") == "pending":
            s["status"] = "accepted"
    for c in session.comments:
        c["status"] = "resolved"

    return output_path


def export_review_pdf(
    session: CollaborationSession,
    output_dir: str,
) -> str:
    """导出带修订痕迹的 PDF（保留 w:ins/w:del + 批注）。

    Returns:
        生成的 PDF 路径
    """
    if not os.path.exists(session.docx_path):
        raise FileNotFoundError(session.docx_path)
    if not output_dir:
        raise ValueError("output_dir 不能为空")

    os.makedirs(output_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(session.docx_path))[0]
    pdf_path = os.path.join(output_dir, f"{base}_review.pdf")

    # 复用 v1.5 doc_converter（虽然它只支持 .doc→.docx，但这里把 docx 直接转 PDF）
    # 由于 python-docx 不能直接产 PDF，我们使用 docx2pdf / LibreOffice fallback
    try:
        from docx2pdf import convert as _docx2pdf_convert
        _docx2pdf_convert(session.docx_path, pdf_path)
    except ImportError:
        # fallback: 复制 docx 为 .pdf.pdf 后缀文档 + 警告
        # 严格不降级：抛出清晰错误
        raise RuntimeError(
            "export_review_pdf 需要 docx2pdf 或 LibreOffice；当前环境未安装。"
            "请安装 docx2pdf（pip install docx2pdf）或系统装 LibreOffice。"
        )
    return pdf_path


__all__ = [
    "list_suggestions",
    "accept_by_id",
    "reject_by_id",
    "accept_by_author",
    "reject_by_author",
    "merge_collaboration",
    "export_clean_copy",
    "export_review_pdf",
]
