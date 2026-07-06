"""comments — DOCX 协作批注 API（v1.5.3 C 能力域）。

- add_review_comment(session, paragraph_index, text, parent_id=None) : 追加批注
- reply_comment(session, parent_comment_id, text) : thread 回复
- list_comments(session) : 列出当前会话所有批注

实现：
- 操作 word/comments.xml part（OXML 2013+ 支持 w15:parentId thread）
- 段落中插入 w:commentRangeStart + w:commentRangeEnd + w:commentReference
"""
from __future__ import annotations

import os
import shutil
import zipfile
from typing import Optional

from lxml import etree

from .session import CollaborationSession

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
WNS = "{" + W + "}"
W15NS = "{" + W15 + "}"
NSMAP = {"w": W, "w15": W15}


def _q(tag: str) -> str:
    return f"{WNS}{tag}"


def _q15(tag: str) -> str:
    return f"{W15NS}{tag}"


def _read_xml_part(docx_path: str, part_name: str) -> Optional[etree._Element]:
    """读指定 part（如 'word/comments.xml'），不存在返回 None。"""
    with zipfile.ZipFile(docx_path, "r") as zf:
        if part_name not in zf.namelist():
            return None
        data = zf.read(part_name)
    return etree.fromstring(data)


def _write_xml_part_replace(
    docx_path: str,
    part_name: str,
    new_root: etree._Element,
    extra_new_parts: Optional[dict[str, bytes]] = None,
) -> None:
    """写回指定 part（替换或新增）。extra_new_parts 用于新增多个 part。"""
    new_data = etree.tostring(new_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = docx_path + ".cmt_tmp"
    extras = extra_new_parts or {}
    with zipfile.ZipFile(docx_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            written = set()
            for item in zin.infolist():
                if item.filename == part_name:
                    zout.writestr(item, new_data)
                    written.add(part_name)
                else:
                    zout.writestr(item, zin.read(item.filename))
                    written.add(item.filename)
            # 写 extras（如果还没写）
            for name, data in extras.items():
                if name not in written:
                    zout.writestr(name, data)
    shutil.move(tmp, docx_path)


def _ensure_content_type(docx_path: str, part_name: str, content_type: str) -> None:
    """确保 [Content_Types].xml 包含 part_name 的 Override。"""
    tmp = docx_path + ".ct_tmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        ct_data = zin.read("[Content_Types].xml")
        ct_root = etree.fromstring(ct_data)
        CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
        # 查 partname
        ns = "{" + CT_NS + "}"
        found = False
        for override in ct_root.findall(f"{ns}Override"):
            if override.get("PartName") == "/" + part_name:
                found = True
                break
        if not found:
            ov = etree.SubElement(ct_root, f"{ns}Override")
            ov.set("PartName", "/" + part_name)
            ov.set("ContentType", content_type)
        new_ct = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "[Content_Types].xml":
                    zout.writestr(item, new_ct)
                else:
                    zout.writestr(item, zin.read(item.filename))
    shutil.move(tmp, docx_path)


def _ensure_relationship(
    docx_path: str,
    rels_part: str,
    rel_type: str,
    target: str,
) -> str:
    """在 rels_part（如 'word/_rels/document.xml.rels'）追加 relationship。

    Returns:
        rId 字符串
    """
    tmp = docx_path + ".rel_tmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        if rels_part in zin.namelist():
            data = zin.read(rels_part)
            root = etree.fromstring(data)
        else:
            # 创建空 rels
            PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
            root = etree.Element(f"{{{PKG_NS}}}Relationships", nsmap={None: PKG_NS})
        REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
        rels = root
        # 算下一个 rId
        used = set()
        for r in rels.findall(f"{{{REL_NS}}}Relationship"):
            rid = r.get("Id", "")
            if rid.startswith("rId"):
                try:
                    used.add(int(rid[3:]))
                except ValueError:
                    pass
        next_id = max(used, default=0) + 1
        new_rid = f"rId{next_id}"
        # 检查是否已存在（按 target + type）
        for r in rels.findall(f"{{{REL_NS}}}Relationship"):
            if r.get("Type") == rel_type and r.get("Target") == target:
                return r.get("Id", new_rid)
        rel = etree.SubElement(rels, f"{{{REL_NS}}}Relationship")
        rel.set("Id", new_rid)
        rel.set("Type", rel_type)
        rel.set("Target", target)
        new_data = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == rels_part:
                    zout.writestr(item, new_data)
                else:
                    zout.writestr(item, zin.read(item.filename))
            if rels_part not in [i.filename for i in zin.infolist()]:
                zout.writestr(rels_part, new_data)
    shutil.move(tmp, docx_path)
    return new_rid


def _next_comment_id(session: CollaborationSession) -> int:
    used = {c.get("id") for c in session.comments}
    base = max(used, default=-1) + 1
    return base


def _get_current_author(session: CollaborationSession) -> tuple[str, str]:
    if not session.collaborators:
        from datetime import datetime
        return "AI-Comment", datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    last = session.collaborators[-1]
    return last.author, last.joined_at


def _read_or_create_comments_root(docx_path: str) -> etree._Element:
    """读 word/comments.xml，没有则创建空 root。"""
    root = _read_xml_part(docx_path, "word/comments.xml")
    if root is None:
        root = etree.Element(_q("comments"), nsmap={"w": W})
    return root


def _add_comment_marker_to_paragraph(
    docx_path: str,
    paragraph_index: int,
    cid: int,
) -> None:
    """在指定段落的 OOXML 中插入 commentRangeStart/End + commentReference。"""
    tmp = docx_path + ".mk_tmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        doc_data = zin.read("word/document.xml")
        doc_root = etree.fromstring(doc_data)
        body = doc_root.find(_q("body"))
        if body is None:
            raise RuntimeError("document.xml 缺少 body 节点")
        p_count = 0
        target = None
        for child in body:
            if child.tag == _q("p"):
                if p_count == paragraph_index:
                    target = child
                    break
                p_count += 1
        if target is None:
            raise IndexError(f"段落索引超出范围: {paragraph_index}")
        # 插入 commentRangeStart（pPr 之后）
        insert_idx = 0
        for i, c in enumerate(target):
            if c.tag == _q("pPr"):
                insert_idx = i + 1
                break
        crs = etree.Element(_q("commentRangeStart"))
        crs.set(_q("id"), str(cid))
        target.insert(insert_idx, crs)
        # commentRangeEnd + commentReference
        cre = etree.Element(_q("commentRangeEnd"))
        cre.set(_q("id"), str(cid))
        target.append(cre)
        ref_run = etree.SubElement(target, _q("r"))
        rpr = etree.SubElement(ref_run, _q("rPr"))
        rstyle = etree.SubElement(rpr, _q("rStyle"))
        rstyle.set(_q("val"), "CommentReference")
        cref = etree.SubElement(ref_run, _q("commentReference"))
        cref.set(_q("id"), str(cid))
        new_doc = etree.tostring(doc_root, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_doc)
                else:
                    zout.writestr(item, zin.read(item.filename))
    shutil.move(tmp, docx_path)


def add_review_comment(
    session: CollaborationSession,
    paragraph_index: int,
    text: str,
    parent_id: Optional[int] = None,
) -> dict:
    """在指定段落追加批注（含可选 parent_id thread）。"""
    if not os.path.exists(session.docx_path):
        raise FileNotFoundError(f"会话文档不存在: {session.docx_path}")
    if not text:
        raise ValueError("comment text 不能为空")

    comments_root = _read_or_create_comments_root(session.docx_path)
    author, date = _get_current_author(session)
    cid = _next_comment_id(session)

    # 构造 w:comment
    ce = etree.SubElement(comments_root, _q("comment"))
    ce.set(_q("id"), str(cid))
    ce.set(_q("author"), author)
    ce.set(_q("date"), date)
    ce.set(_q("initials"), author[:2] if author else "PD")
    if parent_id is not None:
        ce.set(_q15("parentId"), str(parent_id))
    cp = etree.SubElement(ce, _q("p"))
    cr = etree.SubElement(cp, _q("r"))
    ct = etree.SubElement(cr, _q("t"))
    ct.text = text
    ct.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    # 写 comments.xml
    _write_xml_part_replace(session.docx_path, "word/comments.xml", comments_root)
    # 更新 [Content_Types].xml
    _ensure_content_type(
        session.docx_path,
        "word/comments.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    )
    # 写关系
    _ensure_relationship(
        session.docx_path,
        "word/_rels/document.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        "comments.xml",
    )
    # 段落中加 commentReference
    _add_comment_marker_to_paragraph(session.docx_path, paragraph_index, cid)

    cmt = {
        "id": cid,
        "author": author,
        "date": date,
        "text": text,
        "paragraph_index": paragraph_index,
        "parent_id": parent_id,
        "status": "open",
    }
    session.add_comment(cmt)
    return cmt


def reply_comment(
    session: CollaborationSession,
    parent_comment_id: int,
    text: str,
) -> dict:
    """回复父批注（追加为 thread 子批注，parent_id 关联）。"""
    # 找一个段落放 thread reply（默认附在父批注的段落）
    parent = next((c for c in session.comments if c.get("id") == parent_comment_id), None)
    if parent is None:
        raise ValueError(f"找不到父批注 id={parent_comment_id}")
    return add_review_comment(
        session,
        paragraph_index=parent.get("paragraph_index", 0),
        text=text,
        parent_id=parent_comment_id,
    )


def list_comments(session: CollaborationSession) -> list[dict]:
    """列出当前会话所有批注（含 thread reply）。"""
    return list(session.comments)


__all__ = [
    "add_review_comment",
    "reply_comment",
    "list_comments",
]
