"""suggest — DOCX 协作建议 API（v1.5.1 C 能力域）。

基于 OOXML w:ins / w:del 节点构造三种修订建议：
- suggest_insert(session, paragraph_index, text)   : 在指定段落后追加插入建议
- suggest_replace(session, find_text, replace_text) : 全文替换建议（首个匹配）
- suggest_delete(session, paragraph_index)          : 删除整段建议

约束：
- author / date 来自当前协作者（session.collaborators 中最后一个）
- w:ins / w:del 必须含 w:id（全局唯一）
- 不修改 v1.5 稳定文件（pro_docx_gen.engine.renderer 等）
"""
from __future__ import annotations

import os
import shutil
import zipfile
from typing import Optional

from lxml import etree

from .session import CollaborationSession

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WNS = "{" + W + "}"
NSMAP = {"w": W}


def _q(tag: str) -> str:
    return f"{WNS}{tag}"


def _next_rev_id(session: CollaborationSession) -> int:
    """生成下一个修订 ID（在 session 范围内唯一）。"""
    base = len(session.suggestions) + 100
    used: set = set()
    for s in session.suggestions:
        rid = s.get("rev_id")
        if isinstance(rid, list):
            used.update(rid)
        elif rid is not None:
            used.add(rid)
    while base in used:
        base += 1
    return base


def _get_current_author(session: CollaborationSession) -> tuple[str, str]:
    """取当前协作者（最后一个）的 (author, date)。

    没有协作者时返回 ("AI-Suggest", ISO 时间)。
    """
    if not session.collaborators:
        from datetime import datetime
        return "AI-Suggest", datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    last = session.collaborators[-1]
    return last.author, last.joined_at


def _read_document_xml(docx_path: str) -> tuple[etree._ElementTree, etree._Element]:
    """读 word/document.xml 解析为 ElementTree。"""
    with zipfile.ZipFile(docx_path, "r") as zf:
        data = zf.read("word/document.xml")
    root = etree.fromstring(data)
    return etree.ElementTree(root), root


def _write_document_xml(
    docx_path: str,
    new_root: etree._Element,
) -> None:
    """把新 document.xml 写回 docx。"""
    tmp = docx_path + ".sug_tmp"
    new_data = etree.tostring(new_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(docx_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_data)
                else:
                    zout.writestr(item, zin.read(item.filename))
    shutil.move(tmp, docx_path)


def _find_paragraph_by_index(root: etree._Element, index: int) -> Optional[etree._Element]:
    """按正文段落索引找 w:p 元素（跳过表格单元格内段落不计）。"""
    body = root.find(_q("body"))
    if body is None:
        return None
    p_count = 0
    for child in body:
        if child.tag == _q("p"):
            if p_count == index:
                return child
            p_count += 1
    return None


def _make_run(text: str, run_props: Optional[dict] = None) -> etree._Element:
    """构造 w:r 元素（含 w:t）。"""
    r = etree.Element(_q("r"))
    if run_props:
        rpr = etree.SubElement(r, _q("rPr"))
        if run_props.get("bold"):
            etree.SubElement(rpr, _q("b"))
        if run_props.get("italic"):
            etree.SubElement(rpr, _q("i"))
    t = etree.SubElement(r, _q("t"))
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def _wrap_in_ins(run: etree._Element, author: str, date: str, rev_id: int) -> etree._Element:
    """把 w:r 包到 w:ins 中。"""
    ins = etree.Element(_q("ins"))
    ins.set(_q("id"), str(rev_id))
    ins.set(_q("author"), author)
    ins.set(_q("date"), date)
    ins.append(run)
    return ins


def _wrap_runs_in_del(paragraph: etree._Element, author: str, date: str, rev_id: int) -> Optional[etree._Element]:
    """把段落内所有 w:r 包到 w:del（替换为 w:delText）中。"""
    runs = paragraph.findall(_q("r"))
    if not runs:
        return None
    d = etree.Element(_q("del"))
    d.set(_q("id"), str(rev_id))
    d.set(_q("author"), author)
    d.set(_q("date"), date)
    for r in runs:
        # 把 w:t 改名为 w:delText
        for t in r.findall(_q("t")):
            t.tag = _q("delText")
        paragraph.remove(r)
        d.append(r)
    return d


def suggest_insert(
    session: CollaborationSession,
    paragraph_index: int,
    text: str,
) -> dict:
    """在指定段落后追加一段插入建议（w:ins）。"""
    if not os.path.exists(session.docx_path):
        raise FileNotFoundError(f"会话文档不存在: {session.docx_path}")
    if not text:
        raise ValueError("suggest_insert text 不能为空")

    _, root = _read_document_xml(session.docx_path)
    p = _find_paragraph_by_index(root, paragraph_index)
    if p is None:
        raise IndexError(f"段落索引超出范围: {paragraph_index}")

    author, date = _get_current_author(session)
    rev_id = _next_rev_id(session)

    # 构造新段落（包含 w:ins）
    new_p = etree.Element(_q("p"))
    pPr = etree.SubElement(new_p, _q("pPr"))
    pStyle = etree.SubElement(pPr, _q("pStyle"))
    pStyle.set(_q("val"), "Normal")
    run = _make_run(text)
    ins = _wrap_in_ins(run, author, date, rev_id)
    new_p.append(ins)

    # 插入到目标段落后
    p.addnext(new_p)
    _write_document_xml(session.docx_path, root)

    sug = {
        "id": f"sug-{rev_id}",
        "rev_id": rev_id,
        "type": "insert",
        "author": author,
        "date": date,
        "target": {"paragraph_index": paragraph_index},
        "text": text,
        "status": "pending",
    }
    session.add_suggestion(sug)
    return sug


def suggest_replace(
    session: CollaborationSession,
    find_text: str,
    replace_text: str,
) -> dict:
    """全文查找首个匹配 find_text 的段落，把 w:r 包成 w:del + w:ins。"""
    if not os.path.exists(session.docx_path):
        raise FileNotFoundError(f"会话文档不存在: {session.docx_path}")
    if not find_text:
        raise ValueError("find_text 不能为空")

    _, root = _read_document_xml(session.docx_path)
    body = root.find(_q("body"))
    if body is None:
        raise RuntimeError("document.xml 缺少 body 节点")

    author, date = _get_current_author(session)
    rev_id_base = _next_rev_id(session)

    target_para = None
    target_run = None
    for p in body.iter(_q("p")):
        for r in p.findall(_q("r")):
            t = r.find(_q("t"))
            if t is not None and t.text and find_text in t.text:
                target_para = p
                target_run = r
                break
        if target_para is not None:
            break
    if target_para is None:
        raise ValueError(f"未找到包含 {find_text!r} 的段落")

    # 把 target_run 拆成 3 段：[prefix, find_text, suffix]
    t = target_run.find(_q("t"))
    full = t.text
    idx = full.find(find_text)
    prefix = full[:idx]
    suffix = full[idx + len(find_text):]

    # 取出原 run 的 rPr 复制给新 runs
    rPr_orig = target_run.find(_q("rPr"))

    # 清空原 run
    target_run.clear()
    if rPr_orig is not None:
        target_run.append(etree.fromstring(etree.tostring(rPr_orig)))
    # 写 prefix
    if prefix:
        pt = etree.SubElement(target_run, _q("t"))
        pt.text = prefix
        pt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    # 构造 del + ins（用新的 w:r）
    del_run = etree.Element(_q("r"))
    if rPr_orig is not None:
        del_run.append(etree.fromstring(etree.tostring(rPr_orig)))
    del_text = etree.SubElement(del_run, _q("delText"))
    del_text.text = find_text
    del_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    d = etree.Element(_q("del"))
    d.set(_q("id"), str(rev_id_base))
    d.set(_q("author"), author)
    d.set(_q("date"), date)
    d.append(del_run)

    ins_run = _make_run(replace_text)
    ins = _wrap_in_ins(ins_run, author, date, rev_id_base + 1)

    # 在原 run 后面追加 del + ins
    target_run.addnext(d)
    d.addnext(ins)

    # 写 suffix（作为新的独立 run）
    if suffix:
        sfx_run = etree.Element(_q("r"))
        if rPr_orig is not None:
            sfx_run.append(etree.fromstring(etree.tostring(rPr_orig)))
        sfx_t = etree.SubElement(sfx_run, _q("t"))
        sfx_t.text = suffix
        sfx_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        ins.addnext(sfx_run)

    _write_document_xml(session.docx_path, root)

    sug = {
        "id": f"sug-{rev_id_base}",
        "rev_id": [rev_id_base, rev_id_base + 1],
        "type": "replace",
        "author": author,
        "date": date,
        "target": {"find": find_text, "para": target_para.get(_q("rsidR")) or ""},
        "text": {"find": find_text, "replace": replace_text},
        "status": "pending",
    }
    session.add_suggestion(sug)
    return sug


def suggest_delete(
    session: CollaborationSession,
    paragraph_index: int,
) -> dict:
    """把指定段落的所有 w:r 包到 w:del（标记为删除建议）。"""
    if not os.path.exists(session.docx_path):
        raise FileNotFoundError(f"会话文档不存在: {session.docx_path}")
    _, root = _read_document_xml(session.docx_path)
    p = _find_paragraph_by_index(root, paragraph_index)
    if p is None:
        raise IndexError(f"段落索引超出范围: {paragraph_index}")

    author, date = _get_current_author(session)
    rev_id = _next_rev_id(session)

    d = _wrap_runs_in_del(p, author, date, rev_id)
    if d is None:
        raise ValueError(f"段落 {paragraph_index} 为空，无法标记删除")
    # 把 w:del 追加到段落末尾（保留段落骨架，含 pPr）
    p.append(d)
    # 段落本身打上 w:rPr/w:del 标记（让整段变删除线）
    pPr = p.find(_q("pPr"))
    if pPr is None:
        pPr = etree.SubElement(p, _q("pPr"))
        p.insert(0, pPr)
    rPr = pPr.find(_q("rPr"))
    if rPr is None:
        rPr = etree.SubElement(pPr, _q("rPr"))
    del_marker = etree.SubElement(rPr, _q("del"))
    del_marker.set(_q("id"), str(rev_id + 1000))
    del_marker.set(_q("author"), author)
    del_marker.set(_q("date"), date)

    _write_document_xml(session.docx_path, root)

    sug = {
        "id": f"sug-{rev_id}",
        "rev_id": rev_id,
        "type": "delete",
        "author": author,
        "date": date,
        "target": {"paragraph_index": paragraph_index},
        "text": "",
        "status": "pending",
    }
    session.add_suggestion(sug)
    return sug


__all__ = [
    "suggest_insert",
    "suggest_replace",
    "suggest_delete",
]
