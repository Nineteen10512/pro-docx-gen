"""diff — DOCX 双版本对比 API（v1.5.3 C 能力域）。

提供：
- diff_documents(path_a, path_b, output_path, author="AI-Diff") : 双版本对比
- 返回新 docx 路径 + diff 元数据

实现策略：
- 提取两个 docx 的段落纯文本（按 w:p 顺序，单元格内段落也计算）
- 用 difflib.SequenceMatcher 做段落级 diff
- 差异段落标记：
  - 仅 A 有 → w:del author=author（标记为删除）
  - 仅 B 有 → w:ins author=author（标记为新增）
  - 共有且内容不同 → 行内 w:del + w:ins
- 以 B 为基准，输出新 docx（保留 B 的所有结构 + diff 标记）
"""
from __future__ import annotations

import difflib
import os
import shutil
import zipfile
from typing import Optional

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WNS = "{" + W + "}"
NSMAP = {"w": W}


def _q(tag: str) -> str:
    return f"{WNS}{tag}"


def _read_docx_paragraphs(docx_path: str) -> list[str]:
    """读 docx 的所有段落纯文本（含表格单元格内段落）。"""
    if not os.path.exists(docx_path):
        raise FileNotFoundError(docx_path)
    with zipfile.ZipFile(docx_path, "r") as zf:
        if "word/document.xml" not in zf.namelist():
            raise RuntimeError(f"{docx_path} 不是有效 docx（缺少 word/document.xml）")
        data = zf.read("word/document.xml")
    root = etree.fromstring(data)
    paragraphs = []
    for p in root.iter(_q("p")):
        texts = []
        for t in p.iter(_q("t")):
            if t.text:
                texts.append(t.text)
        paragraphs.append("".join(texts))
    return paragraphs


def _read_document_xml(docx_path: str) -> etree._Element:
    with zipfile.ZipFile(docx_path, "r") as zf:
        data = zf.read("word/document.xml")
    return etree.fromstring(data)


def _mark_paragraph_as_inserted(p: etree._Element, author: str, date: str, rev_id: int) -> None:
    """把段落内容包成 w:ins（标记为新增）。"""
    runs = p.findall(_q("r"))
    if runs:
        for r in runs:
            ins = etree.Element(_q("ins"))
            ins.set(_q("id"), str(rev_id))
            ins.set(_q("author"), author)
            ins.set(_q("date"), date)
            idx = list(p).index(r)
            p.remove(r)
            ins.append(r)
            p.insert(idx, ins)
            rev_id += 1
    # 段落标记（pPr/rPr/ins）也加，让整段带"新增"标记（即使没 w:r）
    pPr = p.find(_q("pPr"))
    if pPr is None:
        pPr = etree.Element(_q("pPr"))
        p.insert(0, pPr)
    rPr = pPr.find(_q("rPr"))
    if rPr is None:
        rPr = etree.SubElement(pPr, _q("rPr"))
    ins_marker = etree.SubElement(rPr, _q("ins"))
    ins_marker.set(_q("id"), str(rev_id))
    ins_marker.set(_q("author"), author)
    ins_marker.set(_q("date"), date)


def _mark_paragraph_as_deleted(p: etree._Element, author: str, date: str, rev_id: int) -> None:
    """把段落内 w:t 改名为 w:delText 并用 w:del 包裹。"""
    runs = p.findall(_q("r"))
    if not runs:
        # 段落标记（pPr/rPr/del）也加
        pPr = p.find(_q("pPr"))
        if pPr is None:
            pPr = etree.SubElement(p, _q("pPr"))
            p.insert(0, pPr)
        rPr = pPr.find(_q("rPr"))
        if rPr is None:
            rPr = etree.SubElement(pPr, _q("rPr"))
        d_marker = etree.SubElement(rPr, _q("del"))
        d_marker.set(_q("id"), str(rev_id))
        d_marker.set(_q("author"), author)
        d_marker.set(_q("date"), date)
        return

    # 把段落里所有 w:r 用一个 w:del 包起来
    d = etree.Element(_q("del"))
    d.set(_q("id"), str(rev_id))
    d.set(_q("author"), author)
    d.set(_q("date"), date)
    for r in runs:
        for t in r.findall(_q("t")):
            t.tag = _q("delText")
        idx = list(p).index(r)
        p.remove(r)
        d.append(r)
    p.append(d)


def _write_docx_from_root(
    src_path: str,
    dst_path: str,
    new_root: etree._Element,
) -> None:
    """以 src 为模板，把新 document.xml 写回 dst（保留其他 part）。"""
    os.makedirs(os.path.dirname(dst_path) or ".", exist_ok=True)
    new_data = etree.tostring(new_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = dst_path + ".diff_write_tmp"
    with zipfile.ZipFile(src_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_data)
                else:
                    zout.writestr(item, zin.read(item.filename))
    shutil.move(tmp, dst_path)


def _enable_track_changes(docx_path: str) -> None:
    """在 docx settings.xml 启用 trackChanges。"""
    tmp = docx_path + ".diff_tc_tmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            modified = False
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/settings.xml":
                    root = etree.fromstring(data)
                    if root.find(_q("trackChanges")) is None:
                        etree.SubElement(root, _q("trackChanges"))
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                    modified = True
                zout.writestr(item, data)
            if not modified:
                settings_xml = (
                    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    f'<w:settings xmlns:w="{W}"><w:trackChanges/></w:settings>'
                ).encode("utf-8")
                zout.writestr("word/settings.xml", settings_xml)
    shutil.move(tmp, docx_path)


def diff_documents(
    path_a: str,
    path_b: str,
    output_path: str,
    author: str = "AI-Diff",
) -> dict:
    """对比两个 docx 文档，输出带修订标记的新 docx。

    Args:
        path_a: 原版本路径
        path_b: 新版本路径
        output_path: 输出 docx 路径
        author: diff 标记的 author（默认 "AI-Diff"）

    Returns:
        {
            "output_path": str,
            "stats": {"added": int, "removed": int, "unchanged": int, "modified": int},
            "details": [{"op": "equal"|"insert"|"delete"|"replace", "a": str, "b": str}]
        }

    Raises:
        FileNotFoundError: 任意 docx 不存在
        RuntimeError: 解析失败
    """
    if not os.path.exists(path_a):
        raise FileNotFoundError(f"path_a 不存在: {path_a}")
    if not os.path.exists(path_b):
        raise FileNotFoundError(f"path_b 不存在: {path_b}")
    if not output_path:
        raise ValueError("output_path 不能为空")

    # 1. 读两版段落
    paras_a = _read_docx_paragraphs(path_a)
    paras_b = _read_docx_paragraphs(path_b)

    # 2. SequenceMatcher 段落级 diff
    matcher = difflib.SequenceMatcher(a=paras_a, b=paras_b, autojunk=False)
    opcodes = matcher.get_opcodes()

    # 3. 复制 B 为模板，按 diff 标记
    shutil.copy2(path_b, output_path)
    root = _read_document_xml(output_path)
    body = root.find(_q("body"))
    if body is None:
        raise RuntimeError("path_b 文档结构异常：缺少 body 节点")

    # 收集 body 内的直接子段落（顶层 w:p，不包括表格里的）
    top_paras = [c for c in body if c.tag == _q("p")]

    from datetime import datetime
    date = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    rev_id = 50000

    details = []
    stats = {"added": 0, "removed": 0, "unchanged": 0, "modified": 0}

    for tag, i1, i2, j1, j2 in opcodes:
        if tag == "equal":
            stats["unchanged"] += (i2 - i1)
            for k in range(i1, i2):
                details.append({"op": "equal", "a": paras_a[k], "b": paras_b[j1 + (k - i1)]})
        elif tag == "delete":
            # 标记 A 中这些段落为删除
            for k in range(i1, i2):
                if k < len(top_paras):
                    _mark_paragraph_as_deleted(top_paras[k], author, date, rev_id)
                    rev_id += 1
                stats["removed"] += 1
                details.append({"op": "delete", "a": paras_a[k] if k < len(paras_a) else "", "b": ""})
        elif tag == "insert":
            # 标记 B 中这些段落为新增
            for k in range(j1, j2):
                if k < len(top_paras):
                    _mark_paragraph_as_inserted(top_paras[k], author, date, rev_id)
                    rev_id += 1
                stats["added"] += 1
                details.append({"op": "insert", "a": "", "b": paras_b[k] if k < len(paras_b) else ""})
        elif tag == "replace":
            # 把 A 的段落标记为删除 + B 的段落标记为新增（最直观）
            n = max(i2 - i1, j2 - j1)
            for k in range(n):
                if k < (i2 - i1) and (i1 + k) < len(top_paras):
                    _mark_paragraph_as_deleted(top_paras[i1 + k], author, date, rev_id)
                    rev_id += 1
                    stats["removed"] += 1
                if k < (j2 - j1) and (j1 + k) < len(top_paras):
                    _mark_paragraph_as_inserted(top_paras[j1 + k], author, date, rev_id)
                    rev_id += 1
                    stats["added"] += 1
                stats["modified"] += 1
                details.append({
                    "op": "replace",
                    "a": paras_a[i1 + k] if (i1 + k) < len(paras_a) else "",
                    "b": paras_b[j1 + k] if (j1 + k) < len(paras_b) else "",
                })

    # 4. 写回
    _write_docx_from_root(output_path, output_path, root)
    # 5. 开启 trackChanges 让修订可见
    _enable_track_changes(output_path)

    return {
        "output_path": output_path,
        "stats": stats,
        "details": details,
    }


__all__ = ["diff_documents"]
