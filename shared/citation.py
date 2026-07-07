"""Citation formatter — 参考文献格式化引擎（v1.4 P1-4）。

支持 4 种引用格式：
- apa:       APA 7th
- gb7714:    GB/T 7714-2015（中文国标）
- mla:       MLA 9
- ieee:      IEEE

对外入口 format_reference(item, style, index=None) -> str。
item 字段：type(article/book/webpage/thesis/conference) + 对应字段。
"""
from __future__ import annotations


# 支持的引用格式
SUPPORTED_STYLES = {"apa", "gb7714", "mla", "ieee"}

# GB/T 7714 文献类型标识
_GB_TYPE_CODE = {
    "article": "J",      # 期刊
    "book": "M",         # 专著
    "webpage": "EB/OL",  # 电子资源
    "thesis": "D",       # 学位论文
    "conference": "C",   # 会议论文
    "report": "R",
    "standard": "S",
    "patent": "P",
}


def format_reference(item: dict, style: str = "apa", index: int | None = None) -> str:
    """格式化单条参考文献。

    Args:
        item: 参考文献条目 dict，包含 type + 对应字段。
        style: apa / gb7714 / mla / ieee。
        index: IEEE/GB 序号（1-based），None 时不追加编号（IEEE 会用 [1]）。
    """
    if not isinstance(item, dict):
        return str(item)
    style = (style or "apa").lower()
    if style not in SUPPORTED_STYLES:
        # 不抛异常，回退到 apa，避免渲染中断
        style = "apa"

    rtype = (item.get("type") or "article").lower()
    authors = item.get("authors", "") or ""
    title = item.get("title", "") or ""
    year = str(item.get("year", "") or "").strip()
    journal = item.get("journal", "") or item.get("source", "") or ""
    publisher = item.get("publisher", "") or ""
    place = item.get("place", "") or ""
    volume = item.get("volume", "") or ""
    issue = item.get("issue", "") or ""
    pages = item.get("pages", "") or ""
    doi = item.get("doi", "") or ""
    url = item.get("url", "") or ""
    accessed = item.get("accessed", "") or ""
    school = item.get("school", "") or ""
    degree = item.get("degree", "") or ""
    conference = item.get("conference", "") or item.get("source", "") or ""
    site = item.get("site", "") or ""
    isbn = item.get("isbn", "") or ""

    # DOI/URL 规范化
    doi_url = ""
    if doi:
        d = doi.strip()
        if d.startswith("http"):
            doi_url = d
        elif d.startswith("10."):
            doi_url = f"https://doi.org/{d}"
        else:
            doi_url = f"https://doi.org/{d}"

    if style == "apa":
        return _fmt_apa(rtype, authors, year, title, journal, publisher, place,
                        volume, issue, pages, doi_url, url, school, degree,
                        conference, site, accessed)
    if style == "gb7714":
        idx = index if index is not None else 0
        return _fmt_gb7714(idx, rtype, authors, year, title, journal, publisher, place,
                           volume, issue, pages, doi_url, url, school, degree,
                           conference, site, accessed, isbn)
    if style == "mla":
        return _fmt_mla(rtype, authors, year, title, journal, publisher, place,
                        volume, issue, pages, doi_url, url, school, degree,
                        conference, site, accessed)
    if style == "ieee":
        idx = index if index is not None else 1
        return _fmt_ieee(idx, rtype, authors, year, title, journal, publisher, place,
                         volume, issue, pages, doi_url, url, school, degree,
                         conference, site, accessed)
    return ""


# ─── 作者名格式化辅助 ────────────────────────────────────────────

def _apa_authors(s: str) -> str:
    """APA: '张三, 李四 and Wang, W. and Smith, J. K.' -> 'Zhang San, Li Si, W. Wang, & J. K. Smith'
    简化处理：如果含英文逗号分隔多作者，保留；中文作者原样；插入 Oxford 逗号 &。
    """
    s = s.strip()
    if not s:
        return ""
    # 用 " and " / "，" / ", " / ";" 切分
    import re
    parts = re.split(r"\s+and\s+|；|;", s)
    # 再按中文逗号/英文逗号切分（但英文逗号可能是 'Wang, W.' 内部的，需要保守处理）
    flat = []
    for p in parts:
        # 如果片段里含逗号后紧跟空格+大写字母+点（如 Wang, W.），则视为单一作者
        if re.search(r",\s*[A-Z]\.", p):
            flat.append(p.strip())
        else:
            for sub in re.split(r"[，,]", p):
                sub = sub.strip()
                if sub:
                    flat.append(sub)
    if len(flat) == 1:
        return flat[0]
    if len(flat) == 2:
        return f"{flat[0]} & {flat[1]}"
    return ", ".join(flat[:-1]) + f", & {flat[-1]}"


def _mla_authors(s: str) -> str:
    """MLA: 第一作者姓在前，后续作者名在前。简化实现。"""
    s = s.strip()
    if not s:
        return ""
    import re
    parts = re.split(r"\s+and\s+|；|;|[，,]", s)
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return f"{parts[0]}, and {parts[1]}"
    return ", ".join(parts[:-1]) + f", and {parts[-1]}"


def _ieee_authors(s: str) -> str:
    """IEEE: A. Author, B. Author, and C. Author。简化实现：英文按 'and' 切分，中文/原样保留。"""
    s = s.strip()
    if not s:
        return ""
    import re
    parts = re.split(r"\s+and\s+|；|;|[，,]", s)
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return f"{parts[0]} and {parts[1]}"
    return ", ".join(parts[:-1]) + f", and {parts[-1]}"


def _gb_authors(s: str, max_show: int = 3) -> str:
    """GB/T 7714: 作者列表，超过 max_show 取前 max_show 加 '等'/'et al'。"""
    s = s.strip()
    if not s:
        return ""
    import re
    parts = re.split(r"\s+and\s+|；|;|[，,]", s)
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) <= max_show:
        return ", ".join(parts)
    # 判断首位是否中文决定 et al. 还是 等
    import unicodedata
    first = parts[0]
    is_cn = any('\u4e00' <= ch <= '\u9fff' for ch in first)
    return ", ".join(parts[:max_show]) + (" 等" if is_cn else " et al")


# ─── APA ──────────────────────────────────────────────────────────

def _fmt_apa(rtype, authors, year, title, journal, publisher, place,
             volume, issue, pages, doi_url, url, school, degree,
             conference, site, accessed):
    a = _apa_authors(authors)
    pieces = []
    if a:
        pieces.append(f"{a} ({year})" if year else a)
    if title:
        pieces.append(title)

    if rtype == "article":
        src = journal
        vi = ""
        if volume:
            vi = f", {volume}"
            if issue:
                vi += f"({issue})"
        if pages:
            vi += f", {_apa_pages(pages)}"
        if src:
            pieces.append(f"{src}{vi}")
        elif vi:
            pieces.append(vi.lstrip(", "))
    elif rtype == "book":
        src_parts = []
        if place and publisher:
            src_parts.append(f"{place}: {publisher}")
        elif publisher:
            src_parts.append(publisher)
        if src_parts:
            pieces.append("".join(src_parts))
    elif rtype == "webpage":
        if site:
            pieces.append(site)
    elif rtype == "thesis":
        th = f"{degree} dissertation" if degree else "Thesis"
        if school:
            th += f", {school}"
        pieces.append(th)
    elif rtype == "conference":
        pieces.append(f"In Proceedings of {conference}" if conference else "Conference proceedings")
        if pages:
            pieces[-1] += f", {_apa_pages(pages)}"
        if place:
            pieces[-1] += f". {place}"

    tail = []
    if doi_url and rtype in ("article", "book", "conference"):
        tail.append(doi_url)
    if url and rtype in ("webpage", "thesis"):
        tail.append(url)
    if accessed and rtype == "webpage":
        tail.append(f"Retrieved {accessed}")

    out = ". ".join(pieces) + "."
    if tail:
        out = out.rstrip(".") + " " + " ".join(tail)
    return out


def _apa_pages(pages):
    # APA: pp. xx–yy （期刊 article 省略 pp. 前缀，直接数字）
    return pages


# ─── GB/T 7714 ────────────────────────────────────────────────────

def _fmt_gb7714(index, rtype, authors, year, title, journal, publisher, place,
               volume, issue, pages, doi_url, url, school, degree,
               conference, site, accessed, isbn):
    code = _GB_TYPE_CODE.get(rtype, "J")
    a = _gb_authors(authors)
    # [N] 作者. 题名[类型]. 出处, 年, 卷(期): 页码.
    parts = []
    if index > 0:
        parts.append(f"[{index}]")
    if a:
        parts.append(a + ".")
    if title:
        parts.append(f" {title}[{code}].")
    if rtype == "article":
        src = ""
        if journal:
            src = f" {journal}"
        if year:
            src += f", {year}"
        if volume:
            src += f", {volume}"
            if issue:
                src += f"({issue})"
        if pages:
            src += f": {pages}"
        parts.append(src + ".")
    elif rtype == "book":
        src = ""
        if place and publisher:
            src = f" {place}: {publisher}"
        elif publisher:
            src = f" {publisher}"
        if year:
            src += f", {year}"
        parts.append(src + ".")
    elif rtype == "webpage":
        src = ""
        if site:
            src += f" {site}"
        if accessed:
            src += f"[{accessed}]"
        if url:
            src += f". {url}"
        parts.append(src + ".")
    elif rtype == "thesis":
        src = ""
        if place and school:
            src += f" {place}: {school}"
        elif school:
            src += f" {school}"
        if year:
            src += f", {year}"
        parts.append(src + ".")
    elif rtype == "conference":
        src = ""
        if conference:
            src += f" {conference}"
        if place:
            src += f", {place}"
        if year:
            src += f", {year}"
        if pages:
            src += f": {pages}"
        parts.append(src + ".")
    return "".join(parts)


# ─── MLA ──────────────────────────────────────────────────────────

def _fmt_mla(rtype, authors, year, title, journal, publisher, place,
             volume, issue, pages, doi_url, url, school, degree,
             conference, site, accessed):
    a = _mla_authors(authors)
    pieces = []
    if a:
        pieces.append(a + ".")
    # Title in quotes for shorter works, italic for container (we approximate with quotes)
    if title:
        pieces.append(f' "{title}."')
    if rtype == "article":
        if journal:
            src = f" {journal}"
            if volume:
                src += f", vol. {volume}"
            if issue:
                src += f", no. {issue}"
            if year:
                src += f", {year}"
            if pages:
                src += f", pp. {pages}"
            pieces.append(src + ".")
    elif rtype == "book":
        if publisher:
            pieces.append(f" {publisher}")
            if year:
                pieces[-1] += f", {year}"
            pieces[-1] += "."
    elif rtype == "webpage":
        if site:
            pieces.append(f" {site}")
        if accessed:
            pieces[-1] += f", accessed {accessed}"
        if url:
            pieces[-1] += f". {url}"
        pieces[-1] += "."
    elif rtype == "thesis":
        tp = f" {degree} thesis" if degree else " Thesis"
        if school:
            tp += f", {school}"
        if year:
            tp += f", {year}"
        pieces.append(tp + ".")
    elif rtype == "conference":
        if conference:
            pieces.append(f" Proc. of {conference}")
        if year:
            pieces[-1] += f", {year}"
        if pages:
            pieces[-1] += f", pp. {pages}"
        pieces[-1] += "."
    if doi_url and rtype in ("article", "book"):
        pieces.append(f" {doi_url}")
    return "".join(pieces).strip()


# ─── IEEE ─────────────────────────────────────────────────────────

def _fmt_ieee(index, rtype, authors, year, title, journal, publisher, place,
              volume, issue, pages, doi_url, url, school, degree,
              conference, site, accessed):
    a = _ieee_authors(authors)
    parts = [f"[{index}]"]
    if a:
        parts.append(f" {a},")
    if title:
        parts.append(f' "{title},"')
    if rtype == "article":
        if journal:
            parts.append(f" {journal},")
        if volume:
            parts.append(f" vol. {volume},")
        if issue:
            parts.append(f" no. {issue},")
        if pages:
            parts.append(f" pp. {pages},")
        if year:
            parts.append(f" {year}.")
    elif rtype == "book":
        if place and publisher:
            parts.append(f" {place}: {publisher},")
        elif publisher:
            parts.append(f" {publisher},")
        if year:
            parts.append(f" {year}.")
    elif rtype == "webpage":
        if site:
            parts.append(f" {site}.")
        if year:
            parts.append(f" {year}.")
        if url:
            parts.append(f" {url}")
        if accessed:
            parts.append(f" [Accessed: {accessed}]")
    elif rtype == "thesis":
        tp = f" {degree} thesis" if degree else " Thesis"
        if school:
            tp += f", {school}"
        if year:
            tp += f", {year}"
        parts.append(tp + ".")
    elif rtype == "conference":
        if conference:
            parts.append(f" in Proc. {conference},")
        if place:
            parts.append(f" {place},")
        if year:
            parts.append(f" {year},")
        if pages:
            parts.append(f" pp. {pages}.")
    if doi_url and rtype in ("article", "book", "conference"):
        parts.append(f" doi: {doi_url.replace('https://doi.org/', '')}")
    return "".join(parts)


import re as _re

_RE_CJK = _re.compile(r'[\u4e00-\u9fff\u3000-\u303f\u3400-\u4dbf]')


def _contains_cjk(text: str) -> bool:
    """判断字符串是否包含 CJK（中/日/韩）字符。"""
    return bool(_RE_CJK.search(text or ""))


def _safe(v) -> str:
    """None-safe str() helper。"""
    if v is None:
        return ""
    return str(v)
