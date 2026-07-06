"""Parser — JSON 语义 → 扁平布局节点编译器。

接收高层文档 JSON（经过 validators 验证），展开为扁平的"待渲染节点"列表，
每个节点绑定对应的 design token 引用键（不解析为实际数值，renderer 负责查表）。

v1.2 新增：支持 revision/comment/footnote/endnote/watermark/equation/signature 等节点，
支持 doc-level watermark/page_border/header/footer 增强字段，支持 meta.page_setup。
"""

from typing import Any
import re

from .validators import validate_document

# v1.4 P1-4b: 模块级 meta 引用（parse_document 设置，_expand_node 引用）
_current_meta: dict = {}


# 内联交叉引用：匹配 {ref label_name}（仅在普通正文/quote/list/abstract/callout 文本中识别）
_RE_INLINE_REF = re.compile(r'\{ref\s+([a-zA-Z_][a-zA-Z0-9_-]*)\}')

# v1.4 P2-4b: inline {ref xxx} 自动解析——哪些 node_type/style 允许解析内联 ref
# 注意：code/pre/formula 不解析，避免误识别代码花括号
_INLINE_REF_ELIGIBLE_TYPES = {"paragraph", "list_item", "ref", "callout", "abstract_block", "quote"}
_INLINE_REF_SKIP_STYLES = {"code", "pre"}  # paragraph 节点若 style 为这些则跳过


def _parse_inline_refs(text: str) -> list[dict]:
    """将文本中的 ``{ref label}`` 拆分为 text/ref 段列表。

    返回交替的节点列表::

        [{"type":"text","text":"..."}, {"type":"ref","target":"label","text":null}, ...]

    若文本中没有匹配，返回单元素列表 ``[{"type":"text","text":text}]`` 以保持向后兼容。
    """
    if not text:
        return [{"type": "text", "text": ""}]
    segments: list[dict] = []
    pos = 0
    for m in _RE_INLINE_REF.finditer(text):
        pre = text[pos:m.start()]
        if pre:
            segments.append({"type": "text", "text": pre})
        segments.append({"type": "ref", "target": m.group(1), "text": None})
        pos = m.end()
    tail = text[pos:]
    if tail or not segments:
        segments.append({"type": "text", "text": tail})
    return segments


def _apply_inline_refs_to_node(n: dict) -> None:
    """对单个节点的文本字段做内联 ref 解析，必要时写入 ``inline_segments``。

    - 对 paragraph：style=code/pre 不解析；仅当解析出 ref 段时写入 ``inline_segments``。
    - 对 list_item：text 字段解析。
    - 对 callout：body 字段解析。
    - 对 abstract_block：text 字段解析。
    - 不修改不含 ``{ref ...}`` 的节点，保证 100% 向后兼容。
    """
    ntype = n.get("node_type")
    if ntype not in _INLINE_REF_ELIGIBLE_TYPES:
        return
    if ntype == "paragraph" and n.get("style") in _INLINE_REF_SKIP_STYLES:
        return
    # 确定要解析的文本字段
    text_field = "text"
    if ntype == "callout":
        text_field = "body"
    raw = n.get(text_field, "")
    if not isinstance(raw, str) or not raw:
        return
    # 无 {ref 就直接跳过，保持原状
    if "{ref" not in raw:
        return
    segs = _parse_inline_refs(raw)
    has_ref = any(s.get("type") == "ref" for s in segs)
    if has_ref:
        n["inline_segments"] = segs


# ─── 节点展开 ──────────────────────────────────────────────────────

def parse_document(doc: dict) -> dict:
    """将高层文档 JSON 编译为渲染计划。

    Returns:
        {
            "meta": {...},
            "tokens_theme": str|dict,
            "page_setup": {...}|None,
            "abstract": {...}|None,
            "toc": bool|dict,
            "header": {...}|None,
            "footer": {...}|None,
            "watermark": {...}|None,
            "page_border": {...}|None,
            "nodes": [flat_node, ...],
            "references": [...],
            "appendices": [...],
            "lang": str|None,
        }
    """
    validate_document(doc)

    meta = doc["meta"]
    # v1.4 P1-4b: 暴露 meta 给 _expand_node（用于 references 节点回退 citation_style）
    global _current_meta
    _current_meta = meta
    theme = doc.get("theme", "academic")
    abstract = doc.get("abstract")
    toc = doc.get("toc", False)
    header = doc.get("header")
    footer = doc.get("footer")
    watermark = doc.get("watermark")
    page_border = doc.get("page_border")
    page_setup = meta.get("page_setup")
    sections = doc["sections"]
    references = doc.get("references", [])
    appendices = doc.get("appendices", [])

    nodes = []

    # v1.4 P3-4 / P3-6: 规范化文档级 columns / line_numbers
    meta_columns = _normalize_columns(meta.get("columns"))
    meta_line_numbers = _normalize_line_numbers(meta.get("line_numbers"))

    # 标题页信息（由 renderer 在开头渲染标题块）
    nodes.append(_node("title_block", {
        "title": meta["title"],
        "subtitle": meta.get("subtitle"),
        "author": meta.get("author"),
        "institution": meta.get("institution"),
        "date": meta.get("date"),
    }))

    # 摘要
    if abstract:
        nodes.append(_node("abstract_block", {
            "text": abstract.get("text", ""),
            "keywords": abstract.get("keywords", []),
        }))

    # TOC
    if toc:
        toc_data = toc if isinstance(toc, dict) else {}
        # 规范化 levels：int -> [1..n]；缺省 -> [1,2,3]
        lv = toc_data.get("levels", 3)
        if isinstance(lv, int):
            lv = list(range(1, lv + 1))
        # title 缺省：renderer 根据 lang 选择
        nodes.append(_node("toc", {
            "title": toc_data.get("title"),
            "levels": lv,
        }))

    # 正文章节
    for sec in sections:
        sec_level = sec.get("level", 1)
        nodes.append(_node("heading", {
            "level": sec_level,
            "text": sec["title"],
        }))
        for raw_node in sec["content"]:
            nodes.extend(_expand_node(raw_node, sec_level))

    # 参考文献（doc 级，旧格式；新格式见 sections 内的 references 节点）
    if references:
        nodes.append(_node("references_block", {
            "title": "References" if (meta.get("lang", "en") == "en") else "参考文献",
            "citation_style": doc.get("citation_style", "apa"),
            "items": references,
        }))

    # 附录
    for app in appendices:
        nodes.append(_node("page_break", {}))
        nodes.append(_node("heading", {"level": 1, "text": app["title"]}))
        for raw_node in app["content"]:
            nodes.extend(_expand_node(raw_node, 1))

    # v1.4 P3-9: 智能结构识别（auto_structure）——在 normalize 末尾对 nodes 做二次分类
    as_mode = _coerce_auto_structure(meta.get("auto_structure"))
    auto_structure_meta_updates: dict = {}
    has_explicit_title = bool(meta.get("title"))
    auto_lang: str | None = None
    auto_resolved_mode: str | None = None
    if as_mode:
        nodes, auto_structure_meta_updates = _auto_structure_nodes(
            nodes, mode=as_mode, lang_hint=meta.get("lang"),
            skip_title_detect=has_explicit_title,
        )
        auto_lang = auto_structure_meta_updates.get("_auto_lang")
        auto_resolved_mode = auto_structure_meta_updates.get("_auto_mode")
        # 回写识别出的 title/subtitle/author 到 meta（仅当原 meta 未显式指定）
        if not has_explicit_title and auto_structure_meta_updates.get("title"):
            meta["title"] = auto_structure_meta_updates["title"]
        if not meta.get("subtitle") and auto_structure_meta_updates.get("subtitle"):
            meta["subtitle"] = auto_structure_meta_updates["subtitle"]
        if not meta.get("author") and auto_structure_meta_updates.get("author"):
            meta["author"] = auto_structure_meta_updates["author"]

    plan = {
        "meta": meta,
        "tokens_theme": theme,
        "page_setup": page_setup,
        "abstract": abstract,
        "toc": toc,
        "header": header,
        "footer": footer,
        "watermark": watermark,
        "page_border": page_border,
        # v1.4 P3-4 / P3-6: 文档级排版配置（None 表示不设置，保持 v1.3 默认单栏/无行号）
        "columns": meta_columns,
        "line_numbers": meta_line_numbers,
        # v1.4 P3-9: 智能结构识别模式（False/'academic'/'business'/'auto'）
        "auto_structure": as_mode,
        # v1.4 P3-9c: 自动识别出的语言/模式（用于渲染阶段字体映射）
        "auto_lang": auto_lang,
        "auto_mode": auto_resolved_mode,
        "nodes": nodes,
        "references": references,
        "appendices": appendices,
    }

    # v1.4 P2-4b: 扫描所有节点，对正文/列表/abstract/callout 中的内联 {ref xxx}
    # 做自动拆分（自动识别出的 paragraph/list_item 也被扫描；code 块跳过）
    _finalize_inline_refs(plan)
    return plan


def _finalize_inline_refs(plan: dict) -> None:
    """对 plan 中所有符合条件的节点做 inline ref 解析。

    - paragraph / list_item 文本段扫描
    - paragraph style=code/pre 跳过（代码块不解析 {ref}）
    - callout 的 body 字段、abstract_block 的 text 字段同样扫描
    - 解析结果写入 node['inline_segments']
    - 对不含 ``{ref`` 的文本节点零修改，100% 向后兼容
    """
    nodes = plan.get("nodes", [])
    abstract = plan.get("abstract")
    if isinstance(abstract, dict) and isinstance(abstract.get("text"), str) and "{ref" in abstract["text"]:
        _apply_inline_refs_to_abstract(abstract)
    for n in nodes:
        try:
            _apply_inline_refs_to_node(n)
        except Exception:
            pass


def _apply_inline_refs_to_abstract(abs_node: dict) -> None:
    """对 abstract 节点（doc 级 metadata）做 inline ref 解析。"""
    raw = abs_node.get("text", "")
    if not isinstance(raw, str) or "{ref" not in raw:
        return
    segs = _parse_inline_refs(raw)
    if any(s.get("type") == "ref" for s in segs):
        abs_node["inline_segments"] = segs


def _node(ntype: str, data: dict) -> dict:
    """创建扁平节点。"""
    return {"node_type": ntype, **data}


def _expand_node(raw: dict, parent_heading_level: int) -> list[dict]:
    """将单个语义节点展开为一个或多个扁平节点。"""
    ntype = raw["type"]

    if ntype == "heading":
        level = raw.get("level", parent_heading_level + 1)
        n = _node("heading", {"level": level, "text": raw["text"]})
        if "comment" in raw:
            n["comment"] = raw["comment"]
        return [n]

    if ntype == "paragraph":
        n = _node("paragraph", {
            "text": raw["text"],
            "style": raw.get("style", "normal"),
            "bold": raw.get("bold", False),
            "italic": raw.get("italic", False),
            # v1.4 P3-5: 首字下沉
            "drop_cap": bool(raw.get("drop_cap", False)),
            # v1.4 P3-4: 段落级分栏（段末分节符，可选）
            "columns": _normalize_columns(raw.get("columns")),
        })
        if "comment" in raw:
            n["comment"] = raw["comment"]
        if "runs" in raw:
            # 高级用法：runs 列表，每个 run 可单独指定 bold/italic/revision 等
            n["runs"] = raw["runs"]
        return [n]

    if ntype == "list":
        return _expand_list(raw)

    if ntype == "table":
        caption = raw.get("caption") or raw.get("title")
        tnode = _node("table", {
            "caption": caption,
            "headers": raw["headers"],
            "rows": raw["rows"],
            "col_widths": raw.get("col_widths"),
            "header_repeat": raw.get("header_repeat", True),
            "label": raw.get("label"),  # v1.4 P2-4
        })
        return [tnode]

    if ntype == "figure":
        caption = raw.get("caption") or raw.get("title")
        return [_node("figure", {
            "path": raw["path"],
            "caption": caption,
            "width_inches": raw.get("width_inches"),
            "align": raw.get("align", "center"),
            "label": raw.get("label"),  # v1.4 P2-4
        })]

    if ntype == "chart":
        return [_node("chart", {
            "chart_type": raw.get("chart_type", "column"),
            "title": raw.get("title"),
            "caption": raw.get("caption"),
            "categories": raw.get("categories", []),
            "series": raw.get("series", []),
            "show_legend": raw.get("show_legend", True),
            "legend_position": raw.get("legend_position", "bottom"),
            "show_data_labels": raw.get("show_data_labels"),
            "number_format": raw.get("number_format", "0"),
            "x_title": raw.get("x_title"),
            "y_title": raw.get("y_title"),
            "aspect_ratio": raw.get("aspect_ratio", "4:3"),
            "width_pct": raw.get("width_pct", 1.0),
            "align": raw.get("align", "center"),
            "mode": raw.get("mode", "image"),
        })]

    if ntype == "kpi_card":
        return [_node("kpi_card", {
            "value": raw["value"],
            "label": raw["label"],
            "subtext": raw.get("subtext"),
        })]

    if ntype == "callout":
        return [_node("callout", {
            "title": raw.get("title"),
            "body": raw["body"],
            "variant": raw.get("variant", "info"),
        })]

    if ntype == "page_break":
        return [_node("page_break", {})]

    if ntype == "toc":
        lv = raw.get("levels", 3)
        if isinstance(lv, int):
            lv = list(range(1, lv + 1))
        return [_node("toc", {"title": raw.get("title"), "levels": lv})]

    if ntype == "references":
        # v1.4 P1-4/P1-4b: 传入 citation_style/group_by_type/title
        # 优先读节点自身的 citation_style，其次读 meta 级，最后默认 apa
        parent_meta = globals().get("_current_meta", {}) or {}
        style = raw.get("citation_style") or parent_meta.get("citation_style", "apa")
        title = raw.get("title")  # None 时由 renderer 按 lang 选择
        return [_node("references_block", {
            "title": title,
            "citation_style": style,
            "group_by_type": bool(raw.get("group_by_type", False)),
            "items": raw.get("items", []),
        })]

    if ntype == "appendix":
        nodes = [
            _node("page_break", {}),
            _node("heading", {"level": 1, "text": raw["title"]}),
        ]
        for inner in raw["content"]:
            nodes.extend(_expand_node(inner, 1))
        return nodes

    # ─── v1.2 新节点 ──────────────────────────────────────────────

    if ntype == "revision":
        return [_node("revision", {
            "action": raw.get("action", "insert"),
            "text": raw.get("text", ""),
            "old_text": raw.get("old_text", ""),
            "new_text": raw.get("new_text", ""),
            "author": raw.get("author"),
            "date": raw.get("date"),
        })]

    if ntype == "comment":
        return [_node("comment", {
            "text": raw["text"],
            "author": raw.get("author"),
            "date": raw.get("date"),
        })]

    if ntype == "footnote":
        return [_node("footnote", {
            "text": raw["text"],
            "id": raw.get("id"),
        })]

    if ntype == "endnote":
        return [_node("endnote", {
            "text": raw["text"],
            "id": raw.get("id"),
        })]

    if ntype == "watermark":
        return [_node("watermark", {
            "enabled": raw.get("enabled", True),
            "text": raw.get("text", "DRAFT"),
            "image_path": raw.get("image_path"),
            "rotation": raw.get("rotation", -45),
        })]

    if ntype == "page_border":
        return [_node("page_border", {
            "enabled": raw.get("enabled", True),
            "style": raw.get("style", "single"),
            "offset_from": raw.get("offset_from", "page"),
        })]

    if ntype == "equation":
        # Backward-compat: v1.2 used inline=True/False; v1.3 uses display="inline"/"block"
        if "inline" in raw and "display" not in raw:
            display = "inline" if raw["inline"] else "block"
        else:
            display = raw.get("display", "block")
        caption = raw.get("caption") or raw.get("title")
        return [_node("equation", {
            "latex": raw["latex"],
            "caption": caption,
            "display": display,
            "label": raw.get("label"),  # v1.4 P2-4
        })]

    if ntype == "svg_shape":
        return [_node("svg_shape", {
            "svg": raw["svg"],
            "width": raw.get("width", "5cm"),
            "height": raw.get("height"),
            "align": raw.get("align", "center"),
        })]

    if ntype == "signature_block":
        return [_node("signature_block", {
            "name": raw.get("name", ""),
            "date": raw.get("date", ""),
            "title": raw.get("title", "签字"),
            "lines": raw.get("lines", 1),
        })]

    if ntype == "signature_line":
        return [_node("signature_line", {
            "signer": raw.get("signer", "签字人"),
            "date": raw.get("date", "日期"),
        })]

    # v1.4 P2-4: cross-reference
    if ntype == "ref":
        return [_node("ref", {
            "target": raw["target"],
            "prefix": raw.get("prefix", ""),
            "suffix": raw.get("suffix", ""),
            "style": raw.get("style", "inline"),
        })]

    # 未知类型（validators 应已拦截，这里兜底）
    # English original: f"Unknown node type: {ntype}"
    raise ValueError(
        f"未知节点类型 {ntype!r}。修复建议：请检查节点 type 字段是否拼写正确，"
        f"或使用 generate_from_markdown 让系统自动解析 Markdown 内容"
    )


def _expand_list(raw: dict) -> list[dict]:
    """展开列表（支持嵌套 items）为带 level 标记的 list_item 节点序列。

    v1.4 P2-3:
    - 每个列表节点携带 ``list_id``（自增）和 ``start``（起始编号）。
    - 嵌套：当某个 item 是 dict 且包含 ``items``（新）或 ``sub_items``（旧），
      其子项递归展开，level = parent_level + 1。
    - 单层字符串列表的行为与 v1.3 保持一致（ordered 缺省为 True 是与 PRD
      对齐的新默认；为保持向后兼容，旧的 sub_items 形式继续沿用父节点的 ordered）。
    """
    ordered = raw.get("ordered", True)
    start = int(raw.get("start", 1) or 1)
    result: list[dict] = []
    # Allocate a unique list_id per call (counter on function object)
    _expand_list._counter = getattr(_expand_list, "_counter", 0) + 1
    list_id = f"list{_expand_list._counter}"
    for item in raw["items"]:
        _flatten_list_item(item, ordered, 0, result, list_id=list_id)
    # Attach list metadata on the first list_item for this list so the renderer
    # can create a numbering definition. Subsequent items reference same list_id.
    if result:
        result[0]["_list_start"] = start
        result[0]["_list_ordered"] = ordered
    return result


def _flatten_list_item(item, ordered: bool, depth: int, result: list, *, list_id: str):
    if isinstance(item, str):
        node = _node("list_item", {
            "text": item, "ordered": ordered, "level": depth, "list_id": list_id,
        })
        result.append(node)
        return
    if isinstance(item, dict):
        # Determine nested items key
        nested_key = None
        if "items" in item:
            nested_key = "items"
        elif "sub_items" in item:
            nested_key = "sub_items"
        # This item's own ordering defaults to parent's ordered
        self_ordered = item.get("ordered", ordered)
        lvl = item.get("level", depth)
        text = item.get("text", "")
        node = _node("list_item", {
            "text": text, "ordered": self_ordered, "level": lvl, "list_id": list_id,
        })
        result.append(node)
        if nested_key:
            for sub in item.get(nested_key, []):
                _flatten_list_item(sub, self_ordered, lvl + 1, result, list_id=list_id)


# ─── 大纲提取 ─────────────────────────────────────────────────────

def extract_outline(doc: dict) -> dict:
    """从文档 JSON 提取章节大纲树（阶段 1 用）。"""
    def walk_sections(sections, parent_level=0):
        tree = []
        for sec in sections:
            title = sec["title"]
            level = sec.get("level", parent_level + 1 if parent_level else 1)
            children = []
            for node in sec.get("content", []):
                if isinstance(node, dict) and node.get("type") == "heading":
                    children.append({
                        "title": node["text"],
                        "level": node.get("level", level + 1),
                        "children": [],
                    })
            tree.append({"title": title, "level": level, "children": children})
        return tree

    return {
        "title": doc.get("meta", {}).get("title", ""),
        "subtitle": doc.get("meta", {}).get("subtitle"),
        "sections": walk_sections(doc.get("sections", [])),
    }


# ─── 字数估算 ──────────────────────────────────────────────────────

def count_words(doc: dict, lang: str = "en") -> int:
    """估算正文字数（不含标题页/目录/参考文献/附录）。"""
    total = 0

    if doc.get("abstract"):
        total += _text_len(doc["abstract"].get("text", ""), lang)
        for kw in doc["abstract"].get("keywords", []):
            total += _text_len(kw, lang)

    for sec in doc.get("sections", []):
        total += _text_len(sec.get("title", ""), lang)
        for node in sec.get("content", []):
            total += _count_node_words(node, lang)

    return total


def _count_node_words(node: Any, lang: str) -> int:
    if not isinstance(node, dict):
        return 0
    ntype = node.get("type", "")
    if ntype == "heading":
        return _text_len(node.get("text", ""), lang)
    if ntype == "paragraph":
        return _text_len(node.get("text", ""), lang)
    if ntype == "revision":
        return _text_len(node.get("text", "") or node.get("new_text", ""), lang)
    if ntype == "list":
        total = 0
        for item in node.get("items", []):
            if isinstance(item, str):
                total += _text_len(item, lang)
            elif isinstance(item, dict):
                total += _text_len(item.get("text", ""), lang)
                for si in item.get("sub_items", []):
                    total += _count_node_words_in_list_item(si, lang)
        return total
    if ntype == "table":
        total = 0
        for h in node.get("headers", []):
            total += _text_len(h, lang)
        for row in node.get("rows", []):
            for cell in row:
                total += _text_len(str(cell), lang)
        if node.get("caption"):
            total += _text_len(node["caption"], lang)
        return total
    if ntype == "callout":
        return _text_len(node.get("body", ""), lang) + _text_len(node.get("title", ""), lang)
    if ntype in ("figure", "equation") and node.get("caption"):
        return _text_len(node["caption"], lang)
    if ntype == "chart":
        cap = _text_len(node.get("caption", ""), lang)
        title = _text_len(node.get("title", ""), lang)
        return cap + title + 100
    return 0


def _count_node_words_in_list_item(item: Any, lang: str) -> int:
    if isinstance(item, str):
        return _text_len(item, lang)
    if isinstance(item, dict):
        t = _text_len(item.get("text", ""), lang)
        for si in item.get("sub_items", []):
            t += _count_node_words_in_list_item(si, lang)
        return t
    return 0


def _text_len(text: str, lang: str) -> int:
    """中文按字符计数，英文按空格分词计数。"""
    if not text:
        return 0
    if lang == "cn":
        return len([c for c in text if not c.isspace()])
    return len(text.split())


# ─── v1.4 P3-4 / P3-6: 字段规范化 ────────────────────────────────

def _normalize_columns(cols):
    """int → {count:int, space:720000}；dict 补默认；None → None。"""
    if cols is None:
        return None
    if isinstance(cols, int):
        return {"count": cols, "space": 720000, "sep": False}
    if isinstance(cols, dict):
        return {
            "count": int(cols.get("count", 1)),
            "space": int(cols.get("space", 720000)),
            "sep": bool(cols.get("sep", False)),
        }
    return cols  # validators 已拦截非法类型


def _normalize_line_numbers(ln):
    """补全 line_numbers 默认值；None → None。"""
    if ln is None:
        return None
    return {
        "start": int(ln.get("start", 1)),
        "increment": int(ln.get("increment", 1)),
        "restart": ln.get("restart", "continuous"),
        "distance": int(ln.get("distance", 360000)),
    }



# ─── v1.4 P3-9: 智能结构识别（auto_structure）────────────────────

# v1.4 P3-9c: (lang, mode) × heading level 样式覆盖表（中文学术/中文商业/英文学术/英文商业）
# 参考 GB/T 7713.1-2006 学位论文、GB/T 9704-2012 公文、APA 7th、McKinsey/Harvard 咨询报告。
# 仅对 _auto_structure_nodes 自动识别出来的 heading 生效（显式 heading 节点不注入）。
# alignment: "left"|"center"|"justify"
_AUTO_HEADING_STYLES: dict[tuple[str, str], dict[int, dict]] = {
    ("cn", "academic"): {
        1: {"font_latin": "Times New Roman", "font_eastasia": "SimHei",
            "font_size_pt": 16, "bold": True, "italic": False,
            "alignment": "center", "space_before_pt": 24, "space_after_pt": 18,
            "all_caps": False},
        2: {"font_latin": "Times New Roman", "font_eastasia": "SimHei",
            "font_size_pt": 14, "bold": True, "italic": False,
            "alignment": "left", "space_before_pt": 12, "space_after_pt": 6,
            "all_caps": False},
        3: {"font_latin": "Times New Roman", "font_eastasia": "SimHei",
            "font_size_pt": 12, "bold": True, "italic": False,
            "alignment": "left", "space_before_pt": 6, "space_after_pt": 3,
            "all_caps": False},
        4: {"font_latin": "Times New Roman", "font_eastasia": "SimSun",
            "font_size_pt": 12, "bold": True, "italic": False,
            "alignment": "left", "space_before_pt": 3, "space_after_pt": 3,
            "all_caps": False},
    },
    ("cn", "business"): {
        1: {"font_latin": "Calibri", "font_eastasia": "Microsoft YaHei",
            "font_size_pt": 22, "bold": True, "italic": False,
            "alignment": "center", "space_before_pt": 30, "space_after_pt": 20,
            "all_caps": False},
        2: {"font_latin": "Calibri", "font_eastasia": "Microsoft YaHei",
            "font_size_pt": 16, "bold": True, "italic": False,
            "alignment": "left", "space_before_pt": 18, "space_after_pt": 12,
            "all_caps": False},
        3: {"font_latin": "Calibri", "font_eastasia": "Microsoft YaHei",
            "font_size_pt": 14, "bold": True, "italic": False,
            "alignment": "left", "space_before_pt": 12, "space_after_pt": 6,
            "all_caps": False},
        4: {"font_latin": "Calibri", "font_eastasia": "Microsoft YaHei",
            "font_size_pt": 12, "bold": True, "italic": False,
            "alignment": "left", "space_before_pt": 6, "space_after_pt": 3,
            "all_caps": False},
    },
    ("en", "academic"): {
        1: {"font_latin": "Times New Roman", "font_eastasia": "SimSun",
            "font_size_pt": 14, "bold": True, "italic": False,
            "alignment": "center", "space_before_pt": 24, "space_after_pt": 18,
            "all_caps": False},
        2: {"font_latin": "Times New Roman", "font_eastasia": "SimSun",
            "font_size_pt": 12, "bold": True, "italic": True,
            "alignment": "left", "space_before_pt": 12, "space_after_pt": 6,
            "all_caps": False},
        3: {"font_latin": "Times New Roman", "font_eastasia": "SimSun",
            "font_size_pt": 12, "bold": True, "italic": False,
            "alignment": "left", "space_before_pt": 6, "space_after_pt": 3,
            "all_caps": False},
        4: {"font_latin": "Times New Roman", "font_eastasia": "SimSun",
            "font_size_pt": 11, "bold": True, "italic": True,
            "alignment": "left", "space_before_pt": 3, "space_after_pt": 3,
            "all_caps": False},
    },
    ("en", "business"): {
        1: {"font_latin": "Calibri", "font_eastasia": "Microsoft YaHei",
            "font_size_pt": 16, "bold": True, "italic": False,
            "alignment": "left", "space_before_pt": 18, "space_after_pt": 12,
            "all_caps": True},
        2: {"font_latin": "Calibri", "font_eastasia": "Microsoft YaHei",
            "font_size_pt": 13, "bold": True, "italic": False,
            "alignment": "left", "space_before_pt": 12, "space_after_pt": 6,
            "all_caps": False},
        3: {"font_latin": "Calibri", "font_eastasia": "Microsoft YaHei",
            "font_size_pt": 11, "bold": True, "italic": True,
            "alignment": "left", "space_before_pt": 6, "space_after_pt": 3,
            "all_caps": False},
        4: {"font_latin": "Calibri", "font_eastasia": "Microsoft YaHei",
            "font_size_pt": 11, "bold": True, "italic": False,
            "alignment": "left", "space_before_pt": 3, "space_after_pt": 3,
            "all_caps": False},
    },
}

# auto_structure 模式下正文默认字体（覆盖 tokens 默认，保证中/英×学术/商业各自正确）
_AUTO_BODY_FONT = {
    ("cn", "academic"): {"font_latin": "Times New Roman", "font_eastasia": "SimSun",
                         "font_size_pt": 10.5},
    ("cn", "business"): {"font_latin": "Calibri", "font_eastasia": "Microsoft YaHei",
                         "font_size_pt": 11},
    ("en", "academic"): {"font_latin": "Times New Roman", "font_eastasia": "SimSun",
                         "font_size_pt": 12},
    ("en", "business"): {"font_latin": "Calibri", "font_eastasia": "Microsoft YaHei",
                         "font_size_pt": 11},
}


def _auto_heading_style(lang: str, mode: str, level: int) -> dict | None:
    """返回 (lang, mode, level) 对应的 style_override dict；无匹配返回 None。"""
    table = _AUTO_HEADING_STYLES.get((lang, mode))
    if not table:
        return None
    return table.get(level)


def _auto_body_font(lang: str, mode: str) -> dict | None:
    """返回 auto_structure 正文默认字体配置；无匹配返回 None。"""
    return _AUTO_BODY_FONT.get((lang, mode))


def _coerce_auto_structure(v):
    """将 meta.auto_structure 规范化为 False / 'academic' / 'business' / 'auto'。"""
    if v is None:
        return False
    if v is False:
        return False
    if v is True:
        return "auto"
    if isinstance(v, str):
        if v.lower() in ("false", "off", "none", "0", "no"):
            return False
        if v in ("academic", "business", "auto"):
            return v
    return False


# 中文数字
_CN_NUM = "一二三四五六七八九十百千零〇两"
_CN_NUM_PAT = f"[{_CN_NUM}]+"

# 识别为 front_matter/heading1 的特殊关键词
_CN_FRONT_KEYWORDS = {"摘要", "摘  要", "Abstract", "ABSTRACT", "引言", "绪论",
                      "结论", "参考文献", "致谢", "附录", "目录", "关键词"}
_EN_FRONT_KEYWORDS = {"Abstract", "ABSTRACT", "Introduction", "Conclusion",
                      "References", "Bibliography", "Acknowledgments",
                      "Acknowledgements", "Appendix", "Table of Contents",
                      "Keywords", "KEY WORDS"}

# 中文学术/商业编号正则
# 支持的一级编号：第一章 / 第1节 / 第一部分 / 第2部分（"部分"为双字后缀，单独处理）
_RE_CN_H1_CHAPTER = re.compile(rf"^\s*第[{_CN_NUM}\d]+(?:部分|[章节部篇])\s*[ 　]*(.+)$")
_RE_CN_H1_PURE = re.compile(rf"^\s*[{_CN_NUM}]+[、．.]\s*(.+)$")           # 一、 二、
_RE_CN_H2 = re.compile(rf"^\s*（[{_CN_NUM}]+）\s*(.+)$")                    # （一）（二）
_RE_CN_H2_FULL = re.compile(rf"^\s*\([{_CN_NUM}]+\)\s*(.+)$")              # (一)(二)
_RE_CN_H4 = re.compile(r"^\s*(\d{1,2})\.(\d{1,2})(?:[\.、]|$)\s*(.*)$")    # 1.1
_RE_CN_H3 = re.compile(r"^\s*(\d{1,2})[\.、]\s*(.+)$")                      # 1. 2.
_RE_CN_H4_PAREN = re.compile(r"^\s*（\d{1,2}）\s*(.+)$")                   # （1）（2）

# 英文学术/商业编号正则
_RE_EN_H1_ROMAN = re.compile(r"^\s*([IVXLCDM]+)\.\s+(.+)$")          # I. II. III.
_RE_EN_H2_LETTER = re.compile(r"^\s*([A-Z])\.\s+(.+)$")              # A. B. (IEEE h2)
_RE_EN_H3_PAREN = re.compile(r"^\s*(\d{1,2})\)\s+(.+)$")             # 1) 2) (IEEE h3)
_RE_EN_H1_NUM = re.compile(r"^\s*(\d{1,2})\.\s+([A-Z][A-Za-z].+)$")  # 1. Intro
_RE_EN_H2 = re.compile(r"^\s*(\d{1,2})\.(\d{1,2})\s+(.+)$")          # 1.1 Title
_RE_EN_H3 = re.compile(r"^\s*(\d{1,2})\.(\d{1,2})\.(\d{1,2})\s+(.+)$")  # 1.1.1
_RE_EN_APPX = re.compile(r"^\s*Appendix\s+([A-Z])[\s\.:]\s*(.*)$", re.IGNORECASE)

# 列表前缀
_RE_BULLET = re.compile(r"^\s*[·•\-*\u25A0-\u25CF\u25CB○□■◆▪▫]\s+(.+)$")
_RE_ORDERED = re.compile(r"^\s*(\d+)[\.\)、]\s+(.+)$")

# 图表 caption（不提升为 heading）
_RE_FIG_CAP = re.compile(r"^\s*(图|表|Figure|Table|Fig\.)\s*\d", re.IGNORECASE)

# v1.4 P3-9b: 精确 caption 提取正则（kind, label_id, caption_text）
# 中文图：图 1 / 图1 / 图1-1 后接 ：:、. 空格 等
_RE_CAP_FIG_CN = re.compile(r"^\s*图\s*([A-Za-z0-9][A-Za-z0-9_-]*)[\s:：.、]\s*(.*)$")
# 英文 Fig. / Figure
_RE_CAP_FIG_EN = re.compile(r"^\s*(?:Fig\.?|Figure)\s*([A-Za-z0-9][A-Za-z0-9_-]*)[\s:.、]\s*(.*)$", re.IGNORECASE)
# 中文表
_RE_CAP_TBL_CN = re.compile(r"^\s*表\s*([A-Za-z0-9][A-Za-z0-9_-]*)[\s:：.、]\s*(.*)$")
# 英文 Table / Tab.
_RE_CAP_TBL_EN = re.compile(r"^\s*(?:Tab(?:le)?\.?)\s*([A-Za-z0-9][A-Za-z0-9_-]*)[\s:.、]\s*(.*)$", re.IGNORECASE)
# 公式：公式1 / Eq.1 / Equation 1
_RE_CAP_EQ = re.compile(r"^\s*(?:公式|Eq\.?|Equation)\s*([A-Za-z0-9][A-Za-z0-9_-]*)[\s:.、]?\s*(.*)$", re.IGNORECASE)


def _extract_caption_info(text: str) -> tuple[str | None, str | None, str]:
    """从 caption 文本中提取 (kind, label_id, caption_text)。

    kind ∈ {"figure","table","equation"} 或 None（非 caption 行）。
    label_id 是用户可引用的 label 名（如 ``"fig1"`` / ``"tab2"`` / ``"eq3"``）。
    caption_text 是去掉编号前缀后的正文文本。
    """
    s = (text or "").strip()
    if not s:
        return None, None, ""
    m = _RE_CAP_FIG_EN.match(s)
    if m and s.lower().startswith(("fig", "figure")):
        return "figure", f"fig{m.group(1)}", m.group(2).strip()
    m = _RE_CAP_FIG_CN.match(s)
    if m:
        return "figure", f"fig{m.group(1)}", m.group(2).strip()
    m = _RE_CAP_TBL_EN.match(s)
    if m and s.lower().startswith(("tab", "table")):
        return "table", f"tab{m.group(1)}", m.group(2).strip()
    m = _RE_CAP_TBL_CN.match(s)
    if m:
        return "table", f"tab{m.group(1)}", m.group(2).strip()
    m = _RE_CAP_EQ.match(s)
    if m:
        return "equation", f"eq{m.group(1)}", m.group(2).strip()
    return None, None, s


# 元数据标签行
_RE_META_TAG = re.compile(r"^\s*(关键词|Key\s*words?|Keywords?|作者|Author|单位|编制|审核|批准|版本|Date|Prepared for|Confidential)\s*[:：]\s*(.*)$")


def _detect_lang(texts: list[str], hint: str | None = None) -> str:
    """基于字符统计检测语言：返回 'cn' 或 'en'。"""
    if hint in ("cn", "zh", "zh-CN"):
        return "cn"
    if hint == "en":
        return "en"
    blob = "".join(t for t in texts if t)
    if not blob:
        return "cn"
    cjk = sum(1 for c in blob if "\u4e00" <= c <= "\u9fff" or "\u3000" <= c <= "\u303f")
    ascii_letters = sum(1 for c in blob if c.isascii() and c.isalpha())
    total = max(1, len(blob))
    if cjk / total > 0.3:
        return "cn"
    if ascii_letters / total > 0.5:
        return "en"
    return "cn"


def _guess_mode(mode: str, lang: str) -> str:
    """mode='auto' 时：默认 academic；若检测到商业词汇则 business。"""
    if mode != "auto":
        return mode
    return "academic"


def _is_short(text: str, limit: int = 30) -> bool:
    return len(text.strip()) <= limit and len(text.strip()) > 0


def _looks_like_caption(text: str) -> bool:
    return bool(_RE_FIG_CAP.match(text.strip()))


def _looks_like_list_item(text: str) -> tuple[bool, bool, str]:
    """返回 (is_list, ordered, text_without_prefix)。"""
    s = text.strip()
    m = _RE_BULLET.match(s)
    if m:
        return True, False, m.group(1).strip()
    m = _RE_ORDERED.match(s)
    if m and _is_short(s, 80):
        # 避免把"1. 引言"（可能是 heading）误判为 list；仅当后续看起来是完整句子或
        # 非常短的项目项（无层级编号）才判为 list
        return True, True, m.group(2).strip()
    return False, False, s


def _detect_heading_level_cn(text: str, mode: str) -> int | None:
    """对中文段落检测 heading 级别，返回 1-4 或 None。

    检测顺序从最具体（最深层级）到最宽泛，避免短正则错误吞掉更具体的格式。
    """
    s = text.strip()
    if not s:
        return None
    if _looks_like_caption(s):
        return None
    # front matter 关键词（独占一行）
    for kw in _CN_FRONT_KEYWORDS:
        if s == kw or (s.startswith(kw) and len(s) <= len(kw) + 3):
            return 1
    # 第X章 / 第X节（一级标题）
    m = _RE_CN_H1_CHAPTER.match(s)
    if m and _is_short(s, 40):
        return 1
    # 一、 二、（一级）
    m = _RE_CN_H1_PURE.match(s)
    if m:
        rest = m.group(1).strip()
        if _is_short(s, 40):
            return 1
    # （一）（二级）
    m = _RE_CN_H2.match(s) or _RE_CN_H2_FULL.match(s)
    if m and _is_short(s, 40):
        return 2
    # 1.1（必须在 1. 之前匹配）
    m = _RE_CN_H4.match(s)
    if m and _is_short(s, 40):
        # academic: 1.1 → 4 级；business: 1.1 → 2 级
        return 2 if mode == "business" else 4
    # （1）（四级）
    m = _RE_CN_H4_PAREN.match(s)
    if m and _is_short(s, 40):
        return 4
    # 1. 2.（academic 三级 / business 二级）
    m = _RE_CN_H3.match(s)
    if m and _is_short(s, 40):
        return 2 if mode == "business" else 3
    return None


def _detect_heading_level_en(text: str, mode: str) -> int | None:
    """对英文段落检测 heading 级别，返回 1-4 或 None。

    学术模式支持 APA/MLA（数字点分）+ IEEE（罗马 I./字母 A./括号 1)）；
    商业模式支持数字点分与全大写短行。
    """
    s = text.strip()
    if not s:
        return None
    if _looks_like_caption(s):
        return None
    # front keywords
    for kw in _EN_FRONT_KEYWORDS:
        if s.lower() == kw.lower() or (s.lower().startswith(kw.lower() + " ") and len(s) <= len(kw) + 20):
            return 1
    # Appendix A / Appendix B
    if _RE_EN_APPX.match(s):
        return 1
    # 1.1.1 (h3)
    m = _RE_EN_H3.match(s)
    if m:
        return 3
    # 1.1 (h2)
    m = _RE_EN_H2.match(s)
    if m:
        return 2
    # I. II. III. (h1 academic IEEE)
    m = _RE_EN_H1_ROMAN.match(s)
    if m and _is_short(s, 80):
        return 1 if mode == "academic" else 2
    # A. B. (IEEE h2)
    if mode == "academic":
        m = _RE_EN_H2_LETTER.match(s)
        if m and _is_short(s, 80):
            letter = m.group(1)
            if len(letter) == 1 and letter.isalpha() and letter.isupper():
                # 避免匹配句首大写字母 + "." 句点的正常段落（要求后续紧跟 Title Case 词）
                rest = m.group(2)
                if rest and rest[0].isupper() and not rest.endswith("."):
                    return 2
    # 1) 2) (IEEE h3)
    if mode == "academic":
        m = _RE_EN_H3_PAREN.match(s)
        if m and _is_short(s, 80):
            rest = m.group(2)
            if rest and rest[0].isupper():
                return 3
    # 1. Title (h1)
    m = _RE_EN_H1_NUM.match(s)
    if m and _is_short(s, 80):
        rest = m.group(2)
        if rest[:1].isupper() and _is_short(rest, 60):
            return 1
    # ALL CAPS bold short line -> h1 (business)
    if mode == "business" and _is_short(s, 40):
        letter_count = sum(1 for c in s if c.isalpha())
        upper_count = sum(1 for c in s if c.isupper())
        if letter_count > 3 and upper_count / letter_count > 0.85:
            return 1
    return None


def _merge_bullets_into_list(nodes: list[dict]) -> list[dict]:
    """把连续的 paragraph（以 bullet/序号开头）合并为 list_item 节点序列。"""
    out: list[dict] = []
    i = 0
    _list_counter = getattr(_merge_bullets_into_list, "_counter", 0)
    while i < len(nodes):
        n = nodes[i]
        if n.get("node_type") == "paragraph":
            is_list, ordered, text = _looks_like_list_item(n.get("text", ""))
            if is_list and text:
                _list_counter += 1
                list_id = f"auto_list{_list_counter}"
                # 收集连续 list 项
                items: list[tuple[str, bool]] = [(text, ordered)]
                j = i + 1
                while j < len(nodes):
                    n2 = nodes[j]
                    if n2.get("node_type") != "paragraph":
                        break
                    is_l2, ord2, t2 = _looks_like_list_item(n2.get("text", ""))
                    if is_l2 and t2 and (ord2 == ordered or not ordered):
                        items.append((t2, ord2))
                        j += 1
                    else:
                        break
                # 输出 list_item 节点（与 _expand_list 输出格式一致）
                start_idx = 0
                for k, (txt, _ord) in enumerate(items):
                    li = _node("list_item", {
                        "text": txt, "ordered": ordered, "level": 0, "list_id": list_id,
                    })
                    if k == 0:
                        li["_list_start"] = 1
                        li["_list_ordered"] = ordered
                    out.append(li)
                i = j
                _merge_bullets_into_list._counter = _list_counter
                continue
        out.append(n)
        i += 1
    return out


def _auto_structure_nodes(nodes: list[dict], mode: str,
                          lang_hint: str | None = None,
                          skip_title_detect: bool = False) -> tuple[list[dict], dict]:
    """P3-9 主入口：对扁平节点列表做二次分类。

    仅处理 type=paragraph 且 style=normal（即默认样式）的节点；
    其他节点（heading/list/table/figure/...）保持不变。

    返回 (new_nodes, meta_updates)。
    """
    # 先收集纯文本以检测语言
    body_texts = []
    for n in nodes:
        if n.get("node_type") == "paragraph":
            body_texts.append(n.get("text", ""))
    lang = _detect_lang(body_texts, lang_hint)
    resolved_mode = _guess_mode(mode, lang)
    # v1.4 P3-9c: 记录解析后的 lang/mode 到 meta_updates，renderer 可据此切换默认字体
    meta_updates: dict = {"_auto_lang": lang, "_auto_mode": resolved_mode}

    new_nodes: list[dict] = []
    # 文档头检测窗口（前几个 paragraph 尝试识别 title/subtitle/author）
    head_scanned = bool(skip_title_detect)

    for idx, n in enumerate(nodes):
        ntype = n.get("node_type")
        # 跳过非 paragraph 节点：原样保留
        if ntype != "paragraph":
            new_nodes.append(n)
            continue
        # 只处理默认样式 paragraph，尊重显式指定的 quote/code/abstract/footnote
        if n.get("style", "normal") != "normal":
            new_nodes.append(n)
            continue

        text = n.get("text", "")
        s = text.strip()

        # 阶段一：前 1-2 个非空 paragraph 尝试识别为 title（仅当 meta 还没有 title 时，
        # 且这一段短、后面还有内容、没有编号前缀）
        if not head_scanned and s and not meta_updates.get("title"):
            # 仅当该段短且不含标点长句时视为 title
            if _is_short(s, 40) and not _RE_CN_H3.match(s) and not _RE_CN_H1_PURE.match(s) \
                    and not _looks_like_caption(s) and not _looks_like_list_item(s)[0]:
                # 忽略明显的 front matter 关键词
                if s not in _CN_FRONT_KEYWORDS and s not in _EN_FRONT_KEYWORDS:
                    meta_updates["title"] = s
                    # 不把 title 作为 paragraph 写入 body（避免重复；title_block 会渲染 title）
                    head_scanned = True
                    continue
            head_scanned = True

        # 元数据标签行（关键词/作者/...）——保持为 paragraph 但不加粗提升
        if _RE_META_TAG.match(s):
            new_nodes.append(n)
            # 识别作者
            m = _RE_META_TAG.match(s)
            if m and m.group(1) in ("作者", "Author") and not meta_updates.get("author"):
                meta_updates["author"] = m.group(2).strip()
            continue

        # 空段落直接保留
        if not s:
            new_nodes.append(n)
            continue

        # 分隔线 → page_break
        if re.match(r"^[\-—=_\*·•]{3,}$", s):
            new_nodes.append(_node("page_break", {}))
            continue

        # 项目符号/编号行：先不合并，等合并阶段处理
        is_list, ordered, _ = _looks_like_list_item(s)
        if is_list and len(s) > 2:
            # 进一步判断：如果能匹配到 heading 编号规则，优先作为 heading
            lvl = None
            if lang == "cn":
                lvl = _detect_heading_level_cn(text, resolved_mode)
            else:
                lvl = _detect_heading_level_en(text, resolved_mode)
            if lvl is not None:
                h_text = _strip_heading_number(s, lvl, lang, resolved_mode) or s
                h_node = _node("heading", {"level": lvl, "text": h_text, "auto_structured": True})
                ov = _auto_heading_style(lang, resolved_mode, lvl)
                if ov:
                    h_node["style_override"] = dict(ov)
                new_nodes.append(h_node)
                continue
            # 否则保留为 paragraph，后续合并为 list
            new_nodes.append(n)
            continue

        # 检测 heading
        lvl = None
        if lang == "cn":
            lvl = _detect_heading_level_cn(text, resolved_mode)
        else:
            lvl = _detect_heading_level_en(text, resolved_mode)

        if lvl is not None:
            heading_text = _strip_heading_number(s, lvl, lang, resolved_mode) or s
            h_node = _node("heading", {"level": lvl, "text": heading_text, "auto_structured": True})
            ov = _auto_heading_style(lang, resolved_mode, lvl)
            if ov:
                h_node["style_override"] = dict(ov)
            new_nodes.append(h_node)
            continue

        # 短行 + 全加粗 + 紧跟正文 —— 作为 heading3 候选（保守处理：要求 bold=True 且短）
        if n.get("bold") and _is_short(s, 25) and idx + 1 < len(nodes) and \
                nodes[idx + 1].get("node_type") == "paragraph":
            h3_node = _node("heading", {"level": 3, "text": s, "auto_structured": True})
            ov3 = _auto_heading_style(lang, resolved_mode, 3)
            if ov3:
                h3_node["style_override"] = dict(ov3)
            new_nodes.append(h3_node)
            continue

        # 默认保留为 paragraph
        new_nodes.append(n)

    # 最后合并连续 bullet/numbered 段落为 list 节点
    new_nodes = _merge_bullets_into_list(new_nodes)

    # v1.4 P3-9b: caption 段落自动绑定到上一个 figure/table，或转为 equation 节点
    new_nodes = _bind_caption_nodes(new_nodes)
    return new_nodes, meta_updates


def _bind_caption_nodes(nodes: list[dict]) -> list[dict]:
    """第二阶段：把识别出的 caption paragraph 绑定到前一个 figure/table/equation。

    规则：
    - 如果 caption 段落紧接在一个 figure 节点之后，且该 figure 还没有 caption/label：
      把 caption_text 写入 figure 的 caption，自动分配 label。
    - 如果 caption 段落紧接在一个 table 节点之后，且该 table 还没有 caption/label：同理。
    - 公式 caption：将 paragraph 转换为一个 equation 节点（caption 作为 caption 文本），自动分配 label。
    - 独立 caption（前后无对应对象）：保留为 paragraph 但注入 style="caption" 以获得题注样式。
    - 已经是 figure/table/equation 节点的不会被二次处理；已有 label 的节点不覆盖。
    """
    out: list[dict] = []
    for n in nodes:
        if n.get("node_type") != "paragraph" or n.get("style", "normal") != "normal":
            out.append(n)
            continue
        text = n.get("text", "")
        kind, label_id, caption_text = _extract_caption_info(text)
        if kind is None:
            out.append(n)
            continue
        # 找到上一个非空有效节点（跳过空 paragraph）
        prev = None
        for back in reversed(out):
            if back.get("node_type") in {"figure", "table", "equation"}:
                prev = back
                break
            # 空 paragraph/title_block/toc 等允许被跳过（仅向后找第一个可绑定对象）
            if back.get("node_type") == "paragraph" and not (back.get("text") or "").strip():
                continue
            break

        if kind == "equation":
            # 公式 caption → equation 节点（LaTeX 为空占位，仅 caption + label，可被 {ref} 引用）
            eq_node = _node("equation", {
                "latex": "",
                "caption": caption_text,
                "display": "block",
                "label": label_id,
                "_caption_only": True,
            })
            out.append(eq_node)
            continue

        if kind == "figure" and prev is not None and prev.get("node_type") == "figure" \
                and not prev.get("caption") and not prev.get("label"):
            prev["caption"] = caption_text or ""
            prev["label"] = label_id
            # 不追加 paragraph（caption 由 figure 渲染器自行生成）
            continue

        if kind == "table" and prev is not None and prev.get("node_type") == "table" \
                and not prev.get("caption") and not prev.get("label"):
            prev["caption"] = caption_text or ""
            prev["label"] = label_id
            continue

        # 独立 caption 段落：保留为 paragraph，但打 style=caption 标记获得题注样式
        cap_p = _node("paragraph", {
            "text": text, "style": "caption",
        })
        out.append(cap_p)
    return out


def _strip_heading_number(text: str, level: int, lang: str, mode: str) -> str:
    """去掉标题开头的编号前缀，保留正文。去不掉则原样返回。"""
    s = text.strip()
    # 英文（先匹配以避免 CN 数字正则误吞 "1."）
    m = _RE_EN_APPX.match(s)
    if m:
        return (m.group(2) or f"Appendix {m.group(1)}").strip()
    m = _RE_EN_H3.match(s)
    if m:
        return m.group(4).strip()
    m = _RE_EN_H2.match(s)
    if m:
        return m.group(3).strip()
    m = _RE_EN_H1_NUM.match(s)
    if m:
        return m.group(2).strip()
    m = _RE_EN_H1_ROMAN.match(s)
    if m:
        return m.group(2).strip()
    m = _RE_EN_H2_LETTER.match(s)
    if m:
        return m.group(2).strip()
    m = _RE_EN_H3_PAREN.match(s)
    if m:
        return m.group(2).strip()
    # 中文
    m = _RE_CN_H1_CHAPTER.match(s)
    if m:
        return m.group(1).strip()
    m = _RE_CN_H1_PURE.match(s)
    if m:
        return m.group(1).strip()
    m = _RE_CN_H2.match(s) or _RE_CN_H2_FULL.match(s)
    if m:
        return m.group(1).strip()
    m = _RE_CN_H4_PAREN.match(s)
    if m:
        return m.group(1).strip()
    m = _RE_CN_H4.match(s)
    if m:
        rest = m.group(3).strip()
        return rest if rest else s
    m = _RE_CN_H3.match(s)
    if m:
        return m.group(2).strip()
    return s
