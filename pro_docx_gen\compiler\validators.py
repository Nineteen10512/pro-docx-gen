"""Validators — 输入验证与类型检查。

验证高层文档 JSON 结构、语义节点类型与必填字段。
所有验证错误抛出 ValueError（ValidationError），包含中文错误描述与修复建议。
"""

from typing import Any


# ─── 支持的语义节点类型（v1.2 扩展） ─────────────────────────────────

SUPPORTED_NODE_TYPES = {
    # v1.1
    "heading", "paragraph", "list", "table", "figure", "chart",
    "kpi_card", "callout", "page_break", "toc", "references", "appendix",
    # v1.2 新增
    "revision", "comment", "footnote", "endnote",
    "watermark", "page_border", "equation",
    "signature_block", "signature_line",
    # v1.3 新增
    "svg_shape",
    # v1.4 P2-4 新增
    "ref",
}

PARAGRAPH_STYLES = {"normal", "quote", "code", "abstract", "footnote"}
CALLOUT_VARIANTS = {"info", "warning", "success", "danger"}
LIST_MAX_LEVEL = 4
LINE_NUMBERS_RESTART = {"continuous", "new_page", "new_section"}
AUTO_STRUCTURE_MODES = {False, "academic", "business", "auto", "false"}

SUPPORTED_CHART_TYPES = {
    "column", "bar", "line", "pie", "doughnut",
    "area", "scatter", "radar", "stacked_column", "stacked_bar",
}

SUPPORTED_PAGE_SIZES = {"A4", "Letter", "A3", "B5", "Legal"}
SUPPORTED_ORIENTATIONS = {"portrait", "landscape"}


class ValidationError(ValueError):
    """结构化验证错误（中文错误信息 + 修复建议）。"""

    def __init__(self, path: str, message: str, suggestion: str = ""):
        self.path = path
        self.message = message
        self.suggestion = suggestion
        # Internal English note (preserved as attribute for debugging, not shown to user)
        full = f"[{path}] {message}"
        if suggestion:
            full += f"。修复建议：{suggestion}"
        super().__init__(full)


def _require(data: dict, key: str, path: str, types: tuple | type):
    if key not in data:
        raise ValidationError(
            path, f"缺少必填字段 '{key}'",
            f"请在该节点中添加 '{key}' 字段，类型应为 {_type_name(types)}",
        )
    if not isinstance(data[key], types):
        raise ValidationError(
            path,
            f"字段 '{key}' 类型错误，应为 {_type_name(types)}，实际为 {type(data[key]).__name__}",
            f"请将 '{key}' 的值改为 {_type_name(types)} 类型",
        )


def _type_name(types) -> str:
    if isinstance(types, tuple):
        return " / ".join(t.__name__ for t in types)
    return types.__name__


def _optional(data: dict, key: str, types: tuple | type, default=None):
    if key not in data:
        return default
    if not isinstance(data[key], types):
        return default
    return data[key]


# ─── 文档结构验证 ──────────────────────────────────────────────────

def validate_document(doc: dict) -> None:
    """验证完整文档结构。"""
    if not isinstance(doc, dict):
        # English: Document must be a dict
        raise ValidationError(
            "root", "文档根节点必须是 dict（JSON 对象）",
            "请将传入内容包装为 { \"meta\": {...}, \"sections\": [...] } 的字典结构",
        )

    # meta
    if "meta" not in doc:
        # English: Missing required field 'meta'
        raise ValidationError(
            "root", "缺少必填字段 'meta'",
            "请添加 meta 对象，至少包含 title 字段，如 {\"meta\": {\"title\": \"文档标题\"}}",
        )
    _require(doc, "meta", "root", dict)
    meta = doc["meta"]
    _require(meta, "title", "meta", str)
    _optional(meta, "author", str)
    _optional(meta, "institution", str)
    _optional(meta, "date", str)
    _optional(meta, "subtitle", str)
    # v1.2: 扩展元数据
    for k in ("keywords", "category", "comments", "status", "subject",
              "company", "manager"):
        _optional(meta, k, (str, list))
    # v1.2: page_setup
    if "page_setup" in meta:
        _validate_page_setup(meta["page_setup"], "meta.page_setup")

    # v1.4 P3-4: 文档级分栏
    if "columns" in meta:
        _validate_columns(meta["columns"], "meta.columns")

    # v1.4 P3-6: 文档级行号
    if "line_numbers" in meta:
        _validate_line_numbers(meta["line_numbers"], "meta.line_numbers")

    # v1.4 P3-9: 智能结构识别
    if "auto_structure" in meta:
        _validate_auto_structure(meta["auto_structure"], "meta.auto_structure")

    # theme
    theme = doc.get("theme", "academic")
    if not isinstance(theme, (str, dict)):
        # English: 'theme' must be a string or dict
        raise ValidationError(
            "root", "'theme' 必须是字符串（主题名）或 dict（自定义 tokens）",
            "可选主题：academic/business/teaching 等，或直接传入 tokens 字典",
        )

    # abstract
    if "abstract" in doc:
        abs_data = doc["abstract"]
        if not isinstance(abs_data, dict):
            # English: Abstract must be a dict
            raise ValidationError(
                "abstract", "abstract 字段必须是 dict",
                "请使用 {\"text\": \"摘要内容\", \"keywords\": [...]} 结构",
            )
        _optional(abs_data, "text", str, "")
        if "keywords" in abs_data and not isinstance(abs_data["keywords"], list):
            # English: 'keywords' must be a list of strings
            raise ValidationError(
                "abstract", "'keywords' 必须是字符串列表",
                "请将 keywords 改为 [\"关键词1\", \"关键词2\", ...] 的列表格式",
            )

    # sections
    if "sections" not in doc:
        # English: Missing required field 'sections'
        raise ValidationError(
            "root", "缺少必填字段 'sections'",
            "请添加 sections 数组，每个元素包含 title 和 content 字段",
        )
    _require(doc, "sections", "root", list)
    for i, sec in enumerate(doc["sections"]):
        _validate_section(sec, f"sections[{i}]")

    # v1.2: 文档级配置
    if "watermark" in doc:
        _validate_watermark(doc["watermark"], "watermark")
    if "page_border" in doc:
        _validate_page_border(doc["page_border"], "page_border")

    # references（doc 级，兼容旧格式）
    if "references" in doc:
        refs = doc["references"]
        if not isinstance(refs, list):
            # English: References must be a list
            raise ValidationError(
                "references", "references 必须是列表",
                "请将 references 改为 [{...}, {...}] 数组格式，或使用新版 references 节点",
            )
        for i, ref in enumerate(refs):
            _validate_reference_item(ref, f"references[{i}]")

    # appendices
    if "appendices" in doc:
        if not isinstance(doc["appendices"], list):
            # English: Appendices must be a list
            raise ValidationError(
                "appendices", "appendices 必须是列表",
                "请将 appendices 改为 [{\"title\": \"...\", \"content\": [...]}, ...] 数组",
            )

    # header/footer（v1.2 扩展）
    if "header" in doc and doc["header"] is not None:
        if not isinstance(doc["header"], dict):
            # English: header must be a dict
            raise ValidationError(
                "header", "header 必须是 dict",
                "请使用 {\"left\": \"...\", \"center\": \"...\", \"right\": \"...\"} 结构",
            )
    if "footer" in doc and doc["footer"] is not None:
        if not isinstance(doc["footer"], dict):
            # English: footer must be a dict
            raise ValidationError(
                "footer", "footer 必须是 dict",
                "请使用 {\"left\": \"...\", \"center\": \"...\", \"right\": \"...\"} 结构",
            )


def _validate_page_setup(ps: dict, path: str):
    _optional(ps, "size", str, "A4")
    if "size" in ps and ps["size"] not in SUPPORTED_PAGE_SIZES:
        # English: page size must be in {…}
        raise ValidationError(
            path, f"不支持的页面尺寸 '{ps['size']}'",
            f"请从以下尺寸中选择：{sorted(SUPPORTED_PAGE_SIZES)}",
        )
    if "orientation" in ps and ps["orientation"] not in SUPPORTED_ORIENTATIONS:
        # English: orientation must be in {…}
        raise ValidationError(
            path, f"不支持的页面方向 '{ps['orientation']}'",
            f"请使用 'portrait'（纵向）或 'landscape'（横向）",
        )
    for k in ("margin_top", "margin_bottom", "margin_left", "margin_right", "gutter",
              "header_distance", "footer_distance"):
        if k in ps and not isinstance(ps[k], (int, float)):
            # English: '{k}' must be a number (inches)
            raise ValidationError(
                path, f"字段 '{k}' 必须是数字（单位：英寸）",
                "请传入整数或浮点数，例如 1.0 表示 1 英寸",
            )
    for k in ("different_first_page", "different_odd_even"):
        if k in ps and not isinstance(ps[k], bool):
            # English: '{k}' must be bool
            raise ValidationError(
                path, f"字段 '{k}' 必须是布尔值",
                "请传入 true 或 false",
            )


def _validate_columns(cols: Any, path: str):
    """v1.4 P3-4: columns 字段 — int 或 {count:int, space?:int, sep?:bool}。"""
    if isinstance(cols, int):
        if not (1 <= cols <= 9):
            raise ValidationError(
                path, f"columns 为整数时取值范围应为 1-9，实际为 {cols}",
                "例如 columns=2 表示等宽两栏",
            )
        return
    if isinstance(cols, dict):
        if "count" not in cols:
            raise ValidationError(
                path, "columns 为 dict 时必须包含 'count' 字段",
                "例如 {\"count\": 3, \"space\": 720000} 表示三栏，栏间距 0.5 英寸（720000 EMU）",
            )
        cnt = cols["count"]
        if not isinstance(cnt, int) or not (1 <= cnt <= 9):
            raise ValidationError(
                path, f"columns.count 必须是 1-9 的整数，实际为 {cnt!r}",
                "请传入整数，例如 2、3",
            )
        if "space" in cols:
            sp = cols["space"]
            if not isinstance(sp, int) or sp < 0:
                raise ValidationError(
                    path, f"columns.space 必须是非负整数（EMU 单位），实际为 {sp!r}",
                    "720000 EMU = 0.5 英寸，360000 EMU = 0.25 英寸",
                )
        if "sep" in cols and not isinstance(cols["sep"], bool):
            raise ValidationError(
                path, "columns.sep 必须是布尔值（true=显示分隔线）",
                "sep=true 时栏间显示竖线",
            )
        return
    raise ValidationError(
        path, f"columns 必须是 int 或 dict，实际为 {type(cols).__name__}",
        "整数写法：columns=2；dict 写法：{\"count\": 3, \"space\": 720000}",
    )


def _validate_line_numbers(ln: Any, path: str):
    """v1.4 P3-6: line_numbers 字段 dict。"""
    if not isinstance(ln, dict):
        raise ValidationError(
            path, f"line_numbers 必须是 dict，实际为 {type(ln).__name__}",
            "例如 {\"start\": 1, \"increment\": 5, \"restart\": \"new_page\"}",
        )
    if "start" in ln:
        s = ln["start"]
        if not isinstance(s, int) or s < 1:
            raise ValidationError(
                path, f"line_numbers.start 必须是 ≥1 的整数，实际为 {s!r}",
                "起始行号，默认 1",
            )
    if "increment" in ln:
        inc = ln["increment"]
        if not isinstance(inc, int) or inc < 1:
            raise ValidationError(
                path, f"line_numbers.increment 必须是 ≥1 的整数，实际为 {inc!r}",
                "每隔多少行显示一次行号，1=每行显示，5=每 5 行显示",
            )
    if "restart" in ln:
        r = ln["restart"]
        if r not in LINE_NUMBERS_RESTART:
            raise ValidationError(
                path, f"line_numbers.restart 必须是 continuous/new_page/new_section，实际为 {r!r}",
                "continuous=连续编号，new_page=每页重启，new_section=每节重启",
            )
    if "distance" in ln:
        d = ln["distance"]
        if not isinstance(d, int) or d < 0:
            raise ValidationError(
                path, f"line_numbers.distance 必须是非负整数（EMU 单位），实际为 {d!r}",
                "默认 360000 EMU（0.25 英寸）",
            )


def _validate_section(sec: dict, path: str):
    if not isinstance(sec, dict):
        # English: Section must be a dict
        raise ValidationError(
            path, "section 必须是 dict",
            "请使用 {\"title\": \"章节标题\", \"content\": [...]} 结构",
        )
    _require(sec, "title", path, str)
    if "content" not in sec:
        # English: Missing required field 'content'
        raise ValidationError(
            path, "缺少必填字段 'content'",
            "请添加 content 数组，里面放置段落/列表/表格等语义节点",
        )
    if not isinstance(sec["content"], list):
        # English: 'content' must be a list
        raise ValidationError(
            path, "'content' 必须是列表",
            "请将 content 改为 [{\"type\": \"paragraph\", \"text\": \"...\"}, ...] 数组",
        )
    _optional(sec, "level", int, 1)
    for i, node in enumerate(sec["content"]):
        _validate_node(node, f"{path}.content[{i}]")


def _validate_node(node: Any, path: str):
    """验证单个语义节点。"""
    if not isinstance(node, dict):
        # English: Node must be a dict
        raise ValidationError(
            path, f"节点必须是 dict，实际为 {type(node).__name__}",
            "每个语义节点必须是 {\"type\": \"...\", ...} 结构；字符串请放在 paragraph/list 节点内",
        )
    if "type" not in node:
        # English: Node must have 'type' field
        raise ValidationError(
            path, "节点缺少必填字段 'type'",
            "请添加 type 字段，取值见 SUPPORTED_NODE_TYPES（heading/paragraph/list/table 等）",
        )

    ntype = node["type"]
    if ntype not in SUPPORTED_NODE_TYPES:
        # English: Unknown node type '{ntype}'
        raise ValidationError(
            path, f"未知节点类型 {ntype!r}",
            f"请检查 NODE_TYPES 列表，合法类型：{sorted(SUPPORTED_NODE_TYPES)}；或使用 generate_from_markdown 让系统自动解析",
        )

    validators = {
        "heading": _validate_heading,
        "paragraph": _validate_paragraph,
        "list": _validate_list,
        "table": _validate_table,
        "figure": _validate_figure,
        "chart": _validate_chart,
        "kpi_card": _validate_kpi_card,
        "callout": _validate_callout,
        "page_break": lambda n, p: None,
        "toc": _validate_toc,
        "references": _validate_references_node,
        "appendix": _validate_appendix_node,
        # v1.2
        "revision": _validate_revision,
        "comment": _validate_comment,
        "footnote": _validate_footnote,
        "endnote": _validate_footnote,  # 结构同 footnote
        "watermark": _validate_watermark,
        "page_border": _validate_page_border,
        "equation": _validate_equation,
        "signature_block": _validate_signature_block,
        "signature_line": _validate_signature_line,
        # v1.3
        "svg_shape": _validate_svg_shape,
        # v1.4 P2-4
        "ref": _validate_ref,
    }
    validators[ntype](node, path)


def _validate_svg_shape(n: dict, p: str):
    _require(n, "svg", p, str)
    _optional(n, "width", (str, int, float), "5cm")
    _optional(n, "height", (str, int, float), None)
    _optional(n, "align", str, "center")


def _validate_heading(n: dict, p: str):
    _require(n, "text", p, str)
    lvl = _optional(n, "level", int, 1)
    if not (1 <= lvl <= 5):
        # English: Heading level must be 1-5
        raise ValidationError(
            p, f"标题级别 {lvl} 超出范围",
            "heading.level 必须在 1-5 之间（对应 H1-H5）",
        )
    _optional(n, "comment", dict)


def _validate_paragraph(n: dict, p: str):
    _require(n, "text", p, str)
    style = _optional(n, "style", str, "normal")
    if style not in PARAGRAPH_STYLES:
        # English: Paragraph style '{style}' not in {…}
        raise ValidationError(
            p, f"不支持的段落样式 {style!r}",
            f"请从以下样式中选择：{sorted(PARAGRAPH_STYLES)}",
        )
    _optional(n, "comment", dict)
    _optional(n, "bold", bool)
    _optional(n, "italic", bool)
    # v1.4 P3-5: 首字下沉
    _optional(n, "drop_cap", bool, False)
    # v1.4 P3-4: 段落级分栏（段末分节符）
    if "columns" in n:
        _validate_columns(n["columns"], f"{p}.columns")


def _validate_list(n: dict, p: str):
    if "items" not in n:
        # English: List must have 'items'
        raise ValidationError(
            p, "列表节点缺少 'items' 字段",
            "请添加 items 数组，例如 {\"type\": \"list\", \"items\": [\"项目1\", \"项目2\"]}",
        )
    if not isinstance(n["items"], list):
        # English: 'items' must be a list
        raise ValidationError(
            p, "'items' 必须是列表",
            "请将 items 改为字符串数组或 {text, items} 对象数组（支持嵌套子列表）",
        )
    _optional(n, "ordered", bool, True)
    _optional(n, "start", int, 1)
    for i, item in enumerate(n["items"]):
        _validate_list_item(item, f"{p}.items[{i}]", 0, top_ordered=n.get("ordered", True))


def _validate_list_item(item: Any, p: str, depth: int, top_ordered: bool = True):
    if depth > LIST_MAX_LEVEL:
        # English: List nesting exceeds max depth
        raise ValidationError(
            p, f"列表嵌套层级超过最大深度 {LIST_MAX_LEVEL}",
            f"最多支持 {LIST_MAX_LEVEL} 层嵌套，请扁平化部分列表项",
        )
    if isinstance(item, str):
        return
    if isinstance(item, dict):
        # v1.4 P2-3: support nested list via item.items in addition to legacy sub_items
        has_nested = ("items" in item) or ("sub_items" in item)
        if not has_nested:
            # bare text leaf — require text
            _require(item, "text", p, str)
        else:
            # nested list item: text is optional (may be implicit parent)
            if "text" in item and not isinstance(item["text"], str):
                raise ValidationError(
                    p, "列表项 'text' 字段必须是字符串",
                    "父级列表项的 text 为显示在该层的文本；子列表请放在 items 数组内",
                )
            _optional(item, "text", str, "")
            # nested items
            nested_key = "items" if "items" in item else "sub_items"
            nested = item.get(nested_key)
            if not isinstance(nested, list):
                raise ValidationError(
                    p, f"列表项 '{nested_key}' 必须是列表",
                    "嵌套子列表必须是数组，可递归包含字符串或 {text, items} 对象",
                )
            _optional(item, "ordered", bool, top_ordered)
            for j, si in enumerate(nested):
                _validate_list_item(si, f"{p}.{nested_key}[{j}]", depth + 1, top_ordered=top_ordered)
            # legacy sub_items → normalized downstream
        _optional(item, "level", int, depth)
        return
    # English: List item must be a string or dict
    raise ValidationError(
        p, f"列表项必须是字符串或 dict，实际为 {type(item).__name__}",
        "列表项可以是纯字符串，也可以是 {\"text\": \"...\", \"items\": [...]} 对象（支持嵌套子列表）",
    )


def _validate_table(n: dict, p: str):
    _require(n, "headers", p, list)
    _require(n, "rows", p, list)
    for h in n["headers"]:
        if not isinstance(h, str):
            # English: All headers must be strings
            raise ValidationError(
                p, "表头 headers 中所有元素必须是字符串",
                "请将 headers 改为 [\"列1\", \"列2\", ...] 字符串数组",
            )
    for i, row in enumerate(n["rows"]):
        if not isinstance(row, list):
            # English: Row must be a list
            raise ValidationError(
                f"{p}.rows[{i}]", "表格行必须是列表",
                "每一行应为与 headers 等长的单元格数组",
            )
        if len(row) != len(n["headers"]):
            # English: Row has N cells, expected M
            raise ValidationError(
                f"{p}.rows[{i}]",
                f"第 {i+1} 行有 {len(row)} 个单元格，应与表头列数 {len(n['headers'])} 一致",
                "请补齐或删减该行的单元格，使其与 headers 长度相同",
            )
    if "col_widths" in n:
        if not isinstance(n["col_widths"], list):
            # English: 'col_widths' must be a list
            raise ValidationError(
                p, "'col_widths' 必须是列表",
                "请传入与 headers 等长的数字数组表示各列宽度（英寸）",
            )
    _optional(n, "caption", str)
    _optional(n, "title", str)  # v1.4 P2-4 alias
    _optional(n, "header_repeat", bool, True)
    _optional(n, "label", str)  # v1.4 P2-4: cross-reference label (e.g. "tbl1")


def _validate_figure(n: dict, p: str):
    # v1.4 P2-4: path is optional when label/caption is provided (caption-only figure placeholder)
    if "path" in n:
        _require(n, "path", p, str)
    _optional(n, "caption", str)
    _optional(n, "title", str)  # v1.4 P2-4 alias for caption
    _optional(n, "width_inches", (int, float))
    _optional(n, "align", str, "center")
    # v1.4 P2-4: optional explicit label id for cross-reference
    _optional(n, "label", str)


def _validate_chart(n: dict, p: str):
    chart_type = _optional(n, "chart_type", str, "column")
    if chart_type not in SUPPORTED_CHART_TYPES:
        # English: chart_type '…' not supported
        raise ValidationError(
            p, f"不支持的图表类型 {chart_type!r}",
            f"请从以下类型中选择：{sorted(SUPPORTED_CHART_TYPES)}",
        )

    categories = n.get("categories")
    if chart_type not in ("pie", "doughnut"):
        if "categories" not in n:
            # English: Missing required field 'categories'
            raise ValidationError(
                p, "缺少必填字段 'categories'",
                "非饼图/环形图必须提供 categories 分类轴数据数组",
            )
        if not isinstance(categories, list) or len(categories) == 0:
            # English: 'categories' must be a non-empty list
            raise ValidationError(
                p, "'categories' 必须是非空列表",
                "请传入至少一个分类标签，例如 [\"Q1\", \"Q2\", \"Q3\"]",
            )

    if "series" not in n:
        # English: Missing required field 'series'
        raise ValidationError(
            p, "缺少必填字段 'series'",
            "请添加 series 数组，每个元素包含 name 和 values 字段",
        )
    if not isinstance(n["series"], list) or len(n["series"]) == 0:
        # English: 'series' must be a non-empty list
        raise ValidationError(
            p, "'series' 必须是非空列表",
            "请至少传入一个数据系列",
        )
    for i, s in enumerate(n["series"]):
        if not isinstance(s, dict):
            # English: series item must be a dict
            raise ValidationError(
                f"{p}.series[{i}]", "series 元素必须是 dict",
                "每个系列应为 {\"name\": \"系列名\", \"values\": [...]} 结构",
            )
        if "values" not in s:
            # English: series item must have 'values'
            raise ValidationError(
                f"{p}.series[{i}]", "series 元素缺少 'values' 字段",
                "请添加 values 数组，其长度应与 categories 一致",
            )
        if not isinstance(s["values"], list):
            # English: 'values' must be a list
            raise ValidationError(
                f"{p}.series[{i}]", "'values' 必须是列表",
                "请将 values 改为数字数组",
            )
    _optional(n, "title", str)
    _optional(n, "caption", str)
    _optional(n, "show_legend", bool, True)
    _optional(n, "legend_position", str, "bottom")
    _optional(n, "x_title", str)
    _optional(n, "y_title", str)
    _optional(n, "aspect_ratio", str, "4:3")
    _optional(n, "width_pct", (int, float), 1.0)
    _optional(n, "align", str, "center")
    _optional(n, "show_data_labels", bool)
    _optional(n, "number_format", str, "0")
    _optional(n, "mode", str, "image")


def _validate_kpi_card(n: dict, p: str):
    _require(n, "value", p, (str, int, float))
    _require(n, "label", p, str)
    _optional(n, "subtext", str)


def _validate_callout(n: dict, p: str):
    variant = _optional(n, "variant", str, "info")
    if variant not in CALLOUT_VARIANTS:
        # English: variant must be in {…}
        raise ValidationError(
            p, f"不支持的 callout 样式 {variant!r}",
            f"请从以下 variant 中选择：{sorted(CALLOUT_VARIANTS)}",
        )
    _require(n, "body", p, str)
    _optional(n, "title", str)


def _validate_toc(n: dict, p: str):
    """v1.4 P1-3: TOC 目录节点，支持可选 title 与 levels。"""
    if "title" in n and not isinstance(n["title"], str):
        raise ValidationError(
            p, "'title' 必须是字符串",
            "title 为目录标题文字，例如 \"目  录\" 或 \"Table of Contents\"",
        )
    if "levels" in n:
        lv = n["levels"]
        if isinstance(lv, int):
            if not (1 <= lv <= 5):
                raise ValidationError(
                    p, "'levels' 作为整数时必须在 1-5 之间",
                    "例如 levels=3 表示包含 1-3 级标题",
                )
        elif isinstance(lv, list):
            if not lv or not all(isinstance(x, int) and 1 <= x <= 5 for x in lv):
                raise ValidationError(
                    p, "'levels' 列表元素必须是 1-5 的整数",
                    "例如 [1,2,3] 表示包含 H1-H3",
                )
        else:
            raise ValidationError(
                p, "'levels' 必须是 int 或 list[int]",
                "请传入整数（如 3）或显式列表（如 [1,2,3]）",
            )


def _validate_references_node(n: dict, p: str):
    """v1.4 P1-4: references 语义节点。"""
    if "items" not in n:
        raise ValidationError(
            p, "references 节点缺少 'items' 字段",
            "请添加 items 数组，每个元素为一条参考文献条目（含 type/authors/title 等字段）",
        )
    if not isinstance(n["items"], list):
        raise ValidationError(
            p, "'items' 必须是列表",
            "请将 items 改为参考文献对象数组",
        )
    # citation_style 可选
    _optional(n, "citation_style", str, "apa")
    _optional(n, "title", str)
    for i, ref in enumerate(n["items"]):
        _validate_reference_item(ref, f"{p}.items[{i}]")


def _validate_appendix_node(n: dict, p: str):
    _require(n, "title", p, str)
    _optional(n, "content", list, [])


def _validate_reference_item(ref: dict, p: str):
    if not isinstance(ref, dict):
        # English: Reference must be a dict
        raise ValidationError(
            p, "参考文献条目必须是 dict",
            "每条参考文献应为 {\"type\": \"article/book/...\", \"authors\": \"...\", \"title\": \"...\", ...} 结构",
        )
    _optional(ref, "type", str)
    _optional(ref, "authors", str)
    _optional(ref, "year", (str, int))
    _optional(ref, "title", str)
    _optional(ref, "source", str)
    # v1.4 P1-4: 详细字段（宽松校验，缺字段不报错，渲染时跳过）
    for k in ("journal", "publisher", "school", "conference", "site",
              "volume", "issue", "pages", "doi", "url", "place", "isbn",
              "accessed", "degree"):
        _optional(ref, k, (str, int))


# ─── v1.2 新节点验证 ────────────────────────────────────────────

def _validate_revision(n: dict, p: str):
    action = _optional(n, "action", str, "insert")
    if action not in {"insert", "delete", "replace"}:
        # English: revision.action must be insert|delete|replace
        raise ValidationError(
            p, f"revision.action 必须是 insert/delete/replace，实际为 {action!r}",
            "insert=插入、delete=删除、replace=替换（需同时提供 old_text 和 new_text）",
        )
    if action == "insert":
        _require(n, "text", p, str)
    elif action == "delete":
        _require(n, "text", p, str)
    elif action == "replace":
        _require(n, "old_text", p, str)
        _require(n, "new_text", p, str)
    _optional(n, "author", str)
    _optional(n, "date", str)


def _validate_comment(n: dict, p: str):
    """comment 节点：既可以是行内批注（含 text），也可以包裹段落。"""
    _require(n, "text", p, str)
    _optional(n, "author", str)
    _optional(n, "date", str)


def _validate_footnote(n: dict, p: str):
    _require(n, "text", p, str)
    _optional(n, "id", int)


def _validate_watermark(n: dict, p: str):
    if not isinstance(n, dict):
        # English: watermark must be a dict
        raise ValidationError(
            p, "watermark 必须是 dict",
            "请使用 {\"enabled\": true, \"text\": \"DRAFT\", \"rotation\": -45} 结构",
        )
    _optional(n, "enabled", bool, True)
    _optional(n, "text", str, "DRAFT")
    _optional(n, "image_path", str)
    _optional(n, "rotation", (int, float), -45)


def _validate_page_border(n: dict, p: str):
    if not isinstance(n, dict):
        # English: page_border must be a dict
        raise ValidationError(
            p, "page_border 必须是 dict",
            "请使用 {\"enabled\": true, \"style\": \"single\"} 结构",
        )
    _optional(n, "enabled", bool, True)
    _optional(n, "style", str, "single")
    _optional(n, "offset_from", str, "page")


def _validate_equation(n: dict, p: str):
    _require(n, "latex", p, str)
    _optional(n, "caption", str)
    _optional(n, "title", str)  # v1.4 P2-4 alias
    _optional(n, "inline", bool, False)
    _optional(n, "display", str, "block")
    _optional(n, "label", str)  # v1.4 P2-4 cross-reference label (e.g. "eq1")


def _validate_ref(n: dict, p: str):
    """v1.4 P2-4: cross-reference node {type:"ref", target:"fig1", prefix:"图"}."""
    _require(n, "target", p, str)
    _optional(n, "prefix", str, "")
    _optional(n, "suffix", str, "")
    _optional(n, "style", str, "inline")  # inline / above / below



def _validate_signature_block(n: dict, p: str):
    _optional(n, "name", str, "")
    _optional(n, "date", str, "")
    _optional(n, "title", str, "签字")
    _optional(n, "lines", int, 1)


def _validate_signature_line(n: dict, p: str):
    _optional(n, "signer", str, "签字人")
    _optional(n, "date", str, "日期")


def _validate_auto_structure(v: Any, path: str):
    """v1.4 P3-9: auto_structure 字段 bool | 'academic' | 'business' | 'auto'。"""
    if isinstance(v, bool):
        return
    if isinstance(v, str):
        if v in ("false", "off", "none", "0"):
            return
        if v in ("academic", "business", "auto"):
            return
    raise ValidationError(
        path,
        f"auto_structure 取值非法：{v!r}",
        "请使用 false（关闭）、'academic'（学术论文）、'business'（商业报告）或 'auto'（自动检测）",
    )
