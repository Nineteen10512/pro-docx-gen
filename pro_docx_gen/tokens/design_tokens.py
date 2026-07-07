"""Design Tokens — 设计令牌系统。

所有字号、间距、颜色、缩进等样式数值统一在此定义。
LLM/用户只需引用 token 名称，无需（也不应）传入具体数值。
"""

from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


# ─── 页面尺寸预设 ─────────────────────────────────────────────────

PAGE_SIZES = {
    "A4": (Inches(8.27), Inches(11.69)),
    "Letter": (Inches(8.5), Inches(11)),
    "A3": (Inches(11.69), Inches(16.54)),
    "B5": (Inches(6.93), Inches(9.84)),
    "Legal": (Inches(8.5), Inches(14)),
}


# ─── 基础令牌（academic 主题的默认值） ────────────────────────────────

BASE_TOKENS = {
    "font": {
        "family": {
            "default": "Times New Roman",
            "cn": "宋体",
            "heading": "Times New Roman",
            "cn_heading": "黑体",
            "code": "Consolas",
        },
        "size": {
            "title": Pt(22),
            "subtitle": Pt(14),
            "h1": Pt(16),
            "h2": Pt(14),
            "h3": Pt(13),
            "h4": Pt(12),
            "h5": Pt(12),
            "body": Pt(12),          # 小四
            "abstract": Pt(10.5),   # 五号
            "caption": Pt(10.5),
            "footnote": Pt(9),      # 小五
            "reference": Pt(10.5),
            "code": Pt(10.5),
            "toc": Pt(12),
            "header": Pt(9),
            "footer": Pt(9),
            "watermark": Pt(60),
            "chart_title": Pt(11),  # 图表标题字号
            "chart_caption": Pt(9), # 图表题注字号
            "chart_axis": Pt(9),    # 坐标轴刻度字号
            "chart_legend": Pt(10), # 图例字号
            "comment": Pt(9),
        },
        "weight": {
            "bold": True,
            "normal": False,
        },
    },
    "page": {
        "size": "A4",
        "orientation": "portrait",  # portrait | landscape
        "margin_top": Inches(1),
        "margin_bottom": Inches(1),
        "margin_left": Inches(1),
        "margin_right": Inches(1),
        "gutter": Inches(0),
        "header_distance": Inches(0.5),
        "footer_distance": Inches(0.5),
        "different_first_page": False,
        "different_odd_even": False,
    },
    "spacing": {
        # 保留 page_margin 以兼容旧代码（用于 layout calculator 等）
        "page_margin": Inches(1),
        "page_width": Inches(8.27),   # A4
        "page_height": Inches(11.69),
        "before_title": Pt(0),
        "after_title": Pt(12),
        "before_subtitle": Pt(0),
        "after_subtitle": Pt(24),
        "before_meta": Pt(6),
        "after_meta": Pt(6),
        "before_h1": Pt(24),
        "after_h1": Pt(12),
        "before_h2": Pt(18),
        "after_h2": Pt(6),
        "before_h3": Pt(12),
        "after_h3": Pt(6),
        "before_h4": Pt(6),
        "after_h4": Pt(3),
        "before_h5": Pt(6),
        "after_h5": Pt(3),
        "paragraph_before": Pt(0),
        "paragraph_after": Pt(6),
        "line_spacing": 2.0,          # academic 默认双倍行距
        "first_line_indent": Inches(0.3),  # 约2个中文字符
        "list_indent": Inches(0.25),
        "quote_indent": Inches(0.4),
        "code_indent": Inches(0.2),
    },
    "color": {
        "text": RGBColor(0x33, 0x33, 0x33),
        "title": RGBColor(0x1F, 0x38, 0x64),
        "heading": RGBColor(0x1F, 0x38, 0x64),
        "accent": RGBColor(0x2E, 0x75, 0xB6),
        "muted": RGBColor(0x66, 0x66, 0x66),
        "table_header_bg": RGBColor(0x1F, 0x38, 0x64),
        "table_header_text": RGBColor(0xFF, 0xFF, 0xFF),
        "table_alt_row": RGBColor(0xF2, 0xF2, 0xF2),
        "table_border": RGBColor(0xBF, 0xBF, 0xBF),
        "quote_border": RGBColor(0x2E, 0x75, 0xB6),
        "code_bg": RGBColor(0xF5, 0xF5, 0xF5),
        "callout_info_bg": RGBColor(0xE7, 0xF3, 0xF8),
        "callout_info_border": RGBColor(0x2E, 0x75, 0xB6),
        "callout_warning_bg": RGBColor(0xFF, 0xF4, 0xE5),
        "callout_warning_border": RGBColor(0xED, 0x7D, 0x31),
        "callout_success_bg": RGBColor(0xE8, 0xF5, 0xE9),
        "callout_success_border": RGBColor(0x43, 0xA0, 0x47),
        "callout_danger_bg": RGBColor(0xFD, 0xEA, 0xEA),
        "callout_danger_border": RGBColor(0xE5, 0x39, 0x35),
        # 修订追踪颜色
        "revision_insert": RGBColor(0xC0, 0x00, 0x00),
        "revision_delete": RGBColor(0xC0, 0x00, 0x00),
        "comment": RGBColor(0x80, 0x80, 0x00),
        "watermark": RGBColor(0xCC, 0xCC, 0xCC),
        "header_text": RGBColor(0x66, 0x66, 0x66),
        "footer_text": RGBColor(0x66, 0x66, 0x66),
        "page_border": RGBColor(0xBF, 0xBF, 0xBF),
    },
    "header": {
        "text": "",
        "show_on_first_page": True,
        "different_odd_even": False,
        "include_page_x_of_y": False,
        "image_path": None,
        "image_width_inches": None,
    },
    "footer": {
        "text": "",
        "show_on_first_page": True,
        "different_odd_even": False,
        "page_number": True,
        "page_x_of_y": False,  # "第 X 页 / 共 Y 页"
        "page_x_of_y_cn": False,  # 中文格式"第 X 页，共 Y 页"
    },
    "watermark": {
        "enabled": False,
        "text": "DRAFT",
        "color": RGBColor(0xCC, 0xCC, 0xCC),
        "font_size": Pt(60),
        "rotation": -45,
        "image_path": None,
    },
    "page_border": {
        "enabled": False,
        "offset_from": "page",  # "page" | "text"
        "style": "single",
        "size": 6,  # 1/8 pt
        "color": RGBColor(0xBF, 0xBF, 0xBF),
        "space": 24,
    },
    "revision": {
        "author": "ProDocx Gen",
        "insert_color": RGBColor(0xC0, 0x00, 0x00),
        "delete_color": RGBColor(0xC0, 0x00, 0x00),
    },
    "comment": {
        "author": "ProDocx Gen",
    },
    "chart": {
        "dpi": 300,
        "default_aspect": "4:3",        # 4:3 / 16:9
        "default_width_pct": 1.0,       # 占正文宽度比例
        "alpha": 0.85,                  # 填充透明度
        "gridline_alpha": 0.5,          # 网格线透明度
        "marker_size": 5,
        "line_width": 2.0,
        "bar_width": 0.6,
        "palette": [
            RGBColor(0x1F, 0x38, 0x64),  # deep navy（academic 默认）
            RGBColor(0xC0, 0x50, 0x4D),  # red-brown
            RGBColor(0x2E, 0x75, 0xB6),  # steel blue
            RGBColor(0x7F, 0x60, 0x4F),  # taupe/gray-brown
        ],
        "gridline_color": RGBColor(0xDD, 0xDD, 0xDD),
        "text_color": RGBColor(0x33, 0x33, 0x33),
    },
    "table": {
        "style": "Table Grid",
        "header_bg": True,
        "alt_rows": True,
        "col_min_width": Inches(0.5),
        "header_repeat": True,  # 表头跨页重复
    },
    "reference": {
        "format": "harvard",   # or "gbt7714" / apa
        "hanging_indent": Inches(0.5),
        "font_size": Pt(10.5),
    },
    "alignment": {
        "title": WD_ALIGN_PARAGRAPH.CENTER,
        "subtitle": WD_ALIGN_PARAGRAPH.CENTER,
        "meta": WD_ALIGN_PARAGRAPH.CENTER,
        "h1": WD_ALIGN_PARAGRAPH.LEFT,
        "h2": WD_ALIGN_PARAGRAPH.LEFT,
        "body": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "caption": WD_ALIGN_PARAGRAPH.CENTER,
        "abstract_label": WD_ALIGN_PARAGRAPH.LEFT,
        "toc": WD_ALIGN_PARAGRAPH.LEFT,
        "header": WD_ALIGN_PARAGRAPH.CENTER,
        "footer": WD_ALIGN_PARAGRAPH.CENTER,
    },
    "callout": {
        "variant": "info",
    },
}


def get_token(tokens: dict, *path):
    """按路径获取令牌值，如 get_token(tokens, "font", "size", "body") → Pt(12)。

    支持多级 fallback：如果路径中某级缺失，返回 None。
    """
    cur = tokens
    for key in path:
        if not isinstance(cur, dict) or key not in cur:
            return None
        cur = cur[key]
    return cur


def deep_merge(base: dict, override: dict) -> dict:
    """递归合并字典，override 的值覆盖 base。

    注意：不使用 deepcopy，因为 docx 的 RGBColor/Pt/Inches 等对象
    是不可变的长度/颜色值，浅复制即可安全共享。
    """
    result = {}
    for k, v in base.items():
        if k in override:
            ov = override[k]
            if isinstance(v, dict) and isinstance(ov, dict):
                result[k] = deep_merge(v, ov)
            else:
                result[k] = ov
        else:
            result[k] = v
    # override 中 base 没有的键也加入
    for k, v in override.items():
        if k not in base:
            result[k] = v
    return result
