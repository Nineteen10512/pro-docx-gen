"""Themes — 主题配置。

通过覆盖 BASE_TOKENS 的子集实现预设主题。三套内置主题（v1.1 原版效果不变）：
- academic: 学术论文（Times New Roman + 宋体，双倍行距，Harvard 引用）
- business: 商务报告（Arial + 微软雅黑，1.15 倍行距，专业蓝色调）
- teaching: 教学教案（宋体 + 黑体，1.5 倍行距，清晰易读）

v1.2 新增 10+ 主题（与 PPT v1.2 配色一致，保证跨文档配色统一）：
- tech: 科技蓝（深色导航、霓虹蓝强调）
- dark: 暗黑（深色背景高对比，适合演示/夜间阅读，docx 不做深色页面，仅文字色）
- minimal: 极简白（无衬线、灰黑、无装饰）
- nature: 自然绿（大地色系）
- warm: 暖橙（暖色调，亲和力）
- premium: 高端黑金（黑底金字，仪式感）
- chinese_red: 中国红（朱红+金，正式庆典）
- ocean: 海洋蓝（青蓝渐变感，深蓝系）
- forest: 森林（深绿+木色，自然沉稳）
- sunset: 日落橙（橙+紫渐变，浪漫活力）

@since v1.4.0 颜色 HEX 值统一从 ``skills.shared.themes`` 读取（ARCH-3），
             字体、间距、对齐等排版 token 仍本地定义。
"""

from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from .design_tokens import BASE_TOKENS, deep_merge, PAGE_SIZES


def _rgb(hex_str: str) -> RGBColor:
    """'#RRGGBB' → RGBColor。"""
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _doc_colors(theme_name: str, chart_palette: list[str] | None = None, overrides: dict | None = None) -> dict:
    """从 shared.themes 读取颜色并转换为 docx RGBColor dict。

    保留 DOCX 特有的 key 映射（background 不直接覆盖页面色——docx 页面保持白色，
    颜色仅用于标题、强调、表格等元素）。
    """
    try:
        from ...shared.themes import get_palette, get_chart_palette
    except ImportError:  # pragma: no cover
        try:
            from skills.shared.themes import get_palette, get_chart_palette
        except ImportError:
            from shared.themes import get_palette, get_chart_palette

    pal = get_palette(theme_name)
    color: dict = {}
    # docx color key → shared palette key
    mapping = [
        ("heading", "primary"),
        ("title", "primary"),
        ("accent", "secondary"),
        ("muted", "muted"),
        ("text", "text"),
        ("table_header_bg", "table_header_bg"),
        ("table_header_text", "table_header_text"),
        ("table_alt_row", "table_alt_row"),
        ("quote_border", "secondary"),
        ("divider", "secondary"),
    ]
    for doc_key, shared_key in mapping:
        hv = pal.get(shared_key)
        if hv:
            color[doc_key] = _rgb(hv)

    if chart_palette is None:
        chart_palette = get_chart_palette(theme_name)

    result = {
        "color": color,
        "chart": {
            "palette": [_rgb(hx) for hx in chart_palette],
            "gridline_color": _rgb(pal.get("chart_gridline", "#DDDDDD")),
            "text_color": _rgb(pal.get("text", "#333333")),
        },
    }
    if overrides:
        # 允许局部覆盖（在本文件中手工合并 color/chart 子项）
        for k, v in overrides.items():
            if k in result and isinstance(result[k], dict) and isinstance(v, dict):
                result[k] = {**result[k], **v}
            else:
                result[k] = v
    return result


# ─── academic 主题（基于 BASE_TOKENS，少量覆盖） ──────────────────────

ACADEMIC_OVERRIDES = {
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
            "body": Pt(12),
        },
    },
    "spacing": {
        "line_spacing": 2.0,
        "first_line_indent": Inches(0.3),
    },
    "reference": {"format": "harvard"},
}
ACADEMIC_OVERRIDES.update(_doc_colors(
    "academic",
    chart_palette=["#1F3864", "#C0504D", "#2E75B6", "#7F604F", "#548235", "#7030A0"],
))


# ─── business 主题 ─────────────────────────────────────────────────

BUSINESS_OVERRIDES = {
    "font": {
        "family": {
            "default": "Arial",
            "cn": "微软雅黑",
            "heading": "Arial",
            "cn_heading": "微软雅黑",
            "code": "Consolas",
        },
        "size": {
            "title": Pt(24),
            "subtitle": Pt(14),
            "h1": Pt(18),
            "h2": Pt(14),
            "h3": Pt(12),
            "h4": Pt(11),
            "h5": Pt(11),
            "body": Pt(11),
            "abstract": Pt(10),
            "caption": Pt(10),
            "footnote": Pt(9),
            "reference": Pt(10),
        },
    },
    "spacing": {
        "page_margin": Inches(1),
        "line_spacing": 1.15,
        "first_line_indent": Inches(0),
        "before_h1": Pt(18),
        "after_h1": Pt(8),
        "before_h2": Pt(12),
        "after_h2": Pt(4),
        "before_h3": Pt(8),
        "after_h3": Pt(3),
        "paragraph_after": Pt(4),
    },
    "alignment": {
        "body": WD_ALIGN_PARAGRAPH.LEFT,
    },
    "reference": {"format": "apa"},
}
BUSINESS_OVERRIDES.update(_doc_colors(
    "business",
    chart_palette=["#2C3E50", "#E67E22", "#16A085", "#F1C40F", "#8E44AD", "#C0392B"],
    overrides={"color": {"accent": _rgb("#3498DB")}},
))


# ─── teaching 主题 ─────────────────────────────────────────────────

TEACHING_OVERRIDES = {
    "font": {
        "family": {
            "default": "宋体",
            "cn": "宋体",
            "heading": "黑体",
            "cn_heading": "黑体",
            "code": "Consolas",
        },
        "size": {
            "title": Pt(22),
            "subtitle": Pt(15),
            "h1": Pt(16),
            "h2": Pt(14),
            "h3": Pt(12),
            "h4": Pt(12),
            "h5": Pt(12),
            "body": Pt(12),
            "abstract": Pt(10.5),
            "caption": Pt(10.5),
        },
    },
    "spacing": {
        "line_spacing": 1.5,
        "first_line_indent": Inches(0.3),
        "before_h1": Pt(18),
        "after_h1": Pt(10),
        "before_h2": Pt(12),
        "after_h2": Pt(6),
        "paragraph_after": Pt(4),
    },
    "alignment": {
        "title": WD_ALIGN_PARAGRAPH.CENTER,
    },
    "reference": {"format": "gbt7714"},
}
TEACHING_OVERRIDES.update(_doc_colors(
    "teaching",
    chart_palette=["#2E7D32", "#FFA000", "#42A5F5", "#9CCC65", "#E91E63", "#7B1FA2"],
    overrides={
        "color": {
            "accent": _rgb("#FFA000"),
            "quote_border": _rgb("#FFA000"),
            "callout_info_border": _rgb("#FFA000"),
        },
    },
))


# ─── v1.2 新增主题 ────────────────────────────────────────────────

# 科技蓝
TECH_OVERRIDES = {
    "font": {
        "family": {
            "default": "Segoe UI",
            "cn": "微软雅黑",
            "heading": "Segoe UI",
            "cn_heading": "微软雅黑",
            "code": "Consolas",
        },
        "size": {"title": Pt(26), "subtitle": Pt(13), "h1": Pt(18), "h2": Pt(14), "h3": Pt(12), "body": Pt(11)},
    },
    "spacing": {
        "line_spacing": 1.25, "first_line_indent": Inches(0),
        "before_h1": Pt(18), "after_h1": Pt(8), "before_h2": Pt(12),
        "after_h2": Pt(4), "paragraph_after": Pt(4),
    },
    "alignment": {"body": WD_ALIGN_PARAGRAPH.LEFT},
    "reference": {"format": "apa"},
}
TECH_OVERRIDES.update(_doc_colors(
    "tech",
    chart_palette=["#00D4FF", "#64FFDA", "#7C3AED", "#F472B6", "#F59E0B", "#22D3EE"],
    overrides={"color": {"accent": _rgb("#00C8FF"), "quote_border": _rgb("#009FFF"), "table_alt_row": _rgb("#EEF6FF")}},
))

# 暗黑
DARK_OVERRIDES = {
    "font": {
        "family": {"default": "Segoe UI", "cn": "微软雅黑", "heading": "Segoe UI",
                   "cn_heading": "微软雅黑", "code": "Consolas"},
        "size": {"title": Pt(26), "h1": Pt(18), "h2": Pt(14), "body": Pt(11)},
    },
    "spacing": {"line_spacing": 1.25, "first_line_indent": Inches(0), "paragraph_after": Pt(4)},
    "alignment": {"body": WD_ALIGN_PARAGRAPH.LEFT},
}
DARK_OVERRIDES.update(_doc_colors(
    "dark",
    chart_palette=["#60A5FA", "#F59E0B", "#34D399", "#F472B6", "#A78BFA", "#22D3EE"],
    overrides={
        "color": {
            "heading": _rgb("#111111"), "title": _rgb("#111111"),
            "accent": _rgb("#E91E63"), "quote_border": _rgb("#E91E63"),
            "code_bg": _rgb("#F4F4F8"), "text": _rgb("#222222"),
        },
    },
))

# 极简白
MINIMAL_OVERRIDES = {
    "font": {
        "family": {"default": "Calibri", "cn": "思源黑体", "heading": "Calibri",
                   "cn_heading": "思源黑体", "code": "Consolas"},
        "size": {"title": Pt(22), "h1": Pt(16), "h2": Pt(13), "body": Pt(11)},
    },
    "spacing": {
        "line_spacing": 1.35, "first_line_indent": Inches(0),
        "before_h1": Pt(16), "after_h1": Pt(6), "paragraph_after": Pt(4),
    },
    "alignment": {"body": WD_ALIGN_PARAGRAPH.LEFT},
}
MINIMAL_OVERRIDES.update(_doc_colors(
    "minimal",
    chart_palette=["#2563EB", "#10B981", "#F59E0B", "#EF4444", "#8B5CF6", "#0EA5E9"],
    overrides={
        "color": {
            "heading": _rgb("#000000"), "title": _rgb("#000000"),
            "accent": _rgb("#555555"), "quote_border": _rgb("#999999"),
            "table_border": _rgb("#DDDDDD"),
            "callout_info_bg": _rgb("#F8F8F8"),
            "callout_info_border": _rgb("#BBBBBB"),
        },
    },
))

# 自然绿
NATURE_OVERRIDES = {
    "font": {
        "family": {"default": "Georgia", "cn": "楷体", "heading": "Georgia",
                   "cn_heading": "黑体", "code": "Consolas"},
        "size": {"title": Pt(22), "h1": Pt(16), "body": Pt(11)},
    },
    "spacing": {"line_spacing": 1.4, "first_line_indent": Inches(0.25), "paragraph_after": Pt(5)},
}
NATURE_OVERRIDES.update(_doc_colors(
    "nature",
    chart_palette=["#65A30D", "#A16207", "#CA8A04", "#16A34A", "#B91C1C", "#7C3AED"],
    overrides={
        "color": {
            "heading": _rgb("#2E533F"), "title": _rgb("#2E533F"),
            "accent": _rgb("#88B04B"), "quote_border": _rgb("#88B04B"),
            "table_alt_row": _rgb("#F0F5EC"),
        },
    },
))

# 暖橙
WARM_OVERRIDES = {
    "font": {
        "family": {"default": "Helvetica", "cn": "苹方", "heading": "Helvetica",
                   "cn_heading": "苹方-简中黑体", "code": "Consolas"},
        "size": {"title": Pt(24), "h1": Pt(17), "body": Pt(11)},
    },
    "spacing": {"line_spacing": 1.35, "first_line_indent": Inches(0), "paragraph_after": Pt(5)},
    "alignment": {"body": WD_ALIGN_PARAGRAPH.LEFT},
}
WARM_OVERRIDES.update(_doc_colors(
    "warm",
    chart_palette=["#B45309", "#D97706", "#DC2626", "#92400E", "#65A30D", "#7C3AED"],
    overrides={
        "color": {
            "heading": _rgb("#D85A1F"), "title": _rgb("#D85A1F"),
            "accent": _rgb("#F4A261"), "quote_border": _rgb("#F4A261"),
            "table_alt_row": _rgb("#FDF1E6"),
        },
    },
))

# 高端黑金
PREMIUM_OVERRIDES = {
    "font": {
        "family": {"default": "Times New Roman", "cn": "方正书宋简体",
                   "heading": "Times New Roman", "cn_heading": "方正小标宋简体",
                   "code": "Consolas"},
        "size": {"title": Pt(26), "h1": Pt(18), "h2": Pt(14), "body": Pt(11)},
    },
    "spacing": {"line_spacing": 1.5, "first_line_indent": Inches(0.25), "paragraph_after": Pt(6)},
    "alignment": {"title": WD_ALIGN_PARAGRAPH.CENTER},
}
PREMIUM_OVERRIDES.update(_doc_colors(
    "premium",
    chart_palette=["#D4AF37", "#1C1917", "#B08D57", "#78716C", "#DC2626", "#0EA5E9"],
    overrides={
        "color": {
            "heading": _rgb("#1A1A1A"), "title": _rgb("#B8860B"),
            "accent": _rgb("#C9A227"), "quote_border": _rgb("#C9A227"),
            "table_alt_row": _rgb("#FAF6EB"), "table_header_text": _rgb("#C9A227"),
        },
    },
))

# 中国红
CHINESE_RED_OVERRIDES = {
    "font": {
        "family": {"default": "宋体", "cn": "宋体", "heading": "黑体",
                   "cn_heading": "黑体", "code": "Consolas"},
        "size": {"title": Pt(26), "h1": Pt(18), "h2": Pt(14), "body": Pt(12)},
    },
    "spacing": {"line_spacing": 1.6, "first_line_indent": Inches(0.3), "paragraph_after": Pt(6)},
    "reference": {"format": "gbt7714"},
    "alignment": {"title": WD_ALIGN_PARAGRAPH.CENTER},
}
CHINESE_RED_OVERRIDES.update(_doc_colors(
    "chinese_red",
    chart_palette=["#C41E3A", "#D4AF37", "#8B0000", "#E67E22", "#16A34A", "#0F4C81"],
    overrides={
        "color": {
            "heading": _rgb("#C01D2F"), "title": _rgb("#C01D2F"),
            "accent": _rgb("#D4A017"), "quote_border": _rgb("#D4A017"),
            "table_alt_row": _rgb("#FDEEEE"),
        },
    },
))

# 海洋蓝
OCEAN_OVERRIDES = {
    "font": {
        "family": {"default": "Calibri", "cn": "微软雅黑", "heading": "Calibri",
                   "cn_heading": "微软雅黑", "code": "Consolas"},
        "size": {"title": Pt(24), "h1": Pt(17), "body": Pt(11)},
    },
    "spacing": {"line_spacing": 1.4, "first_line_indent": Inches(0), "paragraph_after": Pt(5)},
    "alignment": {"body": WD_ALIGN_PARAGRAPH.LEFT},
}
OCEAN_OVERRIDES.update(_doc_colors(
    "ocean",
    chart_palette=["#0284C7", "#06B6D4", "#0EA5E9", "#0891B2", "#0369A1", "#0E7490"],
    overrides={
        "color": {
            "heading": _rgb("#054A70"), "title": _rgb("#054A70"),
            "accent": _rgb("#189AB4"), "quote_border": _rgb("#189AB4"),
            "table_alt_row": _rgb("#E3F2F7"),
        },
    },
))

# 森林
FOREST_OVERRIDES = {
    "font": {
        "family": {"default": "Garamond", "cn": "楷体", "heading": "Garamond",
                   "cn_heading": "华文行楷", "code": "Consolas"},
        "size": {"title": Pt(24), "h1": Pt(17), "body": Pt(11.5)},
    },
    "spacing": {"line_spacing": 1.5, "first_line_indent": Inches(0.25), "paragraph_after": Pt(5)},
}
FOREST_OVERRIDES.update(_doc_colors(
    "forest",
    chart_palette=["#059669", "#D97706", "#10B981", "#84CC16", "#DC2626", "#7C3AED"],
    overrides={
        "color": {
            "heading": _rgb("#2D501E"), "title": _rgb("#2D501E"),
            "accent": _rgb("#689F50"), "quote_border": _rgb("#8B5A2B"),
            "table_alt_row": _rgb("#EFF5E8"),
        },
    },
))

# 日落橙
SUNSET_OVERRIDES = {
    "font": {
        "family": {"default": "Helvetica", "cn": "苹方", "heading": "Helvetica",
                   "cn_heading": "苹方-简中黑体", "code": "Consolas"},
        "size": {"title": Pt(26), "h1": Pt(18), "body": Pt(11)},
    },
    "spacing": {"line_spacing": 1.4, "first_line_indent": Inches(0), "paragraph_after": Pt(5)},
    "alignment": {"body": WD_ALIGN_PARAGRAPH.LEFT},
}
SUNSET_OVERRIDES.update(_doc_colors(
    "sunset",
    chart_palette=["#EA580C", "#DB2777", "#F59E0B", "#A855F7", "#0EA5E9", "#16A34A"],
    overrides={
        "color": {
            "heading": _rgb("#C53D1A"), "title": _rgb("#E45F19"),
            "accent": _rgb("#EE964D"), "quote_border": _rgb("#9B3BA1"),
            "table_alt_row": _rgb("#FDEEE1"),
            "callout_warning_bg": _rgb("#FCE4D0"),
        },
    },
))


THEME_OVERRIDES = {
    "academic": ACADEMIC_OVERRIDES,
    "business": BUSINESS_OVERRIDES,
    "teaching": TEACHING_OVERRIDES,
    "tech": TECH_OVERRIDES,
    "dark": DARK_OVERRIDES,
    "minimal": MINIMAL_OVERRIDES,
    "nature": NATURE_OVERRIDES,
    "warm": WARM_OVERRIDES,
    "premium": PREMIUM_OVERRIDES,
    "chinese_red": CHINESE_RED_OVERRIDES,
    "ocean": OCEAN_OVERRIDES,
    "forest": FOREST_OVERRIDES,
    "sunset": SUNSET_OVERRIDES,
}

# high_contrast 主题补一个（与 PPT 对齐）
HIGH_CONTRAST_OVERRIDES = {
    "font": {
        "family": {"default": "Arial", "cn": "黑体", "heading": "Arial",
                   "cn_heading": "黑体", "code": "Consolas"},
        "size": {"title": Pt(24), "h1": Pt(18), "h2": Pt(14), "body": Pt(12)},
    },
    "spacing": {"line_spacing": 1.5, "first_line_indent": Inches(0), "paragraph_after": Pt(4)},
}
HIGH_CONTRAST_OVERRIDES.update(_doc_colors(
    "high_contrast",
    chart_palette=["#000000", "#0047AB", "#B22222", "#006400", "#8B4513", "#4B0082"],
    overrides={
        "color": {
            "heading": _rgb("#000000"), "title": _rgb("#000000"),
            "accent": _rgb("#000000"), "quote_border": _rgb("#000000"),
            "table_alt_row": _rgb("#F0F0F0"), "table_header_text": _rgb("#FFFFFF"),
            "muted": _rgb("#333333"), "text": _rgb("#000000"),
        },
    },
))
THEME_OVERRIDES["high_contrast"] = HIGH_CONTRAST_OVERRIDES

# 自然语言别名 → 标准主题名
THEME_ALIASES = {
    "科技": "tech", "科技蓝": "tech", "tech": "tech", "technology": "tech",
    "暗黑": "dark", "dark": "dark", "深色": "dark", "黑": "dark",
    "极简": "minimal", "minimal": "minimal", "白": "minimal", "简约": "minimal", "light": "minimal",
    "自然": "nature", "nature": "nature", "绿": "nature", "大地": "nature",
    "暖": "warm", "warm": "warm", "橙": "warm", "暖橙": "warm",
    "高端": "premium", "premium": "premium", "黑金": "premium", "商务金": "premium",
    "中国红": "chinese_red", "chinese_red": "chinese_red", "红": "chinese_red", "国风": "chinese_red",
    "海洋": "ocean", "ocean": "ocean", "海蓝": "ocean", "蓝": "ocean",
    "森林": "forest", "forest": "forest", "森系": "forest", "深绿": "forest",
    "日落": "sunset", "sunset": "sunset", "活力橙": "sunset",
    "学术": "academic", "论文": "academic", "academic": "academic",
    "商务": "business", "business": "business", "企业": "business",
    "教学": "teaching", "教案": "teaching", "teaching": "teaching",
    "高对比": "high_contrast", "high_contrast": "high_contrast", "黑白": "high_contrast",
}


def resolve_theme_name(name: str) -> str:
    """将自然语言主题名解析为标准主题 key。"""
    if not isinstance(name, str):
        return "academic"
    key = name.strip().lower().replace("-", "_").replace(" ", "_")
    if key in THEME_OVERRIDES:
        return key
    return THEME_ALIASES.get(name.strip(), name)


def get_theme(name: str = "academic", custom_overrides: dict | None = None) -> dict:
    """按主题名返回合并后的完整 tokens 字典。"""
    resolved = resolve_theme_name(name) if isinstance(name, str) else name
    if isinstance(resolved, str) and resolved not in THEME_OVERRIDES:
        raise ValueError(
            f"Unknown theme '{name}'. Available: {sorted(THEME_OVERRIDES.keys())}"
        )
    if isinstance(resolved, dict):
        tokens = deep_merge(BASE_TOKENS, resolved)
    else:
        tokens = deep_merge(BASE_TOKENS, THEME_OVERRIDES[resolved])

    page = tokens.get("page", {})
    page_size = page.get("size", "A4")
    orientation = page.get("orientation", "portrait")
    pw, ph = PAGE_SIZES.get(page_size, PAGE_SIZES["A4"])
    if orientation == "landscape":
        pw, ph = ph, pw
    tokens["spacing"]["page_width"] = pw
    tokens["spacing"]["page_height"] = ph
    sp = tokens["spacing"]
    default_margin = page.get("margin_top", sp["page_margin"])
    tokens["page"]["margin_top"] = page.get("margin_top", default_margin)
    tokens["page"]["margin_bottom"] = page.get("margin_bottom", default_margin)
    tokens["page"]["margin_left"] = page.get("margin_left", default_margin)
    tokens["page"]["margin_right"] = page.get("margin_right", default_margin)
    sp["page_margin"] = tokens["page"]["margin_left"]

    if custom_overrides:
        tokens = deep_merge(tokens, custom_overrides)
    return tokens


def merge_theme(base_tokens: dict, overrides: dict) -> dict:
    """在现有 tokens 基础上进一步覆盖，用于用户自定义主题。"""
    return deep_merge(base_tokens, overrides)


def list_themes() -> list[dict]:
    """返回所有可用主题的简要信息。"""
    names = sorted(THEME_OVERRIDES.keys())
    result = []
    for n in names:
        toks = deep_merge(BASE_TOKENS, THEME_OVERRIDES[n])
        result.append({
            "name": n,
            "title_color": "#{:02X}{:02X}{:02X}".format(*toks["color"]["title"]),
            "accent_color": "#{:02X}{:02X}{:02X}".format(*toks["color"]["accent"]),
            "font_default": toks["font"]["family"]["default"],
            "line_spacing": toks["spacing"]["line_spacing"],
        })
    return result
