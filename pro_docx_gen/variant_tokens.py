"""v1.6.0 Variant → DOCX Token Overrides.

将 VariantProfile 的抽象视觉特征（font_style / spacing / decoration / colors）
映射为 DOCX 设计令牌的具体数值覆盖。

Usage:
    from .variant_tokens import apply_variant, get_variant_tokens
    tokens = apply_variant(base_tokens, "modern_tech")
    # 或直接获取覆盖 dict
    overrides = get_variant_tokens("bold_impact")

@since v1.6.0
"""
from docx.shared import Pt, Inches, RGBColor

# ─── 5 变体 → 字体族映射 ────────────────────────────────────

_VARIANT_FONT_FAMILY = {
    "corporate_formal": {
        "default": "Calibri",
        "cn": "微软雅黑",
        "heading": "Calibri",
        "cn_heading": "微软雅黑",
        "code": "Consolas",
    },
    "minimal_clean": {
        "default": "Helvetica Neue",
        "cn": "微软雅黑",
        "heading": "Helvetica Neue Light",
        "cn_heading": "微软雅黑 Light",
        "code": "SF Mono",
    },
    "modern_tech": {
        "default": "Inter",
        "cn": "微软雅黑",
        "heading": "Inter SemiBold",
        "cn_heading": "微软雅黑",
        "code": "JetBrains Mono",
    },
    "bold_impact": {
        "default": "Arial",
        "cn": "微软雅黑",
        "heading": "Arial Black",
        "cn_heading": "微软雅黑",
        "code": "Consolas",
    },
    "elegant_luxury": {
        "default": "Georgia",
        "cn": "宋体",
        "heading": "Playfair Display",
        "cn_heading": "楷体",
        "code": "Consolas",
    },
}

# ─── 5 变体 → 字号覆盖 ──────────────────────────────────────

_VARIANT_FONT_SIZE = {
    "corporate_formal": {
        "title": Pt(22), "subtitle": Pt(14), "h1": Pt(16),
        "h2": Pt(14), "h3": Pt(13), "body": Pt(12),
    },
    "minimal_clean": {
        "title": Pt(28), "subtitle": Pt(14), "h1": Pt(18),
        "h2": Pt(14), "h3": Pt(12), "body": Pt(11),
    },
    "modern_tech": {
        "title": Pt(26), "subtitle": Pt(14), "h1": Pt(18),
        "h2": Pt(14), "h3": Pt(13), "body": Pt(11),
    },
    "bold_impact": {
        "title": Pt(32), "subtitle": Pt(16), "h1": Pt(22),
        "h2": Pt(16), "h3": Pt(14), "body": Pt(12),
    },
    "elegant_luxury": {
        "title": Pt(24), "subtitle": Pt(14), "h1": Pt(16),
        "h2": Pt(14), "h3": Pt(13), "body": Pt(11),
    },
}

# ─── 5 变体 → 颜色覆盖 ──────────────────────────────────────

_VARIANT_COLOR = {
    "corporate_formal": {
        "title": RGBColor(0x1F, 0x38, 0x64),
        "heading": RGBColor(0x1F, 0x38, 0x64),
        "accent": RGBColor(0x2E, 0x75, 0xB6),
        "table_header_bg": RGBColor(0x1F, 0x38, 0x64),
        "table_header_text": RGBColor(0xFF, 0xFF, 0xFF),
        "quote_border": RGBColor(0x2E, 0x75, 0xB6),
    },
    "minimal_clean": {
        "title": RGBColor(0x1A, 0x1A, 0x1A),
        "heading": RGBColor(0x33, 0x33, 0x33),
        "accent": RGBColor(0x5B, 0x9B, 0xD5),
        "table_header_bg": RGBColor(0xF5, 0xF5, 0xF5),
        "table_header_text": RGBColor(0x33, 0x33, 0x33),
        "quote_border": RGBColor(0xCC, 0xCC, 0xCC),
    },
    "modern_tech": {
        "title": RGBColor(0x64, 0xFF, 0xDA),
        "heading": RGBColor(0xE0, 0xE7, 0xFF),
        "accent": RGBColor(0x64, 0xFF, 0xDA),
        "table_header_bg": RGBColor(0x1E, 0x2A, 0x4A),
        "table_header_text": RGBColor(0xE0, 0xE7, 0xFF),
        "quote_border": RGBColor(0x64, 0xFF, 0xDA),
        "text": RGBColor(0xCC, 0xD6, 0xF6),
        "muted": RGBColor(0x88, 0x96, 0xB0),
    },
    "bold_impact": {
        "title": RGBColor(0x00, 0x00, 0x00),
        "heading": RGBColor(0x1A, 0x1A, 0x1A),
        "accent": RGBColor(0xE5, 0x39, 0x35),
        "table_header_bg": RGBColor(0x1A, 0x1A, 0x1A),
        "table_header_text": RGBColor(0xFF, 0xFF, 0xFF),
        "quote_border": RGBColor(0xE5, 0x39, 0x35),
    },
    "elegant_luxury": {
        "title": RGBColor(0x1C, 0x19, 0x17),
        "heading": RGBColor(0x2D, 0x2A, 0x26),
        "accent": RGBColor(0xD4, 0xAF, 0x37),
        "table_header_bg": RGBColor(0x1C, 0x19, 0x17),
        "table_header_text": RGBColor(0xD4, 0xAF, 0x37),
        "quote_border": RGBColor(0xD4, 0xAF, 0x37),
        "muted": RGBColor(0x8A, 0x85, 0x7D),
    },
}

# ─── 5 变体 → 间距覆盖 ──────────────────────────────────────

_VARIANT_SPACING = {
    "corporate_formal": {
        "line_spacing": 1.5,
        "before_title": Pt(0),
        "after_title": Pt(12),
        "before_h1": Pt(24),
        "after_h1": Pt(12),
        "paragraph_before": Pt(0),
        "paragraph_after": Pt(6),
    },
    "minimal_clean": {
        "line_spacing": 1.6,
        "before_title": Pt(0),
        "after_title": Pt(24),
        "before_h1": Pt(36),
        "after_h1": Pt(18),
        "paragraph_before": Pt(0),
        "paragraph_after": Pt(10),
    },
    "modern_tech": {
        "line_spacing": 1.5,
        "before_title": Pt(0),
        "after_title": Pt(16),
        "before_h1": Pt(28),
        "after_h1": Pt(14),
        "paragraph_before": Pt(0),
        "paragraph_after": Pt(8),
    },
    "bold_impact": {
        "line_spacing": 1.3,
        "before_title": Pt(0),
        "after_title": Pt(8),
        "before_h1": Pt(20),
        "after_h1": Pt(8),
        "paragraph_before": Pt(0),
        "paragraph_after": Pt(4),
    },
    "elegant_luxury": {
        "line_spacing": 1.8,
        "before_title": Pt(0),
        "after_title": Pt(30),
        "before_h1": Pt(30),
        "after_h1": Pt(16),
        "paragraph_before": Pt(0),
        "paragraph_after": Pt(8),
    },
}

# ─── 5 变体 → 页面设置覆盖 ──────────────────────────────────

_VARIANT_PAGE = {
    "corporate_formal": {
        "margin_top": Inches(1.0),
        "margin_bottom": Inches(1.0),
        "margin_left": Inches(1.0),
        "margin_right": Inches(1.0),
    },
    "minimal_clean": {
        "margin_top": Inches(1.2),
        "margin_bottom": Inches(1.2),
        "margin_left": Inches(1.2),
        "margin_right": Inches(1.2),
    },
    "modern_tech": {
        "margin_top": Inches(0.8),
        "margin_bottom": Inches(0.8),
        "margin_left": Inches(1.0),
        "margin_right": Inches(1.0),
    },
    "bold_impact": {
        "margin_top": Inches(0.6),
        "margin_bottom": Inches(0.6),
        "margin_left": Inches(0.8),
        "margin_right": Inches(0.8),
    },
    "elegant_luxury": {
        "margin_top": Inches(1.2),
        "margin_bottom": Inches(1.2),
        "margin_left": Inches(1.3),
        "margin_right": Inches(1.3),
    },
}


def get_variant_tokens(variant_name: str) -> dict:
    """返回指定 variant 的完整 token 覆盖 dict。

    未识别的 variant 返回空 dict（使用 base tokens 默认值）。
    """
    from .v160_style_registry import VARIANT_PROFILES as _vp
    if variant_name not in _vp:
        return {}

    overrides = {
        "font": {
            "family": _VARIANT_FONT_FAMILY.get(variant_name, {}),
            "size": _VARIANT_FONT_SIZE.get(variant_name, {}),
        },
        "color": _VARIANT_COLOR.get(variant_name, {}),
        "spacing": _VARIANT_SPACING.get(variant_name, {}),
        "page": _VARIANT_PAGE.get(variant_name, {}),
    }
    return overrides


# 延迟导入以避免循环引用
def _get_variant_profiles():
    from .v160_style_registry import VARIANT_PROFILES
    return VARIANT_PROFILES

VARIANT_PROFILES = None  # 占位，模块加载后由 caller 设置


def apply_variant(base_tokens: dict, variant_name: str) -> dict:
    """将 variant token 覆盖合并到 base_tokens。

    Args:
        base_tokens: 基础设计令牌（如 BASE_TOKENS）
        variant_name: 变体名（corporate_formal/minimal_clean/...）

    Returns:
        合并后的 tokens dict（新 dict，不修改 base_tokens）
    """
    from .tokens.design_tokens import deep_merge
    overrides = get_variant_tokens(variant_name)
    if not overrides:
        return base_tokens
    return deep_merge(base_tokens, overrides)


def get_variant_cover_style(variant_name: str) -> dict:
    """返回变体专属的封面页样式覆盖。

    用于 _render_title_block 根据 variant 调整标题字体、间距、装饰线等。
    """
    styles = {
        "corporate_formal": {
            "title_bold": True,
            "title_size": Pt(22),
            "title_color": RGBColor(0x1F, 0x38, 0x64),
            "subtitle_color": RGBColor(0x66, 0x66, 0x66),
            "divider": True,
            "divider_color": RGBColor(0x1F, 0x38, 0x64),
            "divider_width_pct": 0.3,
            "bg_color": None,
        },
        "minimal_clean": {
            "title_bold": False,
            "title_size": Pt(28),
            "title_color": RGBColor(0x1A, 0x1A, 0x1A),
            "subtitle_color": RGBColor(0x99, 0x99, 0x99),
            "divider": False,
            "bg_color": None,
        },
        "modern_tech": {
            "title_bold": True,
            "title_size": Pt(26),
            "title_color": RGBColor(0x64, 0xFF, 0xDA),
            "subtitle_color": RGBColor(0x88, 0x96, 0xB0),
            "divider": True,
            "divider_color": RGBColor(0x64, 0xFF, 0xDA),
            "divider_width_pct": 0.2,
            "bg_color": RGBColor(0x0A, 0x19, 0x2F),
        },
        "bold_impact": {
            "title_bold": True,
            "title_size": Pt(32),
            "title_color": RGBColor(0x00, 0x00, 0x00),
            "subtitle_color": RGBColor(0x66, 0x66, 0x66),
            "divider": True,
            "divider_color": RGBColor(0xE5, 0x39, 0x35),
            "divider_width_pct": 0.15,
            "bg_color": None,
        },
        "elegant_luxury": {
            "title_bold": False,
            "title_size": Pt(24),
            "title_color": RGBColor(0x1C, 0x19, 0x17),
            "subtitle_color": RGBColor(0x8A, 0x85, 0x7D),
            "divider": True,
            "divider_color": RGBColor(0xD4, 0xAF, 0x37),
            "divider_width_pct": 0.4,
            "bg_color": None,
        },
    }
    return styles.get(variant_name, styles["corporate_formal"])