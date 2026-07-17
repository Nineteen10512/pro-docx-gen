"""Shared color palette for PaperJSX PPT + DOCX v1.3.

Centralizes all theme hex values so that PPT and Word skills stay perfectly
color-aligned. PPT uses ``pptx.dml.color.RGBColor``; DOCX uses
``docx.shared.RGBColor``; this module provides the canonical hex constants and
converters for both.

@since v1.3.0 (ARCH-1)
"""
from __future__ import annotations

from typing import Dict, List


# ---------------------------------------------------------------------------
# Canonical hex palette — single source of truth for both skills
# ---------------------------------------------------------------------------
# Values intentionally kept as strings (no pptx/docx dependency) so that any
# consumer can cast to their own RGBColor type.

BASE_PALETTE: Dict[str, str] = {
    # Academic (deep navy base)
    "primary":        "#1F3864",
    "secondary":      "#2E75B6",
    "accent":         "#C0504D",
    "bg":             "#FFFFFF",
    "text":           "#333333",
    "text_light":     "#666666",
    "text_on_primary":"#FFFFFF",
    "title_bar_bg":   "#1F3864",
    "table_header_bg":"#1F3864",
    "table_header_text":"#FFFFFF",
    "table_alt_row":  "#F2F6FC",   # PPT variant
    "table_alt_row_docx": "#F2F2F2",
    "section_bg":     "#1F3864",
    "section_num_color": "#2E75B6",
    "kpi_bg":         "#F2F6FC",
    "quote_border":   "#2E75B6",
    "divider":        "#2E75B6",
    "chart_gridline": "#DDDDDD",
    # DOCX-specific extras
    "muted":          "#666666",
    "heading":        "#1F3864",
    "title":          "#1F3864",
    "table_border":   "#BFBFBF",
    "code_bg":        "#F5F5F5",
    "callout_info_bg":"#E7F3F8",
    "callout_info_border": "#2E75B6",
    "callout_warning_bg": "#FFF4E5",
    "callout_warning_border": "#ED7D31",
    "callout_success_bg": "#E8F5E9",
    "callout_success_border": "#43A047",
    "callout_danger_bg": "#FDECEA",
    "callout_danger_border": "#C0504D",
}

CHART_PALETTE_DEFAULT: List[str] = [
    "#1F3864",  # deep navy
    "#C0504D",  # red-brown accent
    "#2E75B6",  # steel blue
    "#7F604F",  # taupe / gray-brown
]


# ---------------------------------------------------------------------------
# Theme overrides (hex dicts) — same values used in PPT themes.py v1.2
# ---------------------------------------------------------------------------

THEME_OVERRIDES: Dict[str, Dict[str, str]] = {
    "academic": {},  # defaults
    "business": {
        "primary":        "#2C3E50",
        "secondary":      "#3498DB",
        "accent":         "#E67E22",
        "title_bar_bg":   "#2C3E50",
        "table_header_bg":"#2C3E50",
        "section_bg":     "#2C3E50",
        "section_num_color": "#3498DB",
        "quote_border":   "#3498DB",
        "divider":        "#3498DB",
        "kpi_bg":         "#ECF0F1",
        "table_alt_row":  "#ECF0F1",
        "chart_gridline": "#D5DBDB",
    },
    "teaching": {
        "primary":        "#2E7D32",
        "secondary":      "#66BB6A",
        "accent":         "#FFA000",
        "title_bar_bg":   "#2E7D32",
        "table_header_bg":"#2E7D32",
        "section_bg":     "#2E7D32",
        "section_num_color": "#66BB6A",
        "quote_border":   "#66BB6A",
        "divider":        "#66BB6A",
        "kpi_bg":         "#E8F5E9",
        "table_alt_row":  "#E8F5E9",
        "chart_gridline": "#C8E6C9",
    },
}


def hex_to_rgb(hex_str: str):
    """Parse ``#RRGGBB`` → ``(r,g,b)`` tuple (0–255 ints)."""
    h = hex_str.lstrip("#")
    if len(h) != 6:
        raise ValueError(f"Invalid hex color: {hex_str!r}")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def to_pptx_rgb(hex_str: str):
    """Return a ``pptx.dml.color.RGBColor`` instance."""
    from pptx.dml.color import RGBColor as _PptxRGB
    r, g, b = hex_to_rgb(hex_str)
    return _PptxRGB(r, g, b)


def to_docx_rgb(hex_str: str):
    """Return a ``docx.shared.RGBColor`` instance."""
    from docx.shared import RGBColor as _DocxRGB
    r, g, b = hex_to_rgb(hex_str)
    return _DocxRGB(r, g, b)


# ---------------------------------------------------------------------------
# v1.6.6: Global text/heading color enforcement (P0 — root-cause fix).
#
# 任何模板/theme_overrides 把 heading 写成 #D4AF37 / #E0E7FF / #64FFDA /
# 任何浅色（会消失白底上）或饱和度高的颜色，最终输出统一被强制覆盖为
#   HEADING_HEX (#1A1A1A 深黑) — 所有 heading 节点（Heading 1~6 + Title）
#   TEXT_HEX    (#333333)       — 所有正文 run 默认色
#   MUTED_HEX   (#666666)       — 副标题/caption/footer/header 文字
# 强调色（accent/primary/secondary）由各 theme 保留差异（这是表达力所在）。
#
# 该层是"硬约束"——引擎层任何处理 heading/text 的代码都必须从本层取色，
# 不允许再读 theme_overrides["color"]["heading"] 这种会被模板污染的字段。
# ---------------------------------------------------------------------------

HEADING_HEX = "1A1A1A"
TEXT_HEX = "333333"
MUTED_HEX = "666666"

# 用于 self_audit 校验：哪些 hex 不允许出现在 heading/text 节点中
# （典型违规：黑金 #D4AF37、深色科技 #E0E7FF、霓虹 #64FFDA、白/近白等）
FORBIDDEN_HEADING_HEXES = frozenset({
    "D4AF37",  # 黑金 (elegant_luxury)
    "E0E7FF",  # 浅蓝紫 (modern_tech)
    "64FFDA",  # 青蓝霓虹 (modern_tech)
    "E7E7E7",  # 近白
    "FFFFFF",  # 纯白（heading 在白底上不可见）
    "EEF6FF",  # 浅淡
    "F5F5F5",  # 浅灰
    "C9A227",  # 黑金 (premium 旧版)
    "B8860B",  # DarkGoldenRod
    "FFD700",  # 金色
})


def heading_rgb():
    """Return a ``docx.shared.RGBColor`` for headings — always #1A1A1A.

    **Do not read theme_overrides["color"]["heading"] in renderer code.**
    Always call this function instead, so the root-cause color discipline
    survives any future template/theme override changes.
    """
    from docx.shared import RGBColor as _DocxRGB
    r, g, b = hex_to_rgb("#" + HEADING_HEX)
    return _DocxRGB(r, g, b)


def text_rgb():
    """Return a ``docx.shared.RGBColor`` for body text — always #333333."""
    from docx.shared import RGBColor as _DocxRGB
    r, g, b = hex_to_rgb("#" + TEXT_HEX)
    return _DocxRGB(r, g, b)


def muted_rgb():
    """Return a ``docx.shared.RGBColor`` for muted text — always #666666."""
    from docx.shared import RGBColor as _DocxRGB
    r, g, b = hex_to_rgb("#" + MUTED_HEX)
    return _DocxRGB(r, g, b)


def normalize_text_color(theme_color_value):
    """Return a canonical text color (RGBColor) regardless of theme input.

    用于正文段落渲染：忽略 theme_overrides["color"]["text"] 提供的任意值，
    永远返回 #333333。这样模板/主题被注入花哨色时也不会泄漏到正文。
    """
    return text_rgb()


def is_forbidden_heading_hex(hex_or_rgb) -> bool:
    """Check whether the given color is on the forbidden-heading list.

    Args:
        hex_or_rgb: either ``#RRGGBB`` str (with/without #) or an RGBColor-like
            object with a ``__getitem__`` (r,g,b).
    """
    if hex_or_rgb is None:
        return False
    s = ""
    if isinstance(hex_or_rgb, str):
        s = hex_or_rgb.lstrip("#").upper()
    else:
        try:
            r, g, b = hex_or_rgb[0], hex_or_rgb[1], hex_or_rgb[2]
            s = f"{int(r):02X}{int(g):02X}{int(b):02X}"
        except Exception:
            return False
    return s in FORBIDDEN_HEADING_HEXES


def resolve_palette(theme_name: str = "academic") -> Dict[str, str]:
    """Return a fully-resolved hex palette (base + theme overrides).

    v1.6.6: heading / text / muted are normalized to the canonical
    #1A1A1A / #333333 / #666666 (any theme override for those keys is
    intentionally dropped to enforce readability on white backgrounds).
    """
    out = dict(BASE_PALETTE)
    overrides = THEME_OVERRIDES.get(theme_name, {})
    out.update(overrides)
    # Hard enforce canonical text colors (theme may keep accent/primary/secondary)
    out["heading"] = "#" + HEADING_HEX
    out["title"] = "#" + HEADING_HEX
    out["text"] = "#" + TEXT_HEX
    out["muted"] = "#" + MUTED_HEX
    return out


__all__ = [
    "BASE_PALETTE", "CHART_PALETTE_DEFAULT", "THEME_OVERRIDES",
    "hex_to_rgb", "to_pptx_rgb", "to_docx_rgb", "resolve_palette",
    # v1.6.6 — global text/heading color enforcement
    "HEADING_HEX", "TEXT_HEX", "MUTED_HEX",
    "FORBIDDEN_HEADING_HEXES",
    "heading_rgb", "text_rgb", "muted_rgb", "normalize_text_color",
    "is_forbidden_heading_hex",
]
