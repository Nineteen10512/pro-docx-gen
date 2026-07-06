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


def resolve_palette(theme_name: str = "academic") -> Dict[str, str]:
    """Return a fully-resolved hex palette (base + theme overrides)."""
    out = dict(BASE_PALETTE)
    overrides = THEME_OVERRIDES.get(theme_name, {})
    out.update(overrides)
    return out


__all__ = [
    "BASE_PALETTE", "CHART_PALETTE_DEFAULT", "THEME_OVERRIDES",
    "hex_to_rgb", "to_pptx_rgb", "to_docx_rgb", "resolve_palette",
]
