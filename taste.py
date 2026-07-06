"""Taste and craft preflight checks for PRO-DOCX."""

from __future__ import annotations

import copy
import os
from typing import Any

from shared.taste.adapters import docx_units
from shared.taste.core import build_preflight, finalize_report, infer_design_read as _infer_design_read
from shared.taste.rules import check_copy_and_density, check_layout_rhythm, issue

from .compiler.markdown_parser import markdown_to_document
from .tokens import BASE_TOKENS, get_theme, merge_theme, resolve_theme_name


_HARD_FAIL_CODES = {"contrast_iron_law", "chart_color_variety", "placeholder_text"}


def _resolved_theme_name(theme: Any) -> str:
    if isinstance(theme, str):
        return resolve_theme_name(theme)
    return "custom"


def _resolved_tokens(theme: Any) -> dict:
    if isinstance(theme, dict):
        return merge_theme(BASE_TOKENS, theme)
    return get_theme(theme)


def _color_hex(value: Any) -> str | None:
    try:
        from docx.shared import RGBColor as _RGBColor

        if isinstance(value, _RGBColor):
            return "#{:02X}{:02X}{:02X}".format(value[0], value[1], value[2])
    except Exception:
        pass
    if isinstance(value, str) and value.startswith("#") and len(value) == 7:
        return value.upper()
    if hasattr(value, "__iter__") and not isinstance(value, str):
        try:
            r, g, b = [int(v) for v in list(value)[:3]]
            return "#{:02X}{:02X}{:02X}".format(r, g, b)
        except Exception:
            return None
    return None


def _check_contrast_iron_law(theme: Any, issues: list[dict]) -> None:
    from shared.quality import hex_to_rgb, relative_luminance

    tokens = _resolved_tokens(theme)
    colors = tokens.get("color", {})
    page_bg = "#FFFFFF"
    pairs = [
        ("body text", colors.get("text"), page_bg),
        ("heading text", colors.get("heading") or colors.get("title"), page_bg),
    ]

    for label, fg, bg in pairs:
        fg_hex = _color_hex(fg)
        bg_hex = _color_hex(bg)
        if not fg_hex or not bg_hex:
            continue

        fg_lum = relative_luminance(*hex_to_rgb(fg_hex))
        bg_lum = relative_luminance(*hex_to_rgb(bg_hex))
        bg_is_dark = bg_lum <= 0.35
        bg_is_light = bg_lum >= 0.65
        fg_is_light = fg_lum >= 0.55
        fg_is_dark = fg_lum <= 0.45

        violated = (bg_is_dark and not fg_is_light) or (bg_is_light and not fg_is_dark)
        if violated:
            expected = "light text on dark background" if bg_is_dark else "dark text on light background"
            issue(
                issues,
                "error",
                "contrast_iron_law",
                f"{label} violates the contrast iron law: expected {expected}.",
                suggestion="Use light text on dark fills and dark text on light fills before rendering.",
            )


def _check_chart_color_variety(units: list[dict], theme: Any, issues: list[dict]) -> None:
    tokens = _resolved_tokens(theme)
    palette = (tokens.get("chart") or {}).get("palette") or []
    unique_colors = {hx for hx in (_color_hex(color) for color in palette) if hx}

    if not any(unit.get("kind") == "chart" for unit in units):
        return
    if len(unique_colors) < 2:
        issue(
            issues,
            "error",
            "chart_color_variety",
            "Chart palette collapses to a single color, which is disallowed.",
            suggestion="Provide at least two distinct chart colors before generating the document.",
        )


def _normalize_content(content: dict | str) -> dict:
    if isinstance(content, dict):
        return copy.deepcopy(content)
    if isinstance(content, str):
        if os.path.exists(content):
            raise TypeError("taste_check expects semantic JSON or Markdown text, not a file path")
        return markdown_to_document(content)
    raise TypeError("taste_check expects a semantic JSON dict or Markdown string")


def infer_design_read(content: dict | str, theme: str = "academic", lang: str = "cn") -> dict:
    working = _normalize_content(content)
    units = docx_units(working)
    return _infer_design_read(
        units,
        theme=_resolved_theme_name(theme),
        meta=(working.get("meta") or {}),
        lang=lang,
        medium="document",
    )


def taste_check(
    content: dict | str,
    theme: str = "academic",
    lang: str = "cn",
    strict: bool = False,
) -> dict:
    from . import docx_jsx

    working = _normalize_content(content)
    units = docx_units(working)
    issues: list[dict] = []

    design_read = _infer_design_read(
        units,
        theme=_resolved_theme_name(theme),
        meta=(working.get("meta") or {}),
        lang=lang,
        medium="document",
    )
    _check_contrast_iron_law(theme, issues)
    check_copy_and_density(units, issues)
    _check_chart_color_variety(units, theme, issues)
    check_layout_rhythm(units, issues, structural_kinds={"section", "document"})

    base_quality = None
    try:
        base_quality = docx_jsx.quality_check(working, theme=theme, lang=lang)
    except Exception as exc:
        issue(
            issues,
            "warning",
            "base_quality_unavailable",
            f"Base quality_check could not run: {exc}",
            suggestion="Fix schema or theme inputs, then rerun taste_check.",
        )

    preflight = build_preflight(
        issues,
        base_quality=base_quality,
        story=None,
        include_visual_intent=False,
    )
    return finalize_report(
        version="1.5",
        design_read=design_read,
        issues=issues,
        preflight=preflight,
        strict=strict,
        hard_fail_codes=_HARD_FAIL_CODES,
        base_quality=base_quality,
        story=None,
    )


__all__ = ["taste_check", "infer_design_read"]
