from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
from typing import Mapping

from PIL import Image, ImageChops


MIN_EFFECTIVE_TEXT_PT = 9.0
MAX_EDGE_WHITESPACE_RATIO = 0.18
MIN_DISPLAY_PPI = 150.0
# v1.6.6: chart-insertion page-fit contract.
# When a chart's natural height > AVAILABLE_HEIGHT_FILL_RATIO * available
# page-area height, the chart_renderer will proportionally shrink the
# figure so the chart + caption + a one-line caption gap fit on one page
# without producing a stray near-blank page.
AVAILABLE_HEIGHT_FILL_RATIO = 0.90


class FigureAssetGateError(RuntimeError):
    def __init__(self, report: dict):
        self.report = report
        issue_codes = ", ".join(issue["code"] for issue in report["issues"])
        steps = "\n".join(f"{index}. {step}" for index, step in enumerate(report["remediation_steps"], 1))
        super().__init__(
            f"Figure asset rejected before DOCX insertion: {issue_codes}.\n"
            f"Remediation required, then rerender and run the gate again:\n{steps}"
        )


def _foreground_bbox(image: Image.Image, threshold: int = 18) -> tuple[int, int, int, int]:
    if image.mode == "RGBA":
        alpha = image.getchannel("A")
        minimum_alpha, _ = alpha.getextrema()
        if minimum_alpha < 250:
            alpha_bbox = alpha.point(lambda value: 255 if value > 8 else 0).getbbox()
            if alpha_bbox is not None:
                return alpha_bbox
    rgb = image.convert("RGB")
    corners = [
        rgb.getpixel((0, 0)),
        rgb.getpixel((rgb.width - 1, 0)),
        rgb.getpixel((0, rgb.height - 1)),
        rgb.getpixel((rgb.width - 1, rgb.height - 1)),
    ]
    background = tuple(sorted(values)[len(values) // 2] for values in zip(*corners))
    difference = ImageChops.difference(rgb, Image.new("RGB", rgb.size, background)).convert("L")
    bbox = difference.point(lambda value: 255 if value >= threshold else 0).getbbox()
    return bbox or (0, 0, rgb.width, rgb.height)


def audit_figure_asset(
    path: str | Path,
    *,
    display_width_inches: float,
    source_width_inches: float | None = None,
    declared_min_font_pt: float | None = None,
    contains_text: bool = False,
    minimum_effective_text_pt: float = MIN_EFFECTIVE_TEXT_PT,
    maximum_edge_whitespace_ratio: float = MAX_EDGE_WHITESPACE_RATIO,
    minimum_display_ppi: float = MIN_DISPLAY_PPI,
) -> dict:
    asset = Path(path)
    if display_width_inches <= 0:
        raise ValueError("display_width_inches must be positive")
    with Image.open(asset) as opened:
        image = opened.convert("RGBA" if opened.mode == "RGBA" else "RGB")
        width, height = image.size
        left, top, right, bottom = _foreground_bbox(image)

    edge_ratios = {
        "left": left / width,
        "right": (width - right) / width,
        "top": top / height,
        "bottom": (height - bottom) / height,
    }
    maximum_edge = max(edge_ratios.values())
    content_coverage = ((right - left) * (bottom - top)) / (width * height)
    display_ppi = width / display_width_inches
    issues: list[dict] = []
    remediation: list[str] = []

    if maximum_edge > maximum_edge_whitespace_ratio:
        issues.append(
            {
                "code": "excessive_canvas_whitespace",
                "severity": "error",
                "measured_max_edge_ratio": round(maximum_edge, 4),
                "allowed_max_edge_ratio": maximum_edge_whitespace_ratio,
                "foreground_bbox_pixels": [left, top, right, bottom],
            }
        )
        remediation.append(
            "Crop the canvas to foreground bbox "
            f"[{left}, {top}, {right}, {bottom}] plus 3–5% visual padding; preserve the final aspect ratio and rerender with tight bounds."
        )

    if display_ppi < minimum_display_ppi:
        required_width_px = math.ceil(display_width_inches * minimum_display_ppi)
        issues.append(
            {
                "code": "insufficient_display_resolution",
                "severity": "error",
                "display_ppi": round(display_ppi, 1),
                "minimum_display_ppi": minimum_display_ppi,
                "required_pixel_width": required_width_px,
            }
        )
        remediation.append(
            f"Rerender at least {required_width_px}px wide for the planned {display_width_inches:.2f}in insertion (200–300 PPI preferred)."
        )

    effective_font_pt = None
    if contains_text:
        if declared_min_font_pt is None:
            issues.append(
                {
                    "code": "missing_text_size_metadata",
                    "severity": "error",
                    "required_field": "declared_min_font_pt",
                }
            )
            remediation.append(
                "Declare the smallest text size and source canvas width, then rerender; text-bearing figures cannot bypass font-size verification."
            )
        else:
            source_width = source_width_inches or display_width_inches
            scale = display_width_inches / source_width
            effective_font_pt = float(declared_min_font_pt) * scale
            if effective_font_pt < minimum_effective_text_pt:
                required_source_font = minimum_effective_text_pt / scale
                required_display_width = minimum_effective_text_pt * source_width / float(declared_min_font_pt)
                issues.append(
                    {
                        "code": "effective_text_too_small",
                        "severity": "error",
                        "effective_min_font_pt": round(effective_font_pt, 2),
                        "minimum_effective_text_pt": minimum_effective_text_pt,
                        "required_source_font_pt": round(required_source_font, 2),
                        "required_display_width_inches": round(required_display_width, 2),
                    }
                )
                orientation_hint = (
                    " Use a landscape section or simplify/split the figure if that width exceeds the page."
                    if required_display_width > 7.0
                    else ""
                )
                remediation.append(
                    f"Raise the smallest source font to at least {required_source_font:.1f}pt for the current scale, "
                    f"or enlarge insertion width to at least {required_display_width:.2f}in.{orientation_hint}"
                )

    return {
        "asset": str(asset.resolve()),
        "passed": not issues,
        "display_width_inches": round(display_width_inches, 3),
        "source_width_inches": round(source_width_inches or display_width_inches, 3),
        "pixel_size": [width, height],
        "display_ppi": round(display_ppi, 1),
        "foreground_bbox_pixels": [left, top, right, bottom],
        "edge_whitespace_ratio": {key: round(value, 4) for key, value in edge_ratios.items()},
        "content_coverage_ratio": round(content_coverage, 4),
        "effective_min_font_pt": round(effective_font_pt, 2) if effective_font_pt is not None else None,
        "issues": issues,
        "remediation_steps": remediation,
        "next_action": "insert" if not issues else "modify, rerender, and rerun preflight",
    }


def assert_figure_asset_ready(path: str | Path, **kwargs) -> dict:
    report = audit_figure_asset(path, **kwargs)
    if not report["passed"]:
        raise FigureAssetGateError(report)
    return report


# ---------------------------------------------------------------------------
# v1.6.6: Available height computation for chart-insertion page fit.
# ---------------------------------------------------------------------------
#
# A chart that is taller than the remaining page area causes Word to push
# the entire figure to the next page and leaves a near-blank page behind.
# The renderer layer must therefore know, *before* it commits to a figure
# size, how much vertical space is still available on the current page.
#
# Token shape (matches ``pro_docx_gen.engine.layout.LayoutCalculator``):
#   tokens["spacing"]["page_height"]    -> page height in EMU
#   tokens["spacing"]["page_margin"]    -> uniform margin (EMU); 0 if absent
#   tokens["page"]["margin_top"] /
#   tokens["page"]["margin_bottom"]     -> per-side margins (EMU); fall back
#                                           to ``page_margin`` if missing
#   tokens["page"]["gutter"]            -> optional gutter (EMU); 0 if absent
#
# 1 inch == 914400 EMU. We accept either EMU ints or already-converted
# floats (callers in the renderer pass ``.content_width`` which is EMU).

_EMU_PER_INCH = 914400.0


def _emu_to_inches(value: float) -> float:
    """Best-effort EMU → inches conversion. Returns float inches.

    If the value already looks like inches (≤ 30, which exceeds the tallest
    legal A3 page), we return it untouched. This lets callers pass either
    raw tokens (EMU ints) or layout-derived floats (already inches) without
    double-conversion.
    """
    try:
        f = float(value)
    except (TypeError, ValueError):
        return 0.0
    if f <= 0:
        return 0.0
    # Heuristic: anything > 30 inches cannot be inches; treat as EMU.
    if f > 30.0:
        return f / _EMU_PER_INCH
    return f


def compute_available_height_inches(
    tokens: Mapping,
    *,
    consumed_inches: float = 0.0,
) -> float:
    """Return the height (inches) available for the next figure on the page.

    The computation is the textbook page-area formula:

        page_height_in - margin_top_in - margin_bottom_in - consumed_inches

    Args:
        tokens: token dict matching the v1.6.6 layout contract. Pulled from
            ``self.tokens`` in the renderer. Must contain ``spacing`` and
            may contain ``page`` for per-side margins.
        consumed_inches: vertical space (inches) already used on the current
            page (e.g. header height + paragraphs above the insertion point).
            Defaults to 0 for callers that have not yet measured it.

    Returns:
        Available height in inches. Always ≥ 0.5 in (we never promise
        sub-half-inch insertions, which would always cause overflow).

    Notes:
        The function is intentionally defensive: missing keys fall back to
        the safe default of 6.0 in (≈ one A4 content row). The result is
        safe for ``chart_renderer`` to divide by and to use as the upper
        bound for the ``figure_height`` check.
    """
    if not isinstance(tokens, Mapping):
        return 6.0
    spacing = tokens.get("spacing") or {}
    page = tokens.get("page") or {}

    page_height_raw = spacing.get("page_height", 0)
    margin_uniform_raw = spacing.get("page_margin", 0)
    margin_top_raw = page.get("margin_top", margin_uniform_raw)
    margin_bottom_raw = page.get("margin_bottom", margin_uniform_raw)
    gutter_raw = page.get("gutter", 0)

    page_h_in = _emu_to_inches(page_height_raw)
    margin_top_in = _emu_to_inches(margin_top_raw)
    margin_bottom_in = _emu_to_inches(margin_bottom_raw)
    gutter_in = _emu_to_inches(gutter_raw)

    # If page_height is missing or zero, fall back to A4 content height
    if page_h_in <= 0:
        page_h_in = 11.69  # A4 height in inches
    if margin_top_in <= 0:
        margin_top_in = 1.0
    if margin_bottom_in <= 0:
        margin_bottom_in = 1.0

    available = page_h_in - margin_top_in - margin_bottom_in - gutter_in - float(consumed_inches or 0.0)
    if available < 0.5:
        return 0.5
    return round(available, 3)


__all__ = [
    "FigureAssetGateError",
    "audit_figure_asset",
    "assert_figure_asset_ready",
    "MIN_EFFECTIVE_TEXT_PT",
    "MAX_EDGE_WHITESPACE_RATIO",
    "MIN_DISPLAY_PPI",
    "AVAILABLE_HEIGHT_FILL_RATIO",
    "compute_available_height_inches",
]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Block illegible figure/chart assets before DOCX insertion.")
    parser.add_argument("asset", type=Path)
    parser.add_argument("--display-width", type=float, required=True, dest="display_width_inches")
    parser.add_argument("--source-width", type=float, dest="source_width_inches")
    parser.add_argument("--contains-text", action="store_true")
    parser.add_argument("--min-text-pt", type=float, dest="declared_min_font_pt")
    parser.add_argument("--json-report", type=Path)
    args = parser.parse_args(argv)
    report = audit_figure_asset(
        args.asset,
        display_width_inches=args.display_width_inches,
        source_width_inches=args.source_width_inches,
        declared_min_font_pt=args.declared_min_font_pt,
        contains_text=args.contains_text,
    )
    payload = json.dumps(report, ensure_ascii=False, indent=2)
    print(payload)
    if args.json_report:
        args.json_report.parent.mkdir(parents=True, exist_ok=True)
        args.json_report.write_text(payload + "\n", encoding="utf-8")
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
