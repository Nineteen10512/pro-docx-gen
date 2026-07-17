"""Shared quality-check utilities for PaperJSX (ARCH-2 unified QA).

Provides WCAG 2.1 contrast math, CJK-aware word counting, assertive-title
heuristics, information-density grading, and v1.5.2 contrast hard-guard
(``contrast_check``) — used by both PPT and DOCX ``quality_check`` APIs.

@since v1.3.0 (ARCH-2)
@since v1.5.2 新增 contrast_check：对批量 (bg, fg) 对做 WCAG 铁律校验，
               返回 warning/error issues，供 ppt_jsx/taste_check 调用。
"""
from __future__ import annotations

import re
from typing import Iterable, List, Tuple, Union


ColorTriple = Tuple[int, int, int]
ColorLike = Union[str, ColorTriple, None]


# ---------------------------------------------------------------------------
# WCAG 2.1 relative luminance and contrast ratio
# ---------------------------------------------------------------------------

def _linearize(channel_01: float) -> float:
    """sRGB channel 0..1 → linear 0..1."""
    if channel_01 <= 0.03928:
        return channel_01 / 12.92
    return ((channel_01 + 0.055) / 1.055) ** 2.4


def relative_luminance(r: int, g: int, b: int) -> float:
    """Return WCAG relative luminance (0 = black, 1 = white).

    ``r, g, b`` are 0–255 integers.
    """
    R = _linearize(r / 255.0)
    G = _linearize(g / 255.0)
    B = _linearize(b / 255.0)
    return 0.2126 * R + 0.7152 * G + 0.0722 * B


def contrast_ratio(rgb1: Tuple[int, int, int], rgb2: Tuple[int, int, int]) -> float:
    """WCAG contrast ratio. Returns value ≥ 1.0 (up to ~21)."""
    L1 = relative_luminance(*rgb1)
    L2 = relative_luminance(*rgb2)
    if L1 < L2:
        L1, L2 = L2, L1
    return (L1 + 0.05) / (L2 + 0.05)


def wcag_grade(ratio: float, large_text: bool = False) -> str:
    """Return 'AAA', 'AA', or 'FAIL' based on WCAG thresholds."""
    aa = 3.0 if large_text else 4.5
    aaa = 4.5 if large_text else 7.0
    if ratio >= aaa:
        return "AAA"
    if ratio >= aa:
        return "AA"
    return "FAIL"


def hex_to_rgb(h: str) -> Tuple[int, int, int]:
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


# ---------------------------------------------------------------------------
# CJK-aware word/char counting
# ---------------------------------------------------------------------------

_CJK_RE = re.compile(r"[\u4e00-\u9fff\u3400-\u4dbf\u3040-\u30ff\uac00-\ud7af]")


def count_words(text: str, lang: str = "cn") -> int:
    """Hybrid CJK-aware word count.

    - CJK characters count as 1 word each.
    - Runs of Latin characters count as words split by whitespace/punctuation.
    """
    if not text:
        return 0
    total = 0
    buf = []

    def flush():
        nonlocal total
        if buf:
            # join and split Latin by whitespace/punctuation
            s = "".join(buf)
            words = re.findall(r"[A-Za-z0-9]+(?:['-][A-Za-z0-9]+)*", s)
            total += len(words)
            buf.clear()

    for ch in text:
        if _CJK_RE.match(ch):
            flush()
            total += 1
        else:
            buf.append(ch)
    flush()
    return total


def count_chars_cjk(text: str) -> int:
    """Count Chinese characters (useful for 30-char bullet limit)."""
    return sum(1 for ch in text if "\u4e00" <= ch <= "\u9fff")


def text_length_cjk_aware(text: str) -> int:
    """Single-bullet length metric: CJK=1 each, Latin words=1 each.

    The v1.2 spec says ≤30字 for bullets; this is the compatible metric.
    """
    return count_words(text, lang="cn")


# ---------------------------------------------------------------------------
# Assertive-title heuristics
# ---------------------------------------------------------------------------

# Verbs / markers that make a title sound like a conclusion, not a noun phrase
_ASSERTIVE_VERBS = [
    "实现", "提升", "提高", "下降", "降低", "减少", "增长", "增长了", "增加", "达到",
    "突破", "完成", "达成", "推动", "驱动", "促进", "拉动", "带动", "改善", "优化",
    "扩大", "缩小", "拓展", "获得", "取得", "保持", "超越", "领先", "引领", "确立",
    "证明", "表明", "显示", "揭示", "说明", "证实", "验证", "发现",
    "is", "are", "was", "were", "achieves", "increases", "decreases", "grows",
    "drives", "shows", "proves", "reveals", "delivers", "reaches", "improves",
    "reduces", "boosts", "enables", "leads",
]
_CONCLUSION_MARKERS = [
    "结论", "总结", "关键", "核心", "启示", "建议", "方案", "策略",
    "关键在于", "核心是", "本质是",
]
_NUMBER_RE = re.compile(r"\d+(?:\.\d+)?%?")


def is_assertive_title(title: str) -> bool:
    """Heuristic: is the title a conclusion/sentence vs. a noun phrase?"""
    if not title:
        return False
    t = title.strip()
    # Contains a number + verb pattern → likely a data-driven conclusion
    has_number = bool(_NUMBER_RE.search(t))
    has_assertive_verb = any(v in t for v in _ASSERTIVE_VERBS)
    has_conclusion = any(m in t for m in _CONCLUSION_MARKERS)
    # Contains sentence-final punctuation → conclusion
    has_sentence_end = any(p in t for p in ("。", "！", "？", ".", "!", "?"))
    # Simple subject+predicate (含 "是", "为", "有", "将", "能")
    has_predicate = any(v in t for v in ("是", "为", "有", "将", "能", "可", "会", "已", "正"))

    score = 0
    if has_number:
        score += 1
    if has_assertive_verb:
        score += 1
    if has_conclusion:
        score += 1
    if has_sentence_end:
        score += 1
    if has_predicate and len(t) >= 8:
        score += 1
    return score >= 2


def assertive_title_suggestion(title: str, data_point: str = "") -> str:
    """Return a rewrite suggestion for a noun-phrase title (best-effort template)."""
    t = (title or "").strip()
    if not t:
        return "请使用结论式标题，如「XX达成/提升/实现XX」"
    # Noun phrase patterns
    if data_point:
        return f"将「{t}」改写为结论：「{t}达到{data_point}」或「{t}{{动词}}{{结论}}」"
    return f"将名词短语「{t}」改写为动词+结论的断言式标题，例如「{t}实现显著提升」"


# ---------------------------------------------------------------------------
# Information density grading
# ---------------------------------------------------------------------------

def grade_density(score: float) -> str:
    """Map a numeric density score (0..~20) to low/medium/high."""
    if score < 3:
        return "low"
    if score > 14:
        return "high"
    return "medium"


# ---------------------------------------------------------------------------
# Scoring helpers
# ---------------------------------------------------------------------------

def weighted_score(dimensions: dict, weights: dict) -> Tuple[float, dict]:
    """Compute weighted total and per-dimension scores.

    Returns ``(total_0_100, per_dim_scores)``.
    """
    total = 0.0
    wsum = 0.0
    scores = {}
    for name, dim in dimensions.items():
        s = float(dim.get("score", 0))
        w = float(weights.get(name, 1.0))
        total += s * w
        wsum += w
        scores[name] = s
    if wsum <= 0:
        return 0.0, scores
    return round(total / wsum, 1), scores





def _coerce_rgb(c: ColorLike) -> ColorTriple | None:
    """Convert a hex string / 3-tuple / RGBColor-like to (r,g,b) 0-255; return None on failure."""
    if c is None:
        return None
    if isinstance(c, str):
        try:
            return hex_to_rgb(c)
        except Exception:
            return None
    try:
        t = tuple(int(v) for v in tuple(c)[:3])
        if all(0 <= v <= 255 for v in t):
            return t  # type: ignore[return-value]
    except Exception:
        pass
    return None


def safe_text_color(bg: ColorLike, dark: ColorLike = (0x11, 0x11, 0x11),
                    light: ColorLike = (0xFF, 0xFF, 0xFF),
                    threshold: float = 0.5) -> ColorTriple:
    """Defensive color pick: given a bg, return dark or light whichever is readable.

    Uses the same WCAG relative-luminance math as ``relative_luminance``. If bg is
    unparseable, returns ``dark`` (safe default on white slides).

    Used by PPT/DOCX renderers to auto-correct text color when the caller passes
    low-contrast combinations (e.g. black text on dark section bar or vice versa).
    """
    bg_rgb = _coerce_rgb(bg)
    if bg_rgb is None:
        return _coerce_rgb(dark) or (0x11, 0x11, 0x11)
    L = relative_luminance(*bg_rgb)
    return _coerce_rgb(light) if L < threshold else _coerce_rgb(dark)  # type: ignore[return-value]


def rgb_to_hex(rgb: ColorTriple) -> str:
    return "#{:02X}{:02X}{:02X}".format(int(rgb[0]), int(rgb[1]), int(rgb[2]))


def contrast_check(
    pairs: Iterable[Tuple[ColorLike, ColorLike, str]],
    *,
    error_ratio: float = 3.0,
    warning_ratio: float = 4.5,
    large_text: bool = False,
) -> List[dict]:
    """Run WCAG contrast iron-rule on a list of (bg, fg, label) pairs.

    Parameters
    ----------
    pairs : iterable of (bg, fg, label)
        ``bg`` / ``fg`` accept ``#RRGGBB`` str or (r,g,b) tuple or None.
        ``label`` is a short human tag such as ``"slide3.title"`` / ``"kpi_card.value"``.
    error_ratio : float
        Ratio below which an ``error`` issue is emitted (text is effectively invisible).
        Default 3.0 (WCAG AA large-text floor; below this → almost unreadable).
    warning_ratio : float
        Ratio below which a ``warning`` issue is emitted. Default 4.5 (WCAG AA body-text).
    large_text : bool
        If True the warning floor drops to 3.0 (WCAG AA for ≥18pt / ≥14pt bold).

    Returns
    -------
    list[dict]
        Each issue is ``{"level": "error"|"warning", "code": "low_contrast_text",
        "label": str, "ratio": float, "bg": "#rrggbb", "fg": "#rrggbb",
        "suggestion": str}``.
    """
    warn_floor = 3.0 if large_text else warning_ratio
    err_floor = min(error_ratio, warn_floor)
    issues: List[dict] = []
    for bg, fg, label in pairs:
        b = _coerce_rgb(bg)
        f = _coerce_rgb(fg)
        if b is None or f is None:
            continue
        ratio = contrast_ratio(b, f)
        if ratio < err_floor:
            level = "error"
        elif ratio < warn_floor:
            level = "warning"
        else:
            continue
        # Suggest a safe fix
        suggested = safe_text_color(b)
        issues.append({
            "level": level,
            "code": "low_contrast_text",
            "label": label,
            "ratio": round(ratio, 2),
            "bg": rgb_to_hex(b),
            "fg": rgb_to_hex(f),
            "suggestion": (
                f"Replace text color {rgb_to_hex(f)} with {rgb_to_hex(suggested)} "
                f"on background {rgb_to_hex(b)} (ratio {ratio:.2f} < "
                f"{'3.0' if large_text else '4.5'}:1)."
            ),
        })
    return issues


__all__ = [
    "relative_luminance", "contrast_ratio", "wcag_grade", "hex_to_rgb",
    "rgb_to_hex", "safe_text_color", "contrast_check",
    "count_words", "count_chars_cjk", "text_length_cjk_aware",
    "is_assertive_title", "assertive_title_suggestion",
    "grade_density", "weighted_score",
]