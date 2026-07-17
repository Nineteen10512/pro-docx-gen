"""Shared template scanner & theme extractor for PRO-PPTX/DOCX v1.5.1.

Provides:
- scan_local_templates(dirs=None, recursive=True) -> list[TemplateInfo]
- extract_template_theme(path) -> ExtractedTheme  (routes by extension)
- apply_extracted_theme(extracted, base_theme='academic') -> tokens_dict
- contrast_ratio / wcag_contrast utilities

PPT specifics (.pptx/.potx/.dpt) are extracted via python-pptx;
DOCX specifics (.docx/.dotx/.wpt) via python-docx. Lazy imports so that
this module can be imported even if only one side is installed.

@since v1.5.1
"""
from __future__ import annotations
import os
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional
from xml.etree import ElementTree as ET

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class ExtractedTheme:
    """Result of extracting a theme from a local template file."""
    tokens_dict: dict = field(default_factory=dict)
    source_path: str = ""
    source_format: str = ""       # pptx / potx / dpt / docx / dotx / wpt
    confidence: float = 0.0       # 0–100
    warnings: list = field(default_factory=list)


# ---------------------------------------------------------------------------
# Platform default WPS template dirs
# ---------------------------------------------------------------------------

def _default_wps_dirs() -> dict:
    """Return default scan directories per endpoint, keyed by OS."""
    is_win = sys.platform.startswith("win")
    is_mac = sys.platform == "darwin"
    if is_win:
        appdata = os.environ.get("APPDATA", "")
        userprofile = os.environ.get("USERPROFILE", "")
        program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
        program_files_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
        return {
            "ppt": [
                os.path.join(appdata, "Kingsoft", "office6", "templates", "wpp") if appdata else "",
                os.path.join(program_files, "Kingsoft", "WPS Office", "mui", "zh_CN", "templates", "wpp"),
                os.path.join(program_files_x86, "Kingsoft", "WPS Office", "mui", "zh_CN", "templates", "wpp"),
                os.path.join(userprofile, "Documents", "自定义 Office 模板") if userprofile else "",
            ],
            "docx": [
                os.path.join(appdata, "Kingsoft", "office6", "templates", "wps") if appdata else "",
                os.path.join(program_files, "Kingsoft", "WPS Office", "mui", "zh_CN", "templates", "wps"),
                os.path.join(program_files_x86, "Kingsoft", "WPS Office", "mui", "zh_CN", "templates", "wps"),
                os.path.join(userprofile, "Documents", "自定义 Office 模板") if userprofile else "",
            ],
        }
    if is_mac:
        home = os.path.expanduser("~")
        return {
            "ppt": [
                os.path.join(home, "Library", "Application Support", "Kingsoft", "WPS Office", "templates", "wpp"),
                os.path.join(home, "Documents", "自定义 Office 模板"),
            ],
            "docx": [
                os.path.join(home, "Library", "Application Support", "Kingsoft", "WPS Office", "templates", "wps"),
                os.path.join(home, "Documents", "自定义 Office 模板"),
            ],
        }
    # Linux / cloud — no defaults; scan returns empty unless user supplies dirs
    return {"ppt": [], "docx": []}


# ---------------------------------------------------------------------------
# scan_local_templates
# ---------------------------------------------------------------------------

_PPT_EXTS = {".pptx", ".potx", ".dpt"}
_DOCX_EXTS = {".docx", ".dotx", ".wpt"}
_ALL_EXTS = _PPT_EXTS | _DOCX_EXTS


def scan_local_templates(dirs: Optional[list] = None, recursive: bool = True) -> list:
    """Scan local WPS template dirs + user-supplied dirs for template files.

    Returns a list of TemplateInfo (from shared.template_registry).
    On non-Windows with no dirs supplied, returns [] gracefully.
    """
    try:
        from .template_registry import TemplateInfo
    except ImportError:  # pragma: no cover - legacy skill-root layout
        from shared.template_registry import TemplateInfo
    scan_dirs: list = []
    if dirs:
        scan_dirs.extend([d for d in dirs if d])
    else:
        defaults = _default_wps_dirs()
        scan_dirs.extend([d for d in defaults["ppt"] + defaults["docx"] if d])

    found: dict = {}   # path -> TemplateInfo
    for base in scan_dirs:
        if not base or not os.path.isdir(base):
            continue
        if recursive:
            for root, _dirs, files in os.walk(base):
                for fn in files:
                    ext = os.path.splitext(fn)[1].lower()
                    if ext in _ALL_EXTS:
                        fp = os.path.join(root, fn)
                        _add_scan_entry(found, fp, ext)
        else:
            for fn in os.listdir(base):
                ext = os.path.splitext(fn)[1].lower()
                if ext in _ALL_EXTS:
                    fp = os.path.join(base, fn)
                    _add_scan_entry(found, fp, ext)
    return list(found.values())


def _add_scan_entry(found: dict, fp: str, ext: str):
    try:
        from .template_registry import TemplateInfo
    except ImportError:  # pragma: no cover - legacy skill-root layout
        from shared.template_registry import TemplateInfo
    try:
        st = os.stat(fp)
        endpoint = "ppt" if ext in _PPT_EXTS else "docx"
        # .dpt / .wpt are proprietary WPS formats — flag hint
        preview_hint = "native"
        if ext in (".dpt", ".wpt"):
            preview_hint = "wps_proprietary_warning"
        found[fp] = TemplateInfo(
            name=os.path.splitext(os.path.basename(fp))[0],
            display_name=os.path.splitext(os.path.basename(fp))[0],
            endpoint=endpoint,
            scene="local",
            description=f"WPS本地模板 ({ext})",
            preview_prompt=preview_hint,
            path=fp,
        )
    except OSError:
        pass


# ---------------------------------------------------------------------------
# extract_template_theme — routes by file extension
# ---------------------------------------------------------------------------

def extract_template_theme(path: str) -> ExtractedTheme:
    """Extract color/font/background theme tokens from a template file.

    Routes to pro_ppt_gen.theme_extractor or pro_docx_gen.theme_extractor
    based on extension. Returns ExtractedTheme with tokens_dict populated
    with keys matching Design Tokens schema (primary/secondary/accent/bg/
    text/font/etc.). Confidence < 30 triggers fallback in the caller.
    """
    ext = os.path.splitext(path)[1].lower()
    if not os.path.isfile(path):
        return ExtractedTheme(source_path=path, warnings=[f"file not found: {path}"])

    if ext in _PPT_EXTS:
        return _extract_ppt_theme(path, ext)
    if ext in _DOCX_EXTS:
        return _extract_docx_theme(path, ext)
    return ExtractedTheme(source_path=path, warnings=[f"unsupported extension: {ext}"])


# ---- PPT extraction -----------------------------------------------------

def _extract_ppt_theme(path: str, ext: str) -> ExtractedTheme:
    warnings: list = []
    try:
        from pptx import Presentation
    except ImportError as e:
        return ExtractedTheme(source_path=path, source_format=ext,
                              warnings=[f"python-pptx not installed: {e}"])
    # .dpt is a WPS-proprietary zip variant — try opening; if it fails, warn
    try:
        if ext == ".dpt":
            # Probe zip
            if not zipfile.is_zipfile(path):
                warnings.append(".dpt 格式不是标准 zip，请在 WPS 中另存为 .potx/.pptx 后使用")
                return ExtractedTheme(source_path=path, source_format=ext, warnings=warnings)
        prs = Presentation(path)
    except Exception as e:
        warnings.append(f"无法打开 PPT 文件({ext}): {e}")
        if ext == ".dpt":
            warnings.append("提示：请在 WPS 中将 .dpt 另存为 .potx 或 .pptx 后重试")
        return ExtractedTheme(source_path=path, source_format=ext, warnings=warnings)

    tokens: dict = {}
    confidence_parts = 0.0
    master = prs.slide_master

    # --- Colors from slide_master theme part (via rels; theme xml is NOT a
    #     direct child of the slideMaster element in OOXML — must follow the
    #     relationship to ppt/theme/themeN.xml) ---
    theme_el = None
    try:
        for rel in prs.slide_master.part.rels.values():
            if "theme" in rel.reltype:
                theme_xml = rel.target_part.blob
                theme_el = ET.fromstring(theme_xml)
                break
    except Exception as e:
        warnings.append(f"theme rel 解析异常: {e}")

    if theme_el is None:
        # Fallback: search anywhere in the master element tree
        theme_el = master.element.find(
            ".//{http://schemas.openxmlformats.org/drawingml/2006/main}theme"
        )

    try:
        if theme_el is not None:
            A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
            # clrScheme lives under a:themeElements/a:clrScheme, not as direct child of a:theme
            clr = theme_el.find(f"{{{A_NS}}}themeElements/{{{A_NS}}}clrScheme")
            if clr is None:
                clr = theme_el.find(f".//{{{A_NS}}}clrScheme")
            if clr is not None:
                color_map = _parse_clr_scheme(clr)
                tokens.update(_map_ppt_colors(color_map))
                confidence_parts += 40
            font = theme_el.find(f"{{{A_NS}}}themeElements/{{{A_NS}}}fontScheme")
            if font is None:
                font = theme_el.find(f".//{{{A_NS}}}fontScheme")
            if font is not None:
                font_map = _parse_font_scheme(font)
                tokens.update(font_map)
                confidence_parts += 30
    except Exception as e:
        warnings.append(f"theme 解析异常: {e}")

    # --- Background fill from slide_master ---
    try:
        bg_tokens = _extract_ppt_background(prs.slide_master)
        if bg_tokens:
            tokens.update(bg_tokens)
            confidence_parts += 15
    except Exception as e:
        warnings.append(f"background 解析异常: {e}")

    # --- Title bar accent: sample first shape fill on first layout ---
    try:
        if prs.slide_layouts:
            title_bar = _sample_first_shape_fill(prs.slide_layouts[0])
            if title_bar:
                tokens["title_bar_bg"] = title_bar
                confidence_parts += 15
    except Exception as e:
        warnings.append(f"title bar 采样异常: {e}")

    confidence = min(100.0, confidence_parts)
    return ExtractedTheme(
        tokens_dict=tokens,
        source_path=path,
        source_format=ext.lstrip("."),
        confidence=confidence,
        warnings=warnings,
    )


def _parse_clr_scheme(clr) -> dict:
    """Parse a:clrScheme into {name: srgb_hex}."""
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    out = {}
    for child in clr:
        tag = child.tag.replace(ns, "")
        srgb = child.find(f"{ns}srgbClr")
        sysclr = child.find(f"{ns}sysClr")
        if srgb is not None:
            out[tag] = srgb.get("val", "").upper()
        elif sysclr is not None:
            out[tag] = sysclr.get("lastClr", "").upper()
    return out


def _map_ppt_colors(cm: dict) -> dict:
    """Map a:clrScheme 12-color names to Design Token keys."""
    tokens = {}
    if "dk1" in cm:
        tokens["text"] = "#" + cm["dk1"]
    if "lt1" in cm:
        tokens["bg"] = "#" + cm["lt1"]
    if "dk2" in cm:
        tokens["text_secondary"] = "#" + cm["dk2"]
    if "lt2" in cm:
        tokens["bg_secondary"] = "#" + cm["lt2"]
    accents = []
    for k in ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6"):
        if k in cm:
            accents.append("#" + cm[k])
    if accents:
        tokens["primary"] = accents[0]
        tokens["accent"] = accents[0]
        if len(accents) >= 2:
            tokens["secondary"] = accents[1]
        if len(accents) >= 3:
            tokens["accent_palette"] = accents
    if "hlink" in cm:
        tokens["link"] = "#" + cm["hlink"]
    return tokens


def _parse_font_scheme(font) -> dict:
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    major = font.find(f"{ns}majorFont")
    minor = font.find(f"{ns}minorFont")
    out = {}
    if major is not None:
        latin = major.find(f"{ns}latin")
        ea = major.find(f"{ns}ea")
        if latin is not None and latin.get("typeface"):
            out["heading_font_latin"] = latin.get("typeface")
        if ea is not None and ea.get("typeface"):
            out["heading_font_ea"] = ea.get("typeface")
    if minor is not None:
        latin = minor.find(f"{ns}latin")
        ea = minor.find(f"{ns}ea")
        if latin is not None and latin.get("typeface"):
            out["body_font_latin"] = latin.get("typeface")
        if ea is not None and ea.get("typeface"):
            out["body_font_ea"] = ea.get("typeface")
    return out


def _extract_ppt_background(master) -> dict:
    """Extract background fill from slide master as hex color if solid."""
    ns_a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    bg_el = master.element.find(f".//{ns_a}bg")
    if bg_el is None:
        # Try background on master directly
        bgPr = master.element.find(f".//{ns_a}bgPr")
        if bgPr is None:
            return {}
    else:
        bgPr = bg_el.find(f"{ns_a}bgPr")
    if bgPr is None:
        return {}
    solid = bgPr.find(f"{ns_a}solidFill")
    if solid is not None:
        srgb = solid.find(f"{ns_a}srgbClr")
        if srgb is not None:
            return {"bg": "#" + srgb.get("val", "").upper()}
    return {}


def _sample_first_shape_fill(layout) -> Optional[str]:
    """Sample the first shape on a layout that has a solid fill, return hex."""
    ns_a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    for sp in layout.shapes:
        try:
            el = sp.element
            spPr = el.find(f".//{ns_a}spPr")
            if spPr is None:
                continue
            solid = spPr.find(f"{ns_a}solidFill")
            if solid is not None:
                srgb = solid.find(f"{ns_a}srgbClr")
                if srgb is not None:
                    return "#" + srgb.get("val", "").upper()
        except Exception:
            continue
    return None


# ---- DOCX extraction ----------------------------------------------------

def _extract_docx_theme(path: str, ext: str) -> ExtractedTheme:
    warnings: list = []
    try:
        from docx import Document
    except ImportError as e:
        return ExtractedTheme(source_path=path, source_format=ext,
                              warnings=[f"python-docx not installed: {e}"])
    try:
        if ext == ".wpt":
            if not zipfile.is_zipfile(path):
                warnings.append(".wpt 格式不是标准 zip，请在 WPS 中另存为 .dotx/.docx 后使用")
                return ExtractedTheme(source_path=path, source_format=ext, warnings=warnings)
        doc = Document(path)
    except Exception as e:
        warnings.append(f"无法打开 DOCX 文件({ext}): {e}")
        if ext == ".wpt":
            warnings.append("提示：请在 WPS 中将 .wpt 另存为 .dotx 或 .docx 后重试")
        return ExtractedTheme(source_path=path, source_format=ext, warnings=warnings)

    tokens: dict = {}
    confidence_parts = 0.0

    # --- Theme from document part — theme is NOT a direct child of w:document
    #     in OOXML; must walk relationships from the main document part to
    #     find word/theme/themeN.xml (same pattern as PPT slide_master). ---
    try:
        theme_blob = None
        # Try rels from the document part
        try:
            for rel in doc.part.rels.values():
                if "theme" in (rel.reltype or "").lower() and rel.target_ref.endswith(".xml"):
                    theme_blob = rel.target_part.blob
                    break
        except Exception:
            theme_blob = None
        if theme_blob:
            theme_el = ET.fromstring(theme_blob)
            clr = theme_el.find(f"{{{A_NS}}}clrScheme")
            if clr is None:
                for child in theme_el.iter(f"{{{A_NS}}}clrScheme"):
                    clr = child
                    break
            if clr is not None:
                cm = _parse_clr_scheme(clr)
                tokens.update(_map_docx_colors(cm))
                confidence_parts += 40
            font = theme_el.find(f"{{{A_NS}}}fontScheme")
            if font is None:
                for child in theme_el.iter(f"{{{A_NS}}}fontScheme"):
                    font = child
                    break
            if font is not None:
                fm = _parse_font_scheme(font)
                if fm.get("heading_font_ea") or fm.get("body_font_ea"):
                    tokens["heading_font_ea"] = fm.get("heading_font_ea") or fm.get("body_font_ea")
                    tokens["body_font_ea"] = fm.get("body_font_ea") or fm.get("heading_font_ea")
                if fm.get("heading_font_latin") or fm.get("body_font_latin"):
                    tokens["heading_font_latin"] = fm.get("heading_font_latin") or fm.get("body_font_latin")
                    tokens["body_font_latin"] = fm.get("body_font_latin") or fm.get("heading_font_latin")
                confidence_parts += 30
    except Exception as e:
        warnings.append(f"theme 解析异常: {e}")

    # --- Background from document (if any) ---
    try:
        from docx.oxml.ns import qn
        bg = doc.element.find(".//" + qn("w:background"))
        if bg is not None:
            color = bg.get(qn("w:color"))
            if color and len(color) == 6:
                tokens["bg"] = "#" + color.upper()
                confidence_parts += 15
    except Exception as e:
        warnings.append(f"background 解析异常: {e}")

    confidence = min(100.0, confidence_parts)
    return ExtractedTheme(
        tokens_dict=tokens,
        source_path=path,
        source_format=ext.lstrip("."),
        confidence=confidence,
        warnings=warnings,
    )


def _map_docx_colors(cm: dict) -> dict:
    tokens = {}
    if "dk1" in cm:
        tokens["text"] = "#" + cm["dk1"]
    if "lt1" in cm:
        tokens["bg"] = "#" + cm["lt1"]
    if "dk2" in cm:
        tokens["text_secondary"] = "#" + cm["dk2"]
    if "lt2" in cm:
        tokens["bg_secondary"] = "#" + cm["lt2"]
    accents = []
    for k in ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6"):
        if k in cm:
            accents.append("#" + cm[k])
    if accents:
        tokens["primary"] = accents[0]
        tokens["accent"] = accents[0]
        if len(accents) >= 2:
            tokens["secondary"] = accents[1]
        if len(accents) >= 3:
            tokens["accent_palette"] = accents
    return tokens


# ---------------------------------------------------------------------------
# apply_extracted_theme — deep-merge onto base theme + contrast check
# ---------------------------------------------------------------------------

def apply_extracted_theme(extracted: ExtractedTheme, base_theme: str = "academic") -> dict:
    """Deep-merge extracted tokens onto a base theme from shared.themes.

    - If extracted.confidence < 30, returns base theme tokens as-is.
    - After merging, runs WCAG AA contrast check on text/bg; if ratio < 4.5,
      auto-adjusts text color toward black or white and adds warning.
    """
    try:
        from .themes import THEME_PALETTES
    except ImportError:  # pragma: no cover - legacy skill-root layout
        from shared.themes import THEME_PALETTES
    base = dict(THEME_PALETTES.get(base_theme, THEME_PALETTES.get("academic", {})))
    if extracted.confidence < 30:
        base.setdefault("_warnings", []).extend(extracted.warnings)
        base.setdefault("_warnings", []).append("confidence<30, 使用基础主题")
        return base
    merged = dict(base)
    merged.update(extracted.tokens_dict)
    # contrast check
    warnings = list(extracted.warnings)
    text = merged.get("text", "#222222")
    bg = merged.get("bg", "#FFFFFF")
    ratio = wcag_contrast(text, bg)
    if ratio < 4.5:
        # auto-adjust text color
        new_text = _auto_contrast_text(bg)
        warnings.append(f"text/bg 对比度 {ratio:.2f} 低于 WCAG AA(4.5)，已自动将文本颜色调为 {new_text}")
        merged["text"] = new_text
        merged["_contrast_adjusted"] = True
    merged["_warnings"] = warnings
    merged["_confidence"] = extracted.confidence
    return merged


# ---------------------------------------------------------------------------
# WCAG contrast utils
# ---------------------------------------------------------------------------

def _hex_to_rgb(h: str) -> tuple:
    h = h.lstrip("#")
    if len(h) == 3:
        h = "".join(ch * 2 for ch in h)
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _rel_lum(hex_color: str) -> float:
    r, g, b = _hex_to_rgb(hex_color)
    def chan(c):
        c = c / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * chan(r) + 0.7152 * chan(g) + 0.0722 * chan(b)


def wcag_contrast(fg: str, bg: str) -> float:
    l1 = _rel_lum(fg)
    l2 = _rel_lum(bg)
    lighter, darker = max(l1, l2), min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def _auto_contrast_text(bg: str) -> str:
    """Return black or white depending on background luminance."""
    lum = _rel_lum(bg)
    return "#000000" if lum > 0.4 else "#FFFFFF"
