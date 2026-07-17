"""DOCX theme_extractor — v1.5.1 A 能力域。

从 .docx/.dotx/.wpt 文档提取主题（色板/字体），并应用到生成 tokens。

设计原则（v1.5.1 强化）：
- 不降级！任何输入都要真实提取，失败抛清晰错误，不静默 fallback
- 扩展名是 hint，真实判定靠 zipfile + word/document.xml 探测
- .wpt 视为 WPS zip 模板（实际结构与 docx 类似），同路径解析
- 优先调用 ``shared.template_scanner.extract_template_theme``（如果已实现）
- 软降级仅在 shared 层未就绪时（兜底解析）—— 不允许"什么都不做"
- apply_extracted_theme 深合并到现有 tokens

公开 API：
- extract_docx_theme(path) → dict（必须含 tokens/source_path/confidence）
- apply_extracted_theme(tokens, extracted) → dict（合并后的 tokens）

@since v1.5.1
"""
from __future__ import annotations

import os
import re
import zipfile
from typing import Any, Optional
from xml.etree import ElementTree as ET

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

REQUIRED_DOCX_PART = "word/document.xml"


class ThemeExtractionError(Exception):
    """主题提取失败。不允许静默降级，调用方应处理此异常。"""
    pass


def _probe_zip_container(path: str) -> dict[str, Any]:
    """探测文件是否可作为 zip 容器打开，并定位主题 part。

    Returns:
        {
            "is_zip": bool,
            "theme_part": str | None,   # 如 'word/theme/theme1.xml'
            "document_part": str | None,  # 如 'word/document.xml'
            "all_parts": list[str],
        }
    """
    result = {"is_zip": False, "theme_part": None, "document_part": None, "all_parts": []}
    try:
        with zipfile.ZipFile(path, "r") as z:
            names = z.namelist()
            result["all_parts"] = names
            result["is_zip"] = True
            for n in names:
                if re.match(r"word/theme/theme\d*\.xml$", n):
                    result["theme_part"] = n
                    break
            if REQUIRED_DOCX_PART in names:
                result["document_part"] = REQUIRED_DOCX_PART
    except (zipfile.BadZipFile, OSError, KeyError):
        pass
    return result


def _detect_format(path: str, probe: dict[str, Any]) -> str:
    """根据扩展名 + zip 探测综合判定格式。"""
    if not os.path.exists(path):
        return "missing"
    lower = path.lower()
    ext_hint = None
    if lower.endswith(".docx"):
        ext_hint = "docx"
    elif lower.endswith(".dotx"):
        ext_hint = "dotx"
    elif lower.endswith(".wpt"):
        ext_hint = "wpt"
    if not probe["is_zip"]:
        return "non_zip"
    if probe["document_part"] is None:
        return "zip_no_doc"
    return ext_hint or "docx_like"


def _read_theme_xml_from_zip(path: str, theme_part: str) -> bytes:
    """从 zip 中读取指定 theme part 的字节。"""
    with zipfile.ZipFile(path, "r") as z:
        return z.read(theme_part)


def _parse_clr_scheme(theme_xml: bytes) -> dict[str, str]:
    """从 theme*.xml 解析 a:clrScheme 的 12 色，返回 name→hex 映射。"""
    result: dict[str, str] = {}
    if not theme_xml:
        return result
    try:
        root = ET.fromstring(theme_xml)
    except ET.ParseError as e:
        raise ThemeExtractionError(f"theme XML 解析失败: {e}") from e
    clr_scheme = None
    for child in root.iter(f"{{{A_NS}}}clrScheme"):
        clr_scheme = child
        break
    if clr_scheme is None:
        raise ThemeExtractionError("theme XML 中未找到 a:clrScheme")
    for color_el in clr_scheme:
        tag = color_el.tag.split("}", 1)[-1]
        srgb = color_el.find(f"{{{A_NS}}}srgbClr")
        sys_clr = color_el.find(f"{{{A_NS}}}sysClr")
        if srgb is not None:
            val = srgb.get("val", "")
            if val:
                result[tag] = f"#{val.upper()}"
        elif sys_clr is not None:
            val = sys_clr.get("val", "")
            if val == "windowText":
                result[tag] = "#000000"
            elif val == "window":
                result[tag] = "#FFFFFF"
            # 其他 sysClr val（如 'bg1'/'tx1'）需展开，暂不处理
    if not result:
        raise ThemeExtractionError("clrScheme 中未解析到任何颜色")
    return result


def _parse_font_scheme(theme_xml: bytes) -> dict[str, str]:
    """从 theme*.xml 解析 a:fontScheme 的 majorFont/minorFont latin typeface。"""
    result: dict[str, str] = {}
    if not theme_xml:
        return result
    try:
        root = ET.fromstring(theme_xml)
    except ET.ParseError as e:
        raise ThemeExtractionError(f"theme XML 解析失败: {e}") from e
    font_scheme = None
    for child in root.iter(f"{{{A_NS}}}fontScheme"):
        font_scheme = child
        break
    if font_scheme is None:
        raise ThemeExtractionError("theme XML 中未找到 a:fontScheme")
    for kind in ("majorFont", "minorFont"):
        el = font_scheme.find(f"{{{A_NS}}}{kind}")
        if el is None:
            continue
        latin = el.find(f"{{{A_NS}}}latin")
        if latin is not None and latin.get("typeface"):
            result[kind] = latin.get("typeface")
    if not result:
        raise ThemeExtractionError("fontScheme 中未解析到任何字体")
    return result


# 12 色 → 语义映射
_COLOR_MAP = {
    "dk1": "text",
    "lt1": "bg",
    "dk2": "text_strong",
    "lt2": "bg_soft",
    "accent1": "primary",
    "accent2": "secondary",
    "accent3": "accent",
    "accent4": "highlight",
    "accent5": "info",
    "accent6": "warning",
    "hlink": "link",
    "folHlink": "link_visited",
}


def _colors_to_tokens(colors: dict[str, str], fonts: dict[str, str]) -> dict[str, Any]:
    """把 12 色 + 字体映射到 pro_docx_gen 兼容的 tokens 子集。"""
    color_block: dict[str, str] = {}
    for clr_name, hex_val in colors.items():
        semantic = _COLOR_MAP.get(clr_name)
        if semantic:
            color_block[semantic] = hex_val
    # 必须有 text/bg
    color_block.setdefault("text", "#000000")
    color_block.setdefault("bg", "#FFFFFF")

    font_block: dict[str, Any] = {"family": {}}
    major = fonts.get("majorFont")
    minor = fonts.get("minorFont")
    if major:
        font_block["family"]["heading"] = major
        font_block["family"]["default"] = major
    if minor:
        font_block["family"]["body"] = minor
        font_block["family"]["cn"] = minor
    if not font_block["family"]:
        font_block["family"] = {"default": "Calibri", "cn": "宋体", "heading": "Calibri", "body": "Calibri"}

    return {"color": color_block, "font": font_block}


def extract_docx_theme(path: str) -> dict[str, Any]:
    """从 .docx/.dotx/.wpt 文档提取主题（真实提取，不降级）。

    Returns:
        {
            "tokens": dict,        # 至少含 color/bg/text 三字段
            "source_path": str,
            "format": "docx"|"dotx"|"wpt"|"docx_like",
            "theme_part": str,      # 实际解析的 part 名
            "confidence": float,   # 0~1
            "warnings": list[str],  # 警告（不影响结果）
        }

    Raises:
        FileNotFoundError: 文件不存在
        ThemeExtractionError: 文件非 zip、缺 word/document.xml、theme 解析失败
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"文件不存在: {path}")

    probe = _probe_zip_container(path)
    fmt = _detect_format(path, probe)

    if fmt == "missing":
        raise FileNotFoundError(f"文件不存在: {path}")
    if fmt == "non_zip":
        raise ThemeExtractionError(
            f"文件不是有效的 zip 容器（.docx/.dotx/.wpt 必须为 zip 格式）: {path}"
        )
    if fmt == "zip_no_doc":
        raise ThemeExtractionError(
            f"zip 中缺少必需的 word/document.xml part: {path}"
        )

    if not probe["theme_part"]:
        raise ThemeExtractionError(
            f"zip 中找不到 word/theme/theme*.xml，无法提取主题: {path}"
        )

    warnings: list[str] = []

    # 优先尝试 shared.template_scanner（如果主agent 已实现）
    try:
        try:
            from .shared.template_scanner import extract_template_theme as _shared_extract
        except ImportError:  # pragma: no cover - legacy skill-root layout
            from shared.template_scanner import extract_template_theme as _shared_extract
        result = _shared_extract(path)
        tokens = result.get("tokens_dict") or result.get("tokens") or {}
        if not tokens:
            raise ThemeExtractionError("shared.extract_template_theme 返回空 tokens")
        confidence = float(result.get("confidence", 0.0))
        warnings.extend(result.get("warnings", []))
        return {
            "tokens": tokens,
            "source_path": path,
            "format": fmt,
            "theme_part": probe["theme_part"],
            "confidence": confidence,
            "warnings": warnings,
        }
    except (ImportError, AttributeError):
        # shared 层未就绪，用本地解析
        pass

    # 本地解析：真实读取并解析 theme XML
    theme_xml = _read_theme_xml_from_zip(path, probe["theme_part"])
    colors = _parse_clr_scheme(theme_xml)
    fonts = _parse_font_scheme(theme_xml)
    tokens = _colors_to_tokens(colors, fonts)
    # 颜色 + 字体都解析到 → 高 confidence；仅其一 → 较低
    if colors and fonts:
        confidence = 0.9
    elif colors:
        confidence = 0.6
    else:
        confidence = 0.3
    return {
        "tokens": tokens,
        "source_path": path,
        "format": fmt,
        "theme_part": probe["theme_part"],
        "confidence": confidence,
        "warnings": warnings,
    }


def _to_docx_rgb(hex_or_rgb):
    """hex str → docx.shared.RGBColor；已是 RGBColor 则原样返回；None 返回 None。"""
    if hex_or_rgb is None:
        return None
    if isinstance(hex_or_rgb, str):
        from docx.shared import RGBColor
        try:
            return RGBColor.from_string(hex_or_rgb.lstrip("#"))
        except Exception:
            return hex_or_rgb
    return hex_or_rgb  # 已是 RGBColor 或其他类型，保持


def apply_extracted_theme(tokens: dict[str, Any], extracted: dict[str, Any]) -> dict[str, Any]:
    """深合并 extracted["tokens"] 到 tokens（不破坏 BASE_TOKENS 的其他字段）。

    关键：tokens["color"] 的值是 docx.shared.RGBColor 实例（不是 hex str）。
    提取得到的 hex str 会在合并时自动转成 RGBColor。

    Args:
        tokens: 现有 tokens（通常是 get_theme() 返回值或 BASE_TOKENS 副本）
        extracted: extract_docx_theme() 返回的 dict

    Returns:
        合并后的新 tokens dict

    Raises:
        ValueError: extracted 为空或不含 tokens 字段
    """
    if not extracted or "tokens" not in extracted:
        raise ValueError("extracted 必须由 extract_docx_theme() 返回，不能为空或缺少 tokens 字段")
    ext_tokens = extracted["tokens"]
    if not ext_tokens:
        raise ValueError("extracted['tokens'] 为空，无法合并")

    result = dict(tokens)

    # color 浅合并（hex str → RGBColor）
    if "color" in ext_tokens and ext_tokens["color"]:
        new_color = dict(result.get("color", {}))
        for k, v in ext_tokens["color"].items():
            new_color[k] = _to_docx_rgb(v)
        result["color"] = new_color

    # font 浅合并（嵌套 family 浅合并）
    if "font" in ext_tokens and ext_tokens["font"]:
        new_font = dict(result.get("font", {}))
        for k, v in ext_tokens["font"].items():
            if k == "family" and isinstance(v, dict) and isinstance(new_font.get("family"), dict):
                new_font["family"] = {**new_font["family"], **v}
            else:
                new_font[k] = v
        result["font"] = new_font

    # 其他顶层键浅合并
    for k, v in ext_tokens.items():
        if k in ("color", "font"):
            continue
        result[k] = v

    return result


__all__ = [
    "extract_docx_theme",
    "apply_extracted_theme",
    "ThemeExtractionError",
]
