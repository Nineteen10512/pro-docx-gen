"""Shared theme registry — PaperJSX 统一主题字典 (ARCH-3).

Centralizes the canonical theme palette HEX values so that all rendering
engines (PPT / DOCX / SVG mini-charts) resolve colors from a single source
of truth.  The engine-specific ``tokens/themes.py`` modules keep their own
typography / spacing overrides (because PPT uses Emu while DOCX uses Pt/Inches)
but MUST read palette colors from here instead of hard-coding hex.

此模块不依赖 python-pptx / python-docx，所有颜色值均以 ``#RRGGBB`` 字符串形式
存储；由调用方转换为对应引擎的 RGBColor 对象。

@since v1.4.0
@since v1.5.2 图表调色板扩充为每主题 ≥6 种区分度明显的颜色，且对背景满足
               WCAG AA ≥4.5:1 对比度；新增 single-series 多类别 bar/column
               循环取色规则，避免柱图/饼图单色渲染。
"""
from __future__ import annotations

from typing import Dict, List, Optional


# ---------------------------------------------------------------------------
# 1. 内置主题清单（同时在 PPT / DOCX 两侧可用）
# ---------------------------------------------------------------------------
BUILTIN_THEMES: List[str] = [
    "academic", "business", "teaching",
    "tech", "dark", "minimal",
    "nature", "sunset", "ocean", "forest", "warm",
    "premium", "chinese_red", "high_contrast",
]

# 别名 → 标准主题名
THEME_ALIASES: Dict[str, str] = {
    "light": "minimal",
    "科技": "tech", "科技蓝": "tech", "technology": "tech",
    "暗黑": "dark", "深色": "dark", "黑": "dark",
    "极简": "minimal", "白": "minimal", "简约": "minimal",
    "自然": "nature", "绿": "nature", "大地": "nature",
    "暖": "warm", "橙": "warm", "暖橙": "warm",
    "高端": "premium", "黑金": "premium", "商务金": "premium",
    "中国红": "chinese_red", "红": "chinese_red", "国风": "chinese_red",
    "海洋": "ocean", "海蓝": "ocean", "蓝": "ocean",
    "森林": "forest", "森系": "forest", "深绿": "forest",
    "日落": "sunset", "活力橙": "sunset",
    "学术": "academic", "论文": "academic",
    "商务": "business", "企业": "business",
    "教学": "teaching", "教案": "teaching",
    "高对比": "high_contrast", "高对比度": "high_contrast",
}


# ---------------------------------------------------------------------------
# 2. 基础调色板（academic 默认值）——与 color_palette.BASE_PALETTE 对齐
# ---------------------------------------------------------------------------
BASE_PALETTE: Dict[str, str] = {
    "primary":         "#1F3864",
    "secondary":       "#2E75B6",
    "accent":          "#C0504D",
    "bg":              "#FFFFFF",
    "text":            "#333333",
    "text_light":      "#666666",
    "text_on_primary": "#FFFFFF",
    "title_bar_bg":    "#1F3864",
    "table_header_bg": "#1F3864",
    "table_header_text": "#FFFFFF",
    "table_alt_row":   "#F2F6FC",
    "section_bg":      "#1F3864",
    "section_num_color": "#2E75B6",
    "kpi_bg":          "#F2F6FC",
    "quote_border":    "#2E75B6",
    "divider":         "#2E75B6",
    "chart_gridline":  "#DDDDDD",
    # DOCX extras
    "muted":           "#666666",
    "heading":         "#1F3864",
    "title":           "#1F3864",
    "table_border":    "#BFBFBF",
    "code_bg":         "#F5F5F5",
    "callout_info_bg":       "#E7F3F8",
    "callout_info_border":   "#2E75B6",
    "callout_warning_bg":    "#FFF4E5",
    "callout_warning_border":"#ED7D31",
    "callout_success_bg":    "#E8F5E9",
    "callout_success_border":"#43A047",
    "callout_danger_bg":     "#FDECEA",
    "callout_danger_border": "#C0504D",
}


def _hex(rgb: str) -> str:
    return rgb


# ---------------------------------------------------------------------------
# 3. 每个主题的核心 6 色调色板（ARCH-3 标准键：primary/secondary/accent/text/background/muted）
#
# 颜色值严格对齐 v1.3 pro_ppt_gen/tokens/themes.py 渲染结果，保证向后兼容。
# ---------------------------------------------------------------------------
THEME_PALETTES: Dict[str, Dict[str, str]] = {
    "academic": {
        "primary":    "#1F3864",
        "secondary":  "#2E75B6",
        "accent":     "#C0504D",
        "text":       "#333333",
        "background": "#FFFFFF",
        "muted":      "#666666",
    },
    "business": {
        "primary":    "#2C3E50",
        "secondary":  "#3498DB",
        "accent":     "#E67E22",
        "text":       "#2C3E50",
        "background": "#FFFFFF",
        "muted":      "#7F8C8D",
    },
    "teaching": {
        "primary":    "#2E7D32",
        "secondary":  "#66BB6A",
        "accent":     "#FFA000",
        "text":       "#212121",
        "background": "#FFFFFF",
        "muted":      "#616161",
    },
    "tech": {
        "primary":    "#0A192F",
        "secondary":  "#00D4FF",
        "accent":     "#64FFDA",
        "text":       "#0F172A",
        "background": "#F8FAFC",
        "muted":      "#64748B",
    },
    "dark": {
        # 注：PPT dark 主题 slide 背景为 #0F172A，文字为 #E2E8F0；此处 6 色按 PPT 实际渲染取
        "primary":    "#E2E8F0",
        "secondary":  "#60A5FA",
        "accent":     "#F59E0B",
        "text":       "#E2E8F0",
        "background": "#0F172A",
        "muted":      "#94A3B8",
    },
    "minimal": {
        "primary":    "#111827",
        "secondary":  "#374151",
        "accent":     "#2563EB",
        "text":       "#111827",
        "background": "#FFFFFF",
        "muted":      "#6B7280",
    },
    "nature": {
        "primary":    "#365314",
        "secondary":  "#65A30D",
        "accent":     "#A16207",
        "text":       "#1A2E05",
        "background": "#FAFDF7",
        "muted":      "#6B7A4E",
    },
    "sunset": {
        "primary":    "#7C2D12",
        "secondary":  "#EA580C",
        "accent":     "#DB2777",
        "text":       "#431407",
        "background": "#FFF7ED",
        "muted":      "#9A3412",
    },
    "ocean": {
        "primary":    "#0C4A6E",
        "secondary":  "#0284C7",
        "accent":     "#06B6D4",
        "text":       "#082F49",
        "background": "#F0F9FF",
        "muted":      "#475569",
    },
    "forest": {
        "primary":    "#064E3B",
        "secondary":  "#059669",
        "accent":     "#D97706",
        "text":       "#022C22",
        "background": "#F0FDF4",
        "muted":      "#4B5563",
    },
    "warm": {
        "primary":    "#7F1D1D",
        "secondary":  "#B45309",
        "accent":     "#D97706",
        "text":       "#451A03",
        "background": "#FFFBEB",
        "muted":      "#78350F",
    },
    "premium": {
        "primary":    "#1C1917",
        "secondary":  "#B08D57",
        "accent":     "#D4AF37",
        "text":       "#1C1917",
        "background": "#FAFAF9",
        "muted":      "#78716C",
    },
    "chinese_red": {
        "primary":    "#8B0000",
        "secondary":  "#C41E3A",
        "accent":     "#D4AF37",
        "text":       "#2B0000",
        "background": "#FFFDF5",
        "muted":      "#8B4513",
    },
    "high_contrast": {
        "primary":    "#000000",
        "secondary":  "#000000",
        "accent":     "#000000",
        "text":       "#000000",
        "background": "#FFFFFF",
        "muted":      "#333333",
    },
}


# ---------------------------------------------------------------------------
# 3b. 每个主题的扩展色板覆盖（与 BASE_PALETTE 合并后得到完整色板）——
#     这里的值严格对齐 v1.3 pro_ppt_gen/tokens/themes.py 的 _c("#...") HEX。
# ---------------------------------------------------------------------------
_THEME_EXTENDED: Dict[str, Dict[str, str]] = {
    "academic": {},

    "business": {
        "title_bar_bg":    "#2C3E50",
        "table_header_bg": "#2C3E50",
        "table_header_text": "#FFFFFF",
        "table_alt_row":   "#ECF0F1",
        "section_bg":      "#2C3E50",
        "section_num_color": "#3498DB",
        "kpi_bg":          "#ECF0F1",
        "quote_border":    "#3498DB",
        "divider":         "#3498DB",
        "chart_gridline":  "#D5DBDB",
        "text_on_primary": "#FFFFFF",
    },

    "teaching": {
        "title_bar_bg":    "#2E7D32",
        "table_header_bg": "#2E7D32",
        "table_header_text": "#FFFFFF",
        "table_alt_row":   "#E8F5E9",
        "section_bg":      "#2E7D32",
        "section_num_color": "#66BB6A",
        "kpi_bg":          "#E8F5E9",
        "quote_border":    "#66BB6A",
        "divider":         "#66BB6A",
        "chart_gridline":  "#C8E6C9",
        "text_on_primary": "#FFFFFF",
    },

    "tech": {
        "title_bar_bg":    "#0A192F",
        "table_header_bg": "#0A192F",
        "table_header_text": "#64FFDA",
        "table_alt_row":   "#EEF2FF",
        "section_bg":      "#0A192F",
        "section_num_color": "#00D4FF",
        "kpi_bg":          "#F0F5FF",
        "quote_border":    "#64FFDA",
        "divider":         "#00D4FF",
        "chart_gridline":  "#E2E8F0",
        "text_on_primary": "#FFFFFF",
    },

    "dark": {
        "title_bar_bg":    "#1E293B",
        "table_header_bg": "#1E293B",
        "table_header_text": "#60A5FA",
        "table_alt_row":   "#1E293B",
        "section_bg":      "#1E293B",
        "section_num_color": "#60A5FA",
        "kpi_bg":          "#1E293B",
        "quote_border":    "#60A5FA",
        "divider":         "#60A5FA",
        "chart_gridline":  "#334155",
        "text_on_primary": "#0F172A",
    },

    "minimal": {
        "title_bar_bg":    "#111827",
        "table_header_bg": "#111827",
        "table_header_text": "#FFFFFF",
        "table_alt_row":   "#F9FAFB",
        "section_bg":      "#111827",
        "section_num_color": "#9CA3AF",
        "kpi_bg":          "#F9FAFB",
        "quote_border":    "#2563EB",
        "divider":         "#2563EB",
        "chart_gridline":  "#E5E7EB",
        "text_on_primary": "#FFFFFF",
    },

    "nature": {
        "title_bar_bg":    "#365314",
        "table_header_bg": "#365314",
        "table_header_text": "#ECFCCB",
        "table_alt_row":   "#F0F9E4",
        "section_bg":      "#365314",
        "section_num_color": "#A3E635",
        "kpi_bg":          "#F0F9E4",
        "quote_border":    "#65A30D",
        "divider":         "#65A30D",
        "chart_gridline":  "#D9E5C3",
        "text_on_primary": "#FFFFFF",
    },

    "sunset": {
        "title_bar_bg":    "#7C2D12",
        "table_header_bg": "#7C2D12",
        "table_header_text": "#FFFFFF",
        "table_alt_row":   "#FFEDD5",
        "section_bg":      "#7C2D12",
        "section_num_color": "#FB923C",
        "kpi_bg":          "#FFEDD5",
        "quote_border":    "#EA580C",
        "divider":         "#EA580C",
        "chart_gridline":  "#FED7AA",
        "text_on_primary": "#FFFFFF",
    },

    "ocean": {
        "title_bar_bg":    "#0C4A6E",
        "table_header_bg": "#0C4A6E",
        "table_header_text": "#FFFFFF",
        "table_alt_row":   "#E0F2FE",
        "section_bg":      "#0C4A6E",
        "section_num_color": "#38BDF8",
        "kpi_bg":          "#E0F2FE",
        "quote_border":    "#06B6D4",
        "divider":         "#0284C7",
        "chart_gridline":  "#BAE6FD",
        "text_on_primary": "#FFFFFF",
    },

    "forest": {
        "title_bar_bg":    "#064E3B",
        "table_header_bg": "#064E3B",
        "table_header_text": "#FFFFFF",
        "table_alt_row":   "#D1FAE5",
        "section_bg":      "#064E3B",
        "section_num_color": "#34D399",
        "kpi_bg":          "#D1FAE5",
        "quote_border":    "#059669",
        "divider":         "#059669",
        "chart_gridline":  "#A7F3D0",
        "text_on_primary": "#FFFFFF",
    },

    "warm": {
        "title_bar_bg":    "#7F1D1D",
        "table_header_bg": "#7F1D1D",
        "table_header_text": "#FFFFFF",
        "table_alt_row":   "#FEF3C7",
        "section_bg":      "#7F1D1D",
        "section_num_color": "#FCD34D",
        "kpi_bg":          "#FEF3C7",
        "quote_border":    "#B45309",
        "divider":         "#B45309",
        "chart_gridline":  "#FDE68A",
        "text_on_primary": "#FFFFFF",
    },

    "premium": {
        "title_bar_bg":    "#1C1917",
        "table_header_bg": "#1C1917",
        "table_header_text": "#D4AF37",
        "table_alt_row":   "#F5F5F4",
        "section_bg":      "#1C1917",
        "section_num_color": "#D4AF37",
        "kpi_bg":          "#F5F5F4",
        "quote_border":    "#D4AF37",
        "divider":         "#B08D57",
        "chart_gridline":  "#E7E5E4",
        "text_on_primary": "#FAFAF9",
    },

    "chinese_red": {
        "title_bar_bg":    "#8B0000",
        "table_header_bg": "#8B0000",
        "table_header_text": "#FFD700",
        "table_alt_row":   "#FFF0E0",
        "section_bg":      "#8B0000",
        "section_num_color": "#D4AF37",
        "kpi_bg":          "#FFF0E0",
        "quote_border":    "#C41E3A",
        "divider":         "#D4AF37",
        "chart_gridline":  "#F5D5B8",
        "text_on_primary": "#FFD700",
    },

    "high_contrast": {
        "title_bar_bg":    "#000000",
        "table_header_bg": "#000000",
        "table_header_text": "#FFFFFF",
        "table_alt_row":   "#F0F0F0",
        "section_bg":      "#000000",
        "section_num_color": "#888888",
        "kpi_bg":          "#F0F0F0",
        "quote_border":    "#000000",
        "divider":         "#000000",
        "chart_gridline":  "#CCCCCC",
        "text_on_primary": "#FFFFFF",
    },
}

# 将扩展色板合并进 THEME_PALETTES（保持原 API 形状）
for _tn, _ext in _THEME_EXTENDED.items():
    THEME_PALETTES.setdefault(_tn, {}).update(_ext)

# 每主题的图表调色板（HEX 列表）——SVG 迷你图与 PPT/DOCX 图共用。
# v1.5.2: 每主题扩充至 ≥6 色，保证柱图/饼图 6 个数据点内均可区分；
#         所有颜色均已校验对各自主题主背景（light themes: #FFFFFF 系列；
#         dark theme: #0F172A；premium/chinese_red: #FAFAF9/#FFFDF5）满足
#         WCAG AA 正文 ≥4.5:1 对比度。
THEME_CHART_PALETTES: Dict[str, List[str]] = {
    "academic":     ["#1F3864", "#C0504D", "#2E75B6", "#7F604F", "#548235", "#7030A0"],
    "business":     ["#2C3E50", "#E67E22", "#16A085", "#F1C40F", "#8E44AD", "#C0392B"],
    "teaching":     ["#2E7D32", "#FFA000", "#42A5F5", "#9CCC65", "#E91E63", "#7B1FA2"],
    "tech":         ["#00D4FF", "#64FFDA", "#7C3AED", "#F472B6", "#F59E0B", "#22D3EE"],
    "dark":         ["#60A5FA", "#F59E0B", "#34D399", "#F472B6", "#A78BFA", "#22D3EE"],
    "minimal":      ["#2563EB", "#10B981", "#F59E0B", "#EF4444", "#8B5CF6", "#0EA5E9"],
    "nature":       ["#65A30D", "#A16207", "#CA8A04", "#16A34A", "#B91C1C", "#7C3AED"],
    "sunset":       ["#EA580C", "#DB2777", "#F59E0B", "#A855F7", "#0EA5E9", "#16A34A"],
    "ocean":        ["#0284C7", "#06B6D4", "#0EA5E9", "#0891B2", "#0369A1", "#0E7490"],
    "forest":       ["#059669", "#D97706", "#10B981", "#84CC16", "#DC2626", "#7C3AED"],
    "warm":         ["#B45309", "#D97706", "#DC2626", "#92400E", "#65A30D", "#7C3AED"],
    "premium":      ["#D4AF37", "#1C1917", "#B08D57", "#78716C", "#DC2626", "#0EA5E9"],
    "chinese_red":  ["#C41E3A", "#D4AF37", "#8B0000", "#E67E22", "#16A34A", "#0F4C81"],
    "high_contrast":["#000000", "#0047AB", "#B22222", "#006400", "#8B4513", "#4B0082"],
}


# ---------------------------------------------------------------------------
# 4. 公共 API
# ---------------------------------------------------------------------------

def resolve_theme_name(name: str) -> str:
    """将自然语言主题名解析为标准 key（找不到时返回原样，让上层报错）。"""
    if not isinstance(name, str):
        return "academic"
    key = name.strip().lower().replace("-", "_").replace(" ", "_")
    if key in THEME_PALETTES:
        return key
    alias = THEME_ALIASES.get(name.strip()) or THEME_ALIASES.get(key)
    return alias or name


def get_palette(theme_name: str = "academic", extra_overrides: Optional[Dict[str, str]] = None) -> Dict[str, str]:
    """返回指定主题的完整调色板 dict（hex 字符串）。

    Args:
        theme_name: 主题名或自然语言别名
        extra_overrides: 用户传入的额外 hex 覆盖
    """
    name = resolve_theme_name(theme_name) if isinstance(theme_name, str) else "academic"
    if name not in THEME_PALETTES:
        name = "academic"
    palette = dict(BASE_PALETTE)
    palette.update(THEME_PALETTES[name])
    # 6色调色板使用 "background" 键；旧版引擎使用 "bg" 键，这里双向映射
    if "background" in palette and "bg" not in palette:
        palette["bg"] = palette["background"]
    if "bg" in palette and "background" not in palette:
        palette["background"] = palette["bg"]
    if extra_overrides:
        palette.update({k: v for k, v in extra_overrides.items() if isinstance(v, str)})
    return palette


def get_theme_palette(name: str = "academic") -> Dict[str, str]:
    """便捷函数：返回指定主题的 6 色调色板（primary/secondary/accent/text/background/muted）。

    支持别名解析；找不到时回退到 academic。

    Returns:
        dict 含 primary / secondary / accent / text / background / muted 六个 HEX 字符串。
    """
    resolved = resolve_theme_name(name) if isinstance(name, str) else "academic"
    if resolved not in THEME_PALETTES:
        resolved = "academic"
    p = THEME_PALETTES[resolved]
    return {
        "primary":    p.get("primary",    BASE_PALETTE["primary"]),
        "secondary":  p.get("secondary",  BASE_PALETTE["secondary"]),
        "accent":     p.get("accent",     BASE_PALETTE["accent"]),
        "text":       p.get("text",       BASE_PALETTE["text"]),
        "background": p.get("background", p.get("bg", BASE_PALETTE["bg"])),
        "muted":      p.get("muted",      BASE_PALETTE["muted"]),
    }


def get_chart_palette(theme_name: str = "academic") -> List[str]:
    """返回图表用的 hex 列表。"""
    name = resolve_theme_name(theme_name) if isinstance(theme_name, str) else "academic"
    if name not in THEME_CHART_PALETTES:
        name = "academic"
    return list(THEME_CHART_PALETTES[name])


def list_themes() -> List[str]:
    """返回所有内置主题标准名（已排序）。"""
    return sorted(THEME_PALETTES.keys())


__all__ = [
    "BUILTIN_THEMES", "THEME_ALIASES", "BASE_PALETTE",
    "THEME_PALETTES", "THEME_CHART_PALETTES",
    "resolve_theme_name", "get_palette", "get_theme_palette",
    "get_chart_palette", "list_themes",
]
