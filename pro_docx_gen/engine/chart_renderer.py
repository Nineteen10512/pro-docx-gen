"""Chart Renderer — 使用 matplotlib 将 chart 语义节点渲染为高清 PNG 图片。

PaperJSX 原则：LLM 只提供数据与图表类型，不指定颜色/字号/DPI/坐标；
所有样式（色板、字体、网格线、透明度等）均从 design tokens 读取，
与 PPT 技能保持跨文档一致的配色与风格。

支持图表类型：
- column / bar            柱状图 / 条形图
- stacked_column / stacked_bar  堆叠柱状 / 堆叠条形
- line                    折线图（支持 marker）
- area                    面积图（堆积）
- pie / doughnut          饼图 / 环形图
- scatter                 散点图
- radar                   雷达图
"""

from __future__ import annotations

import os

plt = None
rcParams = None
np = None


def _require_matplotlib():
    """Load matplotlib/numpy only when chart rendering is actually requested."""
    global plt, rcParams, np
    if plt is not None and rcParams is not None and np is not None:
        return
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as _plt
        from matplotlib import rcParams as _rcParams
        import numpy as _np
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "DOCX chart rendering requires matplotlib and numpy. "
            "Install requirements.txt or avoid chart nodes in this document."
        ) from exc
    plt = _plt
    rcParams = _rcParams
    np = _np


# ─── 图表类型映射 ─────────────────────────────────────────────────

CHART_TYPE_MAP = {
    "column": "column",
    "bar": "bar",
    "stacked_column": "stacked_column",
    "stacked_bar": "stacked_bar",
    "line": "line",
    "area": "area",
    "pie": "pie",
    "doughnut": "doughnut",
    "scatter": "scatter",
    "radar": "radar",
}

SUPPORTED_CHART_TYPES = set(CHART_TYPE_MAP.keys())


# ─── 工具函数 ────────────────────────────────────────────────────

def _rgb_to_hex(color) -> str:
    """将 docx RGBColor 或 tuple 转为 '#RRGGBB' 形式。"""
    if hasattr(color, "__getitem__"):
        if len(color) == 3:
            return f"#{color[0]:02X}{color[1]:02X}{color[2]:02X}"
    return "#333333"


def _aspect_to_wh(aspect: str, width_in: float) -> tuple[float, float]:
    """将 aspect_ratio 字符串（4:3 / 16:9）换算为 (w, h) 英寸。"""
    if aspect == "16:9":
        return width_in, width_in * 9 / 16
    return width_in, width_in * 3 / 4


def _configure_fonts():
    """设置中文字体与负号。"""
    rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
    rcParams["axes.unicode_minus"] = False


# ─── 主入口 ──────────────────────────────────────────────────────

def render_chart_to_png(
    chart_spec: dict,
    theme_tokens: dict,
    output_path: str,
    dpi: int | None = None,
) -> str:
    """将 chart 语义节点渲染为 PNG 文件。

    Args:
        chart_spec: 由 parser 展开的 chart 扁平节点。
        theme_tokens: 完整 tokens dict（取 chart.* 下的样式）。
        output_path: 输出 PNG 路径。
        dpi: 输出 DPI，缺省从 tokens["chart"]["dpi"] 取（默认 300）。

    Returns:
        输出文件路径（同 output_path）。
    """
    _require_matplotlib()
    _configure_fonts()

    chart_tok = theme_tokens.get("chart", {})
    color_tok = theme_tokens.get("color", {})
    font_tok = theme_tokens.get("font", {})

    palette = [_rgb_to_hex(c) for c in chart_tok.get("palette", [])]
    # v1.5.2: ensure at least 6 colors; otherwise fallback to academic 6-color palette
    if len(palette) < 6:
        palette = ["#1F3864", "#C0504D", "#2E75B6", "#7F604F", "#548235", "#7030A0"]

    grid_color = _rgb_to_hex(chart_tok.get("gridline_color", color_tok.get("muted", (0xDD, 0xDD, 0xDD))))
    text_color = _rgb_to_hex(chart_tok.get("text_color", color_tok.get("text", (0x33, 0x33, 0x33))))
    grid_alpha = chart_tok.get("gridline_alpha", 0.5)
    fill_alpha = chart_tok.get("alpha", 0.85)
    line_width = chart_tok.get("line_width", 2.0)
    bar_width = chart_tok.get("bar_width", 0.6)
    marker_size = chart_tok.get("marker_size", 5)

    axis_size = font_tok["size"].get("chart_axis").pt if hasattr(font_tok["size"].get("chart_axis"), "pt") else 9
    legend_size = font_tok["size"].get("chart_legend").pt if hasattr(font_tok["size"].get("chart_legend"), "pt") else 10
    title_size = font_tok["size"].get("chart_title").pt if hasattr(font_tok["size"].get("chart_title"), "pt") else 11

    # 兼容 Pt 对象
    def _pt(x, default=10):
        return x.pt if hasattr(x, "pt") else float(x) if x is not None else default

    axis_size = _pt(font_tok["size"].get("chart_axis"), 9)
    legend_size = _pt(font_tok["size"].get("chart_legend"), 10)
    title_size = _pt(font_tok["size"].get("chart_title"), 11)

    chart_type = chart_spec.get("chart_type", "column")
    title = chart_spec.get("title")
    categories = chart_spec.get("categories", [])
    series = chart_spec.get("series", [])
    show_legend = chart_spec.get("show_legend", True)
    legend_position = chart_spec.get("legend_position", "bottom")
    show_data_labels = chart_spec.get("show_data_labels")
    number_format = chart_spec.get("number_format", "0")
    x_title = chart_spec.get("x_title")
    y_title = chart_spec.get("y_title")
    aspect_ratio = chart_spec.get("aspect_ratio", chart_tok.get("default_aspect", "4:3"))
    width_pct = chart_spec.get("width_pct", chart_tok.get("default_width_pct", 1.0))
    align = chart_spec.get("align", "center")

    # 饼图默认显示数据标签（百分比）
    if show_data_labels is None:
        show_data_labels = chart_type in ("pie", "doughnut")

    # 计算图片尺寸（英寸）：正文宽度约 6.27"（A4 8.27" - 2*1" margin）
    content_width_in = 6.27
    img_w = content_width_in * float(width_pct)
    img_w, img_h = _aspect_to_wh(aspect_ratio, img_w)

    # 饼图/雷达图用方形画布更好看
    if chart_type in ("pie", "doughnut", "radar"):
        img_h = img_w

    fig, ax = plt.subplots(figsize=(img_w, img_h), dpi=dpi or chart_tok.get("dpi", 300))
    # 透明背景
    fig.patch.set_alpha(0)
    ax.patch.set_alpha(0)

    # ─── 分发具体图表绘制 ────────────────────────────────────────
    if chart_type in ("column", "stacked_column"):
        _draw_column(ax, categories, series, palette, chart_type == "stacked_column",
                     bar_width, fill_alpha, show_data_labels, text_color, axis_size)
        _apply_axis_style(ax, grid_color, grid_alpha, text_color, axis_size,
                          x_title, y_title)
    elif chart_type in ("bar", "stacked_bar"):
        _draw_bar(ax, categories, series, palette, chart_type == "stacked_bar",
                  bar_width, fill_alpha, show_data_labels, text_color, axis_size)
        _apply_axis_style(ax, grid_color, grid_alpha, text_color, axis_size,
                          x_title, y_title, horizontal=True)
    elif chart_type == "line":
        _draw_line(ax, categories, series, palette, line_width, marker_size,
                   fill_alpha, show_data_labels, text_color, axis_size)
        _apply_axis_style(ax, grid_color, grid_alpha, text_color, axis_size,
                          x_title, y_title)
    elif chart_type == "area":
        _draw_area(ax, categories, series, palette, fill_alpha,
                   show_data_labels, text_color, axis_size)
        _apply_axis_style(ax, grid_color, grid_alpha, text_color, axis_size,
                          x_title, y_title)
    elif chart_type in ("pie", "doughnut"):
        _draw_pie(ax, categories, series, palette, chart_type == "doughnut",
                  show_data_labels, text_color, legend_size)
    elif chart_type == "scatter":
        _draw_scatter(ax, series, palette, marker_size, fill_alpha,
                      show_data_labels, text_color, axis_size)
        _apply_axis_style(ax, grid_color, grid_alpha, text_color, axis_size,
                          x_title, y_title)
    elif chart_type == "radar":
        # 雷达图需要 polar 坐标
        plt.close(fig)
        fig = plt.figure(figsize=(img_w, img_h), dpi=dpi or chart_tok.get("dpi", 300))
        fig.patch.set_alpha(0)
        ax = fig.add_subplot(111, polar=True)
        ax.patch.set_alpha(0)
        _draw_radar(ax, categories, series, palette, fill_alpha, line_width,
                    text_color, axis_size, grid_color, grid_alpha)
    else:
        # 兜底画柱状图
        _draw_column(ax, categories, series, palette, False,
                     bar_width, fill_alpha, show_data_labels, text_color, axis_size)
        _apply_axis_style(ax, grid_color, grid_alpha, text_color, axis_size,
                          x_title, y_title)

    # 标题（位于图片上方，作为图片的一部分）
    if title and chart_type not in ("pie", "doughnut") or (title and chart_type in ("pie", "doughnut") and not show_legend):
        ax.set_title(title, fontsize=title_size, color=text_color, pad=12)

    # 图例
    if show_legend and chart_type not in ("pie", "doughnut"):
        _apply_legend(ax, legend_position, legend_size, text_color)
    elif show_legend and chart_type in ("pie", "doughnut"):
        # 饼图图例由 _draw_pie 内部处理
        pass

    # 保存
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    fig.savefig(output_path, dpi=dpi or chart_tok.get("dpi", 300),
                bbox_inches="tight", transparent=True, pad_inches=0.1)
    plt.close(fig)
    return output_path


# ─── 各类图表绘制 ────────────────────────────────────────────────

def _draw_column(ax, categories, series, palette, stacked, bar_width,
                 alpha, show_labels, text_color, axis_size):
    x = np.arange(len(categories))
    n = len(series)
    # v1.5.2: single-series column chart uses palette cycle per category
    multi_series = n >= 2
    if stacked:
        bottom = np.zeros(len(categories))
        for i, s in enumerate(series):
            vals = np.array(s.get("values", []), dtype=float)
            color = palette[i % len(palette)]
            bars = ax.bar(x, vals, bar_width, bottom=bottom, label=s.get("name", f"Series{i+1}"),
                          color=color, alpha=alpha, edgecolor="white", linewidth=0.5)
            if show_labels:
                _add_bar_labels(ax, bars, vals, bottom, text_color, axis_size)
            bottom += vals
    else:
        total_w = bar_width
        bw = total_w / max(n, 1)
        for i, s in enumerate(series):
            vals = np.array(s.get("values", []), dtype=float)
            offset = (i - (n - 1) / 2) * bw
            if multi_series:
                color = palette[i % len(palette)]
                bars = ax.bar(x + offset, vals, bw, label=s.get("name", f"Series{i+1}"),
                              color=color, alpha=alpha, edgecolor="white", linewidth=0.5)
            else:
                # v1.5.2: single-series multi-category → cycle palette per category
                bar_colors = [palette[j % len(palette)] for j in range(len(vals))]
                bars = ax.bar(x + offset, vals, bw, label=s.get("name", f"Series{i+1}"),
                              color=bar_colors, alpha=alpha, edgecolor="white", linewidth=0.5)
            if show_labels:
                _add_bar_labels(ax, bars, vals, 0, text_color, axis_size)
    ax.set_xticks(x)
    ax.set_xticklabels(categories, fontsize=axis_size, color=text_color)


def _draw_bar(ax, categories, series, palette, stacked, bar_width,
              alpha, show_labels, text_color, axis_size):
    y = np.arange(len(categories))
    n = len(series)
    multi_series = n >= 2
    if stacked:
        left = np.zeros(len(categories))
        for i, s in enumerate(series):
            vals = np.array(s.get("values", []), dtype=float)
            color = palette[i % len(palette)]
            bars = ax.barh(y, vals, bar_width, left=left, label=s.get("name", f"Series{i+1}"),
                           color=color, alpha=alpha, edgecolor="white", linewidth=0.5)
            if show_labels:
                _add_barh_labels(ax, bars, vals, left, text_color, axis_size)
            left += vals
    else:
        total_h = bar_width
        bh = total_h / max(n, 1)
        for i, s in enumerate(series):
            vals = np.array(s.get("values", []), dtype=float)
            offset = (i - (n - 1) / 2) * bh
            if multi_series:
                color = palette[i % len(palette)]
                bars = ax.barh(y + offset, vals, bh, label=s.get("name", f"Series{i+1}"),
                               color=color, alpha=alpha, edgecolor="white", linewidth=0.5)
            else:
                bar_colors = [palette[j % len(palette)] for j in range(len(vals))]
                bars = ax.barh(y + offset, vals, bh, label=s.get("name", f"Series{i+1}"),
                               color=bar_colors, alpha=alpha, edgecolor="white", linewidth=0.5)
            if show_labels:
                _add_barh_labels(ax, bars, vals, 0, text_color, axis_size)
    ax.set_yticks(y)
    ax.set_yticklabels(categories, fontsize=axis_size, color=text_color)


def _draw_line(ax, categories, series, palette, line_width, marker_size,
               alpha, show_labels, text_color, axis_size):
    x = np.arange(len(categories))
    for i, s in enumerate(series):
        vals = np.array(s.get("values", []), dtype=float)
        color = palette[i % len(palette)]
        markers = s.get("markers", True)
        ax.plot(x, vals, color=color, linewidth=line_width,
                marker="o" if markers else None, markersize=marker_size,
                markerfacecolor=color, markeredgecolor="white",
                label=s.get("name", f"Series{i+1}"), alpha=alpha)
        if show_labels:
            for xi, vi in zip(x, vals):
                ax.annotate(f"{vi:g}", (xi, vi), textcoords="offset points",
                            xytext=(0, 6), ha="center", fontsize=axis_size - 1,
                            color=text_color)
    ax.set_xticks(x)
    ax.set_xticklabels(categories, fontsize=axis_size, color=text_color)


def _draw_area(ax, categories, series, palette, alpha, show_labels,
               text_color, axis_size):
    x = np.arange(len(categories))
    stack = []
    for s in series:
        stack.append(np.array(s.get("values", []), dtype=float))
    if stack:
        ax.stackplot(x, *stack, labels=[s.get("name", f"Series{i+1}") for i, s in enumerate(series)],
                     colors=[palette[i % len(palette)] for i in range(len(series))],
                     alpha=alpha, edgecolor="white", linewidth=0.5)
    ax.set_xticks(x)
    ax.set_xticklabels(categories, fontsize=axis_size, color=text_color)


def _draw_pie(ax, categories, series, palette, is_doughnut,
              show_labels, text_color, legend_size):
    # 饼图默认取第一个 series
    if not series:
        return
    s = series[0]
    vals = np.array(s.get("values", []), dtype=float)
    # 标签：若提供 categories 且长度匹配则用 categories，否则用 series name
    labels = categories if len(categories) == len(vals) else [f"{v:g}" for v in vals]

    wedgeprops = {"width": 0.45, "edgecolor": "white", "linewidth": 1.5} if is_doughnut else \
                 {"edgecolor": "white", "linewidth": 1.0}

    if show_labels:
        wedges, texts, autotexts = ax.pie(
            vals, labels=labels, colors=[palette[i % len(palette)] for i in range(len(vals))],
            autopct="%1.1f%%", startangle=90, wedgeprops=wedgeprops,
            textprops={"fontsize": legend_size - 1, "color": text_color},
        )
        for at in autotexts:
            at.set_color("white")
            at.set_fontsize(legend_size - 1)
    else:
        ax.pie(
            vals, labels=labels, colors=[palette[i % len(palette)] for i in range(len(vals))],
            startangle=90, wedgeprops=wedgeprops,
            textprops={"fontsize": legend_size - 1, "color": text_color},
        )
    ax.axis("equal")


def _draw_scatter(ax, series, palette, marker_size, alpha,
                  show_labels, text_color, axis_size):
    for i, s in enumerate(series):
        xs = np.array(s.get("x", s.get("values", [])), dtype=float)
        ys = np.array(s.get("y", s.get("y_values", [])), dtype=float)
        color = palette[i % len(palette)]
        ax.scatter(xs, ys, color=color, s=(marker_size * 6), alpha=alpha,
                   edgecolors="white", linewidths=0.5,
                   label=s.get("name", f"Series{i+1}"))
        if show_labels:
            for xv, yv in zip(xs, ys):
                ax.annotate(f"({xv:g},{yv:g})", (xv, yv), textcoords="offset points",
                            xytext=(4, 4), fontsize=axis_size - 1, color=text_color)


def _draw_radar(ax, categories, series, palette, alpha, line_width,
                text_color, axis_size, grid_color, grid_alpha):
    n = len(categories)
    angles = np.linspace(0, 2 * np.pi, n, endpoint=False).tolist()
    angles += angles[:1]
    for i, s in enumerate(series):
        vals = list(s.get("values", []))
        vals += vals[:1]
        color = palette[i % len(palette)]
        ax.plot(angles, vals, color=color, linewidth=line_width,
                label=s.get("name", f"Series{i+1}"))
        ax.fill(angles, vals, color=color, alpha=alpha * 0.5)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontsize=axis_size, color=text_color)
    ax.tick_params(colors=text_color, labelsize=axis_size - 1)
    ax.grid(color=grid_color, alpha=grid_alpha, linestyle="--")
    ax.spines["polar"].set_color(grid_color)


# ─── 坐标轴/图例/网格统一样式 ───────────────────────────────────

def _apply_axis_style(ax, grid_color, grid_alpha, text_color, axis_size,
                      x_title=None, y_title=None, horizontal=False):
    ax.tick_params(axis="both", labelsize=axis_size, colors=text_color,
                   direction="out", length=3)
    # 网格线
    if horizontal:
        ax.yaxis.grid(False)
        ax.xaxis.grid(True, linestyle="--", color=grid_color, alpha=grid_alpha)
    else:
        ax.xaxis.grid(False)
        ax.yaxis.grid(True, linestyle="--", color=grid_color, alpha=grid_alpha)
    # 隐藏上/右边框
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.spines["left"].set_color(grid_color)
    ax.spines["bottom"].set_color(grid_color)
    # 轴标题
    if x_title:
        ax.set_xlabel(x_title, fontsize=axis_size, color=text_color)
    if y_title:
        ax.set_ylabel(y_title, fontsize=axis_size, color=text_color)


def _apply_legend(ax, position, legend_size, text_color):
    loc_map = {
        "bottom": "lower center",
        "top": "upper center",
        "left": "center left",
        "right": "center right",
    }
    loc = loc_map.get(position, "lower center")
    if position == "bottom":
        bbox = (0.5, -0.28)
        ncol = min(len(ax.get_legend_handles_labels()[1]), 4) or 1
    elif position == "top":
        bbox = (0.5, 1.12)
        ncol = min(len(ax.get_legend_handles_labels()[1]), 4) or 1
    else:
        bbox = None
        ncol = 1
    leg = ax.legend(loc=loc, bbox_to_anchor=bbox, ncol=ncol,
                    fontsize=legend_size, frameon=False)
    for text in leg.get_texts():
        text.set_color(text_color)


def _add_bar_labels(ax, bars, vals, bottom, text_color, axis_size):
    for bar, v, b in zip(bars, vals, np.atleast_1d(bottom)):
        ax.annotate(f"{v:g}", xy=(bar.get_x() + bar.get_width() / 2, b + v),
                    xytext=(0, 3), textcoords="offset points",
                    ha="center", va="bottom",
                    fontsize=axis_size - 1, color=text_color)


def _add_barh_labels(ax, bars, vals, left, text_color, axis_size):
    for bar, v, l in zip(bars, vals, np.atleast_1d(left)):
        ax.annotate(f"{v:g}", xy=(l + v, bar.get_y() + bar.get_height() / 2),
                    xytext=(3, 0), textcoords="offset points",
                    ha="left", va="center",
                    fontsize=axis_size - 1, color=text_color)
