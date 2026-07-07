"""Layout — 确定性布局计算器。

负责所有需要确定性计算的尺寸问题：
- 内容区宽度 = page_width - 左/右边距（v1.2 支持四边独立边距）
- 表格列宽自适应（按内容长度加权，每列不小于最小宽度）
- 图片尺寸约束（保持宽高比，不超过内容区宽度）
- 列表缩进层级距离
"""

from typing import Optional


class LayoutCalculator:
    """基于 tokens 的布局计算器。"""

    def __init__(self, tokens: dict):
        self.tokens = tokens
        sp = tokens["spacing"]
        page = tokens.get("page", {})
        self.page_width = sp["page_width"]
        self.page_height = sp["page_height"]
        # v1.2: 四边独立边距，若未设置回退到 page_margin（保持旧行为）
        self.margin_top = page.get("margin_top", sp["page_margin"])
        self.margin_bottom = page.get("margin_bottom", sp["page_margin"])
        self.margin_left = page.get("margin_left", sp["page_margin"])
        self.margin_right = page.get("margin_right", sp["page_margin"])
        self.gutter = page.get("gutter", 0) or 0
        self.margin = sp["page_margin"]  # 兼容旧引用
        self.content_width = self.page_width - self.margin_left - self.margin_right - self.gutter
        self.list_indent = sp["list_indent"]
        self.quote_indent = sp["quote_indent"]
        self.code_indent = sp["code_indent"]
        self.first_line_indent = sp["first_line_indent"]
        self.ref_hanging = tokens["reference"]["hanging_indent"]
        self.table_col_min = tokens["table"]["col_min_width"]

    # ─── 表格列宽 ──────────────────────────────────────────────────

    def compute_table_col_widths(
        self,
        headers: list[str],
        rows: list[list[str]],
        col_widths: Optional[list[float]] = None,
    ) -> list:
        """计算表格列宽（返回 docx.shared.Length 对象列表）。"""
        from docx.shared import Inches

        n = len(headers)

        if col_widths and len(col_widths) == n:
            widths = []
            total = 0
            for w in col_widths:
                w_in = Inches(w)
                widths.append(w_in)
                total += w_in
            if total > self.content_width:
                scale = self.content_width / total
                widths = [w * scale for w in widths]
            return widths

        col_lengths = [len(str(h)) for h in headers]
        for row in rows:
            for i, cell in enumerate(row):
                col_lengths[i] = max(col_lengths[i], len(str(cell)))

        min_total = self.table_col_min * n
        if min_total > self.content_width:
            return [self.content_width / n] * n

        remaining = self.content_width - min_total
        total_len = sum(col_lengths) or 1
        extra_each = [cl / total_len * remaining for cl in col_lengths]

        return [self.table_col_min + Inches(e.emu / 914400) if hasattr(e, 'emu')
                else self.table_col_min + e for e in extra_each]

    # ─── 图片尺寸 ──────────────────────────────────────────────────

    def compute_image_size(
        self,
        native_width,
        native_height,
        width_inches: Optional[float] = None,
    ):
        """根据原始尺寸和约束计算最终宽高，保持宽高比。"""
        from docx.shared import Inches

        max_w = self.content_width
        if width_inches is not None:
            target_w = Inches(width_inches)
            if target_w > max_w:
                target_w = max_w
        else:
            target_w = max_w

        if native_width is None or native_height is None or native_width == 0:
            return target_w, None

        ratio = native_height / native_width
        target_h = int(target_w * ratio)
        return target_w, target_h

    # ─── 列表缩进 ──────────────────────────────────────────────────

    def list_indent_for_level(self, level: int):
        """返回指定层级的列表项左缩进。"""
        from docx.shared import Inches
        base = Inches(0.25)
        return base + self.list_indent * level

    # ─── KPI 卡片列宽 ──────────────────────────────────────────────

    def kpi_col_widths(self, n_cards: int):
        return [self.content_width / n_cards] * n_cards
