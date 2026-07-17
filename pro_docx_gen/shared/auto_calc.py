"""v1.7.0 表格自动计算工具 — YoY 增长 / 汇总行 helper。

为研报/财务/业务表格提供"算"的能力。模板和 LLM 仍写语义数据；本模块负责
最后一步的算术派生（避免 LLM 算错）。同时给 ``renderer.auto_compute_rows``
提供底层算子。

设计约束：
- **修根因，不掩盖**：任何输入异常（空、非数值、零分母）按规则降级到
  原值/0/None，并在 return 时显式标记（``meta`` 字段），调用方可读。
- **零依赖**：纯 stdlib，``renderer`` / 模板可独立 import。
- **可单测**：所有 helper 是纯函数（无 IO / 无 token 依赖）。

API：
- ``compute_yoy(rows, col_index, *, previous_rows=None, fmt='{:.1%}')``
  -> ``(new_rows, meta)``，每行新增一列 YoY% 文本。
  - 默认上一期 rows 是 ``rows`` 自己（only if N=1）；更一般用法是显式传 previous_rows
  - 同比 > 0 → "10.5%"，< 0 → "-3.2%"，0 → "0.0%"
  - 上一期基线为 0 时 → "—" 文本 + meta["zero_base_count"]
- ``compute_summary(rows, *, mode='sum'|'avg', label='合计')``
  -> ``(summary_row, meta)``，返回单行汇总（与 rows 等列数）。
  - 非数值列 → 跳过聚合（不抛错）；仅汇总数值列。
  - 文本列保留 label（"合计"），数值列返回 sum / avg。
- ``auto_compute_rows(rows, *, yoy_col_index=None, summary_mode=None, summary_label='合计')``
  -> ``(new_rows, meta)``：组合上面两个。
  - 先 append 汇总行（如果有），再按 rows+summary 计算 YoY。

@since v1.7.0
"""
from __future__ import annotations

from typing import Iterable, List, Optional, Sequence, Tuple, Union

Number = Union[int, float]
Row = Sequence[Union[Number, str, None]]


# ---------------------------------------------------------------------------
# 内部工具
# ---------------------------------------------------------------------------


def _coerce_number(value) -> Optional[Number]:
    """Best-effort 把字符串/数值转 float。失败返回 None（不算数值）。"""
    if value is None:
        return None
    if isinstance(value, bool):
        # bool 是 int 的子类，但要单独挡掉（True/False 不是 0/1 的研究语义）
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.strip().replace(",", "").replace("，", "")
        # 去百分号 — 这是 8% 形式的字符串
        if s.endswith("%"):
            s = s[:-1]
        # 去前导 ¥/$ 等货币符号
        for prefix in ("¥", "$", "€", "£", "￥"):
            if s.startswith(prefix):
                s = s[len(prefix):]
                break
        try:
            return float(s)
        except ValueError:
            return None
    return None


def _format_pct(ratio: float, fmt: str = "{:.1%}") -> str:
    """Format a ratio as a percentage string. 0.123 -> "12.3%"."""
    try:
        return fmt.format(ratio)
    except Exception:
        # fmt 错误时降级
        return f"{ratio * 100:.1f}%"


# ---------------------------------------------------------------------------
# YoY 同比
# ---------------------------------------------------------------------------


def compute_yoy(
    rows: Sequence[Row],
    col_index: int,
    *,
    previous_rows: Optional[Sequence[Row]] = None,
    fmt: str = "{:.1%}",
    include_baseline: bool = False,
) -> Tuple[List[List[str]], dict]:
    """Append a YoY% column to each row at ``col_index``.

    Args:
        rows: current-period data rows. Each row is a sequence whose
            ``col_index``-th element is a number (or numeric string).
        col_index: the column index in ``rows`` to compute YoY for.
        previous_rows: matching rows from the previous period. If None,
            ``rows`` itself is used (useful for N=1 stress tests).
        fmt: percent format string; default ``"{:.1%}"``.
        include_baseline: if True, the returned row layout is
            ``[..., baseline_value, yoy_pct]``; else ``[..., yoy_pct]``.

    Returns:
        ``(new_rows, meta)``:
        - ``new_rows`` is a list of lists of strings, each with the YoY column
          appended (or inserted at the end of the row).
        - ``meta`` is a dict with counters: ``{"computed": int, "zero_base": int,
          "skipped": int}``.

    Behaviour matrix:
        - cur > 0, prev > 0     → pct text
        - cur > 0, prev == 0    → "—" (zero base), meta.zero_base += 1
        - cur < 0, prev < 0     → pct text (sign kept, e.g. "-15.0%")
        - cur == 0, prev == 0   → "0.0%"
        - non-numeric either    → "—" (skipped), meta.skipped += 1
    """
    if not rows:
        return [], {"computed": 0, "zero_base": 0, "skipped": 0, "rows_in": 0, "rows_out": 0}

    base = previous_rows if previous_rows is not None else rows
    new_rows: List[List[str]] = []
    meta = {"computed": 0, "zero_base": 0, "skipped": 0, "rows_in": len(rows), "rows_out": 0}

    for i, row in enumerate(rows):
        row_list = [str(v) if v is not None else "" for v in row]
        cur_raw = row[col_index] if col_index < len(row) else None
        prev_raw = base[i][col_index] if i < len(base) and col_index < len(base[i]) else None
        cur = _coerce_number(cur_raw)
        prev = _coerce_number(prev_raw)

        pct_text = "—"
        if cur is None or prev is None:
            meta["skipped"] += 1
        elif prev == 0:
            if cur == 0:
                pct_text = "0.0%"
                meta["computed"] += 1
            else:
                pct_text = "—"
                meta["zero_base"] += 1
        else:
            ratio = (cur - prev) / abs(prev)
            pct_text = _format_pct(ratio, fmt)
            meta["computed"] += 1

        if include_baseline:
            row_list.append(str(cur_raw) if cur_raw is not None else "")
        row_list.append(pct_text)
        new_rows.append(row_list)
        meta["rows_out"] += 1
    return new_rows, meta


# ---------------------------------------------------------------------------
# 汇总行
# ---------------------------------------------------------------------------


def compute_summary(
    rows: Sequence[Row],
    *,
    mode: str = "sum",
    label: str = "合计",
    label_col: int = 0,
) -> Tuple[List[str], dict]:
    """Compute a summary row by aggregating numeric columns.

    Args:
        rows: data rows. Text-only columns are passed through unchanged
            (using the value from the **first** row by default, except
            ``label_col`` which uses ``label``).
        mode: ``"sum"`` or ``"avg"``.
        label: label text for the ``label_col`` (default: 合计).
        label_col: column index where ``label`` is placed (default 0).

    Returns:
        ``(summary_row, meta)``:
        - ``summary_row`` is a list of strings, same length as the widest
          row in ``rows``.
        - ``meta`` is a dict: ``{"mode": str, "n_rows": int, "n_numeric_cols": int,
          "skipped_cols": list[int]}``.
    """
    meta = {
        "mode": mode,
        "n_rows": len(rows),
        "n_numeric_cols": 0,
        "skipped_cols": [],
    }
    if mode not in ("sum", "avg"):
        raise ValueError(f"compute_summary mode must be 'sum' or 'avg', got {mode!r}")
    if not rows:
        return [label], meta

    width = max(len(r) for r in rows)
    summary: List[str] = ["" for _ in range(width)]

    # 默认值：除 label_col 外，每列沿用第一行的字符串表示
    first = rows[0]
    for ci in range(width):
        if ci == label_col:
            summary[ci] = label
        else:
            v = first[ci] if ci < len(first) else None
            summary[ci] = "" if v is None else str(v)

    # 数值列聚合
    skipped_cols: List[int] = []
    for ci in range(width):
        if ci == label_col:
            continue
        values: List[float] = []
        is_numeric_col = True
        for r in rows:
            v = r[ci] if ci < len(r) else None
            n = _coerce_number(v)
            if n is None and v is not None and str(v).strip() != "":
                # 此列含非数值（但有内容），整列按非数值处理
                is_numeric_col = False
                break
            if n is not None:
                values.append(n)
        if not is_numeric_col or not values:
            skipped_cols.append(ci)
            continue
        agg = sum(values) if mode == "sum" else (sum(values) / len(values))
        # 数值呈现：整数若无小数则不带 .0
        if abs(agg - round(agg)) < 1e-9 and abs(agg) < 1e15:
            summary[ci] = f"{int(round(agg)):,}"
        else:
            summary[ci] = f"{agg:,.2f}"
        meta["n_numeric_cols"] += 1

    meta["skipped_cols"] = skipped_cols
    return summary, meta


# ---------------------------------------------------------------------------
# 组合：auto_compute_rows
# ---------------------------------------------------------------------------


def auto_compute_rows(
    rows: Sequence[Row],
    *,
    yoy_col_index: Optional[int] = None,
    summary_mode: Optional[str] = None,
    summary_label: str = "合计",
    yoy_fmt: str = "{:.1%}",
    yoy_previous_rows: Optional[Sequence[Row]] = None,
    yoy_insert_baseline: bool = False,
) -> Tuple[List[List[str]], dict]:
    """Convenience: append a summary row and/or a YoY column in one call.

    Args:
        rows: data rows.
        yoy_col_index: if set, append a YoY% column based on this column.
        summary_mode: ``"sum"``, ``"avg"``, or None. If set, the summary row is
            **prepended** so it appears as the first row of the table (matches
            Chinese financial-statement convention "合计 / 汇总在上").
        summary_label: text for the summary row's label cell.
        yoy_fmt: percent format string.
        yoy_previous_rows: passed through to :func:`compute_yoy`.
        yoy_insert_baseline: if True, the YoY row also keeps the raw baseline
            value beside the pct text.

    Returns:
        ``(new_rows, meta)``: see :func:`compute_yoy` / :func:`compute_summary`
        for the meta contents. The combined meta is a flat dict with these
        additional keys: ``"summary_mode": str|None``, ``"yoy_col_index": int|None``.
    """
    combined_meta: dict = {
        "summary_mode": summary_mode,
        "yoy_col_index": yoy_col_index,
    }

    working_rows: List[Row] = [list(r) for r in rows]
    # 1) Summary row first (if requested) — operates on input rows only.
    if summary_mode:
        summary_row, sum_meta = compute_summary(
            working_rows, mode=summary_mode, label=summary_label
        )
        combined_meta.update({"summary": sum_meta})
        working_rows.insert(0, summary_row)
    else:
        combined_meta["summary"] = None

    # 2) YoY column (if requested) — operates on the rows *after* summary insertion
    #    so the summary row itself also gets a YoY value if it has a numeric cell.
    if yoy_col_index is not None:
        yoy_rows, yoy_meta = compute_yoy(
            working_rows,
            yoy_col_index,
            previous_rows=yoy_previous_rows,
            fmt=yoy_fmt,
            include_baseline=yoy_insert_baseline,
        )
        combined_meta.update({"yoy": yoy_meta})
        return yoy_rows, combined_meta

    # 没有 yoy 也要把 row 列表转成 str
    return [[str(v) if v is not None else "" for v in r] for r in working_rows], combined_meta


# ---------------------------------------------------------------------------
# 颜色查找（v1.7.0 配套）
# ---------------------------------------------------------------------------


# v1.7.0：研报同比染色常量。**故意**不进 FORBIDDEN_HEADING_HEXES（这些不是
# heading 色，是数据 run 色）。self_audit 只检查 heading 相关字段，run-level
# 使用这些色不会被审计拦截。
YOY_COLOR_POSITIVE = "0E7C3A"   # 正增长（深绿，#0E7C3A）
YOY_COLOR_NEGATIVE = "B91C1C"   # 负增长（深红，#B91C1C）
YOY_COLOR_ZERO     = "666666"   # 零或不变（muted 灰，与 TEXT 区分保留视觉层级）
YOY_COLOR_NEUTRAL  = "666666"   # 不可计算（"—"，与 zero 共色，但加斜体更明显）


def yoy_color_for(value: Union[Number, str, None]) -> str:
    """Return the run-color hex (no leading ``#``) for a YoY-style value.

    Args:
        value: numeric YoY ratio (e.g. 0.105 for 10.5%), or a pre-formatted
            pct string (e.g. "10.5%" / "-3.2%"), or "—".

    Returns:
        One of ``YOY_COLOR_POSITIVE / YOY_COLOR_NEGATIVE / YOY_COLOR_ZERO /
        YOY_COLOR_NEUTRAL`` (upper-case hex, no ``#``).
    """
    if value is None:
        return YOY_COLOR_NEUTRAL
    n = _coerce_number(value)
    if n is None:
        # value 可能是预格式化 pct 字符串 / "—"
        s = str(value).strip()
        if s in ("—", "-", "—", "n/a", "N/A"):
            return YOY_COLOR_NEUTRAL
        # 预格式化 pct 字符串
        if s.endswith("%"):
            try:
                pct = float(s[:-1].replace(",", ""))
                if pct > 0:
                    return YOY_COLOR_POSITIVE
                if pct < 0:
                    return YOY_COLOR_NEGATIVE
                return YOY_COLOR_ZERO
            except ValueError:
                return YOY_COLOR_NEUTRAL
        return YOY_COLOR_NEUTRAL
    if n > 0:
        return YOY_COLOR_POSITIVE
    if n < 0:
        return YOY_COLOR_NEGATIVE
    return YOY_COLOR_ZERO


__all__ = [
    "compute_yoy",
    "compute_summary",
    "auto_compute_rows",
    "yoy_color_for",
    "YOY_COLOR_POSITIVE",
    "YOY_COLOR_NEGATIVE",
    "YOY_COLOR_ZERO",
    "YOY_COLOR_NEUTRAL",
]
