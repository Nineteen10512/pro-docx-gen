"""research_report — 研报专用模板（v1.7.0 新增）。

场景：equity_research + 变体：institutional
特征：
- 4 字段强制 header：股票代码 / 评级 / 目标价 / 日期
- 评级三色徽章（Overweight=绿 / Hold=黄 / Underweight=红）
- 紫红副标题（Bloomberg / Reuters 风格）+ 深蓝主色
- KPI 卡片行（公司财务 / 估值 / 业绩速览）
- 风险提示固定尾部（披露声明）
- 财务 / 估值 / 业务拆分表格的 YoY + 汇总自动计算

设计要点：
- 16 模板之一（v1.7.0 把模板数从 16 扩到 18）
- theme_overrides 仅用 #1F3864 (navy) / #6B2C91 (purple) / #0E7C3A (绿评级)
  / #B91C1C (红评级) / #B45309 (黄评级) — **全部不在 self_audit 黑名单**
- 评级三色仅用于 cell-shading / run-color；不进 heading 字段
- 风险提示走 muted (#666666) 灰斜体小字
- default_structure 6 段（公司概览 / 投资摘要 / 行业分析 / 财务 / 估值 / 风险）
"""
from ..shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="research_report",
    display_name="研报·机构标准",
    scene="equity_research",
    variant="institutional",
    description=(
        "研报专用（深蓝+紫红强调，4 字段研报 header：股票代码/评级/目标价/日期；"
        "KPI 卡片行；风险提示固定尾部；表格 YoY/汇总自动计算）。"
    ),
    theme_overrides={
        "color": {
            "primary": "#1F3864",    # 深蓝（继承 academic）
            "accent": "#6B2C91",     # 紫红副标题
            "heading": "#1A1A1A",    # 标题色（黑名单外）
        },
    },
    default_structure={
        "toc": True,
        "sections": [
            {"title": "公司概览", "level": 1, "content": []},
            {"title": "投资摘要", "level": 1, "content": []},
            {"title": "行业分析", "level": 1, "content": []},
            {"title": "财务分析", "level": 1, "content": []},
            {"title": "估值与目标价", "level": 1, "content": []},
            {"title": "风险因素", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.0, "line_spacing": 1.3},
))
