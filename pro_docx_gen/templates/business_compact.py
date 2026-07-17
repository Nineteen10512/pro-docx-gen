"""business_compact — 商务紧凑报告模板（v1.7.0 新增）。

场景：business_report + 变体：compact_dense
特征：
- 标题行/段落行高压缩（line_spacing 1.15，margin 2.0cm）
- 单行表格（dense info）支持 multi-table on one page
- 提要框（callout）信息密度高
- 适合：周报 / 月报 / 内部简报 / 销售复盘

设计要点：
- 16 模板之一（v1.7.0 把模板数从 16 扩到 18）
- theme_overrides 仅用 #1A1A1A / #1F3864 体系；不引入黑名单色
- page_setup margins 紧到 2.0cm，line_spacing 1.15，最大化"信息密度"
- default_structure 5 段，符合"执行摘要 / 业务回顾 / 数据 / 风险 / 下一步"
  的常见紧凑商务骨架
"""
from ..shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="business_compact",
    display_name="商务报告·紧凑信息",
    scene="business_report",
    variant="compact_dense",
    description=(
        "商务紧凑报告（标题/段落行高压缩 + 单行表格 + 提要框），"
        "适合周报/月报/销售复盘，最大化单页信息密度。"
    ),
    theme_overrides={"color": {"primary": "#1F3864", "accent": "#2E75B6", "heading": "#1A1A1A"}},
    default_structure={
        "toc": False,  # 紧凑周报不强制 TOC
        "sections": [
            {"title": "执行摘要", "level": 1, "content": []},
            {"title": "业务回顾", "level": 1, "content": []},
            {"title": "关键数据", "level": 1, "content": []},
            {"title": "风险与阻塞", "level": 1, "content": []},
            {"title": "下一步计划", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.0, "line_spacing": 1.15},
))
