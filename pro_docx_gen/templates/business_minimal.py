"""极简商务报告模板 — v1.6.0 variant 模板。

场景：business_report + 变体：minimal_clean
"""
from ..shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="business_minimal",
    display_name="商务报告·极简干净",
    scene="business_report",
    variant="minimal_clean",
    description="极简风格商务报告，大留白，低饱和配色",
    theme_overrides={"color": {"primary": "#1A1A1A", "heading": "#333333"}},
    default_structure={
        "toc": True,
        "sections": [
            {"title": "执行摘要", "level": 1, "content": []},
            {"title": "市场分析", "level": 1, "content": []},
            {"title": "业务回顾", "level": 1, "content": []},
            {"title": "财务数据", "level": 1, "content": []},
            {"title": "下一步计划", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.8, "line_spacing": 1.5},
))
