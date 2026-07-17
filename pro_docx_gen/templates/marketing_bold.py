"""营销方案·大胆冲击模板 — v1.6.0 variant 模板。

场景：marketing + 变体：bold_impact
"""
from ..shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="marketing_bold",
    display_name="营销方案·大胆冲击",
    scene="marketing",
    variant="bold_impact",
    description="强对比黑白+红色强调，超大标题，适合营销提案",
    theme_overrides={"color": {"primary": "#000000", "heading": "#1A1A1A", "accent": "#E53935"}},
    default_structure={
        "toc": False,
        "sections": [
            {"title": "市场洞察", "level": 1, "content": []},
            {"title": "目标受众", "level": 1, "content": []},
            {"title": "核心策略", "level": 1, "content": []},
            {"title": "执行计划", "level": 1, "content": []},
            {"title": "效果预估", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.0, "line_spacing": 1.3},
))
