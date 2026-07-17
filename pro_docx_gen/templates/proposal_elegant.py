"""提案·优雅奢华模板 — v1.6.0 variant 模板。

场景：brand_story + 变体：elegant_luxury
"""
from shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="proposal_elegant",
    display_name="提案·优雅奢华",
    scene="brand_story",
    variant="elegant_luxury",
    description="高端提案文档，黑金配色，衬线字体，精致排版",
    theme_overrides={"color": {"primary": "#1C1917", "heading": "#2D2A26", "accent": "#D4AF37"}},
    default_structure={
        "toc": True,
        "sections": [
            {"title": "项目概述", "level": 1, "content": []},
            {"title": "市场机会", "level": 1, "content": []},
            {"title": "解决方案", "level": 1, "content": []},
            {"title": "执行计划", "level": 1, "content": []},
            {"title": "预期成果", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 3.0, "line_spacing": 1.8},
))