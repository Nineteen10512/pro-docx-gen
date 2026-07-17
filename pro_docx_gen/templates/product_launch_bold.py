"""产品发布·大胆冲击模板 — v1.6.0 variant 模板。

场景：product_launch + 变体：bold_impact
"""
from ..shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="product_launch_bold",
    display_name="产品发布·大胆冲击",
    scene="product_launch",
    variant="bold_impact",
    description="发布会风格产品文档，强视觉冲击，少文字大标题",
    theme_overrides={"color": {"primary": "#000000", "heading": "#1A1A1A", "accent": "#E53935"}},
    default_structure={
        "toc": False,
        "sections": [
            {"title": "产品愿景", "level": 1, "content": []},
            {"title": "核心亮点", "level": 1, "content": []},
            {"title": "用户价值", "level": 1, "content": []},
            {"title": "技术突破", "level": 1, "content": []},
            {"title": "即刻体验", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.0, "line_spacing": 1.3},
))
