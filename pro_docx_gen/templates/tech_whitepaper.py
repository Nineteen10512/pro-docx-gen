"""科技白皮书模板 — v1.6.0 variant 模板。

场景：product_launch + 变体：modern_tech
"""
from shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="tech_whitepaper",
    display_name="科技白皮书·现代科技",
    scene="product_launch",
    variant="modern_tech",
    description="科技感白皮书，深色主题，青蓝霓虹强调",
    theme_overrides={"color": {"primary": "#64FFDA"}},
    default_structure={
        "toc": True,
        "sections": [
            {"title": "产品概述", "level": 1, "content": []},
            {"title": "技术架构", "level": 1, "content": []},
            {"title": "核心功能", "level": 1, "content": []},
            {"title": "性能指标", "level": 1, "content": []},
            {"title": "路线图", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.2, "line_spacing": 1.4},
))