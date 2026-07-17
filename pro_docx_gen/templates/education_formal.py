"""教育课件·企业正式模板 — v1.6.0 variant 模板。

场景：education + 变体：corporate_formal
"""
from ..shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="education_formal",
    display_name="教育课件·企业正式",
    scene="education",
    variant="corporate_formal",
    description="结构化教学课件，蓝白配色，清晰层级",
    theme_overrides={"color": {"primary": "#1F3864", "heading": "#1F3864"}},
    default_structure={
        "toc": True,
        "sections": [
            {"title": "课程目标", "level": 1, "content": []},
            {"title": "知识回顾", "level": 1, "content": []},
            {"title": "新课讲授", "level": 1, "content": []},
            {"title": "案例分析", "level": 1, "content": []},
            {"title": "课后练习", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.54, "line_spacing": 1.5},
))
