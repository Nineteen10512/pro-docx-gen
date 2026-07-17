"""企业正式学术论文模板 — v1.6.0 variant 模板。

场景：academic + 变体：corporate_formal
"""
from ..shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="academic_corporate",
    display_name="学术论文·企业正式",
    scene="academic",
    variant="corporate_formal",
    description="严谨学术论文骨架，蓝白配色，标准学术格式",
    theme_overrides={"color": {"primary": "#1F3864", "heading": "#1F3864"}},
    default_structure={
        "toc": True,
        "abstract": {"text": "", "keywords": []},
        "sections": [
            {"title": "引言", "level": 1, "content": []},
            {"title": "文献综述", "level": 1, "content": []},
            {"title": "研究方法", "level": 1, "content": []},
            {"title": "结果与讨论", "level": 1, "content": []},
            {"title": "结论", "level": 1, "content": []},
        ],
        "references": [],
    },
    page_setup={"margins_cm": 2.54, "line_spacing": 1.5},
))
