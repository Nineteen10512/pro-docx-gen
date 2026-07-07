"""极简报告模板 — v1.6.0 variant 模板。

场景：business_report + 变体：minimal_clean
"""
from shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="report_minimal",
    display_name="报告·极简干净",
    scene="business_report",
    variant="minimal_clean",
    description="通用极简报告，大留白，适合咨询/设计类报告",
    theme_overrides={"color": {"primary": "#1A1A1A", "heading": "#333333", "accent": "#5B9BD5"}},
    default_structure={
        "toc": True,
        "sections": [
            {"title": "背景与目标", "level": 1, "content": []},
            {"title": "核心发现", "level": 1, "content": []},
            {"title": "详细分析", "level": 1, "content": []},
            {"title": "建议方案", "level": 1, "content": []},
            {"title": "下一步", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.8, "line_spacing": 1.6},
))