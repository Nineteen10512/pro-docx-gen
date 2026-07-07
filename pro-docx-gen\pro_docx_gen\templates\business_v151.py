"""Business template (v1.5 迁移) — 商务报告主题。"""
from shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="business",
    display_name="商务报告",
    scene="business",
    description="标准商务报告骨架（执行摘要/背景/分析/建议）。",
    theme_overrides={},
    default_structure={
        "toc": True,
        "sections": [
            {"title": "Executive Summary", "level": 1, "content": []},
            {"title": "Background", "level": 1, "content": []},
            {"title": "Analysis", "level": 1, "content": []},
            {"title": "Recommendations", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.0, "line_spacing": 1.15},
))
