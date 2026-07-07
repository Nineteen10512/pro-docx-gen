"""数据分析·现代科技模板 — v1.6.0 variant 模板。

场景：data_analysis + 变体：modern_tech
"""
from shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="data_analysis_tech",
    display_name="数据分析·现代科技",
    scene="data_analysis",
    variant="modern_tech",
    description="数据分析报告，科技感配色，图表友好布局",
    theme_overrides={"color": {"primary": "#0A192F", "heading": "#E0E7FF", "accent": "#64FFDA"}},
    default_structure={
        "toc": True,
        "sections": [
            {"title": "数据概览", "level": 1, "content": []},
            {"title": "核心指标", "level": 1, "content": []},
            {"title": "趋势分析", "level": 1, "content": []},
            {"title": "异常检测", "level": 1, "content": []},
            {"title": "建议与行动", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.2, "line_spacing": 1.4},
))