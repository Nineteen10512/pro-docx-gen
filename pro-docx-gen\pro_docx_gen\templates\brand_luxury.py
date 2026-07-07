"""品牌故事·优雅奢华模板 — v1.6.0 variant 模板。

场景：brand_story + 变体：elegant_luxury
"""
from shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="brand_luxury",
    display_name="品牌故事·优雅奢华",
    scene="brand_story",
    variant="elegant_luxury",
    description="黑金配色品牌故事，衬线字体，对称精致排版",
    theme_overrides={"color": {"primary": "#1C1917", "heading": "#2D2A26", "accent": "#D4AF37"}},
    default_structure={
        "toc": False,
        "sections": [
            {"title": "品牌起源", "level": 1, "content": []},
            {"title": "核心理念", "level": 1, "content": []},
            {"title": "品牌里程碑", "level": 1, "content": []},
            {"title": "产品哲学", "level": 1, "content": []},
            {"title": "未来愿景", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 3.0, "line_spacing": 1.8},
))