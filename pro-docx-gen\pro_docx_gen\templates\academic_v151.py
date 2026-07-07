"""Academic template (v1.5 迁移) — 学术论文主题。

把 v1.5 学术论文骨架注册为 DOCXTemplate 描述符。
原 academic.new_paper() 函数保留可用。
"""
from shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="academic",
    display_name="学术论文",
    scene="academic",
    description="标准学术论文骨架（摘要/目录/方法/结果/结论/参考文献）。",
    theme_overrides={},
    default_structure={
        "toc": True,
        "abstract": {"text": "", "keywords": []},
        "sections": [
            {"title": "Introduction", "level": 1, "content": []},
            {"title": "Methodology", "level": 1, "content": []},
            {"title": "Results and Discussion", "level": 1, "content": []},
            {"title": "Conclusion", "level": 1, "content": []},
        ],
        "references": [],
    },
    page_setup={"margins_cm": 2.54, "line_spacing": 1.5},
))
