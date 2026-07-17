"""Teaching template (v1.5 迁移) — 中文教学教案主题。"""
from ..shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="teaching",
    display_name="教学教案",
    scene="teaching",
    description="标准中文教案骨架（教学目标/重难点/过程/小结/作业）。",
    theme_overrides={},
    default_structure={
        "toc": False,
        "sections": [
            {"title": "一、教学目标", "level": 1, "content": []},
            {"title": "二、教学重难点", "level": 1, "content": []},
            {"title": "三、教学过程", "level": 1, "content": []},
            {"title": "四、课堂小结", "level": 1, "content": []},
            {"title": "五、作业布置", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.0, "line_spacing": 1.5},
))
