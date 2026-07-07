"""lesson_plan_gaokao — 高考复习教案模板。"""
from shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="lesson_plan_gaokao",
    display_name="高考复习教案",
    scene="gaokao_review",
    description="高考复习专用教案骨架（考点分析/真题演练/错题归因/巩固训练）。",
    theme_overrides={"color": {"primary": "#C00000"}},
    default_structure={
        "toc": False,
        "sections": [
            {"title": "一、考点分析", "level": 1, "content": []},
            {"title": "二、知识梳理", "level": 1, "content": []},
            {"title": "三、真题演练", "level": 1, "content": []},
            {"title": "四、错题归因", "level": 1, "content": []},
            {"title": "五、巩固训练", "level": 1, "content": []},
            {"title": "六、课堂小结", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.0, "line_spacing": 1.5},
))
