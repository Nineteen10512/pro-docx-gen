"""reading_notes — 读书笔记模板。"""
from ..shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="reading_notes",
    display_name="读书笔记",
    scene="reading",
    description="读书笔记骨架（书讯/章节要点/金句摘录/读后感/行动清单）。",
    theme_overrides={"color": {"primary": "#385723"}},
    default_structure={
        "toc": False,
        "sections": [
            {"title": "一、书籍信息", "level": 1, "content": []},
            {"title": "二、作者简介", "level": 1, "content": []},
            {"title": "三、全书结构", "level": 1, "content": []},
            {"title": "四、章节要点", "level": 1, "content": []},
            {"title": "五、金句摘录", "level": 1, "content": []},
            {"title": "六、读后感", "level": 1, "content": []},
            {"title": "七、行动清单", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.0, "line_spacing": 1.5},
))
