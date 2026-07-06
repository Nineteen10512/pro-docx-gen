"""meeting_minutes — 会议纪要模板。"""
from shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="meeting_minutes",
    display_name="会议纪要",
    scene="meeting",
    description="标准会议纪要骨架（会议信息/议程/讨论要点/决议/待办）。",
    theme_overrides={},
    default_structure={
        "toc": False,
        "sections": [
            {"title": "一、会议基本信息", "level": 1, "content": []},
            {"title": "二、参会人员", "level": 2, "content": []},
            {"title": "三、会议议程", "level": 1, "content": []},
            {"title": "四、讨论要点", "level": 1, "content": []},
            {"title": "五、会议决议", "level": 1, "content": []},
            {"title": "六、待办事项", "level": 1, "content": []},
            {"title": "七、下次会议", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.0, "line_spacing": 1.5},
))
