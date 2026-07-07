"""resume_cn — 中文简历模板。"""
from shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="resume_cn",
    display_name="中文简历",
    scene="resume",
    description="标准中文简历骨架（个人信息/教育背景/工作经历/项目经验/技能证书）。",
    theme_overrides={"color": {"primary": "#1F4E79"}},
    default_structure={
        "toc": False,
        "sections": [
            {"title": "个人信息", "level": 1, "content": []},
            {"title": "求职意向", "level": 1, "content": []},
            {"title": "教育背景", "level": 1, "content": []},
            {"title": "工作经历", "level": 1, "content": []},
            {"title": "项目经验", "level": 1, "content": []},
            {"title": "专业技能", "level": 1, "content": []},
            {"title": "获奖证书", "level": 1, "content": []},
            {"title": "自我评价", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 1.8, "line_spacing": 1.25},
))
