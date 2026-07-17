"""business_proposal — 商业提案模板。"""
from ..shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="business_proposal",
    display_name="商业提案",
    scene="business_proposal",
    description="商业提案骨架（项目概述/市场分析/解决方案/预算/时间表/风险）。",
    theme_overrides={},
    default_structure={
        "toc": True,
        "sections": [
            {"title": "1. 项目概述", "level": 1, "content": []},
            {"title": "2. 市场分析", "level": 1, "content": []},
            {"title": "3. 解决方案", "level": 1, "content": []},
            {"title": "4. 预算明细", "level": 1, "content": []},
            {"title": "5. 时间表", "level": 1, "content": []},
            {"title": "6. 风险与对策", "level": 1, "content": []},
            {"title": "7. 结论", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.0, "line_spacing": 1.15},
))
