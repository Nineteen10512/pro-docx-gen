"""thesis_full — 完整学术论文模板（含摘要/关键词/致谢/附录）。"""
from shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="thesis_full",
    display_name="完整论文",
    scene="thesis",
    description="完整学位论文骨架（中英文摘要/目录/正文/致谢/参考文献/附录）。",
    theme_overrides={},
    default_structure={
        "toc": True,
        "abstract": {"text": "", "keywords": [], "english_abstract": True},
        "sections": [
            {"title": "第一章 绪论", "level": 1, "content": []},
            {"title": "1.1 研究背景", "level": 2, "content": []},
            {"title": "1.2 研究意义", "level": 2, "content": []},
            {"title": "第二章 文献综述", "level": 1, "content": []},
            {"title": "第三章 研究方法", "level": 1, "content": []},
            {"title": "第四章 实验与结果", "level": 1, "content": []},
            {"title": "第五章 讨论", "level": 1, "content": []},
            {"title": "第六章 结论与展望", "level": 1, "content": []},
            {"title": "致谢", "level": 1, "content": []},
        ],
        "references": [],
        "appendix": [],
    },
    page_setup={"margins_cm": 2.54, "line_spacing": 1.5},
))
