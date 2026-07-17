"""contract_cn — 中文合同模板。"""
from ..shared.template_registry import DOCXTemplate
from .registry import register

register(DOCXTemplate(
    name="contract_cn",
    display_name="中文合同",
    scene="contract",
    description="标准中文合同骨架（当事人/标的/权利义务/违约/争议解决/签署）。",
    theme_overrides={},
    default_structure={
        "toc": False,
        "sections": [
            {"title": "第一条 合同当事人", "level": 1, "content": []},
            {"title": "第二条 合同标的与数量", "level": 1, "content": []},
            {"title": "第三条 合同价款与支付", "level": 1, "content": []},
            {"title": "第四条 双方权利与义务", "level": 1, "content": []},
            {"title": "第五条 履行期限与地点", "level": 1, "content": []},
            {"title": "第六条 违约责任", "level": 1, "content": []},
            {"title": "第七条 争议解决", "level": 1, "content": []},
            {"title": "第八条 合同生效", "level": 1, "content": []},
            {"title": "签署栏", "level": 1, "content": []},
        ],
    },
    page_setup={"margins_cm": 2.54, "line_spacing": 1.5},
))
