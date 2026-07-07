"""DOCX 模板 registry — v1.5.1 D 能力域。

注册 DOCX 预设模板（10 套 = 3 旧 + 7 新），提供：
- register(tpl)           : 注册一个 DOCXTemplate
- DOCXTEMPLATE_REGISTRY   : dict[name, DOCXTemplate]
- list_templates()        : 摘要列表
- get_template(name)      : 按名取模板

约定（与 shared/template_registry.py 一致）：
- 模板数据文件是**纯数据**（thin），< 30 行，不写逻辑
- 字段全部 optional，默认 None/False
- 旧 API（academic.new_paper / business.new_report / teaching.new_lesson_plan）继续工作

@since v1.5.1
"""
from __future__ import annotations

from typing import Any, Optional

from shared.template_registry import DOCXTemplate

DOCXTEMPLATE_REGISTRY: dict[str, DOCXTemplate] = {}


def register(tpl: DOCXTemplate) -> DOCXTemplate:
    """注册一个 DOCXTemplate。name 重复时覆盖（后注册优先）。"""
    DOCXTEMPLATE_REGISTRY[tpl.name] = tpl
    return tpl


def list_templates() -> list[dict[str, Any]]:
    """返回所有 DOCX 模板的轻量摘要。"""
    return [
        {
            "name": t.name,
            "display_name": t.display_name,
            "scene": t.scene,
            "description": t.description,
        }
        for t in DOCXTEMPLATE_REGISTRY.values()
    ]


def get_template(name: str) -> Optional[DOCXTemplate]:
    """按 name 取模板对象；未找到返回 None。"""
    return DOCXTEMPLATE_REGISTRY.get(name)


__all__ = [
    "DOCXTEMPLATE_REGISTRY",
    "register",
    "list_templates",
    "get_template",
]
