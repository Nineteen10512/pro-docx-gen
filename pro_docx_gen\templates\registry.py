"""DOCX template registry for PRO-DOCX v1.6.0."""

from __future__ import annotations

from typing import Any, Optional

from shared.template_registry import DOCXTemplate

DOCXTEMPLATE_REGISTRY: dict[str, DOCXTemplate] = {}


def register(tpl: DOCXTemplate) -> DOCXTemplate:
    """Register a DOCXTemplate descriptor."""
    DOCXTEMPLATE_REGISTRY[tpl.name] = tpl
    return tpl


def list_templates() -> list[dict[str, Any]]:
    """Return lightweight summary metadata for all DOCX templates."""
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
    """Return a template by name, or ``None`` if missing."""
    return DOCXTEMPLATE_REGISTRY.get(name)


__all__ = [
    "DOCXTEMPLATE_REGISTRY",
    "register",
    "list_templates",
    "get_template",
]
