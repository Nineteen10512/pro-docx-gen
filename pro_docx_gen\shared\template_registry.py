"""Shared template registry dataclasses for PRO-PPTX/DOCX v1.5.1.

Provides PPTTemplate / DOCXTemplate base dataclasses used by both
pro_ppt_gen and pro_docx_gen. Templates are pure-data (< 30 lines each),
no logic here.

@since v1.5.1
"""
from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class PPTTemplate:
    """A scene preset for PPT generation.

    Fields:
        name:           unique snake_case id (e.g. 'thesis_defense')
        display_name:   human readable Chinese name
        scene:          scene tag: academic/business/teaching/general
        base_theme:     underlying theme name (academic/business/teaching)
        theme_overrides: dict merged into base theme Design Tokens
        default_transition: slide transition preset name
        default_animation:  dict for default element animation on content slides
        cover_style:    prompt hint for cover layout
        typical_slides: list[str] of recommended semantic slide types
        description:    one-line scene description shown in list_templates()
    """
    name: str
    display_name: str
    scene: str = "general"
    base_theme: str = "academic"
    theme_overrides: dict = field(default_factory=dict)
    default_transition: Optional[str] = None
    default_animation: Optional[dict] = None
    cover_style: str = ""
    typical_slides: list = field(default_factory=list)
    description: str = ""


@dataclass
class DOCXTemplate:
    """A scene preset for DOCX generation.

    Fields:
        name:             unique snake_case id
        display_name:     human readable Chinese name
        scene:            scene tag
        variant:          v1.6.0 style variant name (corporate_formal/minimal_clean/...)
        base_theme:       underlying theme name
        theme_overrides:  dict merged into base theme Design Tokens
        default_font:     eastAsia + latin font tuple (e.g. {"ea": "宋体", "la": "Times New Roman"})
        default_structure: dict describing default scaffold (sections / toc / abstract / references)
        typical_structure: list[str] of recommended semantic section types
        table_style:      default table style id hint
        page_setup:       dict of page-level settings (margins_cm, line_spacing, columns, etc.)
        citation_style:   default citation style hint (gb7714/apa/mla/ieee)
        description:      one-line scene description
    """
    name: str
    display_name: str
    scene: str = "general"
    variant: str = ""  # v1.6.0
    base_theme: str = "academic"
    theme_overrides: dict = field(default_factory=dict)
    default_font: Optional[dict] = None
    default_structure: Optional[dict] = None
    typical_structure: list = field(default_factory=list)
    table_style: Optional[str] = None
    page_setup: Optional[dict] = None
    citation_style: Optional[str] = None
    description: str = ""


@dataclass
class TemplateInfo:
    """Lightweight info returned by list_templates() / scan_local_templates()."""
    name: str
    display_name: str
    endpoint: str           # 'ppt' | 'docx'
    scene: str
    description: str = ""
    cover_style: str = ""   # ppt only
    preview_prompt: str = ""
    path: str = ""          # for scanned local templates
