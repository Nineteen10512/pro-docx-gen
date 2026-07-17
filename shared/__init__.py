"""PRO-PPTX/DOCX v1.5.1 shared modules - color palette, theme registry, SVG engine, QA utils, citation formatter, template scanner and registry.

Shared across the PRO-PPTX (``pro_ppt_gen``) and PRO-DOCX (``pro_docx_gen``) packages.
Internal module layout:
- color_palette.py   : 颜色调色板与色阶工具
- themes.py          : 统一主题字典 THEME_PALETTES（14 套主题 × 6 色）与别名表 THEME_ALIASES
- svg_engine.py      : SVG 渲染引擎（自由形状 + 迷你图 mini_bar/mini_line/mini_pie）
- quality.py         : QA 质量自检框架（PPT RUBRIC 6 维 / DOCX 7 维）
- citation.py        : 参考文献格式化引擎（apa/gb7714/mla/ieee），双端共享
- template_registry.py : v1.5.1 新增 — PPTTemplate / DOCXTemplate / TemplateInfo dataclass
- template_scanner.py  : v1.5.1 新增 — 本地 WPS 模板扫描 + theme 提取 + WCAG 对比度校验

@since v1.4.0
@updated v1.5.1
"""
from . import color_palette, themes, svg_engine, quality, citation, template_registry, template_scanner

__all__ = ["color_palette", "themes", "svg_engine", "quality", "citation", "template_registry", "template_scanner"]

__version__ = "1.5.1"
