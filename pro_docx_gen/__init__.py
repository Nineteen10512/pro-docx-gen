"""PRO-DOCX v1.6.6 — PaperJSX semantic compilation architecture for professional Word document generation.

Pro Docx Gen (internal package name ``pro_docx_gen``, product name **PRO-DOCX**)
produces native editable ``.docx`` files from high-level semantic JSON or
Markdown. LLMs never write font sizes / colors / spacing — design tokens and a
deterministic renderer compute those automatically.

v1.5.2 新增：
- 共享 charts 调色板扩充至 ≥6 色；matplotlib 单 series 多类别柱/条图自动按调色板循环上色，
  与 PPT 端保持一致，避免单色柱图。
- 共享 quality.contrast_check / safe_text_color 工具函数，DOCX 后续文本节点可复用。

v1.4 新增（v1.4 highlights）：
- 编辑已有文档：load / update_document
- 修订追踪：revision 节点（w:ins / w:del）
- 批注 comment / 脚注 footnote / 尾注 endnote
- 水印 VML / 页面边框 / 页面设置增强
- 页眉页脚增强 / 13 套主题 / 新语义节点 equation / signature_block
- 元数据扩展 / 表格表头跨页 / PDF 预览 / 文本提取

v1.3.0 新增（历史）：
- 原生 OMML 公式 LaTeX → Word 可编辑公式
- 修订 API（list / accept / reject）
- .doc 旧格式自动转 .docx
- svg_shape 节点（共享 svg_engine）
- quality_check 七维 QA 自检

v1.4 新增（v1.4 highlights）：
- Markdown 直入 generate_from_markdown
- 自动目录 TOC（w:sdt + TOC 域代码，多级）
- 参考文献格式化引擎（apa / gb7714 / mla / ieee，shared/citation.py 双端共享）
- References 专业排版（悬挂缩进 / 方括号编号 / DOI 蓝色下划线 / 按类型分组）
- 多级列表自动编号（嵌套 items 3 层）
- 交叉引用 + inline {ref label} 文本自动解析（REF 域代码）
- 多栏排版 + 栏间分隔线（meta.columns / 段落级连续分节）
- 首字下沉 paragraph.drop_cap
- 行号显示 meta.line_numbers
- 智能结构识别四模式（cn×academic / cn×business / en×academic / en×business，字体/字号/对齐按 GB/T 7713 / GB/T 9704 / APA7 / McKinsey 切换）
- caption 自动转 figure/table/equation 节点
- 统一主题字典下沉到 shared/themes.py（与 PRO-PPTX 共享）
- 输出文件默认命名 {title}_v{version}.docx
- 中英双语 docstring + 中文错误提示与修复建议
- 中文字体 eastAsia 全面审计（零缺失）

本文件仅维护对外导入符号与版本号，**不修改包名/import 路径/函数签名**以保证 100% 向后兼容。

v1.7.0 新增：研报全套方案（business_compact + research_report 两套白底模板、
表格引擎 5 项新能力：条件格式 / 进度条 / 星级评分 / 多表头合并 / 自动计算 YoY·汇总、
KPI 卡片行 / 研报 header / 风险提示尾部 3 大模块）。
"""

__version__ = "1.7.0"

from .docx_jsx import (
    generate,
    generate_from_markdown,
    generate_with_collaboration,  # v1.5.1 C-6
    outline, word_count,
    load, update_document,
    to_pdf, to_images, extract_text,
    list_themes,
    NODE_TYPES,
    # v1.3
    list_revisions,
    accept_all_revisions, reject_all_revisions,
    accept_revision_by_id, reject_revision_by_id,
    quality_check,
)

# v1.5.1 A 能力域 — 本地 WPS 模板 theme 提取
from .theme_extractor import (
    extract_docx_theme,
    apply_extracted_theme,
    ThemeExtractionError,
)

# v1.5.1 D 能力域 — DOCX 模板库（10 套场景化模板）
from .templates import (
    list_templates,
    get_template,
    register as register_template,
)

# v1.7.0 — 表格引擎 5 项新能力（模块级公开函数）
from .engine.renderer import (
    render_conditional_cell,
    render_progress_bar,
    render_star_rating,
    render_merged_header,
    auto_compute_rows,
)
# v1.7.0 — KPI 卡片行 / 研报 header / 风险提示尾部
from .engine.kpi_cards import render_kpi_card_row
from .engine.research_modules import (
    render_research_header,
    render_rating_badge,
    render_risk_disclaimer_footer,
)
# v1.7.0 — 自动计算 helper（YoY / 汇总 / 颜色查找）
from .shared.auto_calc import (
    compute_yoy,
    compute_summary,
    auto_compute_rows as auto_compute_rows_helper,
    yoy_color_for,
)

# v1.5.1 C 能力域 — DOCX 多人协作
from .collaboration.session import (
    CollaborationSession,
    start_collaboration,
    add_collaborator,
    list_collaborators,
)
from .collaboration.suggest import (
    suggest_insert,
    suggest_replace,
    suggest_delete,
)
from .collaboration.comments import (
    add_review_comment,
    reply_comment,
    list_comments,
)
from .collaboration.merge import (
    list_suggestions,
    accept_by_id,
    reject_by_id,
    accept_by_author,
    reject_by_author,
    merge_collaboration,
    export_clean_copy,
    export_review_pdf,
)
from .collaboration.diff import diff_documents

# v1.6.0: 风格变体系统
from .v160_style_registry import (
    VariantProfile,
    VARIANT_PROFILES,
    get_variant,
    list_variants,
    match_variant,
)
from .variant_tokens import (
    apply_variant,
    get_variant_tokens,
    get_variant_cover_style,
)
from .docx_taste import taste_check  # v1.6.0: DOCX 审美预检
from .docx_qa import audit_docx_layout, expand_review_pages
from .translation import collect_translation_segments, apply_translation_map, assess_translation_risk

__all__ = [
    "generate",
    "generate_from_markdown",
    "generate_with_collaboration",
    "outline", "word_count",
    "load", "update_document",
    "to_pdf", "to_images", "extract_text",
    "list_themes",
    "NODE_TYPES",
    "list_revisions",
    "accept_all_revisions", "reject_all_revisions",
    "accept_revision_by_id", "reject_revision_by_id",
    "quality_check",
    # v1.5.1
    "extract_docx_theme", "apply_extracted_theme", "ThemeExtractionError",
    "list_templates", "get_template", "register_template",
    "scan_local_templates",
    "CollaborationSession", "start_collaboration",
    "add_collaborator", "list_collaborators",
    "suggest_insert", "suggest_replace", "suggest_delete",
    "add_review_comment", "reply_comment", "list_comments",
    "list_suggestions",
    "accept_by_id", "reject_by_id",
    "accept_by_author", "reject_by_author",
    "merge_collaboration",
    "export_clean_copy", "export_review_pdf",
    "diff_documents",
    "__version__",
    # v1.6.0
    "VariantProfile", "VARIANT_PROFILES", "get_variant", "list_variants", "match_variant",
    "apply_variant", "get_variant_tokens", "get_variant_cover_style",
    "taste_check",
    "audit_docx_layout", "expand_review_pages",
    "collect_translation_segments", "apply_translation_map", "assess_translation_risk",
    # v1.7.0
    "render_conditional_cell", "render_progress_bar", "render_star_rating",
    "render_merged_header", "auto_compute_rows",
    "render_kpi_card_row",
    "render_research_header", "render_rating_badge", "render_risk_disclaimer_footer",
    "compute_yoy", "compute_summary", "auto_compute_rows_helper", "yoy_color_for",
]


# v1.5.1 A-API: scan_local_templates for DOCX endpoint
def scan_local_templates(dirs=None, recursive: bool = True) -> list:
    """v1.5.1: 扫描本地 WPS/Office 模板目录，返回 DOCX 类 TemplateInfo 列表。"""
    try:
        from .shared.template_scanner import scan_local_templates as _slt
    except ImportError:  # pragma: no cover - legacy skill-root layout
        from shared.template_scanner import scan_local_templates as _slt
    return [t for t in _slt(dirs=dirs, recursive=recursive) if getattr(t, "endpoint", None) == "docx"]
