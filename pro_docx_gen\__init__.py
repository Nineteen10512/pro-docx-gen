"""PRO-DOCX v1.6.0 public package surface.

PRO-DOCX generates native editable ``.docx`` files from semantic JSON or
Markdown and can translate existing DOCX files while preserving layout,
tables, and image fidelity as much as possible.

v1.6.0 highlights:
- 5 global style variants plus ``variant`` and ``auto_style``
- DOCX ``taste_check()`` preflight
- Existing DOCX translation APIs:
  ``collect_translation_segments()``,
  ``apply_translation_map()``,
  ``assess_translation_risk()``
- Table translation layer with risk-aware autofit and tightening
- Protected write-back path for image-bearing and field-bearing runs

This module keeps public imports and version metadata stable for backward
compatibility.
"""

__version__ = "1.6.0"

from .docx_jsx import (
    generate,
    generate_from_markdown,
    generate_with_collaboration,
    outline,
    word_count,
    load,
    update_document,
    to_pdf,
    to_images,
    extract_text,
    list_themes,
    NODE_TYPES,
    list_revisions,
    accept_all_revisions,
    reject_all_revisions,
    accept_revision_by_id,
    reject_revision_by_id,
    quality_check,
)

from .theme_extractor import (
    extract_docx_theme,
    apply_extracted_theme,
    ThemeExtractionError,
)

from .templates import (
    list_templates,
    get_template,
    register as register_template,
)

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
from .docx_taste import taste_check
from .translation import (
    collect_translation_segments,
    apply_translation_map,
    assess_translation_risk,
    build_translation_prompt,
    TRANSLATION_SELF_AUDIT_RULES,
)


def scan_local_templates(dirs=None, recursive: bool = True) -> list:
    """Scan local WPS/Office template directories for DOCX endpoints."""
    from shared.template_scanner import scan_local_templates as _slt

    return [
        t for t in _slt(dirs=dirs, recursive=recursive)
        if getattr(t, "endpoint", None) == "docx"
    ]


__all__ = [
    "generate",
    "generate_from_markdown",
    "generate_with_collaboration",
    "outline",
    "word_count",
    "load",
    "update_document",
    "to_pdf",
    "to_images",
    "extract_text",
    "list_themes",
    "NODE_TYPES",
    "list_revisions",
    "accept_all_revisions",
    "reject_all_revisions",
    "accept_revision_by_id",
    "reject_revision_by_id",
    "quality_check",
    "extract_docx_theme",
    "apply_extracted_theme",
    "ThemeExtractionError",
    "list_templates",
    "get_template",
    "register_template",
    "scan_local_templates",
    "CollaborationSession",
    "start_collaboration",
    "add_collaborator",
    "list_collaborators",
    "suggest_insert",
    "suggest_replace",
    "suggest_delete",
    "add_review_comment",
    "reply_comment",
    "list_comments",
    "list_suggestions",
    "accept_by_id",
    "reject_by_id",
    "accept_by_author",
    "reject_by_author",
    "merge_collaboration",
    "export_clean_copy",
    "export_review_pdf",
    "diff_documents",
    "VariantProfile",
    "VARIANT_PROFILES",
    "get_variant",
    "list_variants",
    "match_variant",
    "apply_variant",
    "get_variant_tokens",
    "get_variant_cover_style",
    "taste_check",
    "collect_translation_segments",
    "apply_translation_map",
    "assess_translation_risk",
    "build_translation_prompt",
    "TRANSLATION_SELF_AUDIT_RULES",
    "__version__",
]
