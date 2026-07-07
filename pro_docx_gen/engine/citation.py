"""Backward-compatible re-export of citation formatter.

v1.4 P1-4b: citation engine moved to ``shared/citation.py`` so both PPT and DOCX
sides can share the same implementation. This module re-exports everything from
``shared.citation`` to keep old import paths (``pro_docx_gen.engine.citation``)
fully working.
"""
from __future__ import annotations

# Try import strategies in order:
# 1) Relative import (when installed as proper package with parent)
# 2) Direct `shared.citation` import (when skills/ is on sys.path, the typical setup)
# 3) Fallback: add skills root to sys.path
_imported = False
try:
    from ...shared.citation import (  # type: ignore
        SUPPORTED_STYLES,
        _GB_TYPE_CODE,
        format_reference,
        _fmt_apa,
        _fmt_gb7714,
        _fmt_mla,
        _fmt_ieee,
        _apa_authors,
        _gb_authors,
        _mla_authors,
        _ieee_authors,
        _apa_pages,
    )
    _imported = True
except Exception:
    pass

if not _imported:
    try:
        from shared.citation import (  # type: ignore
            SUPPORTED_STYLES,
            _GB_TYPE_CODE,
            format_reference,
            _fmt_apa,
            _fmt_gb7714,
            _fmt_mla,
            _fmt_ieee,
            _apa_authors,
            _gb_authors,
            _mla_authors,
            _ieee_authors,
            _apa_pages,
        )
        _imported = True
    except Exception:
        pass

if not _imported:
    import sys as _sys, os as _os
    _here = _os.path.dirname(_os.path.abspath(__file__))
    _skills_root = _os.path.normpath(_os.path.join(_here, "..", "..", ".."))
    if _skills_root not in _sys.path:
        _sys.path.insert(0, _skills_root)
    from shared.citation import (  # type: ignore
        SUPPORTED_STYLES,
        _GB_TYPE_CODE,
        format_reference,
        _fmt_apa,
        _fmt_gb7714,
        _fmt_mla,
        _fmt_ieee,
        _apa_authors,
        _gb_authors,
        _mla_authors,
        _ieee_authors,
        _apa_pages,
    )

# The canonical alias used throughout the DOCX renderer
_cite_format = format_reference


__all__ = [
    "SUPPORTED_STYLES",
    "_GB_TYPE_CODE",
    "format_reference",
    "_cite_format",
    "_fmt_apa",
    "_fmt_gb7714",
    "_fmt_mla",
    "_fmt_ieee",
]
