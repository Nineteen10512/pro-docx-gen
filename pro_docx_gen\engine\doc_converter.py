"""DOCX .doc legacy format auto-converter (DOCX-P0-4).

Transparently converts .doc (Word 97-2003) files to .docx using LibreOffice
headless before passing to python-docx loader. Caches converted files in a
temp dir to avoid re-conversion within one session.

@since v1.3.0 (DOCX-P0-4)
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from typing import Optional


_DOC_CONVERT_CACHE: dict[str, str] = {}
_CACHE_DIR: Optional[str] = None


def _get_cache_dir() -> str:
    global _CACHE_DIR
    if _CACHE_DIR is None:
        _CACHE_DIR = tempfile.mkdtemp(prefix="paperjsx_doc_convert_")
    return _CACHE_DIR


def is_doc_path(path: str) -> bool:
    return isinstance(path, str) and path.lower().endswith(".doc") \
        and not path.lower().endswith(".docx")


def convert_doc_to_docx(doc_path: str, out_dir: Optional[str] = None) -> str:
    """Convert a .doc file to .docx using LibreOffice.

    Returns the resulting .docx path. Raises RuntimeError if soffice is missing
    or conversion fails.
    """
    if not os.path.exists(doc_path):
        raise FileNotFoundError(doc_path)
    abs_doc = os.path.abspath(doc_path)
    # cache by (path, mtime, size)
    key = f"{abs_doc}:{os.path.getmtime(abs_doc)}:{os.path.getsize(abs_doc)}"
    if key in _DOC_CONVERT_CACHE:
        p = _DOC_CONVERT_CACHE[key]
        if os.path.exists(p):
            return p

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError(
            "LibreOffice (soffice) not found; cannot convert .doc → .docx. "
            "Please install LibreOffice or provide a .docx file directly."
        )
    target_dir = out_dir or _get_cache_dir()
    os.makedirs(target_dir, exist_ok=True)
    cmd = [soffice, "--headless", "--convert-to", "docx",
           "--outdir", target_dir, abs_doc]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if result.returncode != 0:
        raise RuntimeError(f"soffice .doc→.docx failed: {result.stderr}")
    base = os.path.splitext(os.path.basename(abs_doc))[0]
    docx_path = os.path.join(target_dir, base + ".docx")
    if not os.path.exists(docx_path):
        # Try to find any generated .docx
        import glob
        candidates = glob.glob(os.path.join(target_dir, "*.docx"))
        if candidates:
            docx_path = candidates[-1]
        else:
            raise RuntimeError(f".docx not produced by soffice: {result.stdout} {result.stderr}")
    _DOC_CONVERT_CACHE[key] = docx_path
    print(f"[doc-convert] Auto-converted .doc → .docx: {abs_doc} → {docx_path}")
    return docx_path


def ensure_docx(path: str) -> str:
    """Return a .docx path, converting from .doc if necessary.

    If input is already .docx (or not a path at all), return as-is.
    """
    if not isinstance(path, str):
        return path
    if is_doc_path(path):
        return convert_doc_to_docx(path)
    return path


__all__ = ["is_doc_path", "convert_doc_to_docx", "ensure_docx"]
