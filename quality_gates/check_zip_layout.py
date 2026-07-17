"""Validate PRO-DOCX release zip layout.

Expected zip shape:
    pro-docx-gen/
      pro_docx_gen/
      quality_gates/
      smoke_tests/
      codex_adaptation/
      SKILL.md
      README.md
      CHANGELOG.md

No duplicate root-level package directories are allowed.
"""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path


EXPECTED_ROOT = "pro-docx-gen/"
FORBIDDEN_ROOTS = {
    "pro_docx_gen/",
    "quality_gates/",
    "shared/",
    "__pycache__/",
}
REQUIRED = {
    "pro-docx-gen/pro_docx_gen/__init__.py",
    "pro-docx-gen/SKILL.md",
    "pro-docx-gen/pro_docx_gen/SKILL.md",
    "pro-docx-gen/quality_gates/run_quality_gate.py",
    "pro-docx-gen/quality_gates/check_zip_layout.py",
    "pro-docx-gen/smoke_tests/run_smoke_tests.py",
    "pro-docx-gen/codex_adaptation/CODEX.md",
    "pro-docx-gen/codex_adaptation/compatibility_manifest.json",
    "pro-docx-gen/README.md",
    "pro-docx-gen/CHANGELOG.md",
}
TEXT_SUFFIXES = {".py", ".md", ".txt", ".json", ".yaml", ".yml"}
GENERATED_SUFFIXES = {".docx", ".pptx", ".pdf", ".png", ".jpg", ".jpeg", ".zip", ".ps1"}
TEMP_SUFFIXES = {".bak", ".tmp", ".temp", ".pyc", ".pyo"}
LOCAL_PATH_MARKERS = (
    "C:" + "\\Users\\",
    "C:" + "/Users/",
    "\\AppData" + "\\Local\\Temp\\",
    "/" + "tmp/",
    "/" + "mnt/data/",
    "codex" + "-runtimes",
)


def _top_level(name: str) -> str:
    if "/" not in name:
        return name
    return name.split("/", 1)[0] + "/"


def inspect_zip(path: Path) -> list[str]:
    errors: list[str] = []
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        roots = {_top_level(n) for n in names if n}
        file_names = [n for n in names if n and not n.endswith("/")]

    if EXPECTED_ROOT not in roots:
        errors.append(f"missing expected root: {EXPECTED_ROOT}")
    extra_forbidden = sorted(roots & FORBIDDEN_ROOTS)
    for root in extra_forbidden:
        errors.append(f"forbidden duplicate root: {root}")

    missing = sorted(REQUIRED - names)
    for name in missing:
        errors.append(f"missing required file: {name}")

    junk = sorted(n for n in names if "__pycache__/" in n or n.endswith(".pyc") or n.endswith(".pyo"))
    for name in junk:
        errors.append(f"compiled/cache artifact present: {name}")

    for name in sorted(file_names):
        parts = Path(name).parts
        if any(part == "_work" for part in parts):
            errors.append(f"runtime work directory present: {name}")
        if any(part.startswith("_output") for part in parts):
            errors.append(f"generated test output present: {name}")
        if any(part == "chart_assets" for part in parts):
            errors.append(f"generated chart asset directory present: {name}")
        suffix = Path(name).suffix.lower()
        if suffix in TEMP_SUFFIXES:
            errors.append(f"temporary artifact present: {name}")
        if suffix in GENERATED_SUFFIXES:
            errors.append(f"generated binary/artifact present: {name}")

    with zipfile.ZipFile(path) as zf:
        for name in sorted(file_names):
            suffix = Path(name).suffix.lower()
            if suffix not in TEXT_SUFFIXES:
                continue
            try:
                text = zf.read(name).decode("utf-8")
            except UnicodeDecodeError:
                errors.append(f"text file is not UTF-8: {name}")
                continue
            for marker in LOCAL_PATH_MARKERS:
                if marker in text:
                    errors.append(f"local machine path leaked in {name}: {marker}")
                    break

    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description="Check PRO-DOCX release zip layout.")
    parser.add_argument("zip_path", type=Path)
    args = parser.parse_args()

    errors = inspect_zip(args.zip_path)
    if errors:
        print("FAIL zip layout")
        for error in errors:
            print(f"  ERROR {error}")
        return 1
    print(f"PASS zip layout {args.zip_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
