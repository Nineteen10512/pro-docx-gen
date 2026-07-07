"""Lightweight delivery quality gate for PRO-PPTX / PRO-DOCX outputs.

This script is inspired by the built-in presentation/document skills' delivery
discipline: inspect structure first, render when the local environment supports
it, and return a machine-readable report.

It intentionally uses only the standard library. LibreOffice rendering is
optional and is skipped when `soffice` / `libreoffice` is unavailable.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET


PLACEHOLDER_RE = re.compile(r"TODO|TBD|待补|占位|示例文本|\[chart\]|\[image\]", re.I)
AI_TELL_RE = re.compile(
    r"\b(?:lorem ipsum|john doe|jane doe|acme|nexus|smartflow|cloudly|"
    r"synergy|next-gen|cutting-edge|revolutionize|unlock|elevate|"
    r"seamless|world-class|game-changing)\b|"
    r"\u8d4b\u80fd|\u98a0\u8986|\u91cd\u5851|\u65e0\u7f1d|\u4e0b\u4e00\u4ee3",
    re.I,
)
DECORATIVE_DASH_RE = re.compile(r"[\u2013\u2014]")
LONG_TEXT_LIMIT = 240
PPT_SLIDE_TEXT_LIMIT = 1200
DOCX_PARAGRAPH_LIMIT = 900


def _issue(level: str, code: str, message: str) -> dict:
    return {"level": level, "code": code, "message": message}


def _read_zip_text(zf: zipfile.ZipFile, name: str) -> str:
    return zf.read(name).decode("utf-8", errors="ignore")


def _xml_text(xml: str) -> list[str]:
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return []
    out: list[str] = []
    for el in root.iter():
        if el.text and el.tag.endswith("}t"):
            out.append(el.text)
    return out


def _natural_slide_key(name: str) -> tuple[int, str]:
    m = re.search(r"slide(\d+)\.xml$", name)
    return (int(m.group(1)) if m else 0, name)


def _inspect_pptx_taste(report: dict, slide_texts: list[str]) -> None:
    """Add lightweight Impeccable/Taste-inspired text warnings for PPTX files."""
    repeated = Counter(t.lower().strip() for t in slide_texts if t.strip())
    if any(count > 1 and len(text) > 20 for text, count in repeated.items()):
        report["warnings"].append(
            _issue("warning", "duplicate_slide_text", "Repeated slide text suggests templated or duplicated content")
        )

    for idx, text in enumerate(slide_texts, 1):
        if AI_TELL_RE.search(text):
            report["warnings"].append(
                _issue("warning", "generic_ai_copy", f"Slide {idx} contains generic AI/marketing copy")
            )
        if DECORATIVE_DASH_RE.search(text):
            report["warnings"].append(
                _issue("warning", "dash_tell", f"Slide {idx} contains em/en dash characters")
            )


def inspect_pptx(path: Path) -> dict:
    report = {"path": str(path), "kind": "pptx", "errors": [], "warnings": [], "metrics": {}}
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
            if "ppt/presentation.xml" not in names:
                report["errors"].append(_issue("error", "missing_presentation_xml", "Missing ppt/presentation.xml"))
                return report
            slide_names = sorted(
                [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                key=_natural_slide_key,
            )
            report["metrics"]["slide_count"] = len(slide_names)
            if not slide_names:
                report["errors"].append(_issue("error", "empty_deck", "No slide XML files found"))
                return report

            total_text = 0
            slide_texts = []
            for idx, slide_name in enumerate(slide_names, 1):
                text_runs = _xml_text(_read_zip_text(zf, slide_name))
                slide_text = " ".join(t.strip() for t in text_runs if t.strip())
                slide_texts.append(slide_text)
                total_text += len(slide_text)
                if len(slide_text) > PPT_SLIDE_TEXT_LIMIT:
                    report["warnings"].append(
                        _issue("warning", "dense_slide", f"Slide {idx} has {len(slide_text)} text characters")
                    )
                for run in text_runs:
                    clean = run.strip()
                    if len(clean) > LONG_TEXT_LIMIT:
                        report["warnings"].append(
                            _issue("warning", "long_text_run", f"Slide {idx} contains a long text run ({len(clean)} chars)")
                        )
                    if PLACEHOLDER_RE.search(clean):
                        report["warnings"].append(
                            _issue("warning", "placeholder_text", f"Slide {idx} contains placeholder-like text: {clean[:80]}")
                        )
            report["metrics"]["total_text_chars"] = total_text
            _inspect_pptx_taste(report, slide_texts)
    except zipfile.BadZipFile:
        report["errors"].append(_issue("error", "bad_zip", "File is not a valid PPTX zip package"))
    return report


def inspect_docx(path: Path) -> dict:
    report = {"path": str(path), "kind": "docx", "errors": [], "warnings": [], "metrics": {}}
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
            if "word/document.xml" not in names:
                report["errors"].append(_issue("error", "missing_document_xml", "Missing word/document.xml"))
                return report
            xml = _read_zip_text(zf, "word/document.xml")
            texts = _xml_text(xml)
            body_text = " ".join(t.strip() for t in texts if t.strip())
            report["metrics"]["text_chars"] = len(body_text)
            report["metrics"]["table_count"] = xml.count("<w:tbl>")
            report["metrics"]["image_count"] = xml.count("<w:drawing>")
            if not body_text:
                report["errors"].append(_issue("error", "empty_document", "No visible document text found"))
            for text in texts:
                clean = text.strip()
                if len(clean) > DOCX_PARAGRAPH_LIMIT:
                    report["warnings"].append(
                        _issue("warning", "long_text_run", f"DOCX contains a long text run ({len(clean)} chars)")
                    )
                if PLACEHOLDER_RE.search(clean):
                    report["warnings"].append(
                        _issue("warning", "placeholder_text", f"DOCX contains placeholder-like text: {clean[:80]}")
                    )
            tbl_count = xml.count("<w:tbl>")
            grid_count = xml.count("<w:tblGrid>")
            if tbl_count and grid_count < tbl_count:
                report["warnings"].append(
                    _issue("warning", "table_geometry", f"{tbl_count - grid_count} table(s) missing explicit tblGrid")
                )
    except zipfile.BadZipFile:
        report["errors"].append(_issue("error", "bad_zip", "File is not a valid DOCX zip package"))
    return report


def _find_office() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


def try_render(path: Path, output_dir: Path | None) -> dict:
    office = _find_office()
    if not office:
        return {"status": "skipped", "reason": "LibreOffice/soffice not found"}
    out_dir = output_dir or Path(tempfile.mkdtemp(prefix="pro_quality_render_"))
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        office,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(path),
    ]
    proc = subprocess.run(cmd, text=True, capture_output=True)
    pdf_path = out_dir / f"{path.stem}.pdf"
    if proc.returncode == 0 and pdf_path.exists() and pdf_path.stat().st_size > 1000:
        return {"status": "passed", "pdf": str(pdf_path)}
    return {
        "status": "failed",
        "returncode": proc.returncode,
        "stdout": proc.stdout[-2000:],
        "stderr": proc.stderr[-2000:],
    }


def inspect_file(path: Path, render: bool, output_dir: Path | None) -> dict:
    suffix = path.suffix.lower()
    if suffix == ".pptx":
        report = inspect_pptx(path)
    elif suffix == ".docx":
        report = inspect_docx(path)
    else:
        report = {"path": str(path), "kind": "unknown", "errors": [], "warnings": [], "metrics": {}}
        report["errors"].append(_issue("error", "unsupported_file", f"Unsupported file type: {suffix}"))
    if render and not report["errors"]:
        render_dir = output_dir / "renders" if output_dir else None
        report["render"] = try_render(path, render_dir)
        if report["render"].get("status") == "failed":
            report["warnings"].append(_issue("warning", "render_failed", "LibreOffice render failed"))
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="Run PRO-PPTX/PRO-DOCX delivery quality gates.")
    parser.add_argument("files", nargs="+", type=Path, help="PPTX/DOCX files to inspect")
    parser.add_argument("--json-report", type=Path, help="Write full JSON report to this path")
    parser.add_argument("--output-dir", type=Path, help="Directory for optional render outputs")
    parser.add_argument("--no-render", action="store_true", help="Skip optional LibreOffice PDF rendering")
    parser.add_argument("--strict", action="store_true", help="Treat warnings as failures")
    args = parser.parse_args()

    reports = [inspect_file(path, render=not args.no_render, output_dir=args.output_dir) for path in args.files]
    summary = {
        "files": reports,
        "error_count": sum(len(r["errors"]) for r in reports),
        "warning_count": sum(len(r["warnings"]) for r in reports),
    }
    if args.json_report:
        args.json_report.parent.mkdir(parents=True, exist_ok=True)
        args.json_report.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    for report in reports:
        status = "FAIL" if report["errors"] else "WARN" if report["warnings"] else "PASS"
        print(f"{status} {report['kind']} {report['path']}")
        for err in report["errors"]:
            print(f"  ERROR {err['code']}: {err['message']}")
        for warn in report["warnings"]:
            print(f"  WARN {warn['code']}: {warn['message']}")
        if "render" in report:
            render = report["render"]
            print(f"  render: {render.get('status')} {render.get('reason', render.get('pdf', ''))}")

    if summary["error_count"]:
        return 1
    if args.strict and summary["warning_count"]:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
