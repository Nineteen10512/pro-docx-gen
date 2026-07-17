"""One-command agent entrypoint for PRO-DOCX.

This wrapper keeps agents on the supported path:
- fixed package root and work directory
- UTF-8 text handling by default
- quality gate with stale-output protection
- smoke and release checks without hand-built paths
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import time
import tempfile
import uuid
from datetime import datetime, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parent
WORK_ROOT = ROOT / "_work"
JOBS_ROOT = WORK_ROOT / "jobs"


def _env() -> dict[str, str]:
    env = os.environ.copy()
    env["PYTHONPATH"] = str(ROOT)
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def _safe_print(text: str) -> None:
    enc = sys.stdout.encoding or "utf-8"
    print(text.encode(enc, errors="backslashreplace").decode(enc, errors="replace"))


def _run(cmd: list[str], cwd: Path = ROOT) -> int:
    proc = subprocess.run(cmd, cwd=cwd, env=_env(), text=True, encoding="utf-8", errors="replace")
    return proc.returncode


def _capture(cmd: list[str], cwd: Path = ROOT, env: dict[str, str] | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        cmd,
        cwd=cwd,
        env=env or _env(),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
    )


def _slug(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "-", value.strip()).strip("-._")
    return value[:80] or "docx-job"


def _job_dir(job: str) -> Path:
    return JOBS_ROOT / _slug(job)


def _select_sample_page_previews(
    previews: list[tuple[int, str]], group_size: int = 3
) -> list[tuple[int, str]]:
    """Select one page from each consecutive page group, with page 1 as anchor."""
    if group_size < 1:
        raise ValueError("group_size must be at least 1")
    return previews[::group_size] if previews else []


def _page_previews_from_result(preview: dict) -> list[tuple[int, str]]:
    page_pngs = list(preview.get("page_pngs") or [])
    if not page_pngs:
        for attempt in preview.get("attempts") or []:
            candidate = list(attempt.get("pngs") or [])
            if attempt.get("status") == "passed" and candidate:
                page_pngs = candidate
                break
    if not page_pngs:
        page_pngs = list(preview.get("pngs") or [])
    page_numbers = list(preview.get("page_numbers") or [])
    if page_numbers:
        if len(page_numbers) != len(page_pngs):
            raise ValueError("page_numbers and page_pngs must have equal length")
        return [(int(page_number), str(path)) for page_number, path in zip(page_numbers, page_pngs)]
    return [(page_number, str(path)) for page_number, path in enumerate(page_pngs, 1)]


def _build_four_up_review_bundle(
    sampled: list[tuple[int, str]], output_dir: Path, stem: str
) -> dict:
    """Create 2x2 contact sheets and a multi-page PDF for sampled review pages."""
    if not sampled:
        raise ValueError("at least one sampled page is required")
    output_dir.mkdir(parents=True, exist_ok=True)
    page_numbers = [page for page, _ in sampled]
    if len(sampled) == 1:
        return {
            "pdf": None,
            "sheet_pngs": [sampled[0][1]],
            "pages_per_sheet": 4,
            "sampled_page_numbers": page_numbers,
        }

    from PIL import Image, ImageDraw, ImageOps

    sheet_width, sheet_height = 1800, 2500
    outer_margin, gap, label_height = 48, 36, 54
    cell_width = (sheet_width - 2 * outer_margin - gap) // 2
    cell_height = (sheet_height - 2 * outer_margin - gap) // 2
    sheets: list[Image.Image] = []
    sheet_paths: list[str] = []

    for sheet_index in range(0, len(sampled), 4):
        sheet = Image.new("RGB", (sheet_width, sheet_height), "white")
        draw = ImageDraw.Draw(sheet)
        for slot, (page_number, png_path) in enumerate(sampled[sheet_index : sheet_index + 4]):
            row, column = divmod(slot, 2)
            x = outer_margin + column * (cell_width + gap)
            y = outer_margin + row * (cell_height + gap)
            with Image.open(png_path) as source:
                page = ImageOps.contain(
                    source.convert("RGB"),
                    (cell_width - 24, cell_height - label_height - 24),
                )
                page_x = x + (cell_width - page.width) // 2
                page_y = y + label_height + (cell_height - label_height - page.height) // 2
                sheet.paste(page, (page_x, page_y))
            draw.rectangle((x, y, x + cell_width, y + cell_height), outline="#A6ADB4", width=2)
            draw.text((x + 16, y + 14), f"Source page {page_number}", fill="#18222B")

        sheet_path = output_dir / f"{stem}.review_4up-{len(sheets) + 1:02d}.png"
        sheet.save(sheet_path, "PNG", optimize=True)
        sheets.append(sheet)
        sheet_paths.append(str(sheet_path))

    pdf_path = output_dir / f"{stem}.review_4up.pdf"
    try:
        sheets[0].save(
            pdf_path,
            "PDF",
            resolution=144.0,
            save_all=True,
            append_images=sheets[1:],
        )
    finally:
        for sheet in sheets:
            sheet.close()

    return {
        "pdf": str(pdf_path),
        "sheet_pngs": sheet_paths,
        "pages_per_sheet": 4,
        "sampled_page_numbers": page_numbers,
    }


def _sampled_page_previews_from_quality_data(data: dict) -> list[tuple[int, str]]:
    """Collect already-sampled source pages without sampling them a second time."""
    pages: list[tuple[int, str]] = []
    for item in data.get("files", []):
        preview = item.get("preview_png") or {}
        pages.extend(_page_previews_from_result(preview))
    return pages


def _submission_manifest(
    *,
    target: Path,
    quality_report: Path,
    document_page_count: int,
    sampled_pages: list[tuple[int, str]],
    review_bundle: dict,
    full_figure_review: bool = False,
) -> dict:
    manifest = {
        "deliverable": str(target),
        "quality_report": str(quality_report),
        "render_scope": "sampled_pages_only",
        "document_page_count": int(document_page_count),
        "source_page_pngs": [path for _, path in sampled_pages],
        "preview_pngs": review_bundle["sheet_pngs"],
        "review_pdf": review_bundle["pdf"],
        "visual_review_sampling": {
            "group_size": 3,
            "sample_per_group": 1,
            "minimum_total": 1,
            "sampled_page_numbers": review_bundle["sampled_page_numbers"],
            "pages_per_review_pdf_sheet": review_bundle["pages_per_sheet"],
        },
        "required_agent_review": [
            "The whole DOCX is structurally audited; visual rendering is limited to one source page from every consecutive three-page group.",
            "When more than one sample exists, open the four-up review PDF or its sheet PNGs; each PDF sheet contains up to four sampled source pages.",
            "Check visible layout, table readability, word/price wrapping, image/chart rendering, and obvious content mismatch.",
            "Do not submit if any preview has mojibake, broken layout, missing visuals, audit artifacts in final content, or factual/content mismatch with the user's goal.",
            "HARD GATE: If a sampled page fails, render only adjacent pages and the affected section, repair the smallest source block, then rerun deliver.",
            "If the failure is caused by reusable skill code, prompts, or quality gates, patch the smallest relevant skill area and include a short problem report.",
        ],
        "targeted_repair_policy": {
            "scope": "page_or_section_level",
            "rule": "Fix the smallest source block that produces the bad page; preserve unaffected pages. Aesthetic failure is a delivery failure.",
            "required_after_repair": "rerun agent_docx.py deliver and reopen the sampled four-up review bundle",
            "skill_defect_rule": "When bad output comes from reusable skill behavior, fix the skill with a minimal non-downgrading patch before resubmitting.",
        },
    }
    if full_figure_review:
        manifest["render_scope"] = "full_document_for_figure_review"
        manifest["visual_review_sampling"] = {
            "group_size": 1,
            "sample_per_group": 1,
            "minimum_total": int(document_page_count),
            "sampled_page_numbers": review_bundle["sampled_page_numbers"],
            "pages_per_review_pdf_sheet": review_bundle["pages_per_sheet"],
        }
        manifest["required_agent_review"] = [
            "The DOCX contains figures, so every source page is rendered.",
            "Open every four-up sheet; inspect every figure page and both immediate neighbors.",
            "Check figure legibility, captions, section transitions, blank pages, table borders, fonts, and TOC errors.",
            "Do not submit until all rendered pages have been visually reviewed.",
        ]
    return manifest


def _load_job(job: str) -> dict:
    manifest = _job_dir(job) / "manifest.json"
    if not manifest.exists():
        raise SystemExit(f"missing job manifest: {manifest}")
    return json.loads(manifest.read_text(encoding="utf-8"))


def cmd_doctor(_: argparse.Namespace) -> int:
    import pro_docx_gen

    _safe_print(f"root={ROOT}")
    _safe_print(f"work_root={WORK_ROOT}")
    _safe_print(f"version={pro_docx_gen.__version__}")
    _safe_print(f"templates={len(pro_docx_gen.list_templates())}")
    if _run([sys.executable, str(ROOT / "quality_gates" / "check_text_encoding.py"), str(ROOT)]):
        return 1
    _safe_print("doctor=PASS")
    return 0


def cmd_smoke(args: argparse.Namespace) -> int:
    out = Path(args.output_dir) if args.output_dir else WORK_ROOT / "smoke" / datetime.now().strftime("%Y%m%d-%H%M%S")
    out.mkdir(parents=True, exist_ok=True)
    cmd = [
        sys.executable,
        str(ROOT / "smoke_tests" / "run_smoke_tests.py"),
        "--output-dir",
        str(out),
        "--profile",
        args.profile,
    ]
    for target in args.target or []:
        cmd.extend(["--target", target])
    return _run(cmd)


def cmd_start(args: argparse.Namespace) -> int:
    started_at = time.time()
    job_name = _slug(args.name or datetime.now().strftime("docx-job-%Y%m%d-%H%M%S"))
    job_dir = _job_dir(job_name)
    job_dir.mkdir(parents=True, exist_ok=True)
    manifest = {
        "job": job_name,
        "started_at_epoch": started_at,
        "started_at_iso": datetime.fromtimestamp(started_at, timezone.utc).isoformat(),
        "output_dir": str(job_dir / "outputs"),
        "quality_dir": str(job_dir / "quality"),
        "quality_report": str(job_dir / "quality_report.json"),
    }
    (job_dir / "outputs").mkdir(exist_ok=True)
    (job_dir / "quality").mkdir(exist_ok=True)
    (job_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    _safe_print(str(job_dir / "manifest.json"))
    return 0


def cmd_deliver(args: argparse.Namespace) -> int:
    if args.no_render:
        _safe_print(
            "FAIL --no-render is diagnostic-only and is not allowed through agent_docx.py deliver; "
            "final delivery requires render verification and a PNG preview. "
            "Use quality_gates/run_quality_gate.py --no-render only for local debugging."
        )
        return 1

    target = Path(args.file).resolve()
    if not target.exists():
        raise SystemExit(f"missing deliverable: {target}")

    if args.job:
        manifest = _load_job(args.job)
    else:
        started_at = target.stat().st_mtime - 1
        manifest = {
            "job": "adhoc",
            "started_at_epoch": started_at,
            "quality_dir": str(WORK_ROOT / "adhoc_quality"),
            "quality_report": str(WORK_ROOT / "adhoc_quality_report.json"),
        }

    quality_dir = Path(manifest["quality_dir"])
    quality_dir.mkdir(parents=True, exist_ok=True)
    report = Path(manifest["quality_report"])
    from pro_docx_gen.docx_qa import audit_docx_layout

    layout_report = audit_docx_layout(target)
    layout_report_path = quality_dir / "docx_layout_audit.json"
    layout_report_path.write_text(json.dumps(layout_report, ensure_ascii=False, indent=2), encoding="utf-8")
    if not layout_report["passed"]:
        _safe_print(f"FAIL structural DOCX layout audit: {layout_report_path}")
        return 1
    cmd = [
        sys.executable,
        str(ROOT / "quality_gates" / "run_quality_gate.py"),
        str(target),
        "--json-report",
        str(report),
        "--output-dir",
        str(quality_dir),
        "--created-after",
        str(manifest["started_at_epoch"]),
    ]
    if not args.no_auto_fix:
        cmd.append("--auto-fix")
    rc = _run(cmd)
    if rc == 0:
        data = json.loads(report.read_text(encoding="utf-8"))
        page_previews = _sampled_page_previews_from_quality_data(data)
        if not page_previews and not args.no_render:
            _safe_print("FAIL no PNG previews found in quality report")
            return 1
        page_counts = [
            int((item.get("render") or {}).get("page_count") or 0)
            for item in data.get("files", [])
            if (item.get("render") or {}).get("page_count")
        ]
        document_page_count = max(page_counts) if page_counts else len(page_previews)
        full_figure_review = bool(layout_report["requires_full_figure_review"])
        if full_figure_review:
            from quality_gates import run_quality_gate as gate

            full_render = gate.try_render_specific_pages(
                target,
                quality_dir / "full_figure_review",
                list(range(1, document_page_count + 1)),
            )
            if full_render.get("status") != "passed":
                _safe_print("FAIL full-page render required for figure review")
                return 1
            full_previews = gate._render_sampled_pdf_previews(
                full_render, quality_dir / "full_figure_review_png"
            )
            if full_previews.get("status") != "passed":
                _safe_print("FAIL PNG conversion for full figure review")
                return 1
            page_previews = list(
                zip(full_previews["page_numbers"], full_previews["page_pngs"])
            )
        try:
            review_bundle = _build_four_up_review_bundle(
                page_previews,
                quality_dir / "review",
                target.stem,
            )
        except Exception as exc:
            _safe_print(f"FAIL sampled review bundle generation failed: {exc}")
            return 1
        submission = _submission_manifest(
            target=target,
            quality_report=report,
            document_page_count=document_page_count,
            sampled_pages=page_previews,
            review_bundle=review_bundle,
            full_figure_review=full_figure_review,
        )
        submission_path = quality_dir.parent / "submission_manifest.json"
        submission_path.write_text(json.dumps(submission, ensure_ascii=False, indent=2), encoding="utf-8")
        _safe_print(f"quality_report={report}")
        _safe_print(f"submission_manifest={submission_path}")
        if review_bundle["pdf"]:
            _safe_print(f"review_pdf={review_bundle['pdf']}")
        for png in review_bundle["sheet_pngs"]:
            _safe_print(f"review_preview_png={png}")
    return rc


def cmd_package_check(args: argparse.Namespace) -> int:
    return _run([sys.executable, str(ROOT / "quality_gates" / "check_zip_layout.py"), str(Path(args.zip).resolve())])


def cmd_read_text(args: argparse.Namespace) -> int:
    path = Path(args.file)
    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError as exc:
        _safe_print(f"FAIL utf-8 read: {path}: {exc}")
        return 1
    lines = text.splitlines()
    for line in lines[: args.lines]:
        _safe_print(line)
    if len(lines) > args.lines:
        _safe_print(f"... truncated: {len(lines) - args.lines} more line(s)")
    return 0


def cmd_clean(_: argparse.Namespace) -> int:
    if WORK_ROOT.exists():
        resolved = WORK_ROOT.resolve()
        if resolved == ROOT or ROOT not in resolved.parents:
            raise SystemExit(f"refusing cleanup outside package root: {resolved}")
        shutil.rmtree(WORK_ROOT)
    _safe_print("clean=PASS")
    return 0


def _check_item(results: list[dict], name: str, ok: bool, detail: str, required: bool = True) -> None:
    status = "PASS" if ok else "FAIL" if required else "WARN"
    results.append({"name": name, "status": status, "required": required, "detail": detail})
    _safe_print(f"{status} {name}: {detail}")


def _find_pdf_backend() -> list[str]:
    backends = []
    if os.name == "nt":
        powershell = shutil.which("powershell") or shutil.which("pwsh")
        local_app_data = Path(os.environ.get("LOCALAPPDATA", ""))
        wps_root = local_app_data / "Kingsoft" / "WPS Office"
        if powershell and wps_root.exists() and any(wps_root.glob("*/office6/wps.exe")):
            backends.append("wps_com_selective")
        word_candidates = (
            r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
            r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
        )
        if powershell and any(Path(candidate).exists() for candidate in word_candidates):
            backends.append("word_com_selective")
    return backends


def _find_png_backend() -> list[str]:
    backends = []
    if shutil.which("pdftoppm"):
        backends.append("pdftoppm")
    for parent in Path(sys.executable).resolve().parents:
        if (parent / "native" / "poppler" / "Library" / "bin" / "pdftoppm.exe").exists():
            backends.append("bundled-poppler")
            break
        if (parent / "dependencies" / "native" / "poppler" / "Library" / "bin" / "pdftoppm.exe").exists():
            backends.append("bundled-poppler")
            break
    if importlib.util.find_spec("fitz"):
        backends.append("pymupdf")
    if importlib.util.find_spec("pdf2image"):
        backends.append("pdf2image")
    return sorted(set(backends))


def _pdf_backend_fix_hint() -> str:
    return (
        "missing selective page export backend. Install WPS or Microsoft Word desktop on Windows. "
        "LibreOffice full-document export does not satisfy strict delivery. "
        "Do not skip render verification for delivery."
    )


def _png_backend_fix_hint() -> str:
    return (
        "missing PDF-to-PNG preview backend. Install Poppler/pdftoppm and ensure it is on PATH, "
        "or install PyMuPDF (`python -m pip install pymupdf`). "
        "Do not submit without a generated PNG preview."
    )


def _scan_local_path_leaks() -> list[str]:
    markers = (
        "C:" + "\\Users\\",
        "C:" + "/Users/",
        "\\AppData" + "\\Local\\Temp\\",
        "/" + "tmp/",
        "/" + "mnt/data/",
        "codex" + "-runtimes",
    )
    hits: list[str] = []
    suffixes = {".py", ".md", ".txt", ".json", ".yaml", ".yml"}
    for path in ROOT.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in suffixes:
            continue
        if any(part in {"_work", "__pycache__"} or part.startswith("_output") for part in path.parts):
            continue
        try:
            text = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            hits.append(f"{path}: not valid UTF-8")
            continue
        for marker in markers:
            if marker in text:
                hits.append(f"{path.relative_to(ROOT)} contains local marker {marker}")
                break
    return hits


def _package_only_probe(results: list[dict]) -> None:
    tmp_path = WORK_ROOT / f"pro_docx_pkg_probe_{uuid.uuid4().hex}"
    tmp_path.mkdir(parents=True, exist_ok=False)
    try:
        shutil.copytree(ROOT / "pro_docx_gen", tmp_path / "pro_docx_gen")
        env = os.environ.copy()
        env["PYTHONPATH"] = str(tmp_path)
        env["PYTHONIOENCODING"] = "utf-8"
        code = (
            "import pro_docx_gen; "
            "from pro_docx_gen import generate; "
            "assert callable(generate); "
            "assert len(pro_docx_gen.list_templates()) >= 10; "
            "print(pro_docx_gen.__version__)"
        )
        proc = _capture([sys.executable, "-c", code], cwd=tmp_path, env=env)
        detail = (proc.stdout or proc.stderr or "").strip()[-500:] or f"exit={proc.returncode}"
        _check_item(results, "package_only_import", proc.returncode == 0, detail)
    finally:
        shutil.rmtree(tmp_path, ignore_errors=True)


def cmd_portability_check(args: argparse.Namespace) -> int:
    results: list[dict] = []
    required_modules = ["docx", "matplotlib", "numpy", "lxml", "latex2mathml", "PIL"]

    _safe_print(f"root={ROOT}")
    _safe_print(f"python={sys.executable}")
    _safe_print(f"platform={sys.platform}")

    for module in required_modules:
        _check_item(
            results,
            f"required_module:{module}",
            importlib.util.find_spec(module) is not None,
            "importable" if importlib.util.find_spec(module) else "missing; run pip install -r pro_docx_gen/requirements.txt",
        )

    proc = _capture([sys.executable, str(ROOT / "quality_gates" / "check_text_encoding.py"), str(ROOT)])
    _check_item(results, "text_encoding_gate", proc.returncode == 0, (proc.stdout or proc.stderr).strip()[-500:])

    try:
        import pro_docx_gen

        detail = f"version={pro_docx_gen.__version__} templates={len(pro_docx_gen.list_templates())}"
        _check_item(results, "root_import", True, detail)
    except Exception as exc:
        _check_item(results, "root_import", False, repr(exc))

    _package_only_probe(results)

    leaks = _scan_local_path_leaks()
    _check_item(results, "local_path_leak_scan", not leaks, "; ".join(leaks[:5]) if leaks else "no local path leaks")

    pdf_backends = _find_pdf_backend()
    _check_item(
        results,
        "selective_page_export_backend",
        bool(pdf_backends),
        ", ".join(pdf_backends) if pdf_backends else _pdf_backend_fix_hint(),
    )

    png_backends = _find_png_backend()
    _check_item(
        results,
        "pdf_to_png_backend",
        bool(png_backends),
        ", ".join(png_backends) if png_backends else _png_backend_fix_hint(),
    )

    if args.zip:
        proc = _capture([sys.executable, str(ROOT / "quality_gates" / "check_zip_layout.py"), str(Path(args.zip).resolve())])
        _check_item(results, "zip_layout", proc.returncode == 0, (proc.stdout or proc.stderr).strip()[-500:])

    if args.json_report:
        report = Path(args.json_report)
        report.parent.mkdir(parents=True, exist_ok=True)
        report.write_text(json.dumps({"results": results}, ensure_ascii=False, indent=2), encoding="utf-8")
        _safe_print(f"portability_report={report}")

    failed = [item for item in results if item["status"] == "FAIL"]
    if failed:
        _safe_print(f"portability=FAIL failed={len(failed)}")
        return 1
    _safe_print("portability=PASS")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="PRO-DOCX agent-safe one-command workflow.")
    sub = parser.add_subparsers(dest="cmd", required=True)

    doctor = sub.add_parser("doctor", help="Check imports, templates, and text encoding.")
    doctor.set_defaults(func=cmd_doctor)

    smoke = sub.add_parser("smoke", help="Run gate-render smoke by default; use targeted smoke for feature issues.")
    smoke.add_argument("--output-dir")
    smoke.add_argument("--profile", choices=("gate-render", "full"), default="gate-render")
    smoke.add_argument(
        "--target",
        action="append",
        choices=("chart", "gate", "package", "render", "save", "table"),
        help="Run focused smoke for a known issue area. May be repeated.",
    )
    smoke.set_defaults(func=cmd_smoke)

    start = sub.add_parser("start", help="Create a job manifest for stale-output-safe delivery.")
    start.add_argument("name", nargs="?")
    start.set_defaults(func=cmd_start)

    deliver = sub.add_parser("deliver", help="Run strict delivery QA on a DOCX file.")
    deliver.add_argument("file")
    deliver.add_argument("--job", help="Job name returned by start.")
    deliver.add_argument("--no-auto-fix", action="store_true", help="Disable safe DOCX layout auto-fix.")
    deliver.add_argument("--no-render", action="store_true", help="Rejected by this entrypoint; diagnostic-only flag belongs to quality_gates/run_quality_gate.py.")
    deliver.set_defaults(func=cmd_deliver)

    package = sub.add_parser("package-check", help="Validate release zip layout.")
    package.add_argument("zip")
    package.set_defaults(func=cmd_package_check)

    portability = sub.add_parser("portability-check", help="Check install/import/path/render portability.")
    portability.add_argument("--zip", help="Optional release zip to validate.")
    portability.add_argument("--json-report", help="Optional JSON report path.")
    portability.set_defaults(func=cmd_portability_check)

    read_text = sub.add_parser("read-text", help="Read text as UTF-8 and print safely for Chinese diagnostics.")
    read_text.add_argument("file")
    read_text.add_argument("--lines", type=int, default=40)
    read_text.set_defaults(func=cmd_read_text)

    clean = sub.add_parser("clean", help="Remove runtime _work outputs.")
    clean.set_defaults(func=cmd_clean)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
