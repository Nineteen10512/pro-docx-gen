# PRO-DOCX v1.6.6

Professional DOCX generation, editing, translation, and delivery QA skill.

## Install

Install dependencies from:

```bash
pip install -r pro_docx_gen/requirements.txt
```

Required dependencies:

```text
python-docx>=0.8.11
matplotlib>=3.5
numpy>=1.20
lxml>=4.9
latex2mathml>=3.0.0
Pillow>=9.0
```

## Agent Quick Path

Run from package root. Do not hand-build paths unless the command asks for one.

```bash
python agent_docx.py doctor
python agent_docx.py portability-check
python agent_docx.py start my-report
# generate or edit the DOCX, then:
python agent_docx.py deliver path/to/output.docx --job my-report
```

The entrypoint fixes `PYTHONPATH`, forces `PYTHONIOENCODING=utf-8`, stores runtime files under `_work/`, runs strict delivery QA with `--created-after`, and applies safe DOCX layout auto-fix by default. Warnings and errors still fail.

`deliver` first runs structural layout QA for TOC bookmarks, direct cell borders, redundant page breaks, adjacent section breaks, role font-size drift, and undersized figures. Documents without figures render one page from every consecutive three-page group. Documents with figures render every page so every figure page and both neighbors can be reviewed. Review bundles use 2x2 sheets with up to four source pages.

Hard gate: if any reviewed page fails aesthetically or semantically, repair the smallest affected source block/page/section and rerun `deliver`.

Use `try_render_specific_pages(docx_path, output_dir, [page_numbers])` from `quality_gates/run_quality_gate.py` for adjacent-page expansion. It exports only the explicit pages through WPS/Word COM. PNG conversion uses short sampled-PDF aliases to avoid Windows long-path failures.

If the visual failure is caused by reusable skill behavior, patch the smallest relevant skill code/prompt/gate and include a short problem report. Do not hide a skill defect by manually editing only the final output.

Useful commands:

```bash
python agent_docx.py smoke
python agent_docx.py portability-check --zip pro-docx-gen-v1.6.6-chart-legibility-gate.zip --json-report _work/portability_report.json
python agent_docx.py package-check pro-docx-gen-v1.6.6-chart-legibility-gate.zip
python agent_docx.py read-text pro_docx_gen/SKILL.md --lines 40
python agent_docx.py clean
```

Default smoke is intentionally limited to the gate layer and render layer. If a specific feature has a problem, run a focused target instead:

```bash
python agent_docx.py smoke --target table
python agent_docx.py smoke --target chart
python agent_docx.py smoke --profile full
```

`--profile full` is for release checks or broad refactors. Do not use feature-specific smoke as the default path.

`portability-check` validates required Python modules, UTF-8 text handling, root import, package-only import, local path leaks, release zip layout, selective WPS/Word page export, and PDF-to-PNG preview availability. A strict delivery-capable device needs WPS or Microsoft Word COM plus Poppler/pdftoppm, PyMuPDF, or pdf2image.

If `portability-check` reports a missing render backend, install WPS or Microsoft Word desktop on Windows so the selective page export backend is available. LibreOffice full-document conversion does not satisfy strict delivery. For PDF-to-PNG, install Poppler/pdftoppm or PyMuPDF (`python -m pip install pymupdf`). `--no-render` is diagnostic only and must not be used for final delivery.

For Chinese text: read source files as UTF-8, not console-default ANSI/GBK. PowerShell may display valid UTF-8 Chinese as mojibake; confirm with `python agent_docx.py read-text ...` or `python quality_gates/check_text_encoding.py .` before editing.

## Smoke Test

Run from package root:

```bash
python smoke_tests/run_smoke_tests.py
```

The default smoke suite checks gate-layer behavior and render-layer behavior: imports, package-only import compatibility, version, strict gate failures, zip layout, text encoding, portability, DOCX rendering, one-in-three sampling, and four-up review PDF/PNG generation. Feature-specific smoke is targeted with `--target table`, `--target chart`, `--target save`, `--target package`, `--target gate`, or `--target render`.

If Windows resolves `python` to the Microsoft Store stub, run the same command with the active environment's real interpreter path.

## Text Encoding Gate

Run before packaging:

```bash
python quality_gates/check_text_encoding.py .
```

This blocks common UTF-8/CP936 mojibake and replacement characters. If it fails, fix encoding cause and text content; do not skip garbled regions.

## Delivery Quality Gate

Run before delivering any DOCX:

```bash
python quality_gates/run_quality_gate.py output.docx --json-report output/quality_report.json --output-dir output/quality
```

Warnings fail by default. Render verification is mandatory by default. The gate uses LibreOffice first, then Word COM on Windows.

To repair safe DOCX table layout issues detected by the gate:

```bash
python quality_gates/run_quality_gate.py output.docx --auto-fix --json-report output/quality_report.json --output-dir output/quality
```

`--auto-fix` repairs safe layout defects only: narrow table headers, narrow short body labels, and long rows that can split across pages. It re-runs the quality gate after editing.

DOCX generation and translation save through a temporary sibling file, then replace the target. If Word/WPS has the target open, the skill raises a clear locked-file error instead of risking a partially written document.

## Release Zip Layout

Expected release shape:

```text
pro-docx-gen/
  SKILL.md
  pro_docx_gen/
  quality_gates/
  smoke_tests/
  codex_adaptation/
  README.md
  CHANGELOG.md
```

Validate zip before publishing:

```bash
python quality_gates/check_zip_layout.py pro-docx-gen-v1.6.6-chart-legibility-gate.zip
```

Release zips must not contain runtime `_work`, smoke `_output*` directories, generated documents/images/archives, `chart_assets`, cache files, temporary files, or local absolute paths such as user-home and temp-runtime paths.

Forbidden downgrade paths:

- Do not remove full-file structural verification, conditional render scope, one-in-three sampling for figure-free documents, full-page rendering for documents with figures, or the four-up review bundle from final delivery.
- Do not allow warnings/errors to pass in final gates.
- Do not remove the top-level `shared/` compatibility package; use wrappers if slimming duplicate code.
- Do not make feature-specific smoke the default; default smoke must remain gate plus render.
