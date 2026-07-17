---
name: pro-docx-gen
description: Use when users need to create, edit, translate, verify, or deliver DOCX/Word files with semantic JSON or Markdown, including reports, papers, training docs, tables, charts, images, references, comments, tracked changes, translation with layout preservation, PDF/image export, and strict delivery QA.
---

# PRO-DOCX v1.6.6

Use this skill for professional DOCX creation, editing, translation, visual reports, and delivery QA.

中文触发词：生成 Word、生成 DOCX、修改 Word、翻译 Word、表格翻译、可视化报告、SWOT 报告、审核 DOCX、导出 PDF。

## Required Dependencies

Install from package root:

```bash
pip install -r pro_docx_gen/requirements.txt
```

Required:

- `python-docx`
- `matplotlib`
- `numpy`
- `lxml`
- `latex2mathml`
- `Pillow`

If Windows resolves `python` to Microsoft Store stub, run commands with the real active interpreter path.

## Agent Quick Path

Use the root entrypoint before reading deeper internals:

```bash
python agent_docx.py doctor
python agent_docx.py portability-check
python agent_docx.py start my-report
# create/edit/translate the DOCX, then:
python agent_docx.py deliver path/to/output.docx --job my-report
```

`agent_docx.py` fixes `PYTHONPATH`, forces `PYTHONIOENCODING=utf-8`, stores runtime files under `_work/`, runs strict QA with stale-output protection, and applies safe layout auto-fix by default.

Run `portability-check` on unfamiliar devices before promising delivery. It verifies required Python modules, UTF-8 handling, root/package-only imports, local path leaks, optional release zip layout, selective WPS/Word page export, and PDF-to-PNG preview capability.

If it reports a missing render backend, stop and have the agent/user install one before delivery. Strict DOCX delivery needs a selective page export backend through WPS or Microsoft Word COM; LibreOffice full-document export is not a strict-delivery fallback. PDF-to-PNG needs Poppler/pdftoppm or PyMuPDF. `--no-render` is diagnostic only and cannot be used for final delivery.

Run `audit_docx_layout()` before rendering. It blocks dangling TOC bookmarks, direct cell border palette leaks, redundant page breaks before page-break headings, and adjacent section breaks that can create blank pages. It also reports role-based font-size drift and figures displayed too small.

For documents without figures, inspect one page from every consecutive three-page group (pages 1, 4, 7, ...), with a minimum one page. For documents with figures, `deliver` renders every page: inspect every figure page and both immediate neighbors. Script pass is not enough.

When more than one page must be reviewed, `deliver` automatically creates a 2x2 review PDF with up to four sampled pages per PDF sheet and matching sheet PNGs. Open that bundle before final response and verify visible aesthetics, alignment, table readability, word/price wrapping, image/chart rendering, mojibake, and obvious content mismatch.

Hard gate: if a reviewed page fails visual aesthetics or content review, fix only the smallest affected source page/section/block, then rerun `python agent_docx.py deliver ...` and reopen the four-up bundle. Do not ship the defect.

For targeted defect expansion, import `quality_gates/run_quality_gate.py` and call `try_render_specific_pages(docx_path, output_dir, [page_numbers])`. For figure-heavy QA, pass every page number; single-page PDFs are rasterized through short aliases such as `p0001.pdf` to avoid Windows long-path failures.

If the failed preview reveals a reusable skill defect in code, prompts, or quality gates, patch the smallest relevant skill area and include a short problem report. Do not hide the defect by manually editing only the final output.

Chinese text must be read as UTF-8. If the terminal shows garbled Chinese, confirm with:

```bash
python agent_docx.py read-text pro_docx_gen/SKILL.md --lines 40
python quality_gates/check_text_encoding.py .
```

## Core Rule

Source file and user goal are truth. Never invent facts, reviews, claims, tables, images, citations, or translated content that did not exist in the source or verified research material.

If any mojibake or garbled text appears, stop and investigate encoding cause. Do not skip garbled regions. Fix the text or remove it only when it is obsolete documentation.

## Generate DOCX

Use semantic JSON or Markdown. LLM writes meaning; renderer computes fonts, spacing, colors, table widths, captions, images, and charts.

Chart readability is mandatory. For single-series column/bar comparisons, keep data labels on by default. If category labels are long or dense, let the renderer auto-switch column charts to horizontal bars instead of shipping overlapping axis labels.

### Pre-insertion figure gate

Every rendered chart and text-bearing figure must pass `assert_figure_asset_ready()` before `add_picture`. Failure is a hard stop: never insert the asset, never replace the failure with a placeholder, and never lower thresholds to finish a job.

The gate blocks:

- effective text below 9 pt at the planned DOCX size;
- any outer edge with more than 18% blank canvas;
- display resolution below 150 PPI;
- text-bearing figures without declared font-size metadata.

For external diagrams, declare final and source geometry:

```python
{
    "type": "figure",
    "path": "diagram.png",
    "contains_text": True,
    "min_text_pt": 12,
    "source_width_inches": 9.4,
    "width_inches": 9.4,
}
```

Run an asset alone when diagnosing:

```bash
python figure_preflight.py diagram.png --display-width 9.4 --source-width 9.4 --contains-text --min-text-pt 12 --json-report figure-preflight.json
```

On rejection, follow every returned `remediation_steps` item: crop to the reported foreground box plus 3–5% padding, raise the smallest source font to the reported minimum, enlarge the insertion/use a landscape section, simplify or split dense content, and rerender at 200–300 PPI. Rerun preflight; insert only after `passed: true`.

```python
from pro_docx_gen import generate, generate_from_markdown

doc = {
    "meta": {"title": "Report", "author": "PRO-DOCX"},
    "theme": "premium",
    "sections": [
        {
            "title": "Executive Summary",
            "level": 1,
            "content": [
                {"type": "paragraph", "text": "Concise evidence-backed summary."},
                {
                    "type": "table",
                    "caption": "Action plan",
                    "headers": ["Layer", "Move", "Priority"],
                    "col_widths": [1.5, 4.6, 1.3],
                    "rows": [["Proof", "Publish repeatable tests.", "High"]],
                    "header_repeat": True,
                },
            ],
        }
    ],
}

generate(doc, "output/report.docx", theme="premium", lang="en")
```

Supported common nodes:

- `heading`, `paragraph`, `list`, `table`, `figure`, `chart`, `kpi_card`, `callout`, `page_break`, `toc`, `references`, `appendix`
- `revision`, `comment`, `footnote`, `endnote`, `watermark`, `page_border`, `equation`, `signature_block`, `signature_line`, `svg_shape`

Callout variants are strict: use only `info`, `warning`, `success`, or `danger`. Do not use unsupported aliases such as `important`, `note`, or `tip`; the validator must reject them.

## Translate DOCX

Use structure-preserving translation when the user wants Chinese to English, English to Chinese, or bilingual DOCX output.

1. Treat original DOCX as visual source of truth.
2. Prefer editing a copy over regenerating from scratch.
3. Collect segments with `collect_translation_segments(docx_path)`.
4. Translate only returned text payloads.
5. Preserve numbers, dates, units, formula syntax, citations, figure numbers, brand names, images, tables, headers, footers, and nearby semantic placement.
6. Run `assess_translation_risk(docx_path, translations)` before write-back when table content expands.
7. Write back with `apply_translation_map(..., auto_format_tables=True)`.

Table translation layer:

- Translate table cells; do not skip tables.
- Preserve row and column meaning exactly.
- Let engine judge layout risk from cell growth and column density.
- Medium risk: allow mild font shrink and autofit.
- High risk: allow stronger autofit plus proportionate tightening, but do not delete meaning to save space.

## Delivery Quality Gate

Before delivering any DOCX:

```bash
python quality_gates/run_quality_gate.py output.docx --json-report output/quality_report.json --output-dir output/quality
```

Strict defaults:

- Warnings fail.
- Render failure fails.
- Render skip fails unless diagnostic-only `--allow-render-skip` is explicitly used.
- WPS selective page export runs first; Microsoft Word selective page export is fallback.
- Table visual word breaks fail, including short labels and prices such as `Plane`, `Context`, and `$169.50`.

If a builder/generation command runs before the gate, record generation start time and pass it into the gate:

```bash
python quality_gates/run_quality_gate.py output.docx --created-after 1783512000 --json-report output/quality_report.json --output-dir output/quality
```

`--created-after` fails with `stale_output` when the DOCX was not modified after generation started. This prevents a failed generator from accidentally validating an old output file.

Safe auto-fix:

```bash
python quality_gates/run_quality_gate.py output.docx --auto-fix --json-report output/quality_report.json --output-dir output/quality
```

`--auto-fix` may repair only safe layout defects:

- `table_header_narrow`: narrow table headers likely to wrap badly.
- `table_body_word_narrow`: narrow body labels, short words, or prices likely to split visually.
- `table_row_can_split`: long rows may split across pages.
- `callout_row_can_split`: one-cell callout/note blocks may split across pages.
- `table_orphan_split_risk`: compact prose tables may orphan rows across pages.

After auto-fix, run quality gate again. No warning/error may remain in final delivery.

## Encoding Quality Gate

Run before release:

```bash
python quality_gates/check_text_encoding.py .
```

This blocks common UTF-8/CP936 mojibake markers and Unicode replacement characters. If it fails, investigate source encoding and fix the text. Do not whitelist garbled production text.

## Smoke Test

Run from package root:

```bash
python smoke_tests/run_smoke_tests.py
python smoke_tests/run_smoke_tests.py --target table
python smoke_tests/run_smoke_tests.py --target chart
python smoke_tests/run_smoke_tests.py --profile full
```

Smoke checks:

Default smoke runs only gate-layer and render-layer checks:

- Gate layer: imports/version, package-only import compatibility, warning/fail hard stops, stale output blocking, locked target save handling, UTF-8/mojibake gate, portability check, release zip layout, and shared compatibility wrappers.
- Render layer: DOCX generation, selective one-in-three source-page export, sampled PDF-to-PNG verification, four-up review PDF/PNG generation, and agent quick-path doctor/start/deliver/read-text workflow.

Feature-specific smoke is targeted:

- `--target table`: table header/body/row/callout/orphan auto-fix, short word and price break blocking.
- `--target chart`: chart embedding, chart readability defaults, and no leaked `chart_assets`.
- `--target save`, `--target package`, `--target gate`, `--target render`: focused layers for known issue areas.
- `--profile full`: release checks or broad refactors only.

## Release Zip

Expected zip shape:

```text
pro-docx-gen/
  pro_docx_gen/
  quality_gates/
  smoke_tests/
  README.md
  CHANGELOG.md
```

Validate:

```bash
python quality_gates/check_zip_layout.py pro-docx-gen-v1.6.6-chart-legibility-gate.zip
```

Release zips must not contain runtime `_work`, smoke `_output*` directories, generated documents/images/archives, `chart_assets`, cache files, temporary files, or local absolute paths such as user-home and temp-runtime paths.

Forbidden downgrade paths:

- Do not remove full-file structural verification, conditional render scope, one-in-three sampling for figure-free documents, full-page rendering for documents with figures, or the four-up review bundle from final delivery.
- Do not allow warnings/errors to pass in final gates.
- Do not remove the top-level `shared/` compatibility package; use wrappers if slimming duplicate code.
- Do not make feature-specific smoke the default; default smoke must remain gate plus render.

## v1.6.6 Notes

- Added a hard pre-insertion chart/figure gate for effective font size, outer canvas whitespace, and final display resolution.
- Rejected assets now return numeric repair instructions and must be modified, rerendered, and rechecked before insertion.
- Generated chart text is clamped to a legible minimum at final DOCX size.

## v1.6.5 Notes

- Added structural layout auditing for TOC bookmarks, direct cell borders, redundant page breaks, adjacent section breaks, font-size drift, and undersized figures.
- Documents containing figures now render every page; figure-free documents retain one-in-three sampling.
- Delivery manifests record full-document figure review, and regression tests cover the newly blocked defects.

## v1.6.4 Notes

- Strict delivery now performs full-file structural checks while rendering sampled pages only.
- WPS/Word COM exports source pages 1, 4, 7, ... directly; no complete PDF is created.
- Missing selective page export support fails strict delivery instead of falling back to full-document rendering.
- Sampled PNGs and four-up sheets preserve original source page numbers.

## v1.6.3 Notes

- Unified public version to `1.6.3`.
- Fixed package-only import compatibility by preferring package-local `pro_docx_gen.shared` imports over top-level `shared` path assumptions.
- Tightened release packaging guard to block smoke outputs, generated artifacts, cache/temp files, and leaked local absolute paths.
- Added `agent_docx.py` one-command workflow for doctor, job start, strict delivery QA, release package check, UTF-8 text reading, and runtime cleanup.
- Added `agent_docx.py portability-check` for cross-device install/import/path/render readiness.
- Missing render backends now produce actionable install guidance instead of silent degradation.
- `agent_docx.py deliver --no-render` is rejected so final delivery cannot bypass render verification or PNG preview generation.
- Visual failures caused by reusable skill behavior require a minimal skill patch plus a short problem report.
- Default smoke now runs only gate and render layers; feature-specific smoke is invoked with `--target`.
- Top-level duplicate `shared` modules can be slimmed only through compatibility wrappers, not removal.
- Added mandatory combined submission PNG preview generation and final agent visual review handoff.
- Added targeted repair policy: failed review pages/sections must be locally fixed and rechecked, not ignored or globally regenerated.
- Added multi-backend render fallback: LibreOffice/soffice paths, Word COM on Windows, and PNG preview through `pdftoppm`, PyMuPDF, or pdf2image when available.
- Added chart readability protections: dense single-series column charts auto-switch to horizontal bars.
- Single-series column/bar charts now default to complete data labels.
- Added callout split and compact prose table orphan gates with safe auto-fix.
- Clarified strict callout variants: only `info`, `warning`, `success`, and `danger` are valid.
- Fixed renderer comment drift that mentioned unsupported callout aliases.
- Added strict smoke suite and zip layout checker.
- Warnings, render skips, and render failures block final delivery by default.
- Chart assets now render in temporary internal files, embed into DOCX, and clean up after save.
- Table quality blocks narrow headers, narrow short body labels, and long rows that can split across pages.
- Generation and translation save through temp file before replacing target. If Word/WPS locks target, close it or choose different `output_path`.
- Mojibake/encoding check is mandatory before release.
