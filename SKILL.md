---
name: pro-docx-gen
description: Use when users need to create, edit, translate, verify, or deliver DOCX/Word files with semantic JSON or Markdown, including reports, papers, training docs, tables, charts, images, references, comments, tracked changes, translation with layout preservation, PDF/image export, and strict delivery QA. v1.7.0 adds 2 white-background templates (business_compact, research_report), 5 table-engine capabilities (conditional_format, progress_bar, star_rating, merged_header, auto_compute), KPI card row, research header, and risk disclaimer footer.
---

# PRO-DOCX v1.7.0

Use this skill for professional DOCX creation, editing, translation, visual reports, and delivery QA. v1.7.0 adds full research-report toolchain: 2 new templates, 5 table-engine capabilities, KPI cards, research header, and risk disclaimer footer.

中文触发词：生成 Word、生成 DOCX、修改 Word、翻译 Word、表格翻译、可视化报告、SWOT 报告、审核 DOCX、导出 PDF、研报、股票研报、券商研报、机构研报、评级、目标价、KPI、风险提示。

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

## v1.6.6 Self-Audit & 防复发

v1.6.6 的设计哲学是**修根因 + 自审防复发**。Three-layer defense 防止 heading 出现花哨色（金/黄/橙/亮红/浅蓝紫 #E0E7FF / 青蓝霓虹 #64FFDA / 黑金 #D4AF37 / 纯白等）再次渗透到产出物。

### 颜色铁律
- `heading` 统一 `#1A1A1A`，`text` 统一 `#333333`，`muted` 统一 `#666666`
- 这三条常量在 `pro_docx_gen/shared/color_palette.py`（`HEADING_HEX` / `TEXT_HEX` / `MUTED_HEX`）
- `resolve_palette()` 在生成 token 时**强制覆盖** theme override 中对 heading/text/muted 的赋值
- 任何模板/主题再也写不进去花哨色到正文

### 第 1 层：Renderer 强制覆盖（修根因）
- `pro_docx_gen/engine/renderer.py` 在 `__init__` 里强制把 `color_dict["heading"] / ["title"] / ["text"] / ["muted"]` 覆盖为 `color_palette.heading_rgb() / text_rgb() / muted_rgb()`
- 所有 heading 渲染点（`_render_heading` / `_setup_styles` / `table_header`）的 run color 全部走 `_heading_rgb()` 工厂函数
- 即使模板的 `register(DOCXTemplate(heading="#D4AF37"))` 也无效，token 阶段被覆盖

### 第 2 层：Quality Gate 同级标题色一致检查
- `quality_gates/run_quality_gate.py` 新增 `check_heading_color_consistency(docx_path)`
- 解析 `word/document.xml`，提取所有 heading 段落的 run color
- 检查两条：① 标题色不能命中 `FORBIDDEN_HEADING_HEXES` 黑名单；② 同级标题（H1/H2/...）必须颜色一致
- 任一违例 → ERROR `heading_color_inconsistent`，gate fail

### 第 3 层：Self-Audit 模块（防复发）
- `pro_docx_gen/self_audit.py` 是独立的自审模块，3 个公开函数 + CLI
  - `audit_template_file(path) -> AuditReport` — AST 扫描模板，遍历 `register(DOCXTemplate(...))` 的 kwargs 和 Dict 字面量，提取 heading/title/h1-h6/text/body/muted 字段的 hex 字符串，命中黑名单即 violation
  - `audit_all_templates() -> List[AuditReport]` — 遍历 `pro_docx_gen/templates/` 全部模板
  - `audit_docx_output(docx_path) -> AuditReport` — 解析 docx XML 跑产物级检查（与 gate 复用同一套逻辑）
- 用 `ast` 而非正则：结构性扫描，不可能误报 docstring 注释里的 hex
- CLI：`python -m pro_docx_gen.self_audit [--templates] [--docx PATH]`
- 集成：`python quality_gates/run_quality_gate.py --with-self-audit` 在跑 gate 之前先跑模板自审，禁用模板无法发布

### 模板精简（20 → 16）
- 删除 4 个违规模板：`brand_luxury`（黑金 #D4AF37）、`proposal_elegant`（黑金 #D4AF37）、`data_analysis_tech`（heading #E0E7FF + 深底 #0A192F）、`tech_whitepaper`（heading #E0E7FF）
- 保留 16 套白底模板（academic_corporate / academic_v151 / business_minimal / business_proposal / business_v151 / contract_cn / education_formal / lesson_plan_gaokao / marketing_bold / meeting_minutes / product_launch_bold / reading_notes / report_minimal / resume_cn / teaching_v151 / thesis_full）

### 图表空白页根治
- `pro_docx_gen/figure_preflight.py` 新增 `compute_available_height_inches(tokens, *, consumed_inches=0.0)`，从 page_height / page_margin 计算当前页剩余可用高度
- `pro_docx_gen/engine/chart_renderer.py` 的 `render_chart_to_png(...)` 新增 `available_height_inches` 参数：图表高度超过 `available * 0.9` 时按比例自动缩图（保持宽高比）
- 新增 `compact_mode=False` 参数：紧凑模式取消图表前后多余段间距（`space_before=0, space_after=0`），caption 段落绑定 `keep_with_next=True`
- 空白页根因：原代码只设 `space_before/after` 但没量剩余空间 → 撑爆当前页跳到下一页再溢出空白

### 防复发的运行时保证
- 任何后续往 `pro_docx_gen/templates/` 加新模板的人，如果不小心用了 `#D4AF37` / `#E0E7FF` / `#64FFDA` 等花哨色：
  1. 模板注册时被 `register(DOCXTemplate(...))` 接受
  2. 运行时 `renderer.__init__` 强制覆盖掉花哨色
  3. 跑 `python -m pro_docx_gen.self_audit --templates` 立刻 fail
  4. 跑 `python quality_gates/run_quality_gate.py --with-self-audit` 在生成前就 fail
- 任何后续往 renderer 加新代码绕过 `_heading_rgb()` 工厂的人：
  1. 跑 gate 时 `check_heading_color_consistency` 会在产物里抓出花哨色
  2. 跑 `audit_docx_output` 也会抓

## v1.6.6 Notes

- Added a hard pre-insertion chart/figure gate for effective font size, outer canvas whitespace, and final display resolution.
- Rejected assets now return numeric repair instructions and must be modified, rerendered, and rechecked before insertion.
- Generated chart text is clamped to a legible minimum at final DOCX size.
- **Template cleanup (20 → 16)**: removed `brand_luxury` / `proposal_elegant` (gold #D4AF37) and `data_analysis_tech` / `tech_whitepaper` (heading #E0E7FF on dark background). 16 white-background templates retained.
- **Global colour discipline** via `pro_docx_gen/shared/color_palette.py`: heading → #1A1A1A, text → #333333, muted → #666666 (forced in `resolve_palette`; `renderer.__init__` overrides theme tokens before any heading is rendered).
- **`pro_docx_gen/self_audit.py` module** (new): AST-based template scanner + DOCX XML walker; CLI `python -m pro_docx_gen.self_audit [--templates] [--docx PATH]`. Catches forbidden hex in templates before they ship and in DOCX output after generation.
- **Quality gate `--with-self-audit` flag**: runs `audit_all_templates()` before any DOCX inspection; exits 1 if any template fails the v1.6.6 colour audit.
- **`check_heading_color_consistency()` in quality gate**: scans `word/document.xml` for `<w:color>` on Heading paragraphs; fails if (a) any heading uses a colour in `FORBIDDEN_HEADING_HEXES`, or (b) same-level headings use different colours.
- **Chart page-fit + compact mode**: `figure_preflight.compute_available_height_inches()` exposes the remaining page area; `chart_renderer.render_chart_to_png(..., available_height_inches=, compact_mode=)` auto-shrinks tall charts to fit (preserving aspect ratio) and, in compact mode, drops inter-figure spacing and binds caption `keep_with_next=True`. Root-cause fix for stray blank pages after tall charts.
- **Forbidden-heading hex registry** in `color_palette.FORBIDDEN_HEADING_HEXES`: `#D4AF37`, `#E0E7FF`, `#64FFDA`, `#E7E7E7`, `#FFFFFF`, `#EEF6FF`, `#F5F5F5`, `#C9A227`, `#B8860B`, `#FFD700`. New forbidden colours can be added to this frozenset only.

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
- Warnings, render skips, and render failures block final delivery by default.
- Chart assets now render in temporary internal files, embed into DOCX, and clean up after save.
- Table quality blocks narrow headers, narrow short body labels, and long rows that can split across pages.
- Generation and translation save through temp file before replacing target. If Word/WPS locks target, close it or choose different `output_path`.
- Mojibake/encoding check is mandatory before release.

## v1.7.0 Research-Report Toolchain

v1.7.0 升级包为研报 / 财务报告 / 机构销售场景提供"全套方案"：2 套白底模板 + 5 项表格引擎能力 + 3 大渲染模块（KPI 卡片行 / 研报 header / 风险提示尾部）。保留 v1.6.6 全部三层防复发机制（renderer 强制覆盖 heading/text/muted 颜色、quality_gate 同级标题色一致、self_audit AST 扫描黑名单）。

### 2 套新白底模板（白底合规，self_audit 必过）

| name | scene | variant | 用途 |
|---|---|---|---|
| `business_compact` | business_report | compact_dense | 商务紧凑（周报/月报/销售复盘），行高压缩到 1.15，margin 2.0cm，最大化单页信息密度 |
| `research_report` | equity_research | institutional | 研报标准（券商/机构研报），margin 2.0cm，line_spacing 1.3；与研报 header / KPI / 风险提示尾部模块配套使用 |

两套模板都通过 `register(DOCXTemplate(...))` 注册在 `pro_docx_gen/templates/`，走 `_heading_rgb()` 强制覆盖，**不会** 引入黑名单色。

### 表格引擎 5 项新能力（公开模块级函数）

新增 5 个模块级公开函数（同时以同名方法挂在 `DocxRenderer` 上），均通过 `pro_docx_gen` 顶层导出：

```python
from pro_docx_gen import (
    render_conditional_cell,  # 条件格式（同比标红/标绿 / 阈值染色）
    render_progress_bar,      # 进度条（0-100% 纯色矩形 + 灰底）
    render_star_rating,       # 星级评分（unicode ★/☆，半星支持）
    render_merged_header,     # 多表头合并（行/列二维合并）
    auto_compute_rows,        # 自动计算（YoY 增长列 + sum/avg 汇总行）
    compute_yoy,              # YoY 算子（纯函数，无 IO）
    compute_summary,          # 汇总行算子（纯函数，无 IO）
    yoy_color_for,            # 同比色查找（值 → hex）
)
```

**条件格式 (conditional_format)**：单元格数据级根据数值或 YoY% 决定底色/字色。`type: "yoy"` 默认绑定同比语义：正增长标绿 #0E7C3A / 负增长标红 #B91C1C / 零增长标灰 #666666 / 不可计算标灰；`type: "threshold"` 支持多档阈值。**所有 4 色都不在 self_audit 黑名单**，且为 run/cell-shading 颜色，不是 heading 字段。

**进度条 (progress_bar)**：用 unicode 块字符 `█` / `░`（20 格）画 0-100% 进度条，附 ` 85%` 标签。不依赖 matplotlib，避免插入图片与字体测量；Word / WPS 都正确显示。

**星级评分 (star_rating)**：用 unicode `★` / `☆` / `⯨`（半星）拼 1-10 颗星，默认 5 颗。`D97706` 暖琥珀色（不在黑名单）。

**多表头合并 (merged_header)**：二维表头，行/列自动合并。约定：行内连续空字符串触发水平合并（继承左侧文字）；列内上下相同文字触发垂直合并（继承上方文字）；复杂合并通过 `merges=[(r,c), ...]` 显式指定。

**自动计算 (auto_compute_rows)**：数据驱动，返回新 rows + meta 字典。`yoy_col_index` 加 YoY 增长列（接受显式 `previous_rows`）；`summary_mode="sum"/"avg"` 在数据行前加汇总行。底层 `compute_yoy` / `compute_summary` 是纯函数，可独立 import。

### 3 大渲染模块（公开模块级函数）

```python
from pro_docx_gen import (
    render_kpi_card_row,             # KPI 卡片行（3-5 卡片，label/value/delta/sparkline）
    render_research_header,          # 研报 header（4 字段强制：股票代码/评级/目标价/日期）
    render_rating_badge,             # 评级徽章（Overweight/Hold/Underweight 三色）
    render_risk_disclaimer_footer,   # 风险提示尾部（固定模板 + 灰色斜体）
)
```

**KPI 卡片行 (`render_kpi_card_row`)**：3-5 张横向卡片，每张含 label（灰 9pt）/ value（深黑 #1A1A1A 20pt 加粗）/ delta（▲▼ 同比染色）/ 可选 sparkline（unicode 块字符 7 列趋势图）。卡间用细灰线分隔。颜色：value #1A1A1A / label #666666 / 正向 #0E7C3A / 负向 #B91C1C。

**研报 header (`render_research_header`)**：**4 字段强制** —— `stock_code` / `rating` / `target_price` / `report_date`，缺一即抛 `ValueError`。构造：meta 行（灰斜体）/ 标题行（深蓝大字 + 股票代码 + 评级徽章）/ 副标题（紫红 #6B2C91 斜体）/ 三行 meta（目标价 / 当前价 / 上行空间，自动算 upside%）。评级徽章底色随级别变化：Overweight #0E7C3A / Hold #B45309 / Underweight #B91C1C（**3 色均不在黑名单**）。

**评级徽章 (`render_rating_badge`)**：单格徽章，1.0×0.4 in。接受 Overweight/Hold/Underweight + 别名（Buy → Overweight, Neutral → Hold, Sell → Underweight 等），中英文切换。

**风险提示尾部 (`render_risk_disclaimer_footer`)**：固定模板"本报告仅供研究参考，不构成投资建议……"灰色斜体 8pt，自动出现在每份研报末尾。支持中英文 + 自定义机构名。

### 颜色策略（v1.7.0 研报色板）

| 用途 | hex | 自审 |
|---|---|---|
| 评级 Overweight | #0E7C3A | ✓ 不在黑名单 |
| 评级 Hold | #B45309 | ✓ 不在黑名单 |
| 评级 Underweight | #B91C1C | ✓ 不在黑名单 |
| 研报深蓝（继承 academic primary） | #1F3864 | ✓ 不在黑名单 |
| 紫红副标题 | #6B2C91 | ✓ 不在黑名单 |
| 同比正 / 负 / 零 | #0E7C3A / #B91C1C / #666666 | ✓ 全部不在黑名单 |
| 进度条 / 星级 | #0E7C3A / #D97706 | ✓ 不在黑名单 |
| 风险提示 / meta 灰 | #666666 / #999999 | ✓ muted 同色 |

`color_palette.HEADING_HEX / TEXT_HEX / MUTED_HEX / FORBIDDEN_HEADING_HEXES` 全部**保持 v1.6.6 原值不动**。`FORBIDDEN_HEADING_HEXES` 仍为 10 项（不删不增）。

### 保留 v1.6.6 三层防复发机制

1. **Renderer 强制覆盖（修根因）**：`DocxRenderer.__init__` 仍把 `color_dict["heading"]/["title"]/["text"]/["muted"]` 覆盖为 `_heading_rgb()/_text_rgb()/_muted_rgb()`；`_setup_styles` 与 `_render_heading` 同样走工厂函数。新模板 theme_overrides 写 `#D4AF37` 也无效。
2. **Quality Gate 同级标题色一致**：`quality_gates/run_quality_gate.py::check_heading_color_consistency` 解析 `word/document.xml` 抓 heading run color，命中黑名单 / 同级多色 → ERROR。
3. **Self-Audit AST 扫描**：`pro_docx_gen/self_audit.py` 用 `ast.walk` 扫描 `templates/*.py`，任何新模板用 `register(DOCXTemplate(heading="#D4AF37"))` 都会被 `python -m pro_docx_gen.self_audit --templates` 立刻 fail。`--with-self-audit` 在 `quality_gate` 跑前先跑模板自审。

### 验证产物

`pro_docx_gen/v17_demo/` 放 3 份代表性 demo：
- `business_compact_demo.docx`（商务紧凑报告：单行表格 + 提要框 + 条件格式）
- `research_report_demo.docx`（研报标准：4 字段 header + KPI 卡片行 + YoY/汇总表格 + 风险提示尾部）
- `table_advanced_features_demo.docx`（5 项新能力专项：条件格式 / 进度条 / 星级 / 多表头合并 / 自动计算 YoY·汇总 + 评级徽章 + 风险提示）

每份 demo 都通过 `python -m pro_docx_gen.self_audit --docx` 0 violations、python-docx 正常打开、Heading 颜色统一 #1A1A1A。

`pro_docx_gen/v17_test_report.txt` 记录完整验证报告：18 模板 self_audit 全过 + 3 份 demo self_audit 0 violations + quality_gate 全过 + 黑名单 hex 残留 = 0 + 评级三色 + 同比红绿均不命中黑名单 + 边界测试（合成 #D4AF37 docx 必被拦截）。

### v1.7.0 验收对照

- ✓ 18 模板 self_audit 0 violations（16 旧 + 2 新）
- ✓ 3 份 demo docx 都能用 python-docx 打开、Heading 颜色统一 #1A1A1A
- ✓ 黑名单 hex 残留 = 0
- ✓ 评级三色 + 同比红绿不命中黑名单
- ✓ 合成禁色 docx → self_audit exit 1
- ✓ 任何模板加新禁用色时 `--with-self-audit` 仍能拦截

### v1.7.0 Notes

- **2 套新白底模板**：`business_compact`（compact_dense 商务紧凑）+ `research_report`（institutional 研报标准）。总计 16 → 18 模板，全部走 `_heading_rgb()` 强制覆盖，self_audit 必过。
- **表格引擎 5 项新能力**：`render_conditional_cell`（同比标红/标绿 + 阈值染色）/ `render_progress_bar`（unicode 块字符 0-100% 进度条）/ `render_star_rating`（unicode ★/☆ + 半星）/ `render_merged_header`（二维表头行/列自动合并）/ `auto_compute_rows`（YoY 列 + sum/avg 汇总行）。底层 `compute_yoy` / `compute_summary` 是纯函数 helper。
- **KPI 卡片行**：`render_kpi_card_row` 横向 3-5 卡，label/value/delta/sparkline 四层；delta 用 ▲▼ 符号 + 同比染色（正绿/负红/0灰/不可计算灰）；trend mini 图用 unicode ▁▂▃▄▅▆▇█ sparkline。
- **研报 header**：`render_research_header` 强制 4 字段 `stock_code/rating/target_price/report_date`，缺一抛 `ValueError`；自动算 upside%；副标题紫红斜体；评级徽章三色（绿/琥珀/红）按级别切换。
- **风险提示尾部**：`render_risk_disclaimer_footer` 固定模板文案，灰色斜体 8pt，中英文 + 自定义机构名。
- **v1.6.6 三层防复发保留**：renderer 强制覆盖 / quality_gate 同级标题色一致 / self_audit AST 扫描全部保留；`FORBIDDEN_HEADING_HEXES` 仍 10 项不删不增。
- **颜色铁律不动**：`HEADING_HEX / TEXT_HEX / MUTED_HEX` 三个常量保持 v1.6.6 值；`resolve_palette()` 仍强制覆盖 heading/text/muted；新色（评级三色 / 同比色 / 紫红 / 深蓝 / 暖琥珀）只用于 run-color / cell-shading / cover-style，不进 heading 字段。
- **修根因不掩盖**：所有 helper 显式 raise `ValueError`（不合规参数），不吞错；auto_calc 对零分母 / 非数值 / 空 rows 显式降级 + 写 meta 字段。
- **不删任何 16 模板**：`brand_luxury / proposal_elegant / data_analysis_tech / tech_whitepaper` v1.6.6 已删，本次保持。
- **不改 quality_gates 既有结构**：v1.7.0 仅是新增模块 + 新增模板；`run_quality_gate.py` / `check_zip_layout.py` / `check_text_encoding.py` 全部不动。
- **公开 API 在 `pro_docx_gen` 顶层导出**：`render_conditional_cell / render_progress_bar / render_star_rating / render_merged_header / auto_compute_rows / render_kpi_card_row / render_research_header / render_rating_badge / render_risk_disclaimer_footer / compute_yoy / compute_summary / auto_compute_rows_helper / yoy_color_for`。
- **18 模板注册入口**：`pro_docx_gen.templates.registry.list_templates()` 返回 18 项；`template_name="business_compact"` / `template_name="research_report"` 走 `generate(...)` 直接生效。
- **demo 生成脚本**：`pro_docx_gen/v17_make_demos.py` 重新生成 3 份 demo，便于回归验证。
- **依赖零增加**：v1.7.0 不引入新 Python 依赖；条件格式 / 进度条 / 星级 / 表头合并 / KPI 卡片 / 研报 header 全部用 python-docx 原生 OOXML 实现。
- **版本号**：`pro_docx_gen.__version__ = "1.7.0"`；SKILL.md 顶部改为 v1.7.0；中文触发词新增"研报 / 股票研报 / 券商研报 / 机构研报 / 评级 / 目标价 / KPI / 风险提示"。
