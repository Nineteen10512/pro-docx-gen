# Changelog

## v1.7.0 research-report toolchain - 2026-07-15

- **2 套新白底模板**（总计 16 → 18 套）：`business_compact`（compact_dense 商务紧凑：周报/月报/销售复盘，行高 1.15，margin 2.0cm）+ `research_report`（institutional 研报标准：margin 2.0cm，line_spacing 1.3）。
- **表格引擎 5 项新能力（公开模块级函数）**：`render_conditional_cell`（同比标红/标绿 + 阈值染色）/ `render_progress_bar`（unicode 块字符 0-100% 进度条）/ `render_star_rating`（unicode ★/☆ + 半星 ⯨）/ `render_merged_header`（二维表头行/列自动合并）/ `auto_compute_rows`（YoY 列 + sum/avg 汇总行）。底层 helper `compute_yoy` / `compute_summary` / `yoy_color_for` 是纯函数。
- **3 大渲染模块（公开模块级函数）**：`render_kpi_card_row`（横向 3-5 卡，label/value/delta/sparkline 四层；delta ▲▼+ 同比染色；sparkline 用 ▁▂▃▄▅▆▇█ unicode 块字符无 matplotlib 依赖）/ `render_research_header`（4 字段强制 stock_code/rating/target_price/report_date，缺一抛 `ValueError`；自动算 upside%；副标题紫红斜体）/ `render_rating_badge`（Overweight/Hold/Underweight 三色徽章；接受 Buy/Neutral/Sell 等别名）/ `render_risk_disclaimer_footer`（固定模板"本报告仅供研究参考，不构成投资建议..."灰色斜体 8pt，中英文 + 自定义机构名）。
- **研报色板（v1.7.0）**：评级三色 `#0E7C3A` 绿 / `#B45309` 琥珀 / `#B91C1C` 红；同比 `#0E7C3A / #B91C1C / #666666`（正/负/零）；研报深蓝 `#1F3864`（继承 academic primary）；紫红副标题 `#6B2C91`；暖琥珀星级 `#D97706`。**全部不在 self_audit 黑名单**，且只用于 run-color / cell-shading，不进 heading 字段。
- **公开 API 在 `pro_docx_gen` 顶层导出**：`render_conditional_cell / render_progress_bar / render_star_rating / render_merged_header / auto_compute_rows / render_kpi_card_row / render_research_header / render_rating_badge / render_risk_disclaimer_footer / compute_yoy / compute_summary / auto_compute_rows_helper / yoy_color_for`。
- **v1.6.6 三层防复发保留**：renderer 强制覆盖（`DocxRenderer.__init__` 覆盖 `color_dict[heading/title/text/muted]` 为 `_heading_rgb()/_text_rgb()/_muted_rgb()`，模板/主题写啥色都进不了 heading）+ quality_gate 同级标题色一致（`check_heading_color_consistency` 抓禁色 + 同级多色）+ self_audit AST 扫描（用 `ast.walk` 抓 `register(DOCXTemplate(...))` 命中禁色）。
- **颜色铁律不动**：`HEADING_HEX / TEXT_HEX / MUTED_HEX` 保持 v1.6.6 原值；`FORBIDDEN_HEADING_HEXES` 仍 10 项不删不增。
- **修根因不掩盖**：所有 helper 显式 `raise ValueError`（不合规参数），不吞错；`auto_calc` 对零分母 / 非数值 / 空 rows 显式降级 + 写 meta 字段，调用方能 audit。
- **不删任何 16 模板**：v1.6.6 删的 4 个保持删；v1.7.0 仅新增 2 套。
- **不改 quality_gates 既有结构**：`run_quality_gate.py / check_zip_layout.py / check_text_encoding.py` 全部不动。
- **依赖零增加**：v1.7.0 不引入新 Python 依赖；KPI 卡片 / 研报 header / 风险提示 / 进度条 / 星级 / 表头合并 / 条件格式全部用 python-docx 原生 OOXML 实现。
- **版本号**：`pro_docx_gen.__version__ = "1.7.0"`；SKILL.md 顶部版本号同步更新；中文触发词新增"研报 / 股票研报 / 券商研报 / 机构研报 / 评级 / 目标价 / KPI / 风险提示"。
- **demo 验证产物**：`pro_docx_gen/v17_demo/` 3 份代表性 docx（business_compact / research_report / table_advanced_features），`pro_docx_gen/v17_make_demos.py` 重新生成脚本，`pro_docx_gen/v17_test_report.txt` 完整验证报告。


## v1.6.6 pre-insertion chart legibility gate - 2026-07-11

- Hard-blocks charts and text-bearing figures before DOCX insertion when effective text is below 9 pt, outer blank canvas exceeds 18%, display resolution is below 150 PPI, or text-size metadata is missing.
- Returns numeric remediation steps for cropping, font enlargement, landscape/splitting decisions, and rerender resolution.
- Raises `FigureAssetGateError`; the renderer may not insert a placeholder or bypass the failure.
- Clamps generated chart text to a legible minimum and adds standalone JSON preflight reporting.

## v1.6.5 DOCX layout and figure review hardening - 2026-07-11

- Added structural QA for dangling TOC bookmarks, direct cell-border palette leaks, redundant page breaks before page-break headings, adjacent section breaks, role font-size drift, and undersized figures.
- Changed strict delivery to render every page when a document contains figures, while preserving one-in-three sampling for figure-free documents.
- Required review of every figure page and its immediate neighbors, with full-document render scope recorded in the delivery manifest.
- Added regression coverage for cover-typography false positives and the layout defects found while rebuilding the DNT UK GTM report.

## v1.6.4 Sampled-page strict rendering - 2026-07-10

- Kept full-file DOCX structural and content checks while limiting visual rendering to pages 1, 4, 7, ... only.
- Added selective WPS/Word COM page export and removed full-document PDF fallback from strict delivery.
- Preserved original source page numbers through sampled PNG generation and four-up Agent review bundles.
- Made missing selective page export support a hard delivery failure.
- Added explicit-page COM export for adjacent-page defect expansion without any full-document fallback.
- Added short sampled-PDF aliases before PNG conversion to avoid Windows long-path failures.

## v1.6.3 Codex Windows adaptation - 2026-07-10

- Added mandatory root `SKILL.md` and `codex_adaptation` package metadata.
- Kept full-document rendering while changing visual review to one sample per consecutive three-page group, minimum one sample.
- Added automatic 2x2 review PDF and PNG sheets with up to four sampled source pages per PDF sheet.
- Added escalation from a failed sample to adjacent pages and the full affected section, with full review for systemic defects.
- Added Pillow as a required dependency for review-bundle generation.
- Added WPS Office COM rendering and package-local sandbox-safe temporary directories for Codex on Windows.

## v1.6.3 - 2026-07-08

- Slimmed duplicate top-level `shared` modules with compatibility wrappers for identical package-local implementations.
- Reduced default smoke usage to gate-layer plus render-layer checks; feature-specific checks now run through focused `--target` options.
- Added explicit forbidden downgrade rules: no removal of final render/PNG preview, no warning/error pass-through, no removal of top-level shared compatibility, and no feature-specific smoke as the default path.
- Added `agent_docx.py portability-check` to validate cross-device readiness: required modules, UTF-8 gate, root import, package-only import, local path leaks, optional zip layout, DOCX-to-PDF backend, and PDF-to-PNG preview backend.
- Added actionable missing-render-backend guidance so agents/users know to install LibreOffice/Word COM and Poppler/PyMuPDF instead of skipping final preview verification.
- Rejected `agent_docx.py deliver --no-render` for final delivery so the one-command path cannot bypass render verification or PNG preview generation.
- Added visual-failure escalation rule: if bad output comes from reusable skill behavior, patch the smallest relevant skill code/prompt/gate and report the problem.
- Added smoke coverage for the portability check so release packages fail when they only work in the current agent environment.
- Fixed package-only import compatibility: internal modules now prefer package-local `pro_docx_gen.shared` imports instead of requiring a top-level `shared/` directory on `PYTHONPATH`.
- Added smoke regression for package-only import compatibility and template registry loading.
- Tightened release zip layout guard to block smoke `_output*` directories, generated documents/images/archives, `chart_assets`, cache/temp files, and leaked local absolute paths.
- Cleaned the release package so local smoke outputs and machine-specific absolute paths are not embedded in the published zip.
- Tightened DOCX table visual-break gate so short words and prices such as `Plane`, `Context`, and `$169.50` cannot pass in too-narrow cells.
- Added `--created-after` quality-gate option to fail stale outputs when generation fails but an older DOCX still exists.
- Added smoke regressions for short word/price break blocking and stale-output blocking.
- Fixed package version mismatch: public `pro_docx_gen.__version__` now reports `1.6.3`.
- Unified the shared module version marker to `1.6.3`.
- Added chart readability auto-layout: dense single-series column charts switch to horizontal bars to avoid overlapping category labels.
- Fixed bar/column data-label placement so scalar baselines no longer label only the first bar.
- Single-series column/bar comparison charts now default to complete data labels.
- Added callout row split protection in the DOCX renderer.
- Added compact prose table keep-with-next protection in the DOCX renderer.
- Added quality gates and safe auto-fix for `callout_row_can_split` and `table_orphan_split_risk`.
- Added smoke regressions for dense chart labels, scalar data-label baselines, callout split risk, and compact prose table orphan risk.
- Clarified strict callout variants in `SKILL.md`: only `info`, `warning`, `success`, and `danger` are valid.
- Fixed renderer comment drift that mentioned unsupported aliases such as `important`, `note`, and `tip`.
- Added strict smoke suite at `smoke_tests/run_smoke_tests.py`.
- Added release zip layout checker at `quality_gates/check_zip_layout.py`.
- Fixed quality gate policy: warnings now fail by default.
- Fixed render QA policy: render failure or skipped render now blocks delivery unless explicitly allowed.
- Added Windows Word COM PDF render fallback when LibreOffice is unavailable.
- Fixed chart rendering temp assets: chart PNGs now render in an internal temporary directory, embed into DOCX, and clean up after save instead of leaking `chart_assets` beside the output file.
- Added smoke regression that generates a chart DOCX, verifies embedded media exists, and hard-fails if `chart_assets` leaks next to the output.
- Added DOCX table readability gate: narrow header columns now hard-fail with `table_header_narrow`.
- Added DOCX table body readability gate: narrow short label cells now hard-fail with `table_body_word_narrow` to block visible word breaks such as `Nightforce` splitting across lines.
- Added DOCX table row continuity gate: long table rows without `w:cantSplit` now fail by default through `table_row_can_split`, and generated tables now set no-split rows automatically.
- Added `--auto-fix` for safe DOCX table layout repair; it widens narrow header/body columns, marks long table rows as non-splitting, and re-runs quality gates.
- Added safe DOCX save behavior: generation and translation write through a sibling temp file before replacing the target, with an actionable error if Word/WPS locks the destination.
- Added smoke regressions for narrow body words and locked destination handling.
- Cleaned legacy mojibake from `SKILL.md` and remaining code comments/docstrings found during release work.
- Added `quality_gates/check_text_encoding.py` to block common UTF-8/CP936 mojibake and replacement characters before release.
- Added smoke regression for the text encoding gate.
- Cleaned README so dependency and verification steps are explicit.
- Added Chinese quick-start triggers and strict delivery rules to `SKILL.md`.
# v1.6.5

- Added structural DOCX layout QA for cell-level border leaks, dangling TOC bookmarks, font-size drift, undersized figures, redundant page breaks, and adjacent section blank pages.
- Figure-bearing documents now require full-page rendering and review of every figure page plus immediate neighbors.
- Added regression smoke tests for the defects found in mixed portrait/landscape business reports.
