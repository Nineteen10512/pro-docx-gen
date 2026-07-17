# Codex Windows Adaptation

This package applies the PRO-DOCX v1.6.6 structural, visual, and pre-insertion figure quality gates for Codex on Windows.

## Included Adaptations

- A root `SKILL.md` is included so Codex can discover requirements without entering the Python package.
- WPS Office COM is the preferred selective page export backend; Microsoft Word COM is the fallback.
- Runtime and chart temporary files stay under package-local work directories instead of restricted system temp locations.
- Final delivery structurally audits the full DOCX. Figure-free documents render one page from every consecutive three-page group; documents containing figures render every page.
- Multiple review samples are automatically arranged as 2x2 sheets with up to four sampled source pages per PDF page. Matching sheet PNGs are generated for efficient model inspection.
- Figure review always includes the figure page and both immediate neighbors. A failed sampled page in a figure-free document expands to adjacent pages and the affected section.
- Use `try_render_specific_pages(...)` for explicit-page expansion or full figure-document export; PDF-to-PNG conversion uses short aliases to remain safe on long Windows workspace paths.
- Warnings, render failures, stale outputs, mojibake, and package pollution remain blocking errors.

## Verification Commands

```bash
python agent_docx.py doctor
python agent_docx.py portability-check
python smoke_tests/run_smoke_tests.py --target gate --target render
python agent_docx.py package-check path/to/pro-docx-gen.zip
```

`Pillow` is mandatory for sampled four-up review PDF and PNG generation.
