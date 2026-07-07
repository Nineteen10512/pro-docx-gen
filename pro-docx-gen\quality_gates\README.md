# Delivery Quality Gates

Run this from the extracted bundle root after generating PPTX/DOCX files:

```bash
python quality_gates/run_quality_gate.py output/deck.pptx output/report.docx \
  --json-report output/quality_report.json
```

Useful options:

- `--no-render`: skip optional LibreOffice PDF rendering.
- `--strict`: treat warnings as failures.
- `--output-dir quality_gates/_renders`: keep optional render outputs.

The gate checks the Office package structure, core XML files, visible text,
placeholder-like text, unusually dense slides, long DOCX text runs, and simple
table geometry signals. PPTX files also get a lightweight taste scan for
generic AI-marketing copy, repeated slide text, and em/en dash telltales.
If LibreOffice/`soffice` is available, it also attempts headless PDF conversion
and records the result in the report.
