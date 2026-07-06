# PRO-DOCX v1.5.3 Addendum

This addendum documents behavior that is already implemented in the archive but
was not clearly described in the earlier `SKILL.md`.

## Public taste API

`pro_docx_gen.taste_check()` is now a real public API. It accepts:

- semantic JSON dict
- Markdown text

It returns the same high-level report structure used by PRO-PPTX:

- `version`
- `score`
- `passed`
- `threshold`
- `design_read`
- `preflight`
- `issues`
- `base_quality_score`
- `story_compliant`

Example:

```python
from pro_docx_gen import taste_check

report = taste_check(content, theme="business", lang="cn")
print(report["score"], report["passed"])
for issue in report["issues"]:
    print(issue["level"], issue["code"], issue["message"])
```

## DOCX taste hard-fail rules

The DOCX taste check currently hard-fails on:

- `contrast_iron_law`
- `chart_color_variety`
- `placeholder_text`

This means the DOCX preflight now blocks:

- dark-on-dark or light-on-light text/background pairings
- single-color chart palettes where variety is required
- placeholder copy such as `TODO`, `TBD`, and similar markers

## Shared taste framework

PRO-DOCX now uses the shared taste layer under:

- `shared/taste/core.py`
- `shared/taste/rules.py`
- `shared/taste/adapters.py`

This is shared with PRO-PPTX for:

- design-read inference
- common anti-placeholder / anti-generic-copy checks
- preflight report assembly

## Additional public APIs already shipped

The following APIs exist and should be considered documented:

- `generate_with_collaboration()`
- `scan_local_templates()`
- `export_review_pdf()`

## `generate_with_collaboration()` convenience entrypoint

```python
from pro_docx_gen import generate_with_collaboration

session = generate_with_collaboration(
    content,
    output_path="draft.docx",
    theme="business",
    lang="cn",
    enable_track_changes=True,
)
```

This helper:

- generates the DOCX
- immediately opens a collaboration session
- optionally enables track changes at session start

## `export_review_pdf()` actual fallback behavior

`pro_docx_gen.collaboration.export_review_pdf()` behaves as follows:

1. try `docx2pdf` first
2. if `docx2pdf` is unavailable, fall back to `pro_docx_gen.docx_jsx.to_pdf()`
3. keep returning a PDF path string

This fallback is covered by shipped tests in:

- `tests/test_export_review_pdf.py`

## Template-only skeleton generation

The package already supports generating a document from template metadata alone:

```python
from pro_docx_gen import generate

generate(
    content=None,
    output_path="skeleton.docx",
    template_name="thesis_full",
    theme="academic",
    lang="cn",
)
```

This is useful when you want an editable DOCX skeleton before filling the real
semantic content.

## Local taste-skill reference

The archive also includes a local reference copy of the public `taste-skill`
repository under:

- `docs/references/taste-skill-main/`

It is a reference source only, not a PRO-DOCX runtime dependency.
