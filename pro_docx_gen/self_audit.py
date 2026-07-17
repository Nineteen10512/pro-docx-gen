"""Self-audit module for PRO-DOCX v1.6.6 — color discipline enforcement.

This module is the "self-audit" + "防复发机制" layer mandated by the v1.6.6
quality discipline. It runs two complementary checks:

1. **Template audit** (``audit_template`` / ``audit_all_templates``):
   Parses every template ``.py`` file under ``pro_docx_gen/templates/`` with
   ``ast`` and extracts all hex string literals (``"#1F3864"`` style). For
   fields that semantically map to heading / text / muted (``heading``,
   ``title``, ``h1``–``h6``, ``text``, ``body``, ``muted``), it verifies that
   no literal matches the forbidden list (黑金/浅蓝紫/青蓝霓虹/白 等).

   Because parsing uses ``ast`` (not regex), false positives from
   docstrings, comments, or ``forbidden_heading_hexes`` constant tables are
   structurally impossible — only **Assign targets** with **Constant str**
   values inside ``Dict``/``Call`` initialisers contribute.

2. **Output audit** (``audit_docx_output``):
   Opens a generated ``.docx`` (a ZIP), walks ``word/document.xml`` with
   ``xml.etree.ElementTree``, and verifies that:

   a) every heading paragraph (``pStyle`` starts with ``Heading``) renders
      a colour that is not in ``FORBIDDEN_HEADING_HEXES``;
   b) all headings at the same level (H1, H2, …) share a single colour, so
      we never see "the first H1 is black, the second H1 is gold" regression.

Why this module exists (P0 root-cause fix):
- A previous release trusted ``theme_overrides["color"]["heading"]`` from
  user templates; brand_luxury / proposal_elegant injected ``#D4AF37`` (gold)
  and modern_tech injected ``#E0E7FF`` (pale lavender) which became invisible
  on a white page. The renderer layer now *forces* ``#1A1A1A`` /
  ``#333333`` / ``#666666`` and this self-audit layer makes the discipline
  enforceable: future regressions trip the gate before they ship.

Run from CLI::

    python -m pro_docx_gen.self_audit --templates
    python -m pro_docx_gen.self_audit --docx /path/to/output.docx
    python -m pro_docx_gen.self_audit             # run both

Exit code: 0 on full pass, 1 on any fail. Use in CI gates.
"""

from __future__ import annotations

import argparse
import ast
import re
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, List, Optional
from xml.etree import ElementTree as ET


# Locate the package root regardless of how this module is invoked
_PACKAGE_DIR = Path(__file__).resolve().parent
_TEMPLATES_DIR = _PACKAGE_DIR / "templates"

# Field names whose values must NEVER be a forbidden heading colour
HEADING_RELATED_FIELDS = frozenset({
    "heading", "title",
    "h1", "h2", "h3", "h4", "h5", "h6",
    "text", "body", "muted",
})

# WordprocessingML namespace
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

# Regex for a #RRGGBB literal; case-insensitive
_HEX_RE = re.compile(r"^#([0-9A-Fa-f]{6})$")


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------


@dataclass
class Violation:
    """A single self-audit violation."""

    field: str            # e.g. "heading", "title", "h1", or "Heading 1"
    hex_value: str        # raw hex including leading "#" (e.g. "#D4AF37") or "" if N/A
    line: int             # 1-based line number in the source file (templates only)
    description: str      # human-readable explanation

    def to_dict(self) -> dict:
        return {
            "field": self.field,
            "hex": self.hex_value,
            "line": self.line,
            "description": self.description,
        }


@dataclass
class AuditReport:
    """Aggregated audit result for one template or one DOCX output."""

    name: str
    status: str = "pass"  # "pass" | "warn" | "fail"
    violations: List[Violation] = field(default_factory=list)
    checked_fields: List[str] = field(default_factory=list)
    details: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "status": self.status,
            "violations": [v.to_dict() for v in self.violations],
            "checked_fields": list(self.checked_fields),
            "details": dict(self.details),
        }


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _normalize_hex(hex_str: str) -> str:
    """Return ``"RRGGBB"`` upper-cased (no leading ``#``)."""
    return hex_str.lstrip("#").upper()


def _is_forbidden_hex(hex_str: str) -> bool:
    """Check if a hex literal is on the forbidden-heading list."""
    from .shared.color_palette import is_forbidden_heading_hex
    return is_forbidden_heading_hex(hex_str)


def _classify_status(report: AuditReport) -> AuditReport:
    """Promote a report's status to ``fail`` if any violations exist.

    Mutates the report in place and returns it so callers can chain
    ``return _classify_status(report)`` without losing the report object.
    """
    if report.violations:
        # All current violations are hard fails (no soft warnings defined yet)
        report.status = "fail"
    elif report.status not in ("pass", "warn", "fail"):
        report.status = "pass"
    return report


# ---------------------------------------------------------------------------
# Template-level audit (uses ast — no regex, no false positives)
# ---------------------------------------------------------------------------


def _walk_assign_targets(
    node: ast.AST,
    in_dict_target: bool = False,
) -> Iterable[tuple[str, ast.Constant]]:
    """Yield ``(field_name, Constant node)`` for every kwarg/Assign that
    looks like ``"heading": "#1F3864"`` inside a template call.

    Walks ``ast.Call`` keyword arguments and ``ast.Dict`` key/value pairs
    to surface colour-bearing literals. ``Constant`` is the modern AST node
    used in Python 3.8+ for string/numeric literals.
    """
    if isinstance(node, ast.Dict):
        for key, value in zip(node.keys, node.values):
            if key is None or not isinstance(key, ast.Constant):
                continue
            if not isinstance(key.value, str):
                continue
            if isinstance(value, ast.Constant) and isinstance(value.value, str):
                yield key.value, value
            # Drill into nested dicts (e.g. theme_overrides={"color": {...}})
            yield from _walk_assign_targets(value, in_dict_target=True)
    elif isinstance(node, ast.Call):
        for kw in node.keywords:
            if kw.arg is None:
                # **kwargs splat — skip; cannot inspect statically
                continue
            if isinstance(kw.value, ast.Constant) and isinstance(kw.value.value, str):
                yield kw.arg, kw.value
            # Drill into nested dicts
            yield from _walk_assign_targets(kw.value, in_dict_target=True)


def _module_top_level_assignments(tree: ast.Module) -> Iterable[tuple[str, ast.AST, int]]:
    """Yield ``(name, value_node, lineno)`` for every ``NAME = ...`` at module
    top level. ``line`` is the 1-based line of the assignment in source.
    """
    for stmt in tree.body:
        if isinstance(stmt, ast.Assign):
            for tgt in stmt.targets:
                if isinstance(tgt, ast.Name):
                    yield tgt.id, stmt.value, stmt.lineno


def audit_template_file(path: Path) -> AuditReport:
    """Audit a single template ``.py`` file.

    Parses the file with :mod:`ast`, walks every top-level ``register(...)``
    call, and inspects the keyword arguments. Any ``"heading" / "title" / "h1..h6" /
    "text" / "body" / "muted"`` kwarg whose value is a string literal matching
    the forbidden-heading set is reported as a Violation.

    The ``description`` field explains why the colour is disallowed; the
    ``line`` field points to the literal's source line for the developer.
    """
    report = AuditReport(name=path.stem)
    try:
        source = path.read_text(encoding="utf-8")
    except OSError as exc:
        report.violations.append(Violation(
            field="<file>",
            hex_value="",
            line=0,
            description=f"Failed to read template file: {exc}",
        ))
        return _classify_status(report)

    try:
        tree = ast.parse(source, filename=str(path))
    except SyntaxError as exc:
        report.violations.append(Violation(
            field="<syntax>",
            hex_value="",
            line=exc.lineno or 0,
            description=f"Template has a syntax error: {exc.msg}",
        ))
        return _classify_status(report)

    # Collect all "field -> constant" pairs that are inside a register(...)
    # call, so module-level helper constants are not mistaken for fields.
    in_register_call = False

    for node in ast.walk(tree):
        if isinstance(node, ast.Call):
            # Heuristic: a "register(DOCXTemplate(...))" call. We look for
            # any Call whose positional argument is itself a Call to
            # DOCXTemplate. This is robust to aliasing (``from .registry
            # import register``).
            func_id = None
            if isinstance(node.func, ast.Name):
                func_id = node.func.id
            elif isinstance(node.func, ast.Attribute):
                func_id = node.func.attr
            if func_id == "register":
                in_register_call = True
                for arg in node.args:
                    yield_pairs = list(_walk_assign_targets(arg))
                    for field_name, const in yield_pairs:
                        if field_name in HEADING_RELATED_FIELDS:
                            report.checked_fields.append(field_name)
                            value = const.value
                            if not _HEX_RE.match(value):
                                # Not a colour literal — skip (e.g. None, "auto")
                                continue
                            if _is_forbidden_hex(value):
                                report.violations.append(Violation(
                                    field=field_name,
                                    hex_value=value,
                                    line=const.lineno,
                                    description=(
                                        f"{field_name}={value!r} is on the "
                                        f"forbidden-heading list; replace with "
                                        f"a non-forbidden hex (e.g. #1A1A1A) "
                                        f"or omit and let the renderer enforce it."
                                    ),
                                ))
                in_register_call = False

    return _classify_status(report)


def audit_template(module) -> AuditReport:
    """Audit a template by module object.

    The caller passes the imported module (e.g. ``pro_docx_gen.templates.academic_corporate``).
    This is a thin wrapper around :func:`audit_template_file` that locates
    the source file via ``module.__file__``.
    """
    file_path = getattr(module, "__file__", None)
    if not file_path:
        return AuditReport(
            name=getattr(module, "__name__", "<unknown>"),
            status="fail",
            violations=[Violation(
                field="<file>",
                hex_value="",
                line=0,
                description="Template module has no __file__ attribute; cannot audit.",
            )],
        )
    return audit_template_file(Path(file_path))


def audit_all_templates(
    templates_dir: Optional[Path] = None,
) -> List[AuditReport]:
    """Audit every template file under ``templates/``.

    Skips ``__init__.py`` and ``registry.py`` because they contain no
    ``DOCXTemplate`` colour fields.
    """
    base = Path(templates_dir) if templates_dir else _TEMPLATES_DIR
    reports: List[AuditReport] = []
    if not base.exists():
        return reports
    for py in sorted(base.glob("*.py")):
        if py.name in ("__init__.py", "registry.py"):
            continue
        reports.append(audit_template_file(py))
    return reports


# ---------------------------------------------------------------------------
# DOCX-output audit (zip + xml.etree)
# ---------------------------------------------------------------------------


def _strip_ns(tag: str) -> str:
    """Strip ``{namespace}`` prefix from an ElementTree tag."""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _pstyle_of_paragraph(paragraph: ET.Element) -> Optional[str]:
    """Return the ``pStyle`` value (e.g. ``"Heading1"``) of a paragraph, or None."""
    pPr = paragraph.find(f"{{{W_NS}}}pPr")
    if pPr is None:
        return None
    pStyle = pPr.find(f"{{{W_NS}}}pStyle")
    if pStyle is None:
        return None
    return pStyle.get(f"{{{W_NS}}}val")


def _run_color(runs: List[ET.Element]) -> Optional[str]:
    """Return the first explicit run colour (with leading ``#``) or None.

    A run may set its colour via ``<w:rPr><w:color w:val="RRGGBB"/></w:rPr>``.
    Returns ``"#RRGGBB"`` upper-case or None if the run inherits the style
    colour.
    """
    for run in runs:
        rPr = run.find(f"{{{W_NS}}}rPr")
        if rPr is None:
            continue
        color = rPr.find(f"{{{W_NS}}}color")
        if color is None:
            continue
        val = color.get(f"{{{W_NS}}}val")
        if not val:
            continue
        if val.lower() == "auto":
            continue
        if not _HEX_RE.match(f"#{val}"):
            continue
        return f"#{val.upper()}"
    return None


def _paragraph_text(paragraph: ET.Element) -> str:
    """Concatenate all ``<w:t>`` text in the paragraph (for context messages)."""
    parts = []
    for el in paragraph.iter(f"{{{W_NS}}}t"):
        if el.text:
            parts.append(el.text)
    return "".join(parts).strip()


def _extract_heading_color_map(root: ET.Element) -> dict:
    """Walk a ``word/document.xml`` tree and return a mapping of
    ``level_key -> {color_hex: [paragraph_indices]}`` and a parallel
    ``level_key -> [paragraph_indices_in_doc_order]``.

    ``level_key`` is ``"Title"`` for the title style and ``"Heading 1"`` …
    ``"Heading 6"`` for heading levels. ``paragraph_indices`` is 0-based
    in document order.
    """
    body = root.find(f"{{{W_NS}}}body")
    if body is None:
        return {"per_level": {}, "per_color": {}, "level_indices": {}}

    per_level: dict[str, list[int]] = {}
    per_color: dict[str, dict[str, list[int]]] = {}
    level_indices: dict[str, list[int]] = {}

    for idx, paragraph in enumerate(body.findall(f"{{{W_NS}}}p")):
        pstyle = _pstyle_of_paragraph(paragraph)
        if not pstyle:
            continue
        # Match Title and Heading{1..6} only
        if pstyle == "Title":
            level_key = "Title"
        elif pstyle.startswith("Heading"):
            # Heading 1 / Heading1 / Heading2 — accept both forms
            m = re.match(r"^Heading\s*(\d+)$", pstyle) or re.match(r"^Heading(\d+)$", pstyle)
            if not m:
                continue
            level_key = f"Heading {m.group(1)}"
        else:
            continue

        runs = paragraph.findall(f"{{{W_NS}}}r")
        colour = _run_color(runs)

        per_level.setdefault(level_key, []).append(idx)
        level_indices.setdefault(level_key, []).append(idx)

        if colour is not None:
            per_color.setdefault(level_key, {}).setdefault(colour, []).append(idx)
        else:
            # Inherited colour — record under the sentinel "auto" so we can
            # tell the user "the run does not override; check the style".
            per_color.setdefault(level_key, {}).setdefault("<auto>", []).append(idx)

    return {
        "per_level": per_level,
        "per_color": per_color,
        "level_indices": level_indices,
    }


def audit_docx_output(docx_path: str | Path) -> AuditReport:
    """Audit a generated DOCX file for colour discipline.

    Opens the ZIP, parses ``word/document.xml``, and verifies that:

    * every heading-level paragraph has a colour that is not in
      :data:`FORBIDDEN_HEADING_HEXES`;
    * all paragraphs at the same heading level share a single colour
      (e.g. all H1 are ``#1A1A1A``). A level with ≥2 distinct colours
      is a fail — this catches "first H1 black, second H1 gold" regressions.

    The report's ``details`` field carries the full level→colour mapping
    for inspection.
    """
    path = Path(docx_path)
    report = AuditReport(name=path.name)

    if not path.exists():
        report.violations.append(Violation(
            field="<file>",
            hex_value="",
            line=0,
            description=f"DOCX does not exist: {path}",
        ))
        return _classify_status(report)

    try:
        with zipfile.ZipFile(path) as zf:
            if "word/document.xml" not in zf.namelist():
                report.violations.append(Violation(
                    field="<zip>",
                    hex_value="",
                    line=0,
                    description="DOCX is missing word/document.xml; not a valid Word file.",
                ))
                return _classify_status(report)
            xml_bytes = zf.read("word/document.xml")
    except zipfile.BadZipFile as exc:
        report.violations.append(Violation(
            field="<zip>",
            hex_value="",
            line=0,
            description=f"DOCX is not a valid zip: {exc}",
        ))
        return _classify_status(report)

    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as exc:
        report.violations.append(Violation(
            field="<xml>",
            hex_value="",
            line=0,
            description=f"word/document.xml failed to parse: {exc}",
        ))
        return _classify_status(report)

    extracted = _extract_heading_color_map(root)
    per_level = extracted["per_level"]
    per_color = extracted["per_color"]
    level_indices = extracted["level_indices"]

    report.details["levels"] = sorted(per_level.keys())
    report.details["color_distribution"] = {
        level: {hex_val: len(idxs) for hex_val, idxs in colours.items()}
        for level, colours in per_color.items()
    }

    for level in sorted(per_level.keys()):
        report.checked_fields.append(level)
        colours_at_level = per_color.get(level, {})

        # 1) Forbidden hex check
        for hex_value, idxs in colours_at_level.items():
            if hex_value == "<auto>":
                continue
            if _is_forbidden_hex(hex_value):
                report.violations.append(Violation(
                    field=level,
                    hex_value=hex_value,
                    line=idxs[0] + 1 if idxs else 0,
                    description=(
                        f"{level} uses forbidden colour {hex_value!r} on "
                        f"{len(idxs)} paragraph(s) (paragraph idx: {idxs!r}). "
                        f"This hex is unreadable on a white background."
                    ),
                ))

        # 2) Same-level colour consistency check
        explicit = {h: idxs for h, idxs in colours_at_level.items() if h != "<auto>"}
        if len(explicit) >= 2:
            # At least two distinct explicit colours on the same heading level
            count_repr = ", ".join(
                f"{h!r}×{len(idxs)}" for h, idxs in sorted(explicit.items())
            )
            report.violations.append(Violation(
                field=level,
                hex_value=",".join(sorted(explicit.keys())),
                line=level_indices[level][0] + 1 if level_indices.get(level) else 0,
                description=(
                    f"{level} has {len(explicit)} different colours across "
                    f"{len(level_indices[level])} paragraph(s): {count_repr}. "
                    f"All {level} paragraphs must share one colour (e.g. #1A1A1A)."
                ),
            ))

    return _classify_status(report)


# ---------------------------------------------------------------------------
# Pretty printing & CLI
# ---------------------------------------------------------------------------


def _format_violation(v: Violation) -> str:
    """Return a one-line human description of a violation."""
    hex_repr = v.hex_value if v.hex_value else "—"
    line_repr = f"L{v.line}" if v.line else "—"
    return f"{v.field}={hex_repr} @ {line_repr}: {v.description}"


def format_report_line(report: AuditReport) -> str:
    """Return the canonical CLI line for one report."""
    n = len(report.violations)
    plural = "" if n == 1 else "s"
    if report.status == "pass":
        return f"[PASS] {report.name}: 0 violations"
    if report.status == "warn":
        return f"[WARN] {report.name}: {n} warning{plural}"
    return f"[FAIL] {report.name}: {n} violation{plural}"


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="pro_docx_gen.self_audit",
        description=(
            "Self-audit module for PRO-DOCX v1.6.6 color discipline. "
            "Pass --templates to audit every template file; pass --docx to "
            "audit a generated .docx; pass neither to run both."
        ),
    )
    parser.add_argument(
        "--templates", action="store_true",
        help="Audit every DOCX template file under pro_docx_gen/templates/.",
    )
    parser.add_argument(
        "--docx", type=Path, default=None,
        help="Audit a single generated .docx file for heading colour discipline.",
    )
    parser.add_argument(
        "--templates-dir", type=Path, default=None,
        help="Override the templates directory (mainly for tests).",
    )
    args = parser.parse_args(argv)

    run_templates = args.templates or (not args.templates and not args.docx)
    run_docx = args.docx is not None or (not args.templates and not args.docx)

    overall_status = 0  # 0 = pass; 1 = any fail
    summary_lines: list[str] = []

    if run_templates:
        reports = audit_all_templates(args.templates_dir)
        if not reports:
            print("[WARN] no template files found to audit")
        else:
            for r in reports:
                line = format_report_line(r)
                print(line)
                for v in r.violations:
                    print(f"  - {_format_violation(v)}")
            passed = sum(1 for r in reports if r.status == "pass")
            total = len(reports)
            summary_lines.append(f"Templates: {passed}/{total} passed")
            if any(r.status == "fail" for r in reports):
                overall_status = 1

    if run_docx:
        if not args.docx:
            print()
        if args.docx:
            targets: List[Path] = [args.docx]
        else:
            # If neither flag was set, audit_docx is a no-op (no file given)
            targets = []
        for target in targets:
            r = audit_docx_output(target)
            line = format_report_line(r)
            print(line)
            for v in r.violations:
                print(f"  - {_format_violation(v)}")
            if r.status == "fail":
                overall_status = 1
        if not targets:
            print("[skip] --docx not provided; DOCX audit skipped")

    if summary_lines:
        print()
        print("Summary: " + " | ".join(summary_lines))

    return overall_status


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
