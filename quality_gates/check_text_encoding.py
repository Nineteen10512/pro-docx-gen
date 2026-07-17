"""Block common mojibake in release text files.

The usual root cause is UTF-8 text decoded as CP936/GBK, then saved again.
This gate fails on replacement characters and common mojibake markers instead
of allowing garbled docs, prompts, or comments to ship.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


TEXT_SUFFIXES = {".py", ".md", ".txt", ".json", ".yaml", ".yml"}
SKIP_DIRS = {"__pycache__", ".git", ".pytest_cache"}
HARD_MARKERS = {
    "\ufffd": "replacement character",
    "\u9225": "UTF-8/CP936 mojibake marker",
    "\u951b": "UTF-8/CP936 mojibake marker",
    "\u9983": "emoji mojibake marker",
}
SOFT_MARKERS = {
    "\u93c2",
    "\u93b5",
    "\u6fc2",
    "\u7edb",
    "\u9365",
    "\u741b",
    "\u947a",
    "\u5bf0",
    "\u9357",
}


def _iter_text_files(root: Path):
    for path in root.rglob("*"):
        if any(part in SKIP_DIRS or part.startswith("_output") for part in path.parts):
            continue
        if path.is_file() and path.suffix.lower() in TEXT_SUFFIXES:
            if path.resolve() == Path(__file__).resolve():
                continue
            yield path


def inspect_file(path: Path) -> list[dict]:
    issues: list[dict] = []
    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError as exc:
        return [{"line": 0, "reason": f"not valid UTF-8: {exc}", "text": ""}]

    for line_no, line in enumerate(text.splitlines(), 1):
        for marker, reason in HARD_MARKERS.items():
            if marker in line:
                issues.append({"line": line_no, "reason": reason, "text": line.strip()[:180]})
                break
        else:
            soft_count = sum(line.count(marker) for marker in SOFT_MARKERS)
            if soft_count >= 2:
                issues.append(
                    {
                        "line": line_no,
                        "reason": f"multiple mojibake markers ({soft_count})",
                        "text": line.strip()[:180],
                    }
                )
    return issues


def _console_safe(text: str) -> str:
    return text.encode(sys.stdout.encoding or "utf-8", errors="backslashreplace").decode(
        sys.stdout.encoding or "utf-8",
        errors="replace",
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Check text files for mojibake/garbled encoding.")
    parser.add_argument("root", type=Path, help="File or directory to inspect")
    args = parser.parse_args()

    paths = [args.root] if args.root.is_file() else list(_iter_text_files(args.root))
    failures = []
    for path in paths:
        issues = inspect_file(path)
        for issue in issues:
            failures.append((path, issue))

    for path, issue in failures:
        line = issue["line"]
        where = f"{path}:{line}" if line else str(path)
        print(_console_safe(f"ERROR mojibake: {where}: {issue['reason']}"))
        if issue["text"]:
            print(_console_safe(f"  {issue['text']}"))

    if failures:
        print(f"FAIL text encoding gate: {len(failures)} issue(s)")
        return 1
    print("PASS text encoding gate")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
