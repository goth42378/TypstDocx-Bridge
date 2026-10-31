#!/usr/bin/env python3
"""Preflight a semantic Typst content file before DOCX conversion."""

from __future__ import annotations

import argparse
import json
import re
from collections import Counter
from pathlib import Path


HEADING1_RE = re.compile(r"^=\s+")
LABEL_RE = re.compile(r"<([A-Za-z0-9_:-]+)>")
CROSSREF_RE = re.compile(r"@((?:fig|tab|eqt):[A-Za-z0-9_:-]+)")
COMMAND_RE = re.compile(r"^\s*#([A-Za-z][A-Za-z0-9_-]*)")
SUPPORTED_TOP_LEVEL_COMMANDS = {"import", "include", "capfig", "captab"}


def strip_line_comment(line: str) -> str:
    """Remove a simple Typst // comment while respecting quoted strings."""
    in_string = False
    escaped = False
    index = 0
    while index < len(line) - 1:
        char = line[index]
        if escaped:
            escaped = False
        elif char == "\\" and in_string:
            escaped = True
        elif char == '"':
            in_string = not in_string
        elif char == "/" and line[index + 1] == "/" and not in_string:
            return line[:index]
        index += 1
    return line


def audit(content: Path, text: str | None = None) -> dict[str, object]:
    lines = (content.read_text(encoding="utf-8-sig") if text is None else text).splitlines()
    active_lines = [strip_line_comment(line) for line in lines]
    heading1_lines: list[int] = []
    labels_with_lines: list[tuple[str, int]] = []
    crossrefs: list[str] = []
    review_commands: list[dict[str, object]] = []
    in_math = False
    for index, line in enumerate(active_lines, start=1):
        labels_with_lines.extend((label, index) for label in LABEL_RE.findall(line))
        crossrefs.extend(CROSSREF_RE.findall(line))
        if not in_math:
            if HEADING1_RE.match(line.strip()):
                heading1_lines.append(index)
            command = COMMAND_RE.match(line)
            supported_strong = bool(re.fullmatch(r"\s*#strong\[[^\[\]]*\]\s*", line))
            if command and command.group(1) not in SUPPORTED_TOP_LEVEL_COMMANDS and not supported_strong:
                review_commands.append(
                    {
                        "line": index,
                        "command": command.group(1),
                        "text": line.strip()[:240],
                    }
                )
            if re.match(r'^\s*(?:-\s+|\d+\.\s+)', line):
                review_commands.append({"line": index, "command": "unsupported-list", "text": line.strip()[:240]})
            if re.match(r'^\s+\+\s+', line):
                review_commands.append({"line": index, "command": "nested-ordered-list", "text": line.strip()[:240]})
        if len(re.findall(r"(?<!\\)\$", line)) % 2:
            in_math = not in_math

    label_counts = Counter(label for label, _ in labels_with_lines)
    duplicate_labels = [
        {
            "label": label,
            "lines": [line for candidate, line in labels_with_lines if candidate == label],
        }
        for label, count in sorted(label_counts.items())
        if count > 1
    ]
    targets = set(label_counts)
    targets.update(f"eqt:{label}" for label in label_counts if ":" not in label)
    unresolved_crossrefs = sorted(set(crossrefs) - targets)

    blocking_issues: list[str] = []
    if len(heading1_lines) != 1:
        blocking_issues.append(
            f"expected exactly one level-1 heading for a single-chapter export; found {len(heading1_lines)}"
        )
    if duplicate_labels:
        blocking_issues.append(f"duplicate labels: {len(duplicate_labels)}")
    if unresolved_crossrefs:
        blocking_issues.append(f"unresolved figure/table/equation references: {len(unresolved_crossrefs)}")
    if review_commands:
        blocking_issues.append(f"unsupported top-level Typst commands requiring review: {len(review_commands)}")

    return {
        "content": str(content),
        "status": "passed" if not blocking_issues else "review_required",
        "heading1_count": len(heading1_lines),
        "heading1_lines": heading1_lines,
        "labels": len(labels_with_lines),
        "duplicate_labels": duplicate_labels,
        "crossrefs": len(crossrefs),
        "unresolved_crossrefs": unresolved_crossrefs,
        "review_required_commands": review_commands,
        "blocking_issues": blocking_issues,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("content", type=Path)
    parser.add_argument("--json", type=Path)
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Return exit code 2 when the file needs review before conversion.",
    )
    args = parser.parse_args()

    content = args.content.resolve(strict=True)
    payload = audit(content)
    rendered = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_text(rendered, encoding="utf-8")
    print(rendered)
    return 2 if args.strict and payload["status"] != "passed" else 0


if __name__ == "__main__":
    raise SystemExit(main())
