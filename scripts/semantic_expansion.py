"""Expand supported static Typst inclusions without losing source locations."""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import preflight_typst


INCLUDE = re.compile(r'^\s*#include\s+["\']([^"\']+)["\']\s*;?\s*$')
ANY_INCLUDE = re.compile(r'^\s*#include\b')
NOTE_DEFINITION = re.compile(
    r'(?m)^\s*#let\s+formula-note\s*\(\s*body\s*\)\s*=\s*'
    r'block\s*\(\s*width\s*:\s*100%\s*\)\s*\[\s*'
    r'#set\s+par\s*\(\s*first-line-indent\s*:\s*\(\s*amount\s*:\s*0pt\s*,'
    r'\s*all\s*:\s*true\s*,?\s*\)\s*\)\s*式中，\s*#body\s*\]', re.M)
BLOCK_OPEN = re.compile(r'(?m)^\s*#block\s*\(\s*breakable\s*:\s*false\s*\)\s*\[')
NOTE_OPEN = re.compile(r'(?m)^\s*#formula-note\s*\[')
FORMULA_NOTE_MARKER = '[[TYPSTDOCX_FORMULA_NOTE]]'


@dataclass(frozen=True)
class SourceSegment:
    path: Path
    text: str
    line: int = 1


def source_segments(path: Path, root: Path, stack: tuple[Path, ...] = ()) -> tuple[list[SourceSegment], bool]:
    path = path.resolve(strict=True)
    if not path.is_relative_to(root) or path in stack:
        raise ValueError(f"include escapes staged root or forms a cycle: {path}")
    lines = path.read_text(encoding='utf-8-sig').splitlines(keepends=True)
    plain = [preflight_typst.strip_line_comment(line.rstrip('\r\n')) for line in lines]
    is_wrapper = not stack and any(re.match(r'^\s*#show:\s*nwpu-thesis\.with\s*\(', line) for line in plain)
    if is_wrapper:
        for line in plain:
            command = re.match(r'^\s*#([A-Za-z][A-Za-z0-9_-]*)', line)
            if command and command.group(1) not in {'include', 'import', 'show', 'counter'}:
                raise ValueError(f'unsupported assembly command in {path}: {line.strip()}')
            if command and command.group(1) == 'show' and not re.match(r'^\s*#show:\s*nwpu-thesis\.with\s*\(', line):
                raise ValueError(f'unsupported show rule in {path}: {line.strip()}')
    result: list[SourceSegment] = []
    chunk: list[str] = []
    chunk_start = 1
    found_include = False

    def flush() -> None:
        if not is_wrapper and ''.join(chunk).strip():
            result.append(SourceSegment(path, ''.join(chunk), chunk_start))
        chunk.clear()

    for index, (raw_line, line) in enumerate(zip(lines, plain), start=1):
        if ANY_INCLUDE.match(line):
            match = INCLUDE.fullmatch(line)
            if not match:
                raise ValueError(f'dynamic or unsupported include in {path}:{index}')
            flush()
            raw = match.group(1)
            child = root / raw.lstrip('/\\') if raw.startswith(('/', '\\')) else path.parent / raw
            nested, _ = source_segments(child, root, stack + (path,))
            result.extend(nested)
            found_include = True
            chunk_start = index + 1
        else:
            chunk.append(raw_line)
    flush()
    if is_wrapper and not found_include:
        raise ValueError(f'no static chapter includes in assembly wrapper: {path}')
    return result, is_wrapper or found_include


def closing_bracket(text: str, opening: int) -> int:
    depth = 0
    in_string = False
    escaped = False
    in_comment = False
    index = opening
    while index < len(text):
        char = text[index]
        if in_comment:
            if char == '\n':
                in_comment = False
        elif in_string:
            if escaped:
                escaped = False
            elif char == '\\':
                escaped = True
            elif char == '"':
                in_string = False
        elif char == '"':
            in_string = True
        elif char == '/' and index + 1 < len(text) and text[index + 1] == '/':
            in_comment = True
            index += 1
        elif char == '[':
            depth += 1
        elif char == ']':
            depth -= 1
            if depth == 0:
                return index
        index += 1
    raise ValueError('unclosed supported Typst content block')


def normalize_wrappers(text: str, note_defined: bool) -> str:
    text = NOTE_DEFINITION.sub('', text)
    out: list[str] = []
    position = 0
    while True:
        candidates = [(match.start(), kind, match) for kind, pattern in (('block', BLOCK_OPEN), ('note', NOTE_OPEN))
                      if (match := pattern.search(text, position))]
        if not candidates:
            out.append(text[position:])
            break
        _, kind, match = min(candidates, key=lambda item: item[0])
        if kind == 'note' and not note_defined:
            raise ValueError('formula-note is used without the supported definition')
        opening = match.end() - 1
        end = closing_bracket(text, opening)
        out.append(text[position:match.start()])
        body = normalize_wrappers(text[opening + 1:end], note_defined)
        out.append('\n\n' + (FORMULA_NOTE_MARKER + '式中，' if kind == 'note' else '') + body.strip() + '\n\n')
        position = end + 1
    return ''.join(out)


def expand_supported_wrappers(segments: list[SourceSegment]) -> list[SourceSegment]:
    note_files = {segment.path for segment in segments if NOTE_DEFINITION.search(segment.text)}
    return [SourceSegment(segment.path, normalize_wrappers(segment.text, segment.path in note_files), segment.line)
            for segment in segments]
