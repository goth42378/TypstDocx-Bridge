#!/usr/bin/env python3
"""Copy a static Typst dependency closure without writing to the source tree."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
from collections import deque
from pathlib import Path


PATH_PATTERNS = (
    re.compile(r'#(?:import|include)\s*["\']([^"\']+)["\']'),
    re.compile(r'\b(?:image|read|csv|json|yaml|bibliography)\s*\(\s*["\']([^"\']+)["\']'),
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def resolve_dependency(raw: str, owner: Path, root: Path) -> Path | None:
    if raw.startswith("@") or "://" in raw:
        return None
    candidate = root / raw.lstrip("/\\") if raw.startswith(("/", "\\")) else owner.parent / raw
    try:
        resolved = candidate.resolve(strict=True)
    except FileNotFoundError:
        return None
    try:
        resolved.relative_to(root)
    except ValueError as exc:
        raise ValueError(f"Dependency escapes source root: {raw} from {owner}") from exc
    return resolved


def discover(root: Path, entry: Path) -> tuple[list[Path], list[dict[str, str]]]:
    queue: deque[Path] = deque([entry])
    seen: set[Path] = set()
    missing: list[dict[str, str]] = []

    while queue:
        current = queue.popleft().resolve()
        if current in seen:
            continue
        seen.add(current)
        if current.suffix.lower() != ".typ":
            continue
        text = current.read_text(encoding="utf-8-sig")
        for pattern in PATH_PATTERNS:
            for match in pattern.finditer(text):
                raw = match.group(1)
                resolved = resolve_dependency(raw, current, root)
                if resolved is None:
                    if not raw.startswith("@") and "://" not in raw:
                        missing.append({"owner": str(current.relative_to(root)), "path": raw})
                    continue
                if resolved.is_dir():
                    queue.extend(path for path in resolved.rglob("*") if path.is_file())
                else:
                    queue.append(resolved)

    return sorted(seen), missing


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-root", required=True, type=Path)
    parser.add_argument("--entry", required=True, help="Entry path relative to source root")
    parser.add_argument("--destination", required=True, type=Path)
    args = parser.parse_args()

    root = args.source_root.resolve(strict=True)
    entry = (root / args.entry).resolve(strict=True)
    destination = args.destination.resolve()
    try:
        entry.relative_to(root)
    except ValueError as exc:
        raise SystemExit(f"Entry is outside source root: {entry}") from exc
    if destination == root or root in destination.parents:
        raise SystemExit("Destination must not be inside the read-only source tree")
    if destination.exists():
        raise SystemExit(f"Destination already exists: {destination}")

    files, missing = discover(root, entry)
    before = {
        path: {"sha256": sha256(path), "size": path.stat().st_size, "mtime_ns": path.stat().st_mtime_ns}
        for path in files
    }

    destination.mkdir(parents=True)
    records: list[dict[str, object]] = []
    for source_path in files:
        relative = source_path.relative_to(root)
        target_path = destination / relative
        target_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, target_path)
        record = {"path": relative.as_posix(), **before[source_path]}
        records.append(record)

    changed = []
    for source_path, snapshot in before.items():
        current = {
            "sha256": sha256(source_path),
            "size": source_path.stat().st_size,
            "mtime_ns": source_path.stat().st_mtime_ns,
        }
        if current != snapshot:
            changed.append(str(source_path))
    if changed:
        raise SystemExit("Source files changed during staging: " + ", ".join(changed))

    manifest = {
        "source_root": str(root),
        "entry": str(entry.relative_to(root)).replace("\\", "/"),
        "copied_files": records,
        "missing_static_dependencies": missing,
        "source_mutation_check": "passed",
    }
    (destination / "export-manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"Copied {len(records)} files to {destination}")
    print(f"Missing static dependencies: {len(missing)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
