#!/usr/bin/env python3
"""Verify that staged and read-only source files still match the staging manifest."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def snapshot(path: Path) -> dict[str, object]:
    return {
        "sha256": sha256(path),
        "size": path.stat().st_size,
        "mtime_ns": path.stat().st_mtime_ns,
    }


def verify_tree(root: Path, records: list[dict[str, object]], check_mtime: bool) -> dict[str, object]:
    mismatches: list[dict[str, object]] = []
    unchanged = 0
    for record in records:
        relative = Path(str(record["path"]))
        candidate = (root / relative).resolve()
        try:
            candidate.relative_to(root)
        except ValueError:
            mismatches.append({"path": relative.as_posix(), "reason": "path escapes root"})
            continue
        if not candidate.is_file():
            mismatches.append({"path": relative.as_posix(), "reason": "missing"})
            continue
        actual = snapshot(candidate)
        expected = {"sha256": record["sha256"], "size": record["size"]}
        if check_mtime:
            expected["mtime_ns"] = record["mtime_ns"]
        compared = {key: actual[key] for key in expected}
        if compared != expected:
            mismatches.append(
                {
                    "path": relative.as_posix(),
                    "reason": "metadata or content changed",
                    "expected": expected,
                    "actual": compared,
                }
            )
        else:
            unchanged += 1
    return {
        "root": str(root),
        "checked": len(records),
        "unchanged": unchanged,
        "mismatches": mismatches,
        "status": "passed" if not mismatches else "failed",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", type=Path)
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()

    manifest_path = args.manifest.resolve(strict=True)
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    records = payload.get("copied_files", [])
    source_root = Path(payload["source_root"]).resolve(strict=True)
    staged_root = manifest_path.parent.resolve(strict=True)

    result = {
        "manifest": str(manifest_path),
        "source": verify_tree(source_root, records, check_mtime=True),
        "staged_copy": verify_tree(staged_root, records, check_mtime=False),
    }
    result["status"] = (
        "passed"
        if result["source"]["status"] == "passed" and result["staged_copy"]["status"] == "passed"
        else "failed"
    )
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_text(rendered, encoding="utf-8")
    print(rendered)
    return 0 if result["status"] == "passed" else 2


if __name__ == "__main__":
    raise SystemExit(main())
