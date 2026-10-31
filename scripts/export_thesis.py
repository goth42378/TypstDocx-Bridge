#!/usr/bin/env python3
"""Stage and export supported static Typst chapter includes as one NWPU DOCX."""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from collections import Counter
from pathlib import Path
from types import SimpleNamespace

import build_docx
import front_matter
import preflight_typst
import semantic_expansion


def semantic_sources(path: Path, root: Path, stack: tuple[Path, ...] = ()) -> tuple[list[semantic_expansion.SourceSegment], bool]:
    segments, wrapper = semantic_expansion.source_segments(path, root, stack)
    return semantic_expansion.expand_supported_wrappers(segments), wrapper


def check_sources(sources: list[semantic_expansion.SourceSegment | Path]) -> dict:
    segments = [source if isinstance(source, semantic_expansion.SourceSegment)
                else semantic_expansion.SourceSegment(source, source.read_text(encoding='utf-8-sig')) for source in sources]
    audits = [preflight_typst.audit(segment.path, segment.text) for segment in segments]
    all_labels = []
    all_refs = []
    for segment in segments:
        lines = [preflight_typst.strip_line_comment(line) for line in segment.text.splitlines()]
        all_labels.extend(label for line in lines for label in preflight_typst.LABEL_RE.findall(line))
        all_refs.extend(ref for line in lines for ref in preflight_typst.CROSSREF_RE.findall(line))
    counts = Counter(all_labels)
    duplicates = sorted(label for label, count in counts.items() if count > 1)
    targets = set(all_labels) | {f"eqt:{label}" for label in all_labels if ":" not in label}
    unresolved = sorted(set(all_refs) - targets)
    commands = [{"source": str(segment.path), "source_segment_start_line": segment.line,
                 "normalized_line": issue["line"], "command": issue["command"], "text": issue["text"]}
                for segment, audit in zip(segments, audits) for issue in audit["review_required_commands"]]
    heading_count = sum(audit["heading1_count"] for audit in audits)
    issues = []
    if not heading_count:
        issues.append("no level-1 chapter headings found")
    if duplicates:
        issues.append(f"duplicate labels: {duplicates}")
    if unresolved:
        issues.append(f"unresolved references: {unresolved}")
    if commands:
        issues.append(f"unsupported top-level commands: {commands}")
    return {"heading1_count": heading_count, "duplicate_labels": duplicates,
            "unresolved_crossrefs": unresolved, "review_required_commands": commands, "blocking_issues": issues}


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--source-root", required=True, type=Path)
    p.add_argument("--entry", required=True, help="Entry path relative to source root")
    p.add_argument("--work-dir", required=True, type=Path, help="New directory outside source root")
    p.add_argument("--output", required=True, type=Path, help="New DOCX outside source root")
    p.add_argument("--tinymist", required=True, type=Path)
    p.add_argument("--mml2omml", required=True, type=Path)
    p.add_argument("--node", required=True, type=Path)
    p.add_argument("--template", type=Path, default=Path(__file__).parent.parent / "assets/nwpu-2025-content-template.docx")
    p.add_argument("--even-header", help="Even-page header; defaults to the degree in front matter")
    p.add_argument("--front-matter-json", type=Path, help="Static metadata JSON if entry has no NWPU wrapper")
    p.add_argument("--populate-front-matter", action="store_true", help="Fill cover and abstracts from source or JSON; default leaves them blank")
    p.add_argument("--start-chapter", type=int, default=1)
    args = p.parse_args()
    root = args.source_root.resolve(strict=True)
    work, output = args.work_dir.resolve(), args.output.resolve()
    if work == root or work.is_relative_to(root) or output == root or output.is_relative_to(root):
        p.error("work directory and output must be outside the source root")
    if work.exists() or output.exists():
        p.error("work directory and output must be new; existing files are not overwritten")
    if args.start_chapter < 1:
        p.error("start chapter must be positive")
    script_dir = Path(__file__).resolve().parent
    stage = work / "staged-project"
    stage_result = subprocess.run([sys.executable, str(script_dir / "stage_project.py"), "--source-root", str(root),
                                   "--entry", args.entry, "--destination", str(stage)], check=False)
    if stage_result.returncode:
        return stage_result.returncode
    report = {"status": "failed", "output": str(output), "source_verification": "not_run",
              "warnings": [], "contents": []}
    try:
        sources, _ = semantic_sources(stage / args.entry, stage.resolve())
        report["contents"] = list(dict.fromkeys(str(segment.path.relative_to(stage.resolve())) for segment in sources))
        report["source_segments"] = len(sources)
        if args.front_matter_json:
            front = json.loads(args.front_matter_json.resolve(strict=True).read_text(encoding="utf-8-sig"))
            if not isinstance(front, dict):
                raise ValueError("front-matter JSON must be an object")
        else:
            front = front_matter.read_front_matter(stage / args.entry, include_contents=args.populate_front_matter) or {}
        if args.populate_front_matter:
            front_matter.enrich_major_translation(front, stage)
        degree = front.get("degree", "doctor")
        if degree not in ("doctor", "master"):
            raise ValueError("degree must be doctor or master")
        even_header = args.even_header or ("西北工业大学博士学位论文" if degree == "doctor" else "西北工业大学硕士学位论文")
        report["warnings"].append("Bibliography is not exported; citation keys remain in the text.")
        preflight = check_sources(sources)
        report["preflight"] = preflight
        if preflight["blocking_issues"]:
            raise ValueError("; ".join(preflight["blocking_issues"]))
        opts = SimpleNamespace(content=sources[0].path, segments=sources, full_document=True, staged_root=stage.resolve(),
                               chapter_number=args.start_chapter, even_header=even_header, front_matter=front,
                               populate_front_matter=args.populate_front_matter,
                               template=args.template.resolve(strict=True), output=work / "candidate.docx",
                               report=work / "conversion-report.json", cache=work / "cache",
                               tinymist=args.tinymist.resolve(strict=True), mml2omml=args.mml2omml.resolve(strict=True),
                               node=args.node.resolve(strict=True), svg_script=script_dir / "svg_to_png.mjs")
        log = build_docx.build_document(opts)
        report.update({"counts": log.counts, "unresolved_refs": sorted(set(log.unresolved_refs)),
                       "formula_failures": log.formula_failures, "image_failures": log.image_failures,
                       "front_matter_mode": "populated" if args.populate_front_matter else "blank",
                       "missing_front_matter": getattr(log, "front_missing", [])})
        if report["missing_front_matter"]:
            report["warnings"].append("Some front-matter fields are placeholders and need user review.")
        if log.unresolved_refs or log.formula_failures or log.image_failures:
            raise ValueError("conversion has unresolved references, formula failures or image failures")
        report["status"] = "passed"
    except Exception as exc:
        report["error"] = str(exc)
    finally:
        verify = subprocess.run([sys.executable, str(script_dir / "verify_source_manifest.py"),
                                 str(stage / "export-manifest.json"), "--json", str(work / "source-verification.json")], check=False)
        report["source_verification"] = "passed" if verify.returncode == 0 else "failed"
        if verify.returncode:
            report["status"] = "failed"
        if report["status"] == "passed":
            try:
                if output.exists():
                    raise FileExistsError(f"output appeared during conversion: {output}")
                output.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(work / "candidate.docx"), str(output))
            except Exception as exc:
                report["status"] = "failed"
                report["error"] = f"could not safely publish output: {exc}"
        (work / "conversion-report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "passed" else 2


if __name__ == "__main__":
    raise SystemExit(main())
