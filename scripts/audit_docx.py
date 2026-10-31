#!/usr/bin/env python3
"""Report semantic Word structures produced by export-typst-to-docx."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS = {"w": W, "m": M, "wp": WP}
FIELD_RE = re.compile(r"\b(SEQ|REF|PAGE|PAGEREF|TOC|NUMPAGES|STYLEREF)\b", re.IGNORECASE)


def read_xml(package: zipfile.ZipFile, name: str):
    return etree.fromstring(package.read(name))


def field_counts(roots) -> Counter:
    counts = Counter()
    for root in roots:
        for text in root.xpath(".//w:instrText/text() | .//w:fldSimple/@w:instr", namespaces=NS):
            match = FIELD_RE.search(text)
            if match:
                counts[match.group(1).upper()] += 1
    return counts


def border_value(tbl_pr, edge: str):
    nodes = tbl_pr.xpath(f"./w:tblBorders/w:{edge}", namespaces=NS)
    if not nodes:
        return None
    node = nodes[0]
    return {
        "val": node.get(f"{{{W}}}val"),
        "sz": node.get(f"{{{W}}}sz"),
    }


def audit(path: Path) -> dict:
    with zipfile.ZipFile(path) as package:
        document = read_xml(package, "word/document.xml")
        related_roots = [document]
        for name in package.namelist():
            if re.fullmatch(r"word/(?:header|footer)\d+\.xml", name):
                related_roots.append(read_xml(package, name))

        styles = read_xml(package, "word/styles.xml")
        style_defs = {}
        for style in styles.xpath("./w:style", namespaces=NS):
            name_nodes = style.xpath("./w:name/@w:val", namespaces=NS)
            if name_nodes:
                style_defs[name_nodes[0]] = style.get(f"{{{W}}}styleId")

        pstyles = Counter(
            document.xpath(".//w:pPr/w:pStyle/@w:val", namespaces=NS)
        )
        table_styles = Counter(
            document.xpath(".//w:tbl/w:tblPr/w:tblStyle/@w:val", namespaces=NS)
        )
        tables = document.xpath(".//w:tbl", namespaces=NS)
        border_profiles = []
        for table in tables:
            tbl_pr = table.xpath("./w:tblPr", namespaces=NS)[0]
            border_profiles.append(
                {
                    edge: border_value(tbl_pr, edge)
                    for edge in ("top", "bottom", "left", "right", "insideH", "insideV")
                }
            )

        bookmark_names = document.xpath(
            ".//w:bookmarkStart/@w:name", namespaces=NS
        )
        bookmark_end_count = len(document.xpath(".//w:bookmarkEnd", namespaces=NS))

        return {
            "path": str(path),
            "paragraph_styles": dict(sorted(pstyles.items())),
            "defined_styles": dict(sorted(style_defs.items())),
            "fields": dict(sorted(field_counts(related_roots).items())),
            "editable_omath": len(document.xpath(".//m:oMath", namespaces=NS)),
            "display_omath_paragraphs": len(
                document.xpath(".//m:oMathPara", namespaces=NS)
            ),
            "inline_images": len(document.xpath(".//wp:inline", namespaces=NS)),
            "anchored_images": len(document.xpath(".//wp:anchor", namespaces=NS)),
            "tables": len(tables),
            "table_styles": dict(sorted(table_styles.items())),
            "table_border_profiles": border_profiles,
            "bookmark_starts": len(bookmark_names),
            "bookmark_ends": bookmark_end_count,
            "unique_bookmark_names": len(set(bookmark_names)),
        }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()
    source = args.docx.resolve(strict=True)
    payload = audit(source)
    rendered = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_text(rendered, encoding="utf-8")
    print(rendered)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
