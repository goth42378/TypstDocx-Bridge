#!/usr/bin/env python3
"""Build an editable DOCX from a constrained, semantic Typst chapter source."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from lxml import etree, html as lxml_html
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt, RGBColor


TOKEN_RE = re.compile(r"(\$[^$]+\$|@[A-Za-z0-9_:-]+)")
REF_RE = re.compile(r"^@((?:fig|tab|eqt):[A-Za-z0-9_:-]+)$")
SEPARATOR_CELL_RE = re.compile(r"^:?-{3,}:?$")
MATHML_NS = "http://www.w3.org/1998/Math/MathML"
OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
SEQ_NAMES = {
    "figure": "Figure",
    "table": "Table",
    "equation": "Equation",
}


@dataclass
class Block:
    kind: str
    text: str = ""
    level: int = 0
    path: str = ""
    width_pct: float = 90.0
    caption: str = ""
    label: str = ""
    rows: list[list[str]] = field(default_factory=list)
    number: str = ""
    bookmark: str = ""


@dataclass
class RefTarget:
    kind: str
    number: str
    bookmark: str


class BuildLog:
    def __init__(self) -> None:
        self.formula_failures: list[dict[str, str]] = []
        self.image_failures: list[dict[str, str]] = []
        self.unresolved_refs: list[str] = []
        self.counts: dict[str, int] = {}

    def bump(self, name: str, amount: int = 1) -> None:
        self.counts[name] = self.counts.get(name, 0) + amount


def count_balance(text: str, opening: str, closing: str) -> int:
    depth = 0
    in_string = False
    escaped = False
    for char in text:
        if escaped:
            escaped = False
            continue
        if char == "\\" and in_string:
            escaped = True
            continue
        if char == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if char == opening:
            depth += 1
        elif char == closing:
            depth -= 1
    return depth


def collect_parenthesized(lines: list[str], start: int) -> tuple[str, int]:
    parts: list[str] = []
    depth = 0
    begun = False
    index = start
    while index < len(lines):
        line = lines[index]
        parts.append(line)
        delta = count_balance(line, "(", ")")
        if "(" in line:
            begun = True
        depth += delta
        index += 1
        if begun and depth <= 0:
            break
    return "\n".join(parts), index


def collect_captab(lines: list[str], start: int) -> tuple[str, int]:
    parts: list[str] = []
    paren_depth = 0
    bracket_depth = 0
    body_started = False
    index = start
    while index < len(lines):
        line = lines[index]
        parts.append(line)
        paren_depth += count_balance(line, "(", ")")
        bracket_depth += count_balance(line, "[", "]")
        if re.search(r"\)\s*\[", line):
            body_started = True
        index += 1
        if body_started and paren_depth <= 0 and bracket_depth <= 0:
            break
    return "\n".join(parts), index


def collect_equation(lines: list[str], start: int) -> tuple[str, int]:
    parts: list[str] = []
    count = 0
    index = start
    while index < len(lines):
        line = lines[index]
        parts.append(line)
        count += len(re.findall(r"(?<!\\)\$", line))
        index += 1
        if count >= 2:
            break
    return "\n".join(parts), index


def starts_display_equation(line: str) -> bool:
    """Distinguish a display formula from prose that starts with inline math."""
    stripped = line.strip()
    if not stripped.startswith("$"):
        return False
    dollars = [match.start() for match in re.finditer(r"(?<!\\)\$", stripped)]
    if len(dollars) == 1:
        return True
    trailing = stripped[dollars[1] + 1 :].strip()
    return not trailing or re.fullmatch(r"<[^>]+>", trailing) is not None


def parse_caption(raw: str) -> str:
    match = re.search(r"caption\s*:\s*\[(.*?)\]\s*,?", raw, re.S)
    return re.sub(r"\s+", " ", match.group(1).strip()) if match else ""


def parse_label(raw: str) -> str:
    match = re.search(r"label\s*:\s*<([^>]+)>", raw)
    return match.group(1).strip() if match else ""


def parse_figure(raw: str) -> Block:
    image = re.search(r'image\s*\(\s*"([^"]+)"', raw, re.S)
    width = re.search(r"width\s*:\s*([0-9.]+)%", raw)
    return Block(
        kind="figure",
        path=image.group(1) if image else "",
        width_pct=float(width.group(1)) if width else 90.0,
        caption=parse_caption(raw),
        label=parse_label(raw),
    )


def parse_table(raw: str) -> Block:
    body_match = re.search(r"\)\s*\[(.*)\]\s*$", raw, re.S)
    rows: list[list[str]] = []
    if body_match:
        for line in body_match.group(1).splitlines():
            stripped = line.strip()
            if not stripped.startswith("|"):
                continue
            cells = [cell.strip() for cell in stripped.strip("|").split("|")]
            if cells and all(SEPARATOR_CELL_RE.fullmatch(cell.replace(" ", "")) for cell in cells):
                continue
            rows.append(cells)
    return Block(kind="table", caption=parse_caption(raw), label=parse_label(raw), rows=rows)


def parse_equation(raw: str) -> Block:
    first = raw.find("$")
    last = raw.rfind("$")
    math = raw[first + 1 : last].strip() if first >= 0 and last > first else raw.strip()
    trailing = raw[last + 1 :] if last >= 0 else ""
    label_match = re.search(r"<([^>]+)>", trailing)
    return Block(kind="equation", text=math, label=label_match.group(1) if label_match else "")


def parse_typst(content: Path) -> list[Block]:
    lines = content.read_text(encoding="utf-8-sig").splitlines()
    blocks: list[Block] = []
    paragraph: list[str] = []

    def flush() -> None:
        if paragraph:
            text = re.sub(r"\s+", " ", " ".join(item.strip() for item in paragraph)).strip()
            if text:
                blocks.append(Block(kind="paragraph", text=text))
            paragraph.clear()

    index = 0
    while index < len(lines):
        line = lines[index]
        stripped = line.strip()
        if not stripped:
            flush()
            index += 1
            continue
        if stripped.startswith("//"):
            flush()
            index += 1
            continue
        heading = re.match(r"^(=+)\s+(.*)$", stripped)
        if heading:
            flush()
            # Typst permits a trailing label on a heading, e.g.
            # `=== Controller history <sec:controller-history>`.  The label is
            # document metadata, not visible heading text, so do not leak it
            # into the Word title.
            heading_text = re.sub(r"\s*<[^<>]+>\s*$", "", heading.group(2)).strip()
            blocks.append(Block(kind="heading", level=len(heading.group(1)), text=heading_text))
            index += 1
            continue
        if stripped.startswith("#capfig("):
            flush()
            raw, index = collect_parenthesized(lines, index)
            blocks.append(parse_figure(raw))
            continue
        if stripped.startswith("#captab("):
            flush()
            raw, index = collect_captab(lines, index)
            blocks.append(parse_table(raw))
            continue
        if starts_display_equation(stripped):
            flush()
            raw, index = collect_equation(lines, index)
            blocks.append(parse_equation(raw))
            continue
        if stripped.startswith("#"):
            flush()
            if "(" in stripped:
                _, index = collect_parenthesized(lines, index)
            else:
                index += 1
            continue
        paragraph.append(line)
        index += 1
    flush()
    return blocks


def safe_bookmark(label: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9_]", "_", label)[:18]
    digest = hashlib.sha1(label.encode("utf-8")).hexdigest()[:12]
    return f"b_{stem}_{digest}"[:40]


def assign_numbers(blocks: list[Block], chapter: int) -> dict[str, RefTarget]:
    """Assign cached display numbers and reference metadata.

    The generated document uses Word SEQ fields for live numbering. These
    values remain as deterministic cached field results for headless rendering
    and as the initial cached results of REF fields.
    """
    refs: dict[str, RefTarget] = {}
    counters = {"figure": 0, "table": 0, "equation": 0}
    for block in blocks:
        if block.kind not in counters:
            continue
        counters[block.kind] += 1
        block.number = f"{chapter}-{counters[block.kind]}"
        if not block.label:
            continue
        block.bookmark = safe_bookmark(block.label)
        target = RefTarget(kind=block.kind, number=block.number, bookmark=block.bookmark)
        refs[block.label] = target
        if block.kind == "equation":
            refs[f"eqt:{block.label}"] = target
    return refs


THEME_FONT_ATTRIBUTES = ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme")


def set_explicit_rfonts(rfonts, latin: str, east_asia: str) -> None:
    """Set concrete Word fonts and remove theme references that override them."""
    for attribute in THEME_FONT_ATTRIBUTES:
        qualified = qn(f"w:{attribute}")
        if qualified in rfonts.attrib:
            del rfonts.attrib[qualified]
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    rfonts.set(qn("w:eastAsia"), east_asia)
    rfonts.set(qn("w:cs"), latin)


def set_half_point_size(rpr, size: float) -> None:
    value = str(int(round(size * 2)))
    for tag in ("w:sz", "w:szCs"):
        node = rpr.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            rpr.append(node)
        node.set(qn("w:val"), value)


def set_run_font(run, latin: str = "Times New Roman", east_asia: str = "宋体", size: float | None = None) -> None:
    run.font.name = latin
    if size is not None:
        run.font.size = Pt(size)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    set_explicit_rfonts(rfonts, latin, east_asia)
    if size is not None:
        set_half_point_size(rpr, size)


def set_style_font(style, latin: str, east_asia: str, size: float, color: str | None = None) -> None:
    style.font.name = latin
    style.font.size = Pt(size)
    if color:
        style.font.color.rgb = RGBColor.from_string(color)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    set_explicit_rfonts(rfonts, latin, east_asia)
    set_half_point_size(rpr, size)


def get_or_add_style(document, name: str, style_type: WD_STYLE_TYPE, base: str | None = None):
    try:
        style = document.styles[name]
    except KeyError:
        style = document.styles.add_style(name, style_type)
    if base:
        style.base_style = document.styles[base]
    return style


def set_outline_level(style, level: int) -> None:
    ppr = style.element.get_or_add_pPr()
    outline = ppr.find(qn("w:outlineLvl"))
    if outline is None:
        outline = OxmlElement("w:outlineLvl")
        ppr.append(outline)
    outline.set(qn("w:val"), str(level))


def set_style_tab_stops(style, stops: list[tuple[float, str]]) -> None:
    ppr = style.element.get_or_add_pPr()
    old = ppr.find(qn("w:tabs"))
    if old is not None:
        ppr.remove(old)
    tabs = OxmlElement("w:tabs")
    for inches, alignment in stops:
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), alignment)
        tab.set(qn("w:pos"), str(round(inches * 1440)))
        tabs.append(tab)
    ppr.append(tabs)


def set_table_borders(table) -> None:
    """Apply the NWPU three-line-table rule at table/header level."""
    tbl_pr = table._tbl.tblPr
    old = tbl_pr.find(qn("w:tblBorders"))
    if old is not None:
        tbl_pr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge, val, size in (
        ("top", "single", 12),
        ("left", "nil", 0),
        ("bottom", "single", 12),
        ("right", "nil", 0),
        ("insideH", "nil", 0),
        ("insideV", "nil", 0),
    ):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), val)
        if size:
            node.set(qn("w:sz"), str(size))
            node.set(qn("w:space"), "0")
            node.set(qn("w:color"), "000000")
        borders.append(node)
    tbl_pr.append(borders)

    if table.rows:
        for cell in table.rows[0].cells:
            tc_pr = cell._tc.get_or_add_tcPr()
            old_tc = tc_pr.find(qn("w:tcBorders"))
            if old_tc is not None:
                tc_pr.remove(old_tc)
            tc_borders = OxmlElement("w:tcBorders")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), "8")
            bottom.set(qn("w:space"), "0")
            bottom.set(qn("w:color"), "000000")
            tc_borders.append(bottom)
            tc_pr.append(tc_borders)


def set_cell_margins(cell, top: int = 60, start: int = 80, bottom: int = 60, end: int = 80) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_table_geometry(table, widths: list[int], indent: int = 0) -> None:
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.first_child_found_in("w:tblW")
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(sum(widths)))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_ind = tbl_pr.first_child_found_in("w:tblInd")
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), str(indent))
    tbl_ind.set(qn("w:type"), "dxa")
    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)
    for row in table.rows:
        for index, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.first_child_found_in("w:tcW")
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(widths[index]))
            tc_w.set(qn("w:type"), "dxa")
            cell.width = Inches(widths[index] / 1440)
            set_cell_margins(cell)


def repeat_table_header(row) -> None:
    """Mark a Word table row as a repeating header on continuation pages."""
    tr_pr = row._tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    tr_pr.append(header)


def column_widths(rows: list[list[str]], total: int = 9072) -> list[int]:
    count = max((len(row) for row in rows), default=1)
    weights = []
    for col in range(count):
        longest = max((len(re.sub(r"\$.*?\$", "MATH", row[col])) if col < len(row) else 0 for row in rows), default=1)
        weights.append(max(5, min(longest, 28)))
    raw = [max(720, round(total * weight / sum(weights))) for weight in weights]
    scale = total / sum(raw)
    widths = [round(value * scale) for value in raw]
    widths[-1] += total - sum(widths)
    return widths


def add_bookmark_start(paragraph, name: str, bookmark_id: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    paragraph._p.append(start)


def add_bookmark_end(paragraph, bookmark_id: int) -> None:
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    paragraph._p.append(end)


def add_seq_field(paragraph, sequence_name: str, cached_value: str, restart_heading_level: int = 1) -> None:
    """Append an editable Word SEQ field with a deterministic cached result."""
    run_begin = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    run_begin._r.append(begin)
    run_instr = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" SEQ {sequence_name} \\* ARABIC \\s {restart_heading_level} "
    run_instr._r.append(instr)
    run_sep = paragraph.add_run()
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    run_sep._r.append(separate)
    paragraph.add_run(cached_value)
    run_end = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run_end._r.append(end)


def add_automatic_number(paragraph, block: Block, chapter: int, bookmark_id: int) -> None:
    """Append chapter-SEQ numbering and bookmark the complete visible number."""
    if block.bookmark:
        add_bookmark_start(paragraph, block.bookmark, bookmark_id)
    paragraph.add_run(f"{chapter}-")
    cached_sequence = block.number.rsplit("-", 1)[-1]
    add_seq_field(paragraph, SEQ_NAMES[block.kind], cached_sequence)
    if block.bookmark:
        add_bookmark_end(paragraph, bookmark_id)


def add_ref_field(paragraph, target: RefTarget) -> None:
    run_begin = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    run_begin._r.append(begin)
    run_instr = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" REF {target.bookmark} \\h "
    run_instr._r.append(instr)
    run_sep = paragraph.add_run()
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    run_sep._r.append(separate)
    paragraph.add_run(target.number)
    run_end = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run_end._r.append(end)


class FormulaRenderer:
    """Convert Typst math to editable Word OMML through Typst HTML MathML."""

    def __init__(self, tinymist: Path, mml2omml: Path, cache: Path, log: BuildLog) -> None:
        self.tinymist = tinymist
        self.transform = etree.XSLT(etree.parse(str(mml2omml)))
        self.cache = cache
        self.log = log
        cache.mkdir(parents=True, exist_ok=True)

    def render(self, math: str, display: bool):
        digest = hashlib.sha256(("D" if display else "I").encode() + math.encode("utf-8")).hexdigest()[:20]
        typ = self.cache / f"formula-{digest}.typ"
        html_path = self.cache / f"formula-{digest}.html"
        omml_path = self.cache / f"formula-{digest}.omml.xml"
        try:
            if not omml_path.exists() or not omml_path.stat().st_size:
                source = f"$ {math} $\n" if display else f"Inline ${math}$ text.\n"
                typ.write_text(source, encoding="utf-8")
                command = [str(self.tinymist), "compile", "--root", str(self.cache), str(typ), str(html_path)]
                completed = subprocess.run(
                    command,
                    text=True,
                    capture_output=True,
                    encoding="utf-8",
                    errors="replace",
                )
                if completed.returncode != 0 or not html_path.exists():
                    raise RuntimeError((completed.stderr or completed.stdout)[-1200:])

                html_root = lxml_html.fromstring(html_path.read_text(encoding="utf-8"))
                math_nodes = html_root.xpath("//math")
                if not math_nodes:
                    raise RuntimeError("Tinymist HTML output contained no MathML <math> element")
                mathml = deepcopy(math_nodes[0])
                mathml.tail = None
                for element in mathml.iter():
                    if isinstance(element.tag, str) and not element.tag.startswith("{"):
                        element.tag = f"{{{MATHML_NS}}}{element.tag}"
                transformed = self.transform(mathml)
                root = transformed.getroot()
                if etree.QName(root).localname == "oMath":
                    omath = root
                else:
                    omath = root.find(f"{{{OMML_NS}}}oMath")
                if omath is None:
                    raise RuntimeError("MML2OMML transform produced no m:oMath element")
                omml_path.write_bytes(
                    etree.tostring(omath, encoding="utf-8", xml_declaration=True, with_tail=False)
                )
                self.log.bump("unique_omml_conversions")
            parsed = etree.fromstring(omml_path.read_bytes())
            self.log.bump("editable_formula_occurrences")
            return parsed
        except Exception as exc:
            self.log.formula_failures.append({"math": math, "error": str(exc)[-1200:]})
            return None


def prepare_image(source: Path, cache: Path, node: Path, svg_script: Path, log: BuildLog) -> Path | None:
    if not source.exists():
        log.image_failures.append({"path": str(source), "error": "missing"})
        return None
    if source.suffix.lower() != ".svg":
        return source
    cache.mkdir(parents=True, exist_ok=True)
    output = cache / f"svg-{hashlib.sha1(str(source).encode()).hexdigest()[:16]}.png"
    if output.exists() and output.stat().st_size:
        return output
    completed = subprocess.run(
        [str(node), str(svg_script), str(source), str(output), "2200"],
        text=True,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
    )
    if completed.returncode != 0 or not output.exists():
        log.image_failures.append({"path": str(source), "error": (completed.stderr or completed.stdout)[-1200:]})
        return None
    return output


def add_picture_run(run, path: Path, width_in: float | None = None, height_pt: float | None = None, alt: str = "") -> None:
    if height_pt is not None:
        run.add_picture(str(path), height=Pt(height_pt))
    else:
        run.add_picture(str(path), width=Inches(width_in or 5.8))
    drawings = run._r.xpath(".//wp:docPr")
    if drawings and alt:
        drawings[0].set("descr", alt)


def clean_plain(text: str) -> str:
    text = text.replace("`", "").replace("~", " ")
    text = re.sub(r"(?<!\*)\*([^*]+)\*(?!\*)", r"\1", text)
    return text


def add_inline_content(paragraph, text: str, refs: dict[str, RefTarget], formulas: FormulaRenderer, log: BuildLog) -> None:
    position = 0
    for match in TOKEN_RE.finditer(text):
        if match.start() > position:
            paragraph.add_run(clean_plain(text[position : match.start()]))
        token = match.group(0)
        ref_match = REF_RE.match(token)
        if ref_match:
            key = ref_match.group(1)
            target = refs.get(key)
            if target:
                if target.kind == "figure":
                    prefix, suffix = "图 ", ""
                elif target.kind == "table":
                    prefix, suffix = "表 ", ""
                else:
                    prefix, suffix = "式（", "）"
                paragraph.add_run(prefix)
                add_ref_field(paragraph, target)
                paragraph.add_run(suffix)
                log.bump("resolved_crossrefs")
            else:
                paragraph.add_run(token)
                log.unresolved_refs.append(key)
        elif token.startswith("$"):
            math = token[1:-1].strip()
            omath = formulas.render(math, display=False)
            if omath is not None:
                paragraph._p.append(omath)
            else:
                run = paragraph.add_run(math)
                set_run_font(run, latin="Cambria Math", east_asia="Cambria Math", size=10)
        else:
            paragraph.add_run(f"[{token[1:]}]")
            log.bump("citation_keys_preserved")
        position = match.end()
    if position < len(text):
        paragraph.add_run(clean_plain(text[position:]))


def create_template(path: Path) -> None:
    """Create a content-only NWPU 2025 style template with no guide prose."""
    document = Document()
    section = document.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(25.4)
    section.right_margin = Mm(25)
    section.bottom_margin = Mm(25.4)
    section.left_margin = Mm(25)
    section.header_distance = Mm(15)
    section.footer_distance = Mm(17.5)

    normal = document.styles["Normal"]
    set_style_font(normal, "Times New Roman", "宋体", 12, "000000")
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.first_line_indent = Pt(24)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = Pt(20)
    normal.paragraph_format.widow_control = True

    for level, name, size, alignment, before, after in (
        (0, "Heading 1", 16, WD_ALIGN_PARAGRAPH.CENTER, 10, 10),
        (1, "Heading 2", 14, WD_ALIGN_PARAGRAPH.LEFT, 10, 0),
        (2, "Heading 3", 12, WD_ALIGN_PARAGRAPH.LEFT, 10, 0),
    ):
        style = document.styles[name]
        # The guide applies 黑体 to the complete heading, including Latin text
        # and automatic numbering. Concrete fonts must replace Word theme refs.
        set_style_font(style, "黑体", "黑体", size, "000000")
        style.font.bold = False
        style.paragraph_format.alignment = alignment
        style.paragraph_format.first_line_indent = Pt(0)
        style.paragraph_format.left_indent = Pt(0)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.line_spacing = Pt(20)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.keep_together = True
        set_outline_level(style, level)

        try:
            character_style = document.styles[f"Heading {level + 1} Char"]
        except KeyError:
            character_style = None
        if character_style is not None:
            set_style_font(character_style, "黑体", "黑体", size, "000000")
            character_style.font.bold = False

    caption = document.styles["Caption"]
    caption.base_style = normal
    set_style_font(caption, "Times New Roman", "宋体", 10.5, "000000")
    caption.font.italic = False
    caption.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption.paragraph_format.first_line_indent = Pt(0)
    caption.paragraph_format.space_before = Pt(0)
    caption.paragraph_format.space_after = Pt(0)
    caption.paragraph_format.line_spacing = Pt(20)
    caption.paragraph_format.keep_together = True

    figure = get_or_add_style(document, "NWPU Figure", WD_STYLE_TYPE.PARAGRAPH, "Normal")
    set_style_font(figure, "Times New Roman", "宋体", 12, "000000")
    figure.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    figure.paragraph_format.first_line_indent = Pt(0)
    figure.paragraph_format.space_before = Pt(10)
    figure.paragraph_format.space_after = Pt(0)
    figure.paragraph_format.line_spacing = 1.0
    figure.paragraph_format.keep_with_next = True
    figure.paragraph_format.keep_together = True

    equation = get_or_add_style(document, "NWPU Equation", WD_STYLE_TYPE.PARAGRAPH, "Normal")
    set_style_font(equation, "Times New Roman", "宋体", 12, "000000")
    equation.paragraph_format.first_line_indent = Pt(0)
    equation.paragraph_format.space_before = Pt(3)
    equation.paragraph_format.space_after = Pt(3)
    equation.paragraph_format.line_spacing = 1.0
    equation.paragraph_format.keep_together = True
    set_style_tab_stops(equation, [(3.15, "center"), (6.30, "right")])

    table_text = get_or_add_style(document, "NWPU Table Text", WD_STYLE_TYPE.PARAGRAPH, "Normal")
    set_style_font(table_text, "Times New Roman", "宋体", 10.5, "000000")
    table_text.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table_text.paragraph_format.first_line_indent = Pt(0)
    table_text.paragraph_format.space_before = Pt(0)
    table_text.paragraph_format.space_after = Pt(0)
    table_text.paragraph_format.line_spacing = Pt(20)

    table_header = get_or_add_style(document, "NWPU Table Header", WD_STYLE_TYPE.PARAGRAPH, "NWPU Table Text")
    set_style_font(table_header, "Times New Roman", "宋体", 10.5, "000000")
    table_header.font.bold = False
    table_header.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table_header.paragraph_format.first_line_indent = Pt(0)
    table_header.paragraph_format.space_before = Pt(0)
    table_header.paragraph_format.space_after = Pt(0)
    table_header.paragraph_format.line_spacing = Pt(20)
    table_header.paragraph_format.keep_with_next = True

    spacer = get_or_add_style(document, "NWPU Table Spacer", WD_STYLE_TYPE.PARAGRAPH, "Normal")
    set_style_font(spacer, "Times New Roman", "宋体", 6, "000000")
    spacer.paragraph_format.first_line_indent = Pt(0)
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(0)
    spacer.paragraph_format.line_spacing = Pt(6)

    table_style = get_or_add_style(document, "NWPU Three-Line Table", WD_STYLE_TYPE.TABLE)
    set_style_font(table_style, "Times New Roman", "宋体", 10.5, "000000")

    footer = document.styles["Footer"]
    set_style_font(footer, "Times New Roman", "宋体", 9, "000000")
    footer.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.first_line_indent = Pt(0)
    footer.paragraph_format.space_before = Pt(0)
    footer.paragraph_format.space_after = Pt(0)

    attach_heading_numbering(document, 1)
    document.save(path)


def attach_heading_numbering(document, chapter_number: int) -> int:
    """Link a real three-level outline list to Heading 1/2/3 styles."""
    numbering = document.part.numbering_part.element
    abstract_ids = [
        int(node.get(qn("w:abstractNumId")))
        for node in numbering.findall(qn("w:abstractNum"))
    ]
    num_ids = [int(node.get(qn("w:numId"))) for node in numbering.findall(qn("w:num"))]
    abstract_id = max(abstract_ids, default=0) + 1
    num_id = max(num_ids, default=0) + 1

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "multilevel")
    abstract.append(multi)

    for level, (style_name, level_text, font_size) in enumerate(
        (
            ("Heading1", "第 %1 章", 16),
            ("Heading2", "%1.%2", 14),
            ("Heading3", "%1.%2.%3", 12),
        )
    ):
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(level))
        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        num_fmt = OxmlElement("w:numFmt")
        num_fmt.set(qn("w:val"), "decimal")
        pstyle = OxmlElement("w:pStyle")
        pstyle.set(qn("w:val"), style_name)
        text = OxmlElement("w:lvlText")
        text.set(qn("w:val"), level_text)
        suff = OxmlElement("w:suff")
        suff.set(qn("w:val"), "space")
        jc = OxmlElement("w:lvlJc")
        jc.set(qn("w:val"), "left")
        ppr = OxmlElement("w:pPr")
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), "0")
        ind.set(qn("w:hanging"), "0")
        ppr.append(ind)
        rpr = OxmlElement("w:rPr")
        rfonts = OxmlElement("w:rFonts")
        set_explicit_rfonts(rfonts, "黑体", "黑体")
        rpr.append(rfonts)
        set_half_point_size(rpr, font_size)
        bold = OxmlElement("w:b")
        bold.set(qn("w:val"), "0")
        rpr.append(bold)
        for node in (start, num_fmt, pstyle, text, suff, jc, ppr, rpr):
            lvl.append(node)
        abstract.append(lvl)
    numbering.append(abstract)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(abstract_id))
    num.append(abstract_ref)
    override = OxmlElement("w:lvlOverride")
    override.set(qn("w:ilvl"), "0")
    start_override = OxmlElement("w:startOverride")
    start_override.set(qn("w:val"), str(chapter_number))
    override.append(start_override)
    num.append(override)
    numbering.append(num)

    for level, style_name in enumerate(("Heading 1", "Heading 2", "Heading 3")):
        style = document.styles[style_name]
        ppr = style.element.get_or_add_pPr()
        old = ppr.find(qn("w:numPr"))
        if old is not None:
            ppr.remove(old)
        num_pr = OxmlElement("w:numPr")
        ilvl = OxmlElement("w:ilvl")
        ilvl.set(qn("w:val"), str(level))
        num_ref = OxmlElement("w:numId")
        num_ref.set(qn("w:val"), str(num_id))
        num_pr.append(ilvl)
        num_pr.append(num_ref)
        ppr.append(num_pr)
    return num_id


def add_page_number(paragraph) -> None:
    paragraph.style = "Footer"
    run_begin = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    run_begin._r.append(begin)
    run_instr = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    run_instr._r.append(instr)
    run_sep = paragraph.add_run()
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    run_sep._r.append(separate)
    paragraph.add_run("1")
    run_end = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run_end._r.append(end)


def build_document(args) -> BuildLog:
    log = BuildLog()
    blocks = parse_typst(args.content)
    refs = assign_numbers(blocks, args.chapter_number)
    document = Document(args.template)
    attach_heading_numbering(document, args.chapter_number)
    add_page_number(document.sections[0].footer.paragraphs[0])
    settings = document.settings.element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")

    formula_renderer = FormulaRenderer(args.tinymist, args.mml2omml, args.cache / "formulas", log)
    bookmark_id = 1
    for block in blocks:
        log.bump(block.kind)
        if block.kind == "heading":
            level = min(3, max(1, block.level))
            paragraph = document.add_paragraph(style=f"Heading {level}")
            add_inline_content(paragraph, block.text, refs, formula_renderer, log)
        elif block.kind == "paragraph":
            paragraph = document.add_paragraph(style="Normal")
            add_inline_content(paragraph, block.text, refs, formula_renderer, log)
        elif block.kind == "equation":
            paragraph = document.add_paragraph(style="NWPU Equation")
            paragraph.add_run("\t")
            omath = formula_renderer.render(block.text, display=True)
            if omath is not None:
                paragraph._p.append(omath)
            else:
                run = paragraph.add_run(block.text)
                set_run_font(run, latin="Cambria Math", east_asia="Cambria Math", size=10)
            paragraph.add_run("\t")
            paragraph.add_run("（")
            add_automatic_number(paragraph, block, args.chapter_number, bookmark_id)
            paragraph.add_run("）")
            log.bump("seq_fields")
            if block.bookmark:
                bookmark_id += 1
        elif block.kind == "figure":
            source = (args.content.parent / block.path).resolve()
            image = prepare_image(source, args.cache / "images", args.node, args.svg_script, log)
            if image:
                paragraph = document.add_paragraph(style="NWPU Figure")
                run = paragraph.add_run()
                width = min(6.1, 6.3 * block.width_pct / 100)
                add_picture_run(run, image, width_in=width, alt=f"{block.label}: {block.caption}")
            else:
                paragraph = document.add_paragraph(f"[图片缺失：{block.path}]", style="NWPU Figure")
            caption = document.add_paragraph(style="Caption")
            caption.add_run("图 ")
            add_automatic_number(caption, block, args.chapter_number, bookmark_id)
            log.bump("seq_fields")
            if block.bookmark:
                bookmark_id += 1
            caption.add_run("  ")
            add_inline_content(caption, block.caption, refs, formula_renderer, log)
        elif block.kind == "table":
            caption = document.add_paragraph(style="Caption")
            caption.add_run("表 ")
            add_automatic_number(caption, block, args.chapter_number, bookmark_id)
            log.bump("seq_fields")
            if block.bookmark:
                bookmark_id += 1
            caption.add_run("  ")
            add_inline_content(caption, block.caption, refs, formula_renderer, log)
            if block.rows:
                col_count = max(len(row) for row in block.rows)
                table = document.add_table(rows=len(block.rows), cols=col_count)
                table.style = "NWPU Three-Line Table"
                widths = column_widths(block.rows)
                set_table_geometry(table, widths)
                set_table_borders(table)
                repeat_table_header(table.rows[0])
                for row_index, source_row in enumerate(block.rows):
                    for col_index in range(col_count):
                        cell = table.cell(row_index, col_index)
                        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                        paragraph = cell.paragraphs[0]
                        paragraph.style = "NWPU Table Header" if row_index == 0 else "NWPU Table Text"
                        add_inline_content(
                            paragraph,
                            source_row[col_index] if col_index < len(source_row) else "",
                            refs,
                            formula_renderer,
                            log,
                        )
                document.add_paragraph(style="NWPU Table Spacer")

    args.output.parent.mkdir(parents=True, exist_ok=True)
    document.save(args.output)
    log.counts["reference_targets"] = len({target.bookmark for target in refs.values()})
    return log


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--content", required=True, type=Path)
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--cache", required=True, type=Path)
    parser.add_argument("--tinymist", required=True, type=Path)
    parser.add_argument("--mml2omml", required=True, type=Path)
    parser.add_argument("--node", required=True, type=Path)
    parser.add_argument("--svg-script", type=Path, default=Path(__file__).with_name("svg_to_png.mjs"))
    parser.add_argument("--chapter-number", type=int, default=2)
    parser.add_argument("--create-template", action="store_true")
    args = parser.parse_args()

    args.content = args.content.resolve(strict=True)
    args.tinymist = args.tinymist.resolve(strict=True)
    args.mml2omml = args.mml2omml.resolve(strict=True)
    args.node = args.node.resolve(strict=True)
    args.svg_script = args.svg_script.resolve(strict=True)
    if args.create_template:
        args.template.parent.mkdir(parents=True, exist_ok=True)
        create_template(args.template)
    elif not args.template.exists():
        raise SystemExit(f"Template not found: {args.template}")
    args.cache.mkdir(parents=True, exist_ok=True)
    log = build_document(args)
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(
        json.dumps(
            {
                "content": str(args.content),
                "output": str(args.output),
                "counts": log.counts,
                "unresolved_refs": sorted(set(log.unresolved_refs)),
                "formula_failures": log.formula_failures,
                "image_failures": log.image_failures,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(json.dumps(log.counts, ensure_ascii=False))
    print(f"Unresolved refs: {len(set(log.unresolved_refs))}")
    print(f"Formula failures: {len(log.formula_failures)}")
    print(f"Image failures: {len(log.image_failures)}")
    return 0 if not (log.image_failures or log.formula_failures or log.unresolved_refs) else 2


if __name__ == "__main__":
    raise SystemExit(main())
