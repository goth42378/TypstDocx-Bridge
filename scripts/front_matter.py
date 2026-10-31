"""Read static NWPU Typst front matter and create editable Word front pages."""
from __future__ import annotations

import json
import re
import textwrap
from pathlib import Path

from docx.enum.section import WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from build_docx import add_complex_field, add_inline_content, add_page_number, set_outline_level, set_run_font, set_style_font


def split_top_level(text: str, separator: str = ',') -> list[str]:
    result, stack, start = [], [], 0
    in_string, escaped = False, False
    pairs = {')': '(', ']': '[', '}': '{'}
    for index, char in enumerate(text):
        if in_string:
            if escaped:
                escaped = False
            elif char == '\\':
                escaped = True
            elif char == '"':
                in_string = False
            continue
        if char == '"':
            in_string = True
        elif char in '([{':
            stack.append(char)
        elif char in ')]}':
            if not stack or stack.pop() != pairs[char]:
                raise ValueError('unbalanced Typst front-matter expression')
        elif char == separator and not stack:
            result.append(text[start:index].strip())
            start = index + 1
    if in_string or stack:
        raise ValueError('unbalanced Typst front-matter expression')
    result.append(text[start:].strip())
    return [part for part in result if part]


def parse_value(value: str):
    value = value.strip()
    if value.startswith('"'):
        try:
            return json.loads(value)
        except json.JSONDecodeError as exc:
            raise ValueError(f'unsupported Typst string in front matter: {exc}') from exc
    if value in ('true', 'false'):
        return value == 'true'
    if re.fullmatch(r'\d+', value):
        return int(value)
    if value.startswith('[') and value.endswith(']'):
        content = textwrap.dedent(value[1:-1]).strip()
        if re.search(r'(?<!\\)[#@]', content):
            raise ValueError('abstract uses a Typst command or citation; supply resolved plain text')
        if '[' in content or ']' in content:
            raise ValueError('abstract uses nested Typst content; supply resolved plain text')
        return content
    if value.startswith('(') and value.endswith(')'):
        parts = split_top_level(value[1:-1])
        if parts == [':'] or not parts:
            return {}
        if all(re.match(r'^[A-Za-z][A-Za-z0-9_-]*\s*:', p) for p in parts):
            return parse_map(value[1:-1])
        return [parse_value(part) for part in parts]
    return {'unsupported_expression': value}


def parse_map(text: str, allowed: set[str] | None = None) -> dict:
    result = {}
    for part in split_top_level(text):
        match = re.match(r'^([A-Za-z][A-Za-z0-9_-]*)\s*:\s*(.*)$', part, re.S)
        if not match:
            raise ValueError('unsupported Typst front-matter field')
        key, value = match.group(1), match.group(2).strip()
        if allowed is not None and key not in allowed:
            continue
        if key == 'info' and value.startswith('(') and value.endswith(')'):
            fields = {'title', 'title-en', 'author', 'author-en', 'supervisor', 'supervisor-en',
                      'student-id', 'class-no', 'department', 'major', 'major-en', 'submit-date'}
            result[key] = parse_map(value[1:-1], fields)
        else:
            result[key] = parse_value(value)
    return result


def read_front_matter(entry: Path, include_contents: bool = True) -> dict | None:
    text = entry.read_text(encoding='utf-8-sig')
    match = re.search(r'#show\s*:\s*nwpu-thesis\.with\s*\(', text)
    if not match:
        return None
    start, stack = match.end(), ['(']
    in_string, escaped = False, False
    pairs = {')': '(', ']': '[', '}': '{'}
    for i in range(start, len(text)):
        char = text[i]
        if in_string:
            if escaped:
                escaped = False
            elif char == '\\':
                escaped = True
            elif char == '"':
                in_string = False
            continue
        if char == '"':
            in_string = True
        elif char in '([{':
            stack.append(char)
        elif char in ')]}':
            if not stack or stack.pop() != pairs[char]:
                raise ValueError('unbalanced NWPU thesis wrapper')
            if not stack:
                allowed = {'graduate', 'degree', 'track', 'anonymous', 'colored-cover'}
                if include_contents:
                    allowed |= {'info', 'abstract', 'abstract-en'}
                data = parse_map(text[start:i], allowed)
                if include_contents and data.get('colored-cover', False) is True:
                    raise ValueError('colored cover backgrounds are not yet supported; use an uncolored cover')
                if data.get('graduate', True) is False:
                    raise ValueError('undergraduate cover layout is not supported')
                return data
    raise ValueError('unterminated NWPU thesis wrapper')


def enrich_major_translation(data: dict, staged_root: Path) -> None:
    info = data.get('info')
    if not isinstance(info, dict) or info.get('major-en') or not isinstance(info.get('major'), str):
        return
    utilities = staged_root / 'lib' / 'utils.typ'
    if not utilities.is_file():
        return
    text = utilities.read_text(encoding='utf-8-sig')
    match = re.search(r'#let\s+major-en-map\s*=\s*\((.*?)\n\)', text, re.S)
    if not match:
        return
    pairs = re.findall(r'"([^"\n]+)"\s*:\s*"([^"\n]+)"', match.group(1))
    for chinese, english in pairs:
        if chinese == info['major']:
            info['major-en'] = english
            return


def literal(value, name: str, missing: list[str], anonymous: bool = False) -> str:
    if anonymous:
        return '__________'
    if isinstance(value, str) and value.strip():
        return value.strip()
    if isinstance(value, list) and value and all(isinstance(item, str) for item in value):
        return ' '.join(item for item in value if item.strip())
    if isinstance(value, dict) and 'unsupported_expression' in value:
        raise ValueError(f'front matter {name} is dynamic; supply static values')
    missing.append(name)
    return '【待填写】'


def add_style(document, name: str, latin: str, chinese: str, size: int, bold: bool = False):
    styles = document.styles
    style = styles[name] if name in styles else styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    style.base_style = styles['Normal']
    set_style_font(style, latin, chinese, size, '000000')
    style.font.bold = bold
    style.paragraph_format.first_line_indent = Pt(0)
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    return style


def page_numbering(section, fmt: str, start: int | None):
    node = section._sectPr.find(qn('w:pgNumType'))
    if node is None:
        node = OxmlElement('w:pgNumType')
        section._sectPr.append(node)
    node.set(qn('w:fmt'), fmt)
    if start is None:
        node.attrib.pop(qn('w:start'), None)
    else:
        node.set(qn('w:start'), str(start))


def add_header_rule(paragraph):
    borders = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    for key, value in (('val', 'thinThickSmallGap'), ('sz', '24'), ('space', '1'), ('color', 'auto')):
        bottom.set(qn('w:' + key), value)
    borders.append(bottom)
    paragraph._p.get_or_add_pPr().append(borders)


def configure_front_section(section, title: str, even_header: str, first: bool):
    page_numbering(section, 'upperRoman', 1 if first else None)
    for header, text in ((section.header, title), (section.even_page_header, even_header)):
        header.is_linked_to_previous = False
        p = header.paragraphs[0]
        p.clear()
        p.style = 'Header'
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        set_run_font(run, latin='Times New Roman', east_asia='宋体', size=9)
        add_header_rule(p)
    for footer in (section.footer, section.even_page_footer):
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0]
        p.clear()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_page_number(p)


def add_blank(p, before=0, after=0):
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    return p


def add_cover(document, data: dict, missing: list[str], blank: bool):
    info = data.get('info', {})
    if not isinstance(info, dict) or 'unsupported_expression' in info:
        raise ValueError('NWPU info must be a static dictionary')
    degree = data.get('degree', 'doctor')
    if degree not in ('doctor', 'master'):
        raise ValueError('degree must be doctor or master')
    anonymous = data.get('anonymous', False)
    if not isinstance(anonymous, bool):
        raise ValueError('anonymous must be a static boolean')
    def field(name: str, mask: bool = False) -> str:
        return '__________' if blank else literal(info.get(name), name, missing, mask)
    title = field('title')
    title_en = field('title-en')
    author = field('author', anonymous)
    author_en = field('author-en', anonymous)
    supervisor = field('supervisor', anonymous)
    supervisor_en = field('supervisor-en', anonymous)
    student_id = field('student-id', anonymous)
    class_no = field('class-no')
    department = field('department')
    major = field('major')
    major_en = field('major-en')
    title_map = {'教授': 'Professor', '副教授': 'Associate Professor', '研究员': 'Researcher', '讲师': 'Lecturer'}
    supervisor_items = info.get('supervisor')
    supervisor_title = title_map.get(supervisor_items[1], '') if isinstance(supervisor_items, list) and len(supervisor_items) > 1 else ''
    date = info.get('submit-date', {})
    if blank:
        date_zh, date_en = '__________', '__________'
    elif isinstance(date, dict) and isinstance(date.get('year'), int) and isinstance(date.get('month'), int):
        date_zh = f"{date['year']}年{date['month']}月"
        date_en = f"{date['month']:02d}/{date['year']}"
    else:
        missing.append('submit-date')
        date_zh, date_en = '【待填写】', '【待填写】'
    track = data.get('track', 'academic')
    if track not in ('academic', 'professional'):
        raise ValueError('track must be academic or professional')
    major_label = '专业领域' if track == 'professional' else '学科专业'
    for name, latin, chinese, size, bold in (
        ('NWPU Cover Title', '黑体', '黑体', 22, True),
        ('NWPU Cover Detail', 'Times New Roman', '宋体', 16, True),
        ('NWPU Cover Plain', 'Times New Roman', '宋体', 16, False),
        ('NWPU Cover English', 'Times New Roman', '宋体', 15, False),
        ('NWPU Front Heading', '黑体', '黑体', 16, False),
        ('NWPU Front English Heading', 'Times New Roman', 'Times New Roman', 16, True),
        ('NWPU TOC Heading', '黑体', '黑体', 16, False),
    ):
        add_style(document, name, latin, chinese, size, bold)
    set_outline_level(document.styles['NWPU Front Heading'], 0)
    set_outline_level(document.styles['NWPU Front English Heading'], 0)
    section = document.sections[0]
    for item in (section.header, section.even_page_header, section.footer, section.even_page_footer):
        item.is_linked_to_previous = False
        item.paragraphs[0].clear()
    table = document.add_table(rows=4, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    table.autofit = False
    for column in table.columns:
        column.width = Cm(2.6)
    for row, (label, value) in zip(table.rows, [('学校代码', '10699'), ('分类号', class_no), ('密级', '__________' if blank else '公开'), ('学号', student_id)]):
        for cell, val in zip(row.cells, (label, value)):
            cell.width = Cm(2.6)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            run = p.add_run(val)
            set_run_font(run, latin='Times New Roman', east_asia='黑体', size=10.5)
            run.bold = True
        # The cover's small metadata table follows the bordered guide specimen.
    table.style = 'Table Grid'
    add_blank(document.add_paragraph(), after=110)
    add_blank(document.add_paragraph('题目  ' + title, style='NWPU Cover Title'), after=35)
    add_blank(document.add_paragraph('作者  ' + author, style='NWPU Cover Detail'), after=24)
    for label, value in ((major_label, major), ('指导教师', supervisor), ('培养单位', department), ('申请日期', date_zh)):
        add_blank(document.add_paragraph(f'{label}  {value}', style='NWPU Cover Detail'), after=10)
    document.add_page_break()
    add_blank(document.add_paragraph('西 北 工 业 大 学', style='NWPU Cover Plain'), before=45, after=12)
    add_blank(document.add_paragraph('博 士 学 位 论 文' if degree == 'doctor' else '硕 士 学 位 论 文', style='NWPU Cover Title'), after=100)
    add_blank(document.add_paragraph('题目： ' + title, style='NWPU Cover Plain'), after=65)
    for label, value in ((major_label, major), ('作    者', author), ('指导教师', supervisor), ('日    期', date_zh)):
        add_blank(document.add_paragraph(f'{label}： {value}', style='NWPU Cover Plain'), after=12)
    document.add_page_break()
    add_blank(document.add_paragraph('Title: ' + title_en, style='NWPU Cover English'), before=55, after=50)
    for text, after in (('By', 4), (author_en, 12), (f'Under the Supervision of {supervisor_title}'.strip(), 4), (supervisor_en, 55),
                        ('A Dissertation Submitted to', 4), ('Northwestern Polytechnical University', 40),
                        ('In Partial Fulfillment of The Requirement For The Degree of', 4),
                        (f"{'Doctor' if degree == 'doctor' else 'Master'} of {major_en}", 45), ("Xi’an, P. R. China", 4), (date_en, 0)):
        add_blank(document.add_paragraph(text, style='NWPU Cover English'), after=after)


def add_abstract(document, data: dict, key: str, missing: list[str], english: bool, renderer, log, blank: bool):
    abstract = data.get(key, {})
    if not isinstance(abstract, dict) or 'unsupported_expression' in abstract:
        raise ValueError(f'{key} must be a static dictionary')
    title = 'Abstract' if english else '摘  要'
    style = 'NWPU Front English Heading' if english else 'NWPU Front Heading'
    add_blank(document.add_paragraph(title, style=style), after=18)
    content = '' if blank else literal(abstract.get('content'), key + '.content', missing)
    for text in re.split(r'\n\s*\n', content):
        text = re.sub(r'\s*\n\s*', ' ', text).strip()
        if text:
            p = document.add_paragraph(style='Normal')
            add_inline_content(p, text, {}, renderer, log)
            if english:
                p.paragraph_format.first_line_indent = Pt(24)
    if blank:
        for _ in range(3):
            document.add_paragraph('　', style='Normal')
    keywords = [] if blank else abstract.get('keywords')
    if not isinstance(keywords, list) or not all(isinstance(item, str) for item in keywords):
        missing.append(key + '.keywords')
        keywords = ['【待填写】']
    p = document.add_paragraph(style='Normal')
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(12)
    label = p.add_run('Key words: ' if english else '关键词：')
    label.bold = True
    p.add_run(('; ' if english else '；').join(keywords))


def add_front_matter(document, data: dict, even_header: str, renderer, log, blank: bool = True) -> list[str]:
    missing = []
    add_cover(document, data, missing, blank)
    section = document.add_section(WD_SECTION_START.ODD_PAGE)
    configure_front_section(section, '摘要', even_header, True)
    add_abstract(document, data, 'abstract', missing, False, renderer, log, blank)
    section = document.add_section(WD_SECTION_START.NEW_PAGE)
    configure_front_section(section, 'Abstract', even_header, False)
    add_abstract(document, data, 'abstract-en', missing, True, renderer, log, blank)
    section = document.add_section(WD_SECTION_START.NEW_PAGE)
    configure_front_section(section, '目录', even_header, False)
    add_blank(document.add_paragraph('目  录', style='NWPU TOC Heading'), after=18)
    p = document.add_paragraph(style='Normal')
    p.paragraph_format.first_line_indent = Pt(0)
    add_complex_field(p, ' TOC \\o "1-3" \\h \\z \\u ', '在 Word 中更新目录')
    return missing
