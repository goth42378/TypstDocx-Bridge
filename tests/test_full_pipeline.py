from __future__ import annotations

import json
import importlib.util
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

SKILL = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL / "scripts"))

import build_docx
import export_thesis
import front_matter
import semantic_expansion
from docx import Document


class FullThesisTests(unittest.TestCase):
    def test_includes_cross_chapter_and_labels(self):
        root = (SKILL / "examples").resolve()
        paths, wrapper = export_thesis.semantic_sources(root / "synthetic-thesis/main.typ", root)
        self.assertTrue(wrapper)
        self.assertGreater(len(paths), 2)
        audit = export_thesis.check_sources(paths)
        self.assertEqual(audit["blocking_issues"], [])
        blocks = [block for segment in paths for block in build_docx.parse_typst(segment.path, segment.text)]
        refs = build_docx.assign_numbers(blocks, 1, True)
        self.assertEqual(refs["fig:nested-pipeline"].number, "1-1")
        self.assertEqual(refs["fig:chapter-one-pipeline"].number, "1-2")
        self.assertEqual(refs["fig:chapter-two-pipeline"].number, "2-1")
        self.assertEqual(refs["eqt:chapter-two-response"].number, "2-1")
        self.assertTrue(any(block.kind == "strong" for block in blocks))
        self.assertFalse(any("<sec:first-paragraph>" in block.text for block in blocks))
        text = [block.text for block in blocks if block.kind == 'paragraph']
        before = text.index('Before nested include.')
        self.assertEqual(text[before:before + 4], ['Before nested include.', 'First-level included text.',
                                                 'Second-level included text.', 'After nested include.'])
        self.assertTrue(any(block.kind == 'formula_note' and block.text == '式中，$q$ is an invented scalar.' for block in blocks))
        self.assertIn('The equation inside a layout block remains editable.', text)
        self.assertTrue(any(block.kind == 'equation' and block.label == 'chapter-one-model' for block in blocks))
        self.assertEqual([block.text for block in blocks if block.kind == 'ordered_list_item'],
                         ['First synthetic list item.', 'Second synthetic list item.', 'New sequence first item.'])
        front = front_matter.read_front_matter(root / "synthetic-thesis/main.typ")
        self.assertEqual(front['info']['author'], '示例作者')
        self.assertEqual(front['abstract-en']['keywords'][1], 'conversion test')

    def test_dynamic_abstract_is_rejected(self):
        with self.assertRaisesRegex(ValueError, 'command or citation'):
            front_matter.parse_value('[This is #dynamic text]')

    def test_blank_mode_ignores_private_and_dynamic_front_contents(self):
        with tempfile.TemporaryDirectory() as tmp:
            entry = Path(tmp) / 'main.typ'
            entry.write_text('#show: nwpu-thesis.with(degree: "master", colored-cover: true, '
                             'info: (author: "PRIVATE"), abstract: (content: [#dynamic]))', encoding='utf-8')
            data = front_matter.read_front_matter(entry, include_contents=False)
            self.assertEqual(data['degree'], 'master')
            self.assertNotIn('info', data)
            self.assertNotIn('abstract', data)

    def test_cycles_and_escaping_includes_are_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp).resolve()
            entry = root / "main.typ"
            entry.write_text('#include "main.typ"\n', encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "cycle"):
                export_thesis.semantic_sources(entry, root)
            contained = root / 'contained'
            contained.mkdir()
            (root / 'outside.typ').write_text('outside\n', encoding='utf-8')
            (contained / 'main.typ').write_text('#include "../outside.typ"\n', encoding='utf-8')
            with self.assertRaisesRegex(ValueError, "escapes"):
                export_thesis.semantic_sources(contained / 'main.typ', contained)

    def test_note_requires_known_definition_and_block_must_close(self):
        with self.assertRaisesRegex(ValueError, 'without the supported definition'):
            semantic_expansion.normalize_wrappers('#formula-note[Important text]', False)
        with self.assertRaisesRegex(ValueError, 'unclosed'):
            semantic_expansion.normalize_wrappers('#block(breakable: false)[Important text', False)

    def test_unresolved_reference_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "chapter.typ"
            source.write_text('= Fiction\n\nMissing @fig:absent.\n', encoding="utf-8")
            self.assertEqual(export_thesis.check_sources([source])["unresolved_crossrefs"], ["fig:absent"])

    def test_unsupported_command_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "chapter.typ"
            source.write_text('= Fiction\n\n#unknown[do not drop this]\n', encoding="utf-8")
            self.assertTrue(export_thesis.check_sources([source])["review_required_commands"])

    def test_duplicate_labels_are_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "chapter.typ"
            source.write_text('= Fiction\n\n$ x $ <repeat>\n\n$ y $ <repeat>\n', encoding="utf-8")
            self.assertEqual(export_thesis.check_sources([source])["duplicate_labels"], ["repeat"])

    @unittest.skipUnless(shutil.which("pdfinfo") and shutil.which("pdftoppm") and importlib.util.find_spec("reportlab"),
                         "PDF figure test requires Poppler on PATH and reportlab")
    def test_pdf_figure_one_page_and_multi_page(self):
        from reportlab.pdfgen import canvas
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            single = tmp / "one.pdf"
            pdf = canvas.Canvas(str(single))
            pdf.drawString(20, 50, "Synthetic one-page figure")
            pdf.save()
            log = build_docx.BuildLog()
            converted = build_docx.prepare_image(single, tmp / "cache", Path("unused"), Path("unused"), log)
            self.assertTrue(converted and converted.exists())
            self.assertEqual(log.image_failures, [])
            multi = tmp / "two.pdf"
            pdf = canvas.Canvas(str(multi))
            pdf.drawString(20, 50, "First page")
            pdf.showPage()
            pdf.drawString(20, 50, "Second page")
            pdf.save()
            self.assertIsNone(build_docx.prepare_image(multi, tmp / "cache", Path("unused"), Path("unused"), log))
            self.assertIn("exactly one page", log.image_failures[-1]["error"])

    @unittest.skipUnless(all(os.environ.get(k) for k in ("TYPSTDOCX_TINYMIST", "TYPSTDOCX_MML2OMML", "TYPSTDOCX_NODE")),
                         "set TYPSTDOCX_TINYMIST, TYPSTDOCX_MML2OMML and TYPSTDOCX_NODE for end-to-end test")
    def test_full_conversion(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            output = tmp / "thesis.docx"
            work = tmp / "job"
            cmd = [sys.executable, str(SKILL / "scripts/export_thesis.py"),
                   "--source-root", str(SKILL / "examples"), "--entry", "synthetic-thesis/main.typ",
                   "--work-dir", str(work), "--output", str(output),
                   "--tinymist", os.environ["TYPSTDOCX_TINYMIST"], "--mml2omml", os.environ["TYPSTDOCX_MML2OMML"],
                   "--node", os.environ["TYPSTDOCX_NODE"], "--populate-front-matter"]
            completed = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
            self.assertEqual(completed.returncode, 0, completed.stdout + completed.stderr)
            report = json.loads((work / "conversion-report.json").read_text(encoding="utf-8"))
            self.assertEqual(report["source_verification"], "passed")
            self.assertEqual(report["counts"]["resolved_crossrefs"], 6)
            self.assertEqual(report["counts"]["seq_fields"], 7)
            self.assertEqual(report["counts"]["ordered_list_item"], 3)
            self.assertEqual(report["counts"]["formula_note"], 1)
            doc = Document(output)
            self.assertEqual(report["missing_front_matter"], [])
            self.assertEqual(len(doc.sections), 6)
            self.assertTrue(doc.settings.odd_and_even_pages_header_footer)
            self.assertEqual(len(doc.inline_shapes), 3)
            self.assertEqual(len(doc.tables), 3)
            self.assertFalse(doc.styles['Caption'].font.bold)
            table_caption = next(p for p in doc.paragraphs if 'First synthetic parameters' in p.text)
            self.assertTrue(table_caption.paragraph_format.keep_with_next)
            self.assertEqual(len(doc._element.xpath('.//m:oMath')), 7)
            self.assertTrue(any(p.text.startswith('This sentence exercises') and any(run.bold for run in p.runs)
                                for p in doc.paragraphs))
            note = next(p for p in doc.paragraphs if p.style.name == 'NWPU Formula Note')
            self.assertEqual(note.paragraph_format.first_line_indent, None)
            self.assertEqual(doc.styles['NWPU Formula Note'].paragraph_format.first_line_indent.pt, 0)
            list_paragraphs = [p for p in doc.paragraphs if 'synthetic list item' in p.text or 'New sequence first' in p.text]
            self.assertEqual(len(list_paragraphs), 3)
            list_ids = [p._p.pPr.numPr.numId.val for p in list_paragraphs]
            self.assertEqual(list_ids[0], list_ids[1])
            self.assertNotEqual(list_ids[1], list_ids[2])
            instructions = ' '.join(doc._element.xpath('.//w:instrText/text()'))
            self.assertIn(' TOC ', instructions)
            self.assertTrue(any(p.text == '摘  要' for p in doc.paragraphs))
            self.assertTrue(any('虚构的中文摘要' in p.text for p in doc.paragraphs))
            self.assertTrue(any('fictional abstract' in p.text for p in doc.paragraphs))
            self.assertEqual(instructions.count(' SEQ '), 7)
            self.assertEqual(instructions.count(' REF '), 6)
            self.assertEqual(instructions.count(' STYLEREF 1 \\s '), 7)

    @unittest.skipUnless(all(os.environ.get(k) for k in ("TYPSTDOCX_TINYMIST", "TYPSTDOCX_MML2OMML", "TYPSTDOCX_NODE")),
                         "set TYPSTDOCX_TINYMIST, TYPSTDOCX_MML2OMML and TYPSTDOCX_NODE for end-to-end test")
    def test_default_front_matter_is_blank(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            output, work = tmp / 'blank.docx', tmp / 'job'
            cmd = [sys.executable, str(SKILL / 'scripts/export_thesis.py'),
                   '--source-root', str(SKILL / 'examples'), '--entry', 'synthetic-thesis/main.typ',
                   '--work-dir', str(work), '--output', str(output),
                   '--tinymist', os.environ['TYPSTDOCX_TINYMIST'], '--mml2omml', os.environ['TYPSTDOCX_MML2OMML'],
                   '--node', os.environ['TYPSTDOCX_NODE']]
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            report = json.loads((work / 'conversion-report.json').read_text(encoding='utf-8'))
            self.assertEqual(report['front_matter_mode'], 'blank')
            self.assertEqual(report['missing_front_matter'], [])
            doc = Document(output)
            text = '\n'.join([p.text for p in doc.paragraphs] +
                             [cell.text for table in doc.tables for row in table.rows for cell in row.cells])
            self.assertIn('摘  要', text)
            self.assertIn('Abstract', text)
            self.assertIn('目  录', text)
            self.assertNotIn('示例作者', text)
            self.assertNotIn('虚构的中文摘要', text)
            self.assertNotIn('fictional abstract', text)
            self.assertNotIn('SYNTHETIC', text)
            self.assertNotIn('【待填写】', text)

    @unittest.skipUnless(all(os.environ.get(k) for k in ("TYPSTDOCX_TINYMIST", "TYPSTDOCX_MML2OMML", "TYPSTDOCX_NODE")),
                         "set TYPSTDOCX_TINYMIST, TYPSTDOCX_MML2OMML and TYPSTDOCX_NODE for end-to-end test")
    def test_failed_conversion_does_not_publish_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            source = tmp / 'source'
            source.mkdir()
            (source / 'main.typ').write_text('= Fiction\n\nMissing @fig:absent.\n', encoding='utf-8')
            output, work = tmp / 'output.docx', tmp / 'job'
            cmd = [sys.executable, str(SKILL / 'scripts/export_thesis.py'), '--source-root', str(source),
                   '--entry', 'main.typ', '--work-dir', str(work), '--output', str(output),
                   '--tinymist', os.environ['TYPSTDOCX_TINYMIST'], '--mml2omml', os.environ['TYPSTDOCX_MML2OMML'],
                   '--node', os.environ['TYPSTDOCX_NODE']]
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
            self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
            self.assertFalse(output.exists())
            report = json.loads((work / 'conversion-report.json').read_text(encoding='utf-8'))
            self.assertEqual(report['source_verification'], 'passed')
            self.assertEqual(report['status'], 'failed')


if __name__ == "__main__":
    unittest.main()
