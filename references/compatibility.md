# Compatibility notes

## Full thesis and chapter boundary

- `export_thesis.py` expands static `#include "..."` directives in place, including recursively nested files and mixed chapter body/includes. Each included figure keeps its own source file's relative path. It builds one Word document and resolves supported figure/table/equation references across chapters. It rejects dynamic includes, include cycles, duplicate labels, unresolved targets and unsupported top-level commands.
- The assembly wrapper may contain `#import`, `#counter` and `#show: nwpu-thesis.with(...)`. Other `#show` rules and assembly commands stop for review.
- Full-thesis output creates blank outer/inner title pages, Chinese/English abstract slots, and an updating TOC by default. `--populate-front-matter` is opt-in for static wrapper metadata or `--front-matter-json`. Metadata is not automatically included in blank mode. A colored-cover background is unsupported. The reference list is not rendered; citation keys stay as text.
- `build_docx.py` remains a single-chapter converter and its strict preflight expects one level-1 heading.
- The full-thesis entry point runs its own aggregate preflight after expansion; use `scripts/preflight_typst.py --strict` for the single-chapter path. Unknown top-level commands are review items because the constrained parser otherwise has no safe way to decide whether they contain content or only layout.
- Use only a staged copy of the Typst project. `scripts/verify_source_manifest.py` must pass after conversion.
- Use a blank content-only DOCX as the template. Existing body content in a template is retained and the converted chapter is appended after it.

## Supported in the current prototype

- Typst headings beginning with `=`, `==`, or `===`.
- Plain paragraphs with inline `$...$` math.
- Display equations delimited by `$ ... $`, optionally followed by a label.
- `capfig(...)` with a literal `image("...")`, percentage width, caption, and label.
- `captab(...)` whose body is a Markdown-like pipe table.
- `@fig:...`, `@tab:...`, and `@eqt:...` cross-references.
- Other `@Key` citations preserved as bracketed citation keys.
- PNG, JPEG, SVG, and single-page PDF figure sources (PDF needs Poppler `pdfinfo` and `pdftoppm` on PATH).
- Single-line top-level `#strong[...]` paragraphs without nested brackets and trailing paragraph labels.
- Top-level Typst `+` numbered lists (one level), rendered as editable Word numbered lists; nested lists and other list types still require review.
- `#block(breakable: false)[...]` content is expanded in place, including prose, figures, tables and editable formulas. The Typst page-break constraint itself is not reproduced.
- The explicitly recognized `formula-note(body)` wrapper defined as `block(width: 100%)` containing `#set par(first-line-indent: (amount: 0pt, all: true))` and `式中，#body` becomes a named Word formula-note style, with its text and inline math preserved. Different definitions are not guessed.

## Conversion semantics

- Word paragraphs and tables remain editable.
- Equations are editable Word OMML objects produced from Tinymist HTML MathML through Office's `MML2OMML.XSL`.
- Formula bodies use Cambria Math for Word's native mathematical layout; equation numbers and ordinary Western text use Times New Roman.
- Figure/table/equation numbering uses Word `SEQ` fields for the sequence component. Full-thesis output uses `STYLEREF` for the chapter prefix and `REF` fields for references; the single-chapter path retains a fixed `--chapter-number` prefix. Sequence fields reset at Heading 1.
- The chapter list instance has a start value fixed at export; after adding, deleting or reordering chapters, regenerate the output or correct each affected Word list start before updating fields. Field refresh does not independently renumber the chapter list instances.
- The complete visible `chapter-sequence` number is bookmarked, and Typst cross-references become clickable Word `REF` fields.
- Typst layout instructions are intentionally not reproduced. Word page geometry and named styles come from the selected DOCX template.
- The bundled NWPU 2025 profile uses real Heading 1/2/3, Normal, Caption, NWPU Figure, NWPU Equation, NWPU Table Text/Header, and NWPU Three-Line Table styles. Chapter/section numbering is a real Word multilevel list.

## Inspect before use

Adapt `scripts/build_docx.py` when the project relies on:

- custom figure or table macros other than `capfig` and `captab`;
- dynamically generated paths or content;
- `capsubfig`, nested subfigures, grids, or complex table cell spanning;
- bibliography rendering instead of citation-key preservation;
- equations that use Typst constructs not represented by the current Tinymist HTML MathML exporter;
- dynamic includes; `include` embedded inside another content wrapper; or wrappers other than the explicitly recognized variants above.

Also stop for review rather than claiming a complete conversion when the semantic body contains unsupported list types, footnotes, code blocks, block quotes, raw HTML, bibliography output, merged or multi-line table cells, nested brackets in macro captions, or other constructs not listed as supported above.

For a single-chapter invocation, identify the included file that contains the semantic body and pass that file with `--content`. Use the full-thesis entry point when chapter body and includes are interleaved.
