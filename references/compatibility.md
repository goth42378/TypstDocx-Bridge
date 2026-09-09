# Compatibility notes

## Version 1 boundary

- Convert one semantic chapter per invocation. Exactly one level-1 heading is required so chapter numbering and `SEQ ... \\s 1` resets stay coherent.
- Use `scripts/preflight_typst.py --strict` before conversion. Unknown top-level commands are review items because the constrained parser otherwise has no safe way to decide whether they contain content or only layout.
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
- PNG, JPEG, and SVG figure sources.

## Conversion semantics

- Word paragraphs and tables remain editable.
- Equations are editable Word OMML objects produced from Tinymist HTML MathML through Office's `MML2OMML.XSL`.
- Formula bodies use Cambria Math for Word's native mathematical layout; equation numbers and ordinary Western text use Times New Roman.
- Figure/table/equation numbering uses Word `SEQ` fields for the sequence component and the `--chapter-number` value for the chapter prefix. The sequence fields reset at Heading 1.
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
- content distributed across multiple includes that should be expanded automatically.

Also stop for review rather than claiming a complete conversion when the semantic body contains lists, footnotes, code blocks, block quotes, raw HTML, bibliography output, merged or multi-line table cells, nested brackets in macro captions, or other constructs not listed as supported above.

For modular projects, identify the included file that contains the semantic body and pass that file with `--content`.
