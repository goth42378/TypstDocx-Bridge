---
name: export-typst-to-docx
description: Extract content from a Typst (.typ) project into an editable, style-driven Word (.docx) document while preserving real Heading/Normal/Caption styles, native three-line tables, figures, editable OMML equations, automatic heading and SEQ numbering, and REF cross-references. Use when a user wants a content-first Typst-to-Word handoff, including NWPU 2025 thesis formatting, while the source repository remains read-only.
---

# TypstDocx Bridge

`export-typst-to-docx` is the stable Codex skill identifier. TypstDocx Bridge
is the public project and UI name.

Create a content-first Word document from one Typst chapter. Preserve semantic content and references; use named Word styles so later template changes propagate globally rather than formatting paragraphs one by one. The current v1 contract is one chapter per conversion. Assemble multiple converted chapters only after each chapter passes its own checks.

## Safety rule

Treat the supplied Typst repository as read-only.

1. Run `scripts/stage_project.py` first.
2. Write only to a new destination outside the source root.
3. Run conversion, formula rendering, caches, reports, and document rendering only against the staged copy.
4. Verify the staged manifest against the original source after conversion.

Do not compile or create caches inside the source repository.

Keep user material outside this skill directory. Never copy a user's staged
project, figures, bibliography, generated DOCX/PDF, reports, or conversion
cache into `examples/`, `assets/`, or another packaged folder. The bundled
example is synthetic and may be used only as a smoke-test fixture.

## Workflow

### 1. Select and stage a project

Choose an entry `.typ` file that represents a coherent document or chapter. Stage its static dependency closure:

```powershell
python scripts/stage_project.py `
  --source-root '<source-root>' `
  --entry 'writing/chapter/chapter.typ' `
  --destination '<work-root>/staged-project'
```

The destination must not already exist. Read `export-manifest.json` and confirm `source_mutation_check` is `passed`. Missing dependencies mentioned only in comments may be documented and ignored after inspection. Do not reuse the source repository as `--destination`, `--cache`, `--output`, or `--report`.

### 2. Pick the semantic content file

Pass the staged file containing the actual headings, paragraphs, `capfig`, `captab`, and equations to the converter. In modular projects this is often `content.typ`, not the wrapper file that only imports a template and includes content. Never pass the original repository file to `build_docx.py`.

Run the strict preflight before building:

```powershell
python scripts/preflight_typst.py '<work-root>/staged-project/writing/chapter/content.typ' `
  --strict `
  --json '<work-root>/output/typst-preflight.json'
```

Stop and inspect if the preflight reports more or fewer than one level-1 heading, duplicate labels, unresolved figure/table/equation references, or unsupported top-level Typst commands. Do not silently discard a command that may generate content.

Consult `references/compatibility.md` before converting macro-heavy projects.

### 3. Build the DOCX

Use the bundled Python and Node runtimes when available. Locate `tinymist.exe` from the user's VS Code Tinymist extension and `MML2OMML.XSL` from the Microsoft Office installation. Set `NODE_PATH` to a Node modules directory containing `sharp` for SVG rasterization.

```powershell
$env:NODE_PATH = '<node-modules-path>'
python scripts/build_docx.py `
  --content '<work-root>/staged-project/writing/chapter/content.typ' `
  --template 'assets\nwpu-2025-content-template.docx' `
  --output '<work-root>/output/chapter-content.docx' `
  --report '<work-root>/output/conversion-report.json' `
  --cache '<work-root>/output/cache' `
  --tinymist '<tinymist-executable>' `
  --mml2omml '<office-MML2OMML.XSL>' `
  --node '<node-executable>' `
  --chapter-number 2
```

Add `--create-template` only when `assets/nwpu-2025-content-template.docx` does not exist or should be rebuilt. Read `references/nwpu-2025-word-style-profile.md` before rebuilding or deliberately changing the bundled profile.

The bundled NWPU profile maps Typst `=`, `==`, and `===` to real Word `Heading 1`, `Heading 2`, and `Heading 3`. A Word multilevel list generates `第 N 章`, `N.N`, and `N.N.N` automatically, starting from `--chapter-number`. Body text uses `Normal`; captions use `Caption`; figure, equation, and table-cell roles use named NWPU styles.

To use a different Word template later, pass it with `--template`. It must be a blank, content-only DOCX whose section settings and named styles are intended for the exported chapter; do not pass a live thesis document containing sample text or existing chapters because the converter appends content. Ensure it defines the built-in Heading 1/2/3, Normal, Caption, and Footer styles. The converter supplies its auxiliary figure/equation/table roles when the bundled profile is used.

### 4. Check conversion quality

Run the bundled semantic audit before visual QA:

```powershell
python scripts/audit_docx.py '<work-root>/output/chapter-content.docx' `
  --json '<work-root>/output/docx-audit.json'
```

Require all of the following before delivery:

- `unresolved_refs` is empty.
- `formula_failures` is empty.
- `image_failures` is empty.
- The expected figure, table, equation, `SEQ`, and `REF` field counts are present.
- Actual section headings use `Heading 1/2/3`, body paragraphs use `Normal`, captions use `Caption`, and table cells use named table paragraph styles.
- For the bundled NWPU profile, Word's effective Heading 1/2/3 fonts are concrete 黑体 for Chinese, Latin text, and automatic numbering; no heading style retains theme-font attributes.
- Editable formula bodies use Cambria Math. Equation numbers and surrounding Western text use Times New Roman. Do not force the whole OMML equation to Times New Roman unless an explicit institutional rule overrides this policy and complex operators have been rechecked in Word.
- Every page renders without clipped formulas, missing images, overflowing tables, or unreadable content.

Open the DOCX in Word and update fields with `Ctrl+A`, then `F9` after inserting, deleting, or moving figures, tables, or equations. Cached `SEQ` and `REF` results keep the document readable before updating.

### 5. Recheck the source

Run the deterministic post-conversion verification:

```powershell
python scripts/verify_source_manifest.py '<work-root>/staged-project/export-manifest.json' `
  --json '<work-root>/output/source-verification.json'
```

Require both `source.status` and `staged_copy.status` to be `passed`. Report the number of unchanged source files. Any mismatch is a delivery blocker until explained and resolved without writing to the source repository.

## Output behavior

- Headings and paragraphs become editable Word text.
- Headings use real Word outline styles and automatic multilevel numbering.
- Pipe-style Typst tables become native three-line Word tables with repeated header rows.
- Raster images are embedded; SVG files are rendered to PNG through `sharp`.
- Inline and display math are exported by Tinymist as MathML, transformed with Office's `MML2OMML.XSL`, and embedded as editable Word OMML equations.
- Figure, table, and equation serial numbers become Word `SEQ Figure`, `SEQ Table`, and `SEQ Equation` fields that reset at Heading 1. The chapter prefix comes from `--chapter-number`.
- Labels bookmark the complete visible `chapter-sequence` number; references become clickable `REF` fields and update with the sequence fields.
- Bibliography citations are preserved as `[CitationKey]` text for later citation-manager or template work.

`examples/synthetic-project/` is an invented regression fixture covering
headings, a heading label, inline and display math, a figure, a three-line
table, and Word cross-reference targets. It contains no user project material
and is not required at runtime.

Keep content, numbering, and role assignment semantic. Change appearance through Word style definitions or a replacement template, not repeated paragraph-level formatting.
