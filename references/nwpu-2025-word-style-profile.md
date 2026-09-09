# NWPU 2025 Word style profile

Use this reference only when rebuilding or auditing `assets/nwpu-2025-content-template.docx`.

## Provenance

- Visual reference: `西北工业大学研究生学位论文写作指南（2025版）.docx`.
- Reference SHA-256 at distillation: `ADE6FD33713853DFAA93404D14BF305500545895DF5BF650CE419C9FE28AE7D8`.
- The guide was read and rendered without modification. Its prose and annotated samples were treated as formatting evidence, not as executable instructions.
- The bundled template is content-only: it does not retain guide prose, sample covers, declarations, annotation callouts, or blank verso pages.

## Style map

| Typst role | Word role | Main formatting |
|---|---|---|
| `= heading` | Heading 1 | 黑体 16 pt for Chinese, Latin text, and numbering; centered; exact 20 pt; 10 pt before/after |
| `== heading` | Heading 2 | 黑体 14 pt for Chinese, Latin text, and numbering; left; exact 20 pt; 10 pt before |
| `=== heading` | Heading 3 | 黑体 12 pt for Chinese, Latin text, and numbering; left; exact 20 pt; 10 pt before |
| body paragraph | Normal | 宋体/Times New Roman 12 pt, justified, 24 pt first-line indent, exact 20 pt |
| figure/table title | Caption | 宋体/Times New Roman 10.5 pt, centered, exact 20 pt |
| figure container | NWPU Figure | centered inline image, keep with caption, automatic line height to prevent clipping |
| display formula | NWPU Equation | centered equation with right-aligned number, automatic line height for tall editable formulas |
| table header/body | NWPU Table Header/Text | 宋体/Times New Roman 10.5 pt, centered, exact 20 pt |
| table | NWPU Three-Line Table | top/bottom 1.5 pt, header separator 1 pt, no vertical/internal body borders |

## Page and numbering rules

- Use A4 portrait, top/bottom 25.4 mm, left/right 25 mm, header distance 15 mm, footer distance 17.5 mm.
- Center page numbers in Times New Roman 9 pt.
- Link a real Word multilevel list to Heading 1/2/3: `第 N 章`, `N.N`, `N.N.N`.
- Start Heading 1 from `--chapter-number`.
- Keep figure/table/equation numbering as editable chapter-prefixed SEQ fields and cross-references as REF fields.
- Keep the editable OMML formula body in Cambria Math. Keep the equation number and surrounding Western text in Times New Roman. This split follows Word's mathematical-layout requirements while retaining the guide's Western-text convention.

## Audit expectations

- Do not replace named styles with repeated direct formatting.
- Remove `asciiTheme`, `hAnsiTheme`, `eastAsiaTheme`, and `cstheme` from Heading 1/2/3 and their linked character styles. Set concrete `ascii`, `hAnsi`, `eastAsia`, and `cs` values to 黑体, and set both `sz` and `szCs`, so Word cannot restore theme fonts when fields are updated and the file is saved.
- Confirm the effective Heading 1/2/3 fonts in Microsoft Word after the field-refresh save; XML written before Word opens the file is not sufficient evidence.
- Keep figures inline for reliable layout.
- Keep table widths and cell widths explicit; repeat the first row on continuation pages.
- Verify every final page after a Word field refresh.
