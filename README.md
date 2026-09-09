# TypstDocx Bridge

将 Typst 论文项目中的语义内容桥接到可继续编辑、可统一套用 Word 样式的
`.docx` 文档。

Codex skill 标识为 `$export-typst-to-docx`。项目关注的是内容交接，而不是把
Typst/PDF 的页面外观逐像素复刻到 Word。

## 能做什么

- 将 Typst 的一至三级标题转换为真正的 Word `Heading 1/2/3`。
- 将正文和题注写入 `Normal`、`Caption` 等命名样式，便于在 Word 中统一修改。
- 将 PNG、JPEG 和 SVG 图片嵌入 Word；SVG 会通过 `sharp` 转为 PNG。
- 将行内和独立公式转换为可编辑的 Word OMML 公式，而不是公式截图。
- 将管道式表格转换为可编辑的 Word 三线表。
- 使用 Word `SEQ` 字段自动编号图、表和公式，并使用 `REF` 字段生成交叉引用。
- 转换前把所需文件暂存到源库之外，转换后用哈希清单确认源文件未被修改。

当前版本一次转换一个章节，并要求语义内容文件恰好包含一个一级标题。

## 从哪里转换

本项目最初用于把
[modern-npu-thesis](https://github.com/1195343015/modern-npu-thesis)
工作流编写的 Typst 学位论文内容转换为 Word，典型输入是项目中的
`content.typ` 或其他实际包含标题、正文、公式、图表和引用的 `.typ` 文件。

转换路线是：

```text
Typst 源项目
  -> 只读依赖暂存
  -> 受约束的语义内容解析
  -> Tinymist HTML/MathML
  -> Office OMML + Word 命名样式/字段
  -> 可编辑 DOCX
```

这不是 PDF 转 Word，也不是 `modern-npu-thesis` 的分支或官方组件。仓库不包含
该模板的源码、字体或任何用户论文内容。内置的 NWPU Word 内容模板是为语义
导出独立生成的样式载体，不是学校发布的官方 Word 模板。其他 Typst 项目只要
使用当前支持的语义结构也可以转换。

## 当前支持范围

- `=`, `==`, `===` 标题；
- 普通正文和 `$...$` 行内公式；
- 独立公式及其标签；
- `capfig(...)` 图片；
- `captab(...)` 包裹的管道式表格；
- `@fig:...`、`@tab:...`、`@eqt:...` 交叉引用；
- 其他 `@CitationKey` 以 `[CitationKey]` 文本保留，供后续文献工具处理。

复杂宏、子图、跨行/跨列表格、脚注、代码块、自动展开的多文件正文和完整参考
文献渲染目前需要人工检查或扩展解析器。详细边界见
[`references/compatibility.md`](references/compatibility.md)。

## 环境要求

当前流程在 Windows 11 上验证。需要：

- Python 3.10 或更高版本；
- Python 包 `python-docx` 和 `lxml`；
- Node.js 和 `sharp`，仅在处理 SVG 时使用；
- Tinymist CLI；可直接使用 VS Code Tinymist 扩展目录中的 `tinymist.exe`；
- Microsoft Office 安装目录中的 `MML2OMML.XSL`；
- Microsoft Word，推荐用于更新字段和最终目视检查。

仓库不会再分发 Tinymist、Microsoft Office 文件或第三方字体，使用者应自行安装
并遵守它们各自的许可证。

## 安装为 Codex skill

可以直接把仓库克隆到 Codex skills 目录：

```powershell
$SkillTarget = Join-Path $env:USERPROFILE '.codex\skills\export-typst-to-docx'
git clone https://github.com/goth42378/TypstDocx-Bridge.git $SkillTarget
```

重启或重新加载 Codex 后，可这样请求：

```text
使用 $export-typst-to-docx，把这个 Typst 章节转换为可编辑 Word。
源项目必须保持只读，输出和缓存放到源库之外。
```

也可以不安装 skill，直接运行仓库中的脚本。

## 安装脚本依赖

在 `export-typst-to-docx` 目录中执行：

```powershell
py -3 -m venv .venv
& '.\.venv\Scripts\python.exe' -m pip install -r requirements.txt
npm install
```

## 完整转换示例

以下示例使用占位路径。`$SourceRoot` 是只读 Typst 项目，`$WorkRoot` 必须位于
源项目之外，且 `$StageRoot` 在执行暂存命令前不能已经存在。

```powershell
$SkillRoot = (Resolve-Path '.').Path
$Python = (Resolve-Path '.\.venv\Scripts\python.exe').Path
$SourceRoot = 'D:\path\to\typst-project'
$WorkRoot = 'D:\path\to\conversion-work\chapter-02'
$StageRoot = Join-Path $WorkRoot 'staged-project'
$OutputRoot = Join-Path $WorkRoot 'output'

$Entry = 'writing\drafts\chapter_02\chapter.typ'
$ContentRelative = 'writing\drafts\chapter_02\content.typ'
$Tinymist = 'C:\path\to\tinymist.exe'
$Mml2Omml = 'C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL'
$Node = (Get-Command node).Source
$env:NODE_PATH = (Resolve-Path '.\node_modules').Path

& $Python 'scripts\stage_project.py' `
  --source-root $SourceRoot `
  --entry $Entry `
  --destination $StageRoot

$Content = Join-Path $StageRoot $ContentRelative

& $Python 'scripts\preflight_typst.py' $Content `
  --strict `
  --json (Join-Path $OutputRoot 'typst-preflight.json')

& $Python 'scripts\build_docx.py' `
  --content $Content `
  --template (Join-Path $SkillRoot 'assets\nwpu-2025-content-template.docx') `
  --output (Join-Path $OutputRoot 'chapter-content.docx') `
  --report (Join-Path $OutputRoot 'conversion-report.json') `
  --cache (Join-Path $OutputRoot 'cache') `
  --tinymist $Tinymist `
  --mml2omml $Mml2Omml `
  --node $Node `
  --chapter-number 2

& $Python 'scripts\audit_docx.py' `
  (Join-Path $OutputRoot 'chapter-content.docx') `
  --json (Join-Path $OutputRoot 'docx-audit.json')

& $Python 'scripts\verify_source_manifest.py' `
  (Join-Path $StageRoot 'export-manifest.json') `
  --json (Join-Path $OutputRoot 'source-verification.json')
```

转换后检查：

1. `conversion-report.json` 中没有公式、图片或引用失败；
2. `docx-audit.json` 中存在预期的 OMML、`SEQ`、`REF`、图片和表格数量；
3. `source-verification.json` 中源文件和暂存副本均为 `passed`；
4. 在 Word 中按 `Ctrl+A`，再按 `F9` 更新全部编号和引用；
5. 目视检查公式、图片尺寸、表格分页和标题层级。

如果章节正文并不在入口文件中，应把暂存副本中真正承载内容的文件传给
`--content`，不要把只有模板导入和 `include` 的包装文件直接交给转换器。

## 使用自己的 Word 模板

把 `--template` 指向一个不含示例正文的空白 `.docx`。模板至少应定义 Word
内置的 `Heading 1/2/3`、`Normal`、`Caption` 和 `Footer` 样式。转换器会向模板
追加内容，因此不要直接传入已经含有论文正文的文件。

内置样式的说明见
[`references/nwpu-2025-word-style-profile.md`](references/nwpu-2025-word-style-profile.md)。
后续修改字体、字号、段距或编号时，应修改 Word 样式定义，而不是逐段调整。

## 合成测试样例

[`examples/synthetic-project/content.typ`](examples/synthetic-project/content.typ)
是完全虚构的回归测试材料，覆盖标题、图片、三线表、行内/独立公式以及图表公式
交叉引用，不包含真实论文内容。可以把它直接传给 `preflight_typst.py` 和
`build_docx.py` 做环境冒烟测试。

## 隐私与源文件安全

- 真实论文、图片、参考文献、转换缓存和输出文件必须放在 skill 目录之外；
- 不在源项目中编译 Typst、创建缓存或生成 DOCX；
- 暂存前后分别记录文件哈希，转换完成后再次验证；
- `.gitignore` 会阻止常见项目副本、DOCX/PDF、报告和缓存被误提交，但它不能
  代替发布前的人工检查。

## 与其他转换方式的区别

Typst 本身目前不直接输出 DOCX。Pandoc 适合较通用的文档转换；PDF 转 Word 更
接近页面外观。TypstDocx Bridge 选择另一条路线：保留标题、图表、公式、编号和
引用的语义，并让 Word 模板负责最终格式。它尤其适合“先取得可编辑内容，再在
Word 中按单位模板统一排版”的论文交付流程。

## 许可证

本项目采用 [MIT License](LICENSE)。输入 Typst 项目、Microsoft Office、Tinymist、
Node/Python 依赖和用户选择的字体仍受各自许可证约束。
