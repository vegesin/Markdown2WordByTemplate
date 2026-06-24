---
name: md-to-template-docx
description: Convert Markdown into a Word DOCX using a supplied DOCX template or the bundled default template while preserving LaTeX as literal $...$ and \[...\] text for a later selective MathType conversion pass. Use when headings must retain their Markdown levels, Markdown tables must become three-line Word tables, Mermaid blocks must become stable Visio placeholders with extracted source files, figure/table captions must use Word Caption style and SEQ 图/表 fields, formula paragraphs must avoid fixed-line-height clipping, and Word-native OMML equations must not be created.
---

# Markdown to Template DOCX

Build the stage-1 DOCX for a multi-stage Word workflow. Preserve the supplied template's styles, page setup, headers, footers, theme, and numbering when a template is supplied. Use the bundled default template when no template is supplied. Leave equations as text that `word-mathtype-selective` can convert into `Equation.DSMT4` objects later.

## Run

Use the PowerShell wrapper on Windows. It validates all paths, selects the bundled Python runtime, and adds `.docx` when omitted.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File `
  scripts/run_md_to_template_docx.ps1 `
  -MarkdownPath "input.md" `
  -TemplatePath "template.docx" `
  -OutputPath "output_stage1.docx"
```

Alternatively, resolve the bundled workspace Python runtime and invoke the builder directly. Pass Pandoc explicitly when it is not on `PATH`.

```powershell
& $python scripts/build_stage1_docx.py `
  input.md `
  template.docx `
  output.docx `
  --pandoc "C:\path\to\pandoc.exe"
```

Without a user template, pass only Markdown and output:

```powershell
& $python scripts/build_stage1_docx.py `
  input.md `
  output.docx `
  --pandoc "C:\path\to\pandoc.exe"
```

The script also creates `<output-stem>_visio_sources/` with one `.mmd` file per Mermaid block and a `manifest.json` mapping each source to its Word placeholder.

Run the structural audit after generation. Expected counts are optional:

```powershell
& $python scripts/audit_stage1_docx.py output.docx
```

## Contract

- Map Markdown heading levels directly: `##` becomes `Heading 2`, `#####` becomes `Heading 5`.
- Preserve inline formulas as literal `$...$`.
- Flatten display formulas into one Word paragraph as literal `\[...\]` so a MathType converter can select only the formula range.
- Apply `MTDisplayEquation` when that style exists.
- Center each display placeholder, set its line spacing to at least 30 pt, and defer equation numbering until after MathType conversion.
- Set paragraphs containing inline formulas to at least 25 pt rather than exact 25 pt so tall MathType objects can expand the line.
- Never create `m:oMath`, `m:oMathPara`, or Word `OMath` objects.
- Convert Markdown tables to borderless three-line tables: top rule, header-bottom rule, and bottom rule only.
- Apply Word `Caption` style to figure and table captions. Use updateable `SEQ 图` and `SEQ 表` fields. Put figure captions after image/Visio placeholders and table captions before tables.
- Derive figure caption text from the adjacent `图1 ...` paragraph when present; otherwise use the Markdown image alt text stored in the DOCX drawing metadata. Derive table caption text from the immediately preceding `表1 ...` paragraph when present.
- Strip any old manual caption number from the derived text, including `图1 ...`, `表1 ...`, and Pandoc image-alt variants such as `1 ...`; the Word `SEQ` field supplies the live number.
- If no caption text exists, insert `图N 请填写图题` or `表N 请填写表题` as a Caption-style placeholder.
- If the template does not define `Caption`, create a default centered Caption paragraph style.
- Replace Mermaid blocks with `[VISIO:diagram-NN]`; do not rasterize diagrams.
- Keep the original Markdown and template unchanged.

## Verification

Treat the script's JSON summary as a minimum structural gate. Verify:

- source and DOCX formula counts match;
- `word/document.xml` contains no OMML nodes;
- no `SEQ Equation` fields exist before MathType conversion;
- every Mermaid block has one placeholder and one `.mmd` source;
- heading counts and levels match;
- every table uses three-line borders;
- no MathType OLE objects exist yet.

Render the DOCX and inspect all pages before delivery. If LibreOffice is unavailable, report that visual QA could not be completed and rely on structural checks.

Read [references/pipeline-contract.md](references/pipeline-contract.md) before handing the output to MathType or Visio automation.
