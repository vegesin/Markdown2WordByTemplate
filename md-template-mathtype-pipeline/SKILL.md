---
name: md-template-mathtype-pipeline
description: Orchestrate Markdown-to-Word conversion and selective MathType conversion into one reusable Windows workflow. Use when the user wants to input a Markdown file and optional Word template and receive a final DOCX with template styling, three-line tables, Word Caption-style figure/table captions, Visio placeholders, and editable MathType Equation.DSMT4 formula objects.
---

# Markdown template MathType pipeline

Use `scripts/convert_md_to_mathtype_docx.ps1` to run the two-stage pipeline:

1. Run `md-to-template-docx` to create a stage-1 DOCX. Preserve LaTeX as `$...$` and `\[...\]`, format tables, extract Mermaid sources, and create Word Caption-style figure/table captions.
2. Run `word-mathtype-selective` to convert residual LaTeX into editable MathType `Equation.DSMT4` OLE objects.

Do not merge the underlying scripts into this skill. Treat this skill as an orchestrator so fixes in either base skill remain reusable.

## Preconditions

- Require the `md-to-template-docx` and `word-mathtype-selective` skills to be installed, or pass their directories explicitly.
- Require Pandoc and the bundled Python runtime for stage 1.
- Require Windows, Microsoft Word, and licensed MathType 6 for stage 2.
- Open Microsoft Word manually, keep one blank document open, and verify MathType Toggle TeX works manually before running the full pipeline.

## Run

With a user template:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File `
  scripts/convert_md_to_mathtype_docx.ps1 `
  -MarkdownPath "C:\path\chapter.md" `
  -TemplatePath "C:\path\template.docx" `
  -OutputPath "C:\path\final.docx" `
  -Force
```

Without a user template, omit `-TemplatePath`; the stage-1 converter uses its bundled default template.

The script creates a stage-1 intermediate DOCX beside the final output unless `-Stage1Path` is passed. Keep this file for debugging if MathType conversion fails.

## Failure handling

- If stage 1 fails, inspect the Pandoc/Python error and the Markdown source.
- If stage 2 fails, inspect the clear MathType error block printed by `word-mathtype-selective`.
- If the environment does not allow the agent to control Word, print the exact PowerShell command for the user to run manually.
