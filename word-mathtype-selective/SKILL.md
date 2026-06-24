---
name: word-mathtype-selective
description: Convert literal LaTeX formulas in Microsoft Word DOCX files into editable MathType Equation.DSMT4 OLE objects on Windows. Use for reusable Word formula conversion when inline formulas use $...$ and display formulas use \[...\], especially when whole-document Toggle TeX could alter Chinese text or document structure.
---

# Selective Word MathType conversion

Use `scripts/convert_word_latex_to_mathtype.ps1`. Do not substitute WPS or Word-native OMath.

## Preconditions

- Require Windows, Microsoft Word, and licensed MathType 6.
- Require MathType **Toggle TeX** to work manually in Word.
- Require `$...$` for inline formulas and `\[...\]` for display formulas. Do not use `$$...$$`.
- Close the input and previous output DOCX files. Leave one blank Word window open.

## Run

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File `
  scripts/convert_word_latex_to_mathtype.ps1 `
  -InputPath "C:\path\input.docx" `
  -Force
```

Default output: `{input basename}_mathtype.docx` beside the input. Pass `-OutputPath` to override it.

The script detects formula counts automatically. Use `-ExpectedInline` and `-ExpectedDisplay` only when fixed counts are an explicit acceptance requirement. Keep equation numbering disabled until formula conversion has been reviewed; `-AddEquationNumbers` is optional.

## Required workflow

1. Attach to the manually opened Microsoft Word instance.
2. Preflight both `$x^2$` and `\[x^2\]` through MathType Toggle TeX.
3. Copy the input to a separate output file.
4. Resolve formulas in Word-native paragraph coordinates.
5. Process formulas from document end to start.
6. Select each exact target range and invoke `MTCommand_OnTexToggle` directly in the target document.
7. For inline formulas inside tables, start a separate UI Automation watcher before Toggle TeX and invoke **Create Inline Style Equation** if MathType shows its inline-versus-display modal dialog. Never insert or delete guard text in table cells.
8. Require the MathType object count to increase by one after every formula.
9. If MathType ignores a plain numeric scalar, numeric interval, or trailing prime notation, retry safe equivalent forms such as `\mathrm{2048}`, `\left[200,\,3000\right]`, or `c^{\prime}`.
10. Require zero residual LaTeX and zero Word OMath objects.
11. Repair display-equation paragraphs: trim text around the MathType object, clear indents and tab stops, force center alignment, set at-least 30 pt line spacing, and remove one empty paragraph immediately after each display equation if MathType inserted it.
12. Save a recovery checkpoint after the core formula audit.
13. Preserve table count, allow at most one added paragraph boundary per display formula, and compare normalized non-formula text.

Never call `Selection.WholeStory()` and never create an equation in a temporary document for copying into the target.

## Failure handling

- Preserve the input unmodified.
- Use `-Force` only to replace an existing output copy.
- If a post-conversion integrity check fails after the checkpoint, retain the output for inspection.
- If any formula fails before the core audit, save the current partially converted output. Resume by passing that DOCX as a new input; existing `Equation.DSMT4` objects are counted and preserved while only residual LaTeX is processed.
- Treat unexpected MathType object counts, residual delimiters, table changes, or non-formula text changes as failures.
- Print a clear error block before throwing for common user-action failures: Word not open, WPS/Kingsoft detected, MathType Toggle TeX unavailable, license/demo dialogs, delimiter mismatch, and partial-output recovery.
- When Word is not found, instruct the user to open Microsoft Word manually, keep one blank document open, use the same privilege level for Word and PowerShell, verify MathType Toggle TeX manually, and rerun.
