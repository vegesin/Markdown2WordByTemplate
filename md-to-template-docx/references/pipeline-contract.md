# Pipeline contract

## Stage 1 output

The DOCX intentionally contains literal LaTeX. It is not the final document.

- Inline: `$latex$`
- Display: centered `\[latex\]` with at-least 30 pt line spacing
- Diagram: `[VISIO:diagram-NN]`

Do not add equation-number fields before MathType conversion. Add right-aligned numbering only after every LaTeX range has been replaced by an `Equation.DSMT4` object.

## MathType handoff

Use `word-mathtype-selective` on a copy. Never run `MTCommand_OnTexToggle` on `Selection.WholeStory()` for a Chinese document. Resolve one exact target range at a time and run Toggle TeX directly on that selection. The expected final MathType object ProgID is `Equation.DSMT4`.

After conversion, require:

- expected MathType object count;
- zero residual `$...$` and `\[...\]` fragments;
- zero Word OMath objects;
- preserved equation-number fields and Visio placeholders.

## Visio handoff

Use the generated `manifest.json`. Generate one editable `.vsdx` for every `.mmd`, then replace the matching `[VISIO:diagram-NN]` paragraph in Word. Preserve any following figure-caption paragraph.
