# Manifest format

Paths are resolved relative to the JSON manifest. The input and output must be different `.docx` files. The optional `pdf` is exported by Microsoft Word and is useful for visual review.

```json
{
  "input": "draft.docx",
  "output": "draft-with-equations.docx",
  "pdf": "draft-with-equations.pdf",
  "equations": [
    {
      "anchor": "[[MT:energy]]",
      "latex": "E=mc^2",
      "mode": "inline"
    },
    {
      "anchor": "[[MT:integral]]",
      "latex": "\\int_0^1 x^2\\,dx=\\frac{1}{3}",
      "mode": "display"
    },
    {
      "anchor": "[[MT:quadratic]]",
      "latex": "\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}",
      "mode": "right-numbered"
    },
    {
      "bookmark": "mt_momentum",
      "latex": "p=mv",
      "mode": "inline"
    },
    {
      "bookmark": "mt_existing_display",
      "operation": "number-existing"
    },
    {
      "equationIndex": 4,
      "operation": "number-existing"
    }
  ]
}
```

The array accepts two operation forms. Existing manifests without `operation` remain compatible and mean `insert-latex`.

### Insert LaTeX

An insertion accepts:

- `anchor`: unique literal text to replace. Use either `anchor` or `bookmark`.
- `bookmark`: existing Word bookmark whose range is replaced. Use either `bookmark` or `anchor`.
- `latex`: nonempty, self-contained LaTeX math content. In JSON, write each backslash as `\\`.
- `mode`: `inline`, `display`, or `right-numbered`.
- `operation`: optional `insert-latex`; omit it in ordinary manifests.

An inline locator may appear within text or alone in a paragraph. A display or right-numbered locator must be the only non-whitespace content in its paragraph. The script suppresses MathType's interactive display suggestion when a standalone inline equation was explicitly requested.

When an insertion uses a bookmark, the output bookmark is restored around the generated MathType OLE object. Use bookmarks for display equations that may later receive native numbers.

### Number an existing display equation

Use `operation: "number-existing"` with exactly one locator:

- `bookmark`: a bookmark containing exactly one existing `Equation.DSMT4` object; or
- `equationIndex`: the one-based position among MathType `Equation.DSMT4` inline objects in the main document story.

Do not supply `latex`, `mode`, or `anchor` for this operation. The target must occupy its own paragraph, must not be inside a table, and must not already contain native MathType number fields. The operation preserves the existing OLE payload and adds only `MTPlaceRef`/`MTEqn` fields. Prefer a bookmark because indices change when equations are inserted or removed.

All locators, including `equationIndex`, are resolved against the input before edits begin. Operations are then executed in document order regardless of their array order. This keeps native sequence numbers aligned with the document.

Top-level fields are `input`, `output`, optional `pdf`, and `equations`. Unknown fields are rejected so misspellings cannot silently change behavior.

The output DOCX, optional PDF, and sibling report must not already exist. Use a new output name when retrying so the original and earlier results remain recoverable.

The report includes per-operation conversion and numbering times plus overall phase timings. A `number-existing` entry should report `source: "existing"` and `conversionMilliseconds: 0`. `processingMilliseconds` ends after save/export and excludes publication and Word shutdown; the repository performance runner also measures full process wall time.
