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
    }
  ]
}
```

Each equation accepts exactly these fields:

- `anchor`: unique literal text to replace. Use either `anchor` or `bookmark`.
- `bookmark`: existing Word bookmark whose range is replaced. Use either `bookmark` or `anchor`.
- `latex`: nonempty, self-contained LaTeX math content. In JSON, write each backslash as `\\`.
- `mode`: `inline`, `display`, or `right-numbered`.

Top-level fields are `input`, `output`, optional `pdf`, and `equations`. Unknown fields are rejected so misspellings cannot silently change behavior.

An inline locator may appear within text or alone in a paragraph. A display or right-numbered locator must be the only non-whitespace content in its paragraph. The script suppresses MathType's interactive display suggestion when a standalone inline equation was explicitly requested.

The output DOCX, optional PDF, and sibling report must not already exist. Use a new output name when retrying so the original and earlier results remain recoverable.
