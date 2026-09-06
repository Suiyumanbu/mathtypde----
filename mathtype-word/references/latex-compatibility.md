# MathType LaTeX compatibility

MathType's TeX translator accepts equation input, not a complete LaTeX document. Codex should generate the smallest self-contained expression that preserves the requested notation.

Prefer standard constructs such as fractions, roots, scripts, sums, integrals, Greek letters, delimiters, accents, arrays, and ordinary style commands. Do not send `\documentclass`, `\usepackage`, file includes, macro definitions, environments that require a package, or document prose around the equation. The script rejects document-level commands and custom definitions before Word starts.

MathType's translator is not a full TeX engine. Package macros, user-defined commands, and some spacing or typography commands may be rejected or converted differently. When generated notation is complex:

1. Expand custom macros into standard TeX.
2. Remove purely cosmetic spacing commands when they are not essential.
3. Split a large aligned construction only when that matches the user's intended layout.
4. Inspect the rendered formula symbol by symbol, including limits, indices, brackets, accents, matrices, and text inside math.

JSON requires doubled backslashes. For example, the LaTeX `\frac{a}{b}` is written as `"\\frac{a}{b}"` in the manifest. Do not double braces or mathematical backslashes beyond JSON escaping.
