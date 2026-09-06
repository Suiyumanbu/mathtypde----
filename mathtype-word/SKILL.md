---
name: mathtype-word
description: Insert editable MathType equations from LaTeX at exact locations in Microsoft Word DOCX files on Windows. Use for inline, centered display, or MathType-native right-numbered equations in a Codex Word-generation workflow.
---

# MathType equations in Word

Use this skill when the requested result must contain editable `Equation.DSMT4` MathType OLE objects. It requires desktop Microsoft Word and MathType 7 for Windows. For native Word equations (OMML) or rendered equation images, use the documents workflow instead.

The bundled script drives a separate Word instance and invokes MathType's own `MTCommand_TeXToggle` and equation-number insertion commands. It does not use keyboard automation or the clipboard. The input DOCX is opened read-only, and the output must have a different path.

## Integrate with Word authoring

Create or edit the ordinary document content first. Put a unique text anchor at each formula location, or add a Word bookmark. Prefer anchors during generated-document workflows:

- Put an inline anchor inside its surrounding sentence, such as `Energy is [[MT:energy]] here.` An anchor alone in a paragraph is also supported when an inline OLE object is intentional.
- Put a display or right-numbered anchor alone in its paragraph.
- Use distinctive ASCII anchors that occur exactly once. Do not place them in headers, footers, text boxes, fields, existing equations, or tracked revisions.

Generate a manifest as described in [references/manifest.md](references/manifest.md). Supply a self-contained LaTeX math body for each formula; outer `$...$`, `\(...\)`, `$$...$$`, or `\[...\]` delimiters are optional because the script chooses them from `mode`.

Run the preflight before launching Word:

```powershell
& scripts/insert-equations.ps1 -Manifest path\to\equations.json -ValidateOnly
& scripts/insert-equations.ps1 -Probe
```

Then run the conversion:

```powershell
& scripts/insert-equations.ps1 -Manifest path\to\equations.json
```

The script rejects an existing output instead of overwriting it. It publishes the DOCX only after all equations and optional PDF export succeed. A sibling `.report.json` records every formula's locator, mode, OLE ProgID, dimensions, and native number-field count.

## Numbered equations

For `right-numbered`, preserve the document's current MathType equation-number format. In a document with no MathType chapter or section field, the script invokes MathType's noninteractive initial chapter command; the installed default commonly yields `(1.1)`, `(1.2)`, and so on. If the user needs a different format, set it in Word's MathType equation-number settings before running this skill, then use that prepared DOCX as the input.

The output number is a MathType-compatible `MTPlaceRef` field containing `MTEqn` sequence fields. Do not substitute a typed number, a plain Word `SEQ` label, or a table-only imitation when MathType-native numbering was requested.

During native-number insertion, the script temporarily enables MathType's document-level deferred field updates and then restores the original setting. This prevents MathType's legacy numbering macro from updating unrelated Word fields throughout the document; only the newly inserted number fields are updated.

## Validate the result

Treat a successful command as structural evidence, not final visual approval.

1. Read the generated report and require `progId` to be `Equation.DSMT4` for every item. For each `right-numbered` item, require a positive `nativeNumberFields` count.
2. Inspect the DOCX package when correctness matters: `word/embeddings/` should contain one new OLE object per inserted formula, all anchors should be gone, and `word/document.xml` should contain `MTPlaceRef` and `SEQ MTEqn` for numbered formulas.
3. Export a PDF through the manifest and inspect every rendered page. If the documents skill is available, its render workflow can also be used. Check equation content, baseline alignment, centering, right-number placement, clipping, and page breaks. LaTeX conversion is not complete until this visual review passes.
4. After the page review passes, replace the report's `visualReview` value with a short record such as `passed: 3 pages inspected`.

If MathType leaves a blank, malformed, or altered formula, simplify the LaTeX according to [references/latex-compatibility.md](references/latex-compatibility.md), rerun from the unchanged source DOCX to a new output path, and inspect again.

## Constraints

- Use only a local Windows desktop session. Word COM and MathType OLE are unavailable in a headless container or cloud-only task.
- Do not run two MathType automation jobs concurrently. MathType's Word commands serialize access to their API.
- Resolve document protection and tracked changes before inserting equations; the script stops on either condition.
- Keep the source, manifest, outputs, and reports in paths available to the local task. Do not attach to or close the user's already-open Word documents.
- Never describe an equation as editable MathType unless the output report and DOCX structure confirm `Equation.DSMT4`.
