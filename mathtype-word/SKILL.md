---
name: mathtype-word
description: Insert editable MathType equations from LaTeX or add native right-side numbers to existing MathType display equations in Microsoft Word DOCX files on Windows.
---

# MathType equations in Word

Use this skill when the requested result must contain editable `Equation.DSMT4` MathType OLE objects. It requires desktop Microsoft Word and MathType 7 for Windows. For native Word equations (OMML) or rendered equation images, use the documents workflow instead.

The bundled script drives a separate Word instance and invokes MathType's own conversion and numbering commands. It does not use keyboard automation or the clipboard. The input DOCX is opened read-only, and the output must have a different path.

## Integrate with Word authoring

Create or edit the ordinary document content first. Put a unique text anchor at each formula location, or add a Word bookmark. Prefer anchors during generated-document workflows:

- Put an inline anchor inside its surrounding sentence, such as `Energy is [[MT:energy]] here.` An anchor alone in a paragraph is also supported when an inline OLE object is intentional.
- Put a display or right-numbered anchor alone in its paragraph.
- Use distinctive ASCII anchors that occur exactly once. Do not place them in headers, footers, text boxes, fields, existing equations, or tracked revisions.
- Prefer a Word bookmark when a display equation may later need numbering. The script restores that bookmark around the inserted OLE object, enabling a fast follow-up pass.

Generate a manifest as described in [references/manifest.md](references/manifest.md). For new formulas, supply a self-contained LaTeX math body; outer `$...$`, `\(...\)`, `$$...$$`, or `\[...\]` delimiters are optional because the script chooses them from `mode`.

Batch all requested formulas in one manifest rather than launching Word once per formula. If the final result needs right numbers, insert new formulas directly as `right-numbered`; use `number-existing` only for OLE equations already present in the input.

Run the conversion directly; it validates the manifest and environment before Word is launched:

```powershell
& scripts/insert-equations.ps1 -Manifest path\to\equations.json
```

Use the standalone checks only for CI, first-time setup, or troubleshooting:

```powershell
& scripts/insert-equations.ps1 -Manifest path\to\equations.json -ValidateOnly
& scripts/insert-equations.ps1 -Probe
```

The script rejects an existing output instead of overwriting it. It publishes the DOCX only after all operations and optional PDF export succeed. A sibling `.report.json` records each operation, whether the OLE was reused, structural results, and phase timings.

## Numbered equations

For `right-numbered`, preserve the document's current MathType equation-number format. In a document with no MathType chapter or section field, the script invokes MathType's noninteractive initial chapter command; the installed default commonly yields `(1.1)`, `(1.2)`, and so on. If the user needs a different format, set it in Word's MathType equation-number settings before running this skill, then use that prepared DOCX as the input.

When converting an existing display equation to right-numbered form, use `operation: "number-existing"` with either a bookmark around that equation or its one-based index among MathType OLE equations in the main document. Do not extract or regenerate its LaTeX. The script keeps the original `Equation.DSMT4` payload and inserts only the native number fields. Prefer bookmarks because an index can change when equations are added or removed.

The output number is a MathType-compatible `MTPlaceRef` field containing `MTEqn` sequence fields. Do not substitute a typed number, a plain Word `SEQ` label, or a table-only imitation when MathType-native numbering was requested.

During native-number insertion, the script temporarily enables MathType's document-level deferred field updates and then restores the original setting. It refreshes MathType number fields from the earliest new number onward so downstream sequence values remain correct, without updating unrelated Word fields.

## Validate the result

Match validation cost to the stage of work:

1. For every run, read the report and require `progId: Equation.DSMT4`; require positive `nativeNumberFields` for right-numbered results. For `number-existing`, also require `source: existing` and `conversionMilliseconds: 0`.
2. During iteration, omit `pdf` unless layout evidence is needed. Structural inspection is sufficient to catch missing OLE objects, anchors, or native number fields.
3. For review, export a PDF and inspect pages containing changed equations or downstream numbers. Inspect every page only for a final deliverable, a document-wide layout change, or an explicit user request.
4. Record visual approval in `visualReview` only when a PDF was actually inspected.

For `number-existing`, the embedding count and OLE payloads should remain unchanged while `MTPlaceRef` and `SEQ MTEqn` fields increase. For LaTeX insertion, expect one new embedding per formula.

If MathType leaves a blank, malformed, or altered formula, simplify the LaTeX according to [references/latex-compatibility.md](references/latex-compatibility.md), rerun from the unchanged source DOCX to a new output path, and inspect again.

## Constraints

- Use only a local Windows desktop session. Word COM and MathType OLE are unavailable in a headless container or cloud-only task.
- Do not run two MathType automation jobs concurrently. MathType's Word commands serialize access to their API.
- Resolve document protection and tracked changes before inserting equations; the script stops on either condition.
- Keep the source, manifest, outputs, and reports in paths available to the local task. Do not attach to or close the user's already-open Word documents.
- Never describe an equation as editable MathType unless the output report and DOCX structure confirm `Equation.DSMT4`.
