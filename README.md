# MathType Word skill for Codex

[简体中文](README.zh-CN.md)

`mathtype-word` inserts Codex-generated LaTeX as editable MathType equations at exact locations in a Word document. It supports equations within a sentence, centered display equations, and right-numbered equations managed by MathType's native numbering commands.

The repository follows the [OpenAI skill folder format](https://learn.chatgpt.com/docs/build-skills): the installable skill lives in `mathtype-word/`, starts with `SKILL.md`, and keeps its scripts and references beside it.

## Features

- Converts LaTeX through MathType's own TeX translator into editable `Equation.DSMT4` OLE objects.
- Locates insertion points by unique text anchors or Word bookmarks.
- Supports `inline`, `display`, and `right-numbered` modes.
- Uses MathType-native `MTEqn`/`MTPlaceRef` numbering so MathType's number-management commands remain compatible.
- Opens the source read-only, refuses to overwrite output, and publishes only after the whole job succeeds.
- Produces a structural JSON report and can export a PDF for visual review.
- Avoids keyboard automation and the clipboard.

## Requirements

- Windows desktop session
- Microsoft Word desktop
- MathType 7 with its Word commands installed
- Codex with local skills enabled

The current release has been exercised with 64-bit Word 2024 and MathType 7.11. Other current Word/MathType 7 combinations should work, but should be verified with `-Probe` and the included integration test.

## Install

### Release bundle

Download the ZIP from the repository's **Releases** page, extract it, and run:

```powershell
& .\install.ps1
```

The installer uses `$CODEX_HOME/skills` when `CODEX_HOME` is set. Otherwise it prefers an existing `~/.agents/skills`, then an existing `~/.codex/skills`, and falls back to `~/.agents/skills`. Pass `-DestinationRoot` to choose another skill root. Existing installations are left untouched unless `-Upgrade` is supplied; upgrades keep a timestamped backup.

### Manual install

Copy the complete `mathtype-word` folder into your Codex skills directory, preserving this structure:

```text
skills/
└── mathtype-word/
    ├── SKILL.md
    ├── VERSION
    ├── LICENSE
    ├── THIRD_PARTY_NOTICES.md
    ├── agents/
    ├── references/
    └── scripts/
```

Restart Codex or start a new task if the skill is not discovered immediately.

## Use from Codex

Put unique anchors in the generated DOCX, then ask Codex to use the skill. For example:

```text
Use $mathtype-word to replace [[MT:energy]] with E=mc^2 as an inline
MathType equation, and replace [[MT:quadratic]] with the quadratic formula
as a right-numbered MathType equation. Export a PDF and inspect the result.
```

Codex creates a JSON manifest and runs the bundled PowerShell bridge. You can also call it directly:

```powershell
& .\mathtype-word\scripts\insert-equations.ps1 `
  -Manifest .\examples\equations.example.json `
  -ValidateOnly

& .\mathtype-word\scripts\insert-equations.ps1 `
  -Manifest .\examples\equations.example.json
```

Manifest paths are relative to the JSON file. Copy `examples/equations.example.json` beside your input DOCX or adjust its paths first.

| Mode | Result | Locator rule |
|---|---|---|
| `inline` | Equation within surrounding text, or a standalone inline OLE object | Anchor or bookmark |
| `display` | Centered MathType display equation | Locator must occupy its paragraph |
| `right-numbered` | Display equation plus a MathType-managed number at the right margin | Locator must occupy its paragraph |

See [the manifest reference](mathtype-word/references/manifest.md) and [LaTeX compatibility notes](mathtype-word/references/latex-compatibility.md) for the complete contract.

## Validate and test

Static checks do not require Word or MathType:

```powershell
python tools/validate_release.py
& .\scripts\validate.ps1
```

The integration test creates a DOCX fixture, inserts six equations, exports a PDF, and checks the OLE objects and native numbering fields. It requires a local Word/MathType installation:

```powershell
python -m pip install -r tests/requirements.txt
& .\tests\run-integration.ps1
```

## Build a release ZIP

```powershell
& .\scripts\package-release.ps1
```

The command creates `dist/mathtype-word-<version>.zip` and a matching SHA-256 file. Tagging a GitHub commit as `v<version>` runs the included release workflow and attaches both files to a GitHub Release.

To check a downloaded bundle in PowerShell, compare `Get-FileHash .\mathtype-word-<version>.zip -Algorithm SHA256` with the value in the adjacent `.sha256` file.

## License and product names

The project is licensed under the [MIT License](LICENSE). Microsoft Word and MathType are required third-party products and are not distributed here; see [Third-party notices](THIRD_PARTY_NOTICES.md).
