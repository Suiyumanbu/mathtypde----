# Contributing

Issues and pull requests are welcome. Keep the distributable skill inside `mathtype-word/`; repository tooling, tests, and release files belong at the repository root.

Before opening a pull request, run:

```powershell
python tools/validate_release.py
& .\scripts\validate.ps1
```

Changes to Word or MathType automation should also pass the local integration test on a Windows desktop with both applications installed:

```powershell
python -m pip install -r tests/requirements.txt
& .\tests\run-integration.ps1
```

For performance-related changes, run `tests/run-performance.ps1`; its alternating repeated runs compare OLE reuse with regeneration on the same formula content without PDF export. Keep full-process wall time separate from the runtime's save/export processing time. After packaging, `tests/test-release.ps1` checks source parity and safe backup placement.

Do not commit generated DOCX, PDF, report, or distribution files. Do not add proprietary Microsoft or Wiris binaries to the repository. When changing MathType calls, document the public macro or observable Word behavior on which the implementation relies.
