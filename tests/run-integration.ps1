[CmdletBinding()]
param(
    [string]$PythonCommand = 'python'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path $PSScriptRoot
$output = Join-Path $PSScriptRoot 'output'
$runtime = Join-Path $root 'mathtype-word\scripts\insert-equations.ps1'

$null = New-Item -ItemType Directory -Path $output -Force
foreach ($name in @(
    'input.docx',
    'equations.json',
    'number-existing.json',
    'result.docx',
    'result.pdf',
    'result.report.json',
    'renumbered.docx',
    'renumbered.pdf',
    'renumbered.report.json'
)) {
    $path = Join-Path $output $name
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

& $PythonCommand (Join-Path $PSScriptRoot 'create_fixture.py')
if ($LASTEXITCODE -ne 0) { throw 'Fixture generation failed.' }

& $runtime -Manifest (Join-Path $output 'equations.json')
& $runtime -Manifest (Join-Path $output 'number-existing.json')

& $PythonCommand (Join-Path $PSScriptRoot 'inspect_result.py')
if ($LASTEXITCODE -ne 0) { throw 'DOCX structure inspection failed.' }

Write-Host "Integration test passed. Inspect the exported PDFs in: $output"
