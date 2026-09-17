[CmdletBinding()]
param(
    [string]$PythonCommand = 'python',
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path $PSScriptRoot
$output = if ($OutputDirectory) { [IO.Path]::GetFullPath($OutputDirectory) } else {
    Join-Path $PSScriptRoot ('output\integration-' + [Guid]::NewGuid().ToString('N'))
}
$runtime = Join-Path $root 'mathtype-word\scripts\insert-equations.ps1'

if (Test-Path -LiteralPath $output) { throw 'Use a new integration output directory; prior results are preserved.' }
$null = New-Item -ItemType Directory -Path $output

& $PythonCommand -B (Join-Path $PSScriptRoot 'create_fixture.py') --output $output
if ($LASTEXITCODE -ne 0) { throw 'Fixture generation failed.' }

& $runtime -Manifest (Join-Path $output 'equations.json')
& $runtime -Manifest (Join-Path $output 'number-existing.json')

& $PythonCommand -B (Join-Path $PSScriptRoot 'inspect_result.py') --output $output
if ($LASTEXITCODE -ne 0) { throw 'DOCX structure inspection failed.' }

& $PythonCommand -B (Join-Path $PSScriptRoot 'regression_cases.py') prepare --output $output
if ($LASTEXITCODE -ne 0) { throw 'Regression fixture preparation failed.' }
$sourceHashes = @{}
foreach ($sourceName in @('result.docx', 'renumbered.docx')) {
    $sourceHashes[$sourceName] = (Get-FileHash (Join-Path $output $sourceName) -Algorithm SHA256).Hash
}

foreach ($case in @(
    @{ Name = 'already-numbered'; Message = 'already right-numbered' },
    @{ Name = 'missing-bookmark'; Message = 'Missing bookmark' },
    @{ Name = 'out-of-range'; Message = 'out of range' },
    @{ Name = 'overlap'; Message = 'overlap or resolve to the same range' }
)) {
    $rejected = $false
    try { & $runtime -Manifest (Join-Path $output ($case.Name + '.json')) }
    catch {
        if ($_.Exception.Message -notmatch $case.Message) { throw }
        $rejected = $true
    }
    if (-not $rejected) { throw "The runtime did not reject $($case.Name)." }
    foreach ($suffix in @('.docx', '.report.json', '.pdf')) {
        if (Test-Path -LiteralPath (Join-Path $output ($case.Name + $suffix))) {
            throw "A failed $($case.Name) job published output."
        }
    }
}
foreach ($sourceName in $sourceHashes.Keys) {
    if ((Get-FileHash (Join-Path $output $sourceName) -Algorithm SHA256).Hash -ne $sourceHashes[$sourceName]) {
        throw 'A rejected job modified its source document.'
    }
}

& $runtime -Manifest (Join-Path $output 'mixed.json')
& $PythonCommand -B (Join-Path $PSScriptRoot 'regression_cases.py') inspect --output $output
if ($LASTEXITCODE -ne 0) { throw 'Mixed operation regression inspection failed.' }

Write-Host "Integration test passed. Inspect the exported PDFs in: $output"
