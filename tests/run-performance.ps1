[CmdletBinding()]
param(
    [string]$PythonCommand = 'python',
    [ValidateRange(1, 1000)][int[]]$Counts = @(10, 50),
    [ValidateRange(1, 20)][int]$Repeats = 3,
    [ValidateRange(0, 10000)][int]$FillerParagraphs = 500
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = Split-Path $PSScriptRoot
$runtime = Join-Path $root 'mathtype-word\scripts\insert-equations.ps1'
$fixture = Join-Path $PSScriptRoot 'performance_fixture.py'
$output = Join-Path $PSScriptRoot ('output\performance-' + [Guid]::NewGuid().ToString('N'))
$shellPath = Join-Path $PSHOME 'pwsh.exe'
if (-not (Test-Path -LiteralPath $shellPath)) { $shellPath = Join-Path $PSHOME 'powershell.exe' }
$rows = [Collections.Generic.List[object]]::new()

& $PythonCommand -B $fixture prepare --output $output --counts $Counts --repeats $Repeats --filler-paragraphs $FillerParagraphs
if ($LASTEXITCODE -ne 0) { throw 'Performance fixture generation failed.' }
foreach ($count in $Counts) {
    $caseRoot = Join-Path $output "n-$count"
    & $shellPath -NoProfile -File $runtime -Manifest (Join-Path $caseRoot 'seed.json')
    if ($LASTEXITCODE -ne 0) { throw 'Display-equation seed conversion failed.' }
    for ($repeat = 1; $repeat -le $Repeats; $repeat++) {
        # Alternate order so the reuse strategy is not always the warmer run.
        $strategies = if ($repeat % 2 -eq 1) { @('recreate', 'reuse') } else { @('reuse', 'recreate') }
        foreach ($strategy in $strategies) {
            Write-Host "Benchmark: $count equations / $strategy / repeat $repeat"
            $timer = [Diagnostics.Stopwatch]::StartNew()
            & $shellPath -NoProfile -File $runtime -Manifest (Join-Path $caseRoot "$strategy-$repeat.json")
            $timer.Stop()
            if ($LASTEXITCODE -ne 0) { throw "Benchmark $strategy-$repeat failed." }
            $rows.Add([pscustomobject]@{
                equations = $count
                strategy = $strategy
                repeat = $repeat
                fillerParagraphs = $FillerParagraphs
                wallMilliseconds = [Math]::Round($timer.Elapsed.TotalMilliseconds, 1)
            })
        }
    }
}
ConvertTo-Json -InputObject $rows.ToArray() -Depth 5 | Set-Content -LiteralPath (Join-Path $output 'measurements.json') -Encoding UTF8
& $PythonCommand -B $fixture inspect --output $output
if ($LASTEXITCODE -ne 0) { throw 'Performance output verification failed.' }
Write-Host "Performance evidence: $(Join-Path $output 'performance-summary.json')"
