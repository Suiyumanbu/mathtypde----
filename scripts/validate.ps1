[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path $PSScriptRoot
$runtimeScript = Join-Path $root 'mathtype-word\scripts\insert-equations.ps1'
$bridgeSource = Join-Path $root 'mathtype-word\scripts\WordComBridge.cs'

$tokens = $null
$parseErrors = $null
$null = [Management.Automation.Language.Parser]::ParseFile(
    $runtimeScript,
    [ref]$tokens,
    [ref]$parseErrors
)
if ($parseErrors.Count -gt 0) {
    $messages = ($parseErrors | ForEach-Object Message) -join [Environment]::NewLine
    throw "PowerShell parse failed:`n$messages"
}

if (-not ('WordComBridge' -as [type])) {
    Add-Type -Path $bridgeSource
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('mathtype-word-validation-' + [Guid]::NewGuid().ToString('N'))
try {
    $null = New-Item -ItemType Directory -Path $tempRoot
    $inputPath = Join-Path $tempRoot 'input.docx'
    $manifestPath = Join-Path $tempRoot 'equations.json'
    $numberExistingPath = Join-Path $tempRoot 'number-existing.json'
    $invalidPath = Join-Path $tempRoot 'invalid.json'
    $invalidOperationPath = Join-Path $tempRoot 'invalid-operation.json'
    $null = New-Item -ItemType File -Path $inputPath

    $manifest = [ordered]@{
        input = 'input.docx'
        output = 'output.docx'
        equations = @(
            [ordered]@{
                anchor = '[[MT:smoke]]'
                latex = 'E=mc^2'
                mode = 'inline'
            }
        )
    }
    [IO.File]::WriteAllText(
        $manifestPath,
        ($manifest | ConvertTo-Json -Depth 5),
        [Text.UTF8Encoding]::new($false)
    )
    & $runtimeScript -Manifest $manifestPath -ValidateOnly

    $numberExisting = [ordered]@{
        input = 'input.docx'
        output = 'numbered.docx'
        equations = @(
            [ordered]@{
                equationIndex = 1
                operation = 'number-existing'
            }
        )
    }
    [IO.File]::WriteAllText(
        $numberExistingPath,
        ($numberExisting | ConvertTo-Json -Depth 5),
        [Text.UTF8Encoding]::new($false)
    )
    & $runtimeScript -Manifest $numberExistingPath -ValidateOnly

    $invalid = [ordered]@{
        input = 'input.docx'
        output = 'output.docx'
        unsupported = $true
        equations = $manifest.equations
    }
    [IO.File]::WriteAllText(
        $invalidPath,
        ($invalid | ConvertTo-Json -Depth 5),
        [Text.UTF8Encoding]::new($false)
    )
    $rejected = $false
    try {
        & $runtimeScript -Manifest $invalidPath -ValidateOnly
    }
    catch {
        if ($_.Exception.Message -match 'Unknown manifest property') {
            $rejected = $true
        }
        else {
            throw
        }
    }
    if (-not $rejected) {
        throw 'The runtime accepted an unknown manifest property.'
    }

    $invalidOperation = [ordered]@{
        input = 'input.docx'
        output = 'invalid-operation-output.docx'
        equations = @(
            [ordered]@{
                bookmark = 'mt_existing'
                operation = 'number-existing'
                latex = 'E=mc^2'
            }
        )
    }
    [IO.File]::WriteAllText(
        $invalidOperationPath,
        ($invalidOperation | ConvertTo-Json -Depth 5),
        [Text.UTF8Encoding]::new($false)
    )
    $rejected = $false
    try {
        & $runtimeScript -Manifest $invalidOperationPath -ValidateOnly
    }
    catch {
        if ($_.Exception.Message -match 'does not accept latex or mode') {
            $rejected = $true
        }
        else {
            throw
        }
    }
    if (-not $rejected) {
        throw 'The runtime accepted LaTeX for number-existing.'
    }
}
finally {
    if (Test-Path -LiteralPath $tempRoot) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force
    }
}

Write-Host 'Static PowerShell, C#, and manifest checks passed.'
