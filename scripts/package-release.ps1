[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path $PSScriptRoot
$version = (Get-Content -LiteralPath (Join-Path $root 'VERSION') -Raw).Trim()
if ($version -notmatch '^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$') {
    throw "VERSION is not a supported semantic version: $version"
}

$dist = Join-Path $root 'dist'
$bundleName = "mathtype-word-$version"
$zipPath = Join-Path $dist "$bundleName.zip"
$checksumPath = "$zipPath.sha256"
if ((Test-Path -LiteralPath $zipPath) -or (Test-Path -LiteralPath $checksumPath)) {
    if (-not $Force) {
        throw "Release output already exists. Re-run with -Force to replace version $version."
    }
    foreach ($path in @($zipPath, $checksumPath)) {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
        }
    }
}

$null = New-Item -ItemType Directory -Path $dist -Force
$stagingRoot = Join-Path $dist ('.staging-' + [Guid]::NewGuid().ToString('N'))
$bundleRoot = Join-Path $stagingRoot $bundleName
try {
    $null = New-Item -ItemType Directory -Path $bundleRoot
    foreach ($name in @('README.md', 'README.zh-CN.md', 'LICENSE', 'THIRD_PARTY_NOTICES.md', 'VERSION', 'install.ps1')) {
        Copy-Item -LiteralPath (Join-Path $root $name) -Destination $bundleRoot
    }
    Copy-Item -LiteralPath (Join-Path $root 'mathtype-word') -Destination $bundleRoot -Recurse

    Compress-Archive -LiteralPath $bundleRoot -DestinationPath $zipPath -CompressionLevel Optimal
    $hash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    [IO.File]::WriteAllText(
        $checksumPath,
        "$hash  $([IO.Path]::GetFileName($zipPath))`n",
        [Text.UTF8Encoding]::new($false)
    )
}
finally {
    if (Test-Path -LiteralPath $stagingRoot) {
        Remove-Item -LiteralPath $stagingRoot -Recurse -Force
    }
}

Write-Host "Created: $zipPath"
Write-Host "SHA-256: $checksumPath"
