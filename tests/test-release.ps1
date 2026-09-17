[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = Split-Path $PSScriptRoot
$version = (Get-Content -Raw (Join-Path $root 'VERSION')).Trim()
$zipPath = Join-Path $root "dist\mathtype-word-$version.zip"
$expectedHash = (Get-Content ($zipPath + '.sha256')).Split(' ')[0]
if ((Get-FileHash $zipPath -Algorithm SHA256).Hash -ine $expectedHash) { throw 'Release checksum mismatch.' }
$output = Join-Path $PSScriptRoot ('output\release-' + [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $output
Expand-Archive -LiteralPath $zipPath -DestinationPath $output
$bundle = Join-Path $output "mathtype-word-$version"
foreach ($name in @('README.md', 'README.zh-CN.md', 'LICENSE', 'THIRD_PARTY_NOTICES.md', 'VERSION', 'install.ps1')) {
    if ((Get-FileHash (Join-Path $root $name)).Hash -ne (Get-FileHash (Join-Path $bundle $name)).Hash) {
        throw "Stale packaged file: $name"
    }
}
foreach ($directory in @('mathtype-word', 'examples')) {
    $sourceRoot = Join-Path $root $directory
    $bundleRoot = Join-Path $bundle $directory
    foreach ($sourceFile in Get-ChildItem -LiteralPath $sourceRoot -File -Recurse) {
        $relativePath = $sourceFile.FullName.Substring($sourceRoot.Length).TrimStart('\')
        $bundleFile = Join-Path $bundleRoot $relativePath
        if (-not (Test-Path -LiteralPath $bundleFile)) { throw "Missing packaged file: $relativePath" }
        if ((Get-FileHash $sourceFile.FullName).Hash -ne (Get-FileHash $bundleFile).Hash) {
            throw "Stale packaged file: $relativePath"
        }
    }
}
$install = Join-Path $bundle 'install.ps1'
$rejected = $false
try { & $install -DestinationRoot $bundle -Upgrade }
catch {
    if ($_.Exception.Message -notmatch 'must differ from the source') { throw }
    $rejected = $true
}
if (-not $rejected) { throw 'Installing onto the source directory must fail.' }
$destination = Join-Path $output 'skills'
& $install -DestinationRoot $destination
$target = Join-Path $destination 'mathtype-word'
$originalHash = (Get-FileHash (Join-Path $target 'scripts\insert-equations.ps1')).Hash
$rejected = $false
try { & $install -DestinationRoot $destination }
catch {
    if ($_.Exception.Message -notmatch 'already installed') { throw }
    $rejected = $true
}
if (-not $rejected) { throw 'Installing twice without -Upgrade must fail.' }
& $install -DestinationRoot $destination -Upgrade
$backups = @(Get-ChildItem -LiteralPath (Join-Path $output 'skill-backups') -Directory)
if ($backups.Count -ne 1) { throw 'An upgrade must create one recoverable backup.' }
if ((Get-FileHash (Join-Path $backups[0].FullName 'scripts\insert-equations.ps1')).Hash -ne $originalHash) {
    throw 'The backup differs from the previous installation.'
}
$discovered = @(Get-ChildItem -LiteralPath $destination -File -Recurse -Filter 'SKILL.md')
if ($discovered.Count -ne 1) { throw 'Backup SKILL.md leaked into the discovery root.' }
Write-Host "Release checksum, package/source parity, repeat-install rejection, backup preservation, and single discovery passed: $output"
