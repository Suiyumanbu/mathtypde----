[CmdletBinding()]
param(
    [string]$DestinationRoot,
    [switch]$Upgrade
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$skillName = 'mathtype-word'
$source = Join-Path $PSScriptRoot $skillName
$required = @(
    'SKILL.md',
    'VERSION',
    'LICENSE',
    'THIRD_PARTY_NOTICES.md',
    'agents\openai.yaml',
    'references\manifest.md',
    'references\latex-compatibility.md',
    'scripts\insert-equations.ps1',
    'scripts\WordComBridge.cs'
)

foreach ($relativePath in $required) {
    $candidate = Join-Path $source $relativePath
    if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
        throw "The release bundle is incomplete. Missing: $relativePath"
    }
}

if ([string]::IsNullOrWhiteSpace($DestinationRoot)) {
    $userProfile = [Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)
    if (-not [string]::IsNullOrWhiteSpace($env:CODEX_HOME)) {
        $DestinationRoot = Join-Path $env:CODEX_HOME 'skills'
    }
    elseif (Test-Path -LiteralPath (Join-Path $userProfile '.agents\skills') -PathType Container) {
        $DestinationRoot = Join-Path $userProfile '.agents\skills'
    }
    elseif (Test-Path -LiteralPath (Join-Path $userProfile '.codex\skills') -PathType Container) {
        $DestinationRoot = Join-Path $userProfile '.codex\skills'
    }
    else {
        $DestinationRoot = Join-Path $userProfile '.agents\skills'
    }
}

$destinationFull = [IO.Path]::GetFullPath($DestinationRoot)
$null = [IO.Directory]::CreateDirectory($destinationFull)
$target = [IO.Path]::GetFullPath((Join-Path $destinationFull $skillName))
$targetParent = [IO.DirectoryInfo]::new($target).Parent.FullName.TrimEnd('\')
if ($targetParent -ine $destinationFull.TrimEnd('\')) {
    throw 'The resolved skill target is outside the requested destination root.'
}
if ($target.TrimEnd('\') -ieq [IO.Path]::GetFullPath($source).TrimEnd('\')) {
    throw 'The installation target must differ from the source skill directory.'
}

# Backups must not live under the discovery root: their SKILL.md would otherwise
# register another copy of the same skill on the next Codex refresh.
$destinationParent = [IO.DirectoryInfo]::new($destinationFull).Parent
if ($null -eq $destinationParent) { throw 'The skills destination must not be a filesystem root.' }
$backupRoot = [IO.Path]::GetFullPath((Join-Path $destinationParent.FullName 'skill-backups'))
if ($backupRoot -ieq $destinationFull.TrimEnd('\') -or
    $backupRoot.StartsWith($destinationFull.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'The backup directory must be outside the skills discovery root.'
}

if ((Test-Path -LiteralPath $target) -and -not $Upgrade) {
    throw "The skill is already installed at '$target'. Re-run with -Upgrade to keep a backup and install this version."
}

$staged = Join-Path $destinationFull ('.{0}.install-{1}' -f $skillName, [Guid]::NewGuid().ToString('N'))
$backup = $null
try {
    Copy-Item -LiteralPath $source -Destination $staged -Recurse

    if (Test-Path -LiteralPath $target) {
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
        $null = [IO.Directory]::CreateDirectory($backupRoot)
        $backup = Join-Path $backupRoot ("$skillName.backup-$stamp")
        if (Test-Path -LiteralPath $backup) {
            throw "Backup path already exists: $backup"
        }
        Move-Item -LiteralPath $target -Destination $backup
    }

    Move-Item -LiteralPath $staged -Destination $target
}
catch {
    if ((-not (Test-Path -LiteralPath $target)) -and $backup -and (Test-Path -LiteralPath $backup)) {
        Move-Item -LiteralPath $backup -Destination $target
        $backup = $null
    }
    if (Test-Path -LiteralPath $staged) {
        Remove-Item -LiteralPath $staged -Recurse -Force
    }
    throw
}

Write-Host "Installed $skillName at: $target"
if ($backup) {
    Write-Host "Previous installation backed up at: $backup"
}
Write-Host 'Restart Codex or begin a new task if the skill is not visible yet.'
