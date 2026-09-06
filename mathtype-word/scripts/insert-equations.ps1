[CmdletBinding()]
param(
    [string]$Manifest,
    [switch]$Probe,
    [switch]$ValidateOnly,
    [string]$MathTypeRoot
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Resolve-FullPath([string]$Path, [string]$BaseDirectory) {
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    return [IO.Path]::GetFullPath((Join-Path $BaseDirectory $Path))
}

function Get-NormalizedLatex([string]$Latex) {
    $value = $Latex.Trim()
    if ($value.StartsWith('$$') -and $value.EndsWith('$$') -and $value.Length -ge 4) {
        return $value.Substring(2, $value.Length - 4).Trim()
    }
    if ($value.StartsWith('$') -and $value.EndsWith('$') -and $value.Length -ge 2) {
        return $value.Substring(1, $value.Length - 2).Trim()
    }
    if ((($value.StartsWith('\[') -and $value.EndsWith('\]')) -or
         ($value.StartsWith('\(') -and $value.EndsWith('\)'))) -and $value.Length -ge 4) {
        return $value.Substring(2, $value.Length - 4).Trim()
    }
    return $value
}

function Find-Anchor($Document, $Item) {
    if ($null -ne $Item.PSObject.Properties['bookmark']) {
        $name = [string]$Item.bookmark
        if (-not $Document.Bookmarks.Exists($name)) { throw "Missing bookmark: $name" }
        $range = $Document.Bookmarks.Item($name).Range.Duplicate
        if ($range.StoryType -ne 1) { throw "Bookmark '$name' is outside the main document story." }
        return $range
    }

    $anchor = [string]$Item.anchor
    $range = $Document.Content.Duplicate
    $range.Find.ClearFormatting()
    $found = $range.Find.Execute($anchor, $true, $false, $false, $false, $false, $true, 0, $false)
    if (-not $found) { throw "Missing anchor: $anchor" }
    $next = $Document.Range($range.End, $Document.Content.End)
    if ($next.Find.Execute($anchor, $true, $false, $false, $false, $false, $true, 0, $false)) {
        throw "Ambiguous anchor (must occur exactly once): $anchor"
    }
    return $range
}

function Invoke-MathTypeMacro($Word, [string]$Name) {
    Write-Host "MathType macro: $Name"
    $null = $Word.Run($Name)
}

function Find-InsertedEquation($Document, [int]$ExpectedStart, [int]$ExpectedIndex) {
    if ($ExpectedIndex -lt 1 -or $ExpectedIndex -gt $Document.InlineShapes.Count) {
        throw "MathType did not create an Equation.DSMT4 object near position $ExpectedStart."
    }
    $shape = $Document.InlineShapes.Item($ExpectedIndex)
    if ($shape.OLEFormat.ProgID -ne 'Equation.DSMT4' -or
        [Math]::Abs([int]$shape.Range.Start - $ExpectedStart) -gt 8) {
        throw "The new inline object near position $ExpectedStart is not the expected MathType equation."
    }
    return $shape
}

function Get-PeBitness([string]$Path) {
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $reader = [IO.BinaryReader]::new($stream)
    try {
        $stream.Position = 0x3c
        $peOffset = $reader.ReadInt32()
        $stream.Position = $peOffset + 4
        $machine = $reader.ReadUInt16()
        if ($machine -eq 0x8664) { return 64 }
        if ($machine -eq 0x014c) { return 32 }
        throw "Unsupported WINWORD architecture: 0x$($machine.ToString('X4'))"
    }
    finally {
        $reader.Dispose()
        $stream.Dispose()
    }
}

if (-not $Manifest -and -not $Probe) { throw 'Specify -Manifest equations.json, or use -Probe.' }
$config = $null
$manifestPath = $null
$manifestDirectory = $null
$inputPath = $null
$outputPath = $null
$pdfPath = $null
$reportPath = $null
$items = @()

if ($Manifest) {
$manifestPath = (Resolve-Path -LiteralPath $Manifest).Path
$manifestDirectory = Split-Path $manifestPath
$config = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$topLevelAllowed = @('input', 'output', 'pdf', 'equations')
foreach ($property in $config.PSObject.Properties.Name) {
    if ($property -notin $topLevelAllowed) { throw "Unknown manifest property: $property" }
}
foreach ($required in @('input', 'output', 'equations')) {
    if ($null -eq $config.PSObject.Properties[$required]) { throw "Manifest is missing '$required'." }
}
$inputPath = Resolve-FullPath ([string]$config.input) $manifestDirectory
$outputPath = Resolve-FullPath ([string]$config.output) $manifestDirectory
if ($inputPath -eq $outputPath) { throw 'Output must differ from input; the source document is never overwritten.' }
if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw "Missing input: $inputPath" }
if (Test-Path -LiteralPath $outputPath) { throw "Output already exists: $outputPath" }
if ([IO.Path]::GetExtension($inputPath) -ne '.docx' -or [IO.Path]::GetExtension($outputPath) -ne '.docx') {
    throw 'Use .docx input and output files.'
}
$reportPath = [IO.Path]::ChangeExtension($outputPath, '.report.json')
if (Test-Path -LiteralPath $reportPath) { throw "Report output already exists: $reportPath" }
if ($null -ne $config.PSObject.Properties['pdf']) {
    $pdfPath = Resolve-FullPath ([string]$config.pdf) $manifestDirectory
    if ([IO.Path]::GetExtension($pdfPath) -ne '.pdf') { throw 'The optional pdf output must use a .pdf extension.' }
    if (Test-Path -LiteralPath $pdfPath) { throw "PDF output already exists: $pdfPath" }
}

$items = @($config.equations)
if ($items.Count -eq 0) { throw 'The equations array is empty.' }
$seen = @{}
foreach ($item in $items) {
    $itemAllowed = @('anchor', 'bookmark', 'latex', 'mode')
    foreach ($property in $item.PSObject.Properties.Name) {
        if ($property -notin $itemAllowed) { throw "Unknown equation property: $property" }
    }
    if ($null -eq $item.PSObject.Properties['mode'] -or $item.mode -notin @('inline', 'display', 'right-numbered')) {
        throw 'mode must be inline, display, or right-numbered.'
    }
    if ($null -eq $item.PSObject.Properties['latex'] -or -not ($item.latex -is [string])) {
        throw 'Each equation needs a LaTeX string.'
    }
    $latex = Get-NormalizedLatex ([string]$item.latex)
    if ([string]::IsNullOrWhiteSpace($latex)) { throw 'LaTeX cannot be empty.' }
    if ($latex -match '\\(documentclass|usepackage|newcommand|renewcommand|def|input|include)\b') {
        throw 'Pass a self-contained equation body, without a LaTeX document or custom macro definitions.'
    }
    $hasAnchor = $null -ne $item.PSObject.Properties['anchor']
    $hasBookmark = $null -ne $item.PSObject.Properties['bookmark']
    if ($hasAnchor -eq $hasBookmark) { throw 'Specify exactly one locator: anchor or bookmark.' }
    $locator = if ($hasAnchor) { [string]$item.anchor } else { [string]$item.bookmark }
    if ([string]::IsNullOrWhiteSpace($locator) -or $locator -match '[\r\n\^]') {
        throw 'Locator is empty or contains unsupported characters.'
    }
    if ($seen.ContainsKey($locator)) { throw "Duplicate locator: $locator" }
    $seen[$locator] = $true
}
}

if ($ValidateOnly) {
    Write-Output "Manifest valid: $($items.Count) equations. Word was not launched."
    return
}

if (-not $MathTypeRoot) {
    $key = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MathType.exe' -ErrorAction SilentlyContinue
    if ($key) { $MathTypeRoot = Split-Path $key.'(default)' }
}
if (-not $MathTypeRoot) { throw 'Desktop MathType was not found. MathType 7 for Windows is required.' }
$wordKey = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE' -ErrorAction SilentlyContinue
if (-not $wordKey) { throw 'Microsoft Word was not found.' }
$wordPath = [string]$wordKey.'(default)'
$bitness = Get-PeBitness $wordPath
$wll = Join-Path $MathTypeRoot "MathPage\$bitness\MathPage.wll"
$blank = Join-Path $MathTypeRoot 'Office Support\BlankEqn.doc'
$template = Join-Path $MathTypeRoot "Office Support\$bitness\MathType Commands 2016.dotm"
foreach ($path in @($wll, $blank, $template)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing MathType component: $path" }
}
if ($Probe) {
    [pscustomobject]@{
        mathTypeRoot = $MathTypeRoot
        wordPath = $wordPath
        wordBits = $bitness
        mathTypeCommands = $template
        wordRegistered = ($null -ne [type]::GetTypeFromProgID('Word.Application'))
        mathTypeRegistered = ($null -ne [type]::GetTypeFromProgID('Equation.DSMT4'))
    } | ConvertTo-Json
    return
}

$word = $null
$document = $null
$loadedAddIn = $null
$completed = [Collections.Generic.List[object]]::new()
$numberedCount = @($items | Where-Object mode -eq 'right-numbered').Count
$hadDeferredUpdates = $false
$originalDeferredUpdates = $null
$tempId = [Guid]::NewGuid().ToString('N')
$outputDirectory = Split-Path $outputPath
$tempDocx = Join-Path $outputDirectory ".$([IO.Path]::GetFileNameWithoutExtension($outputPath)).$tempId.docx"
$tempReport = Join-Path $outputDirectory ".$([IO.Path]::GetFileNameWithoutExtension($reportPath)).$tempId.json"
$tempPdf = $null
if ($pdfPath) {
    $tempPdf = Join-Path (Split-Path $pdfPath) ".$([IO.Path]::GetFileNameWithoutExtension($pdfPath)).$tempId.pdf"
}
try {
    # Use a separate Word instance and never attach to the user's open documents.
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.Options.SaveNormalPrompt = $false
    $word.Options.UpdateLinksAtOpen = $false
    $alreadyLoaded = @($word.Templates | Where-Object { $_.FullName -eq $template }).Count -gt 0
    if (-not $alreadyLoaded) { $loadedAddIn = $word.AddIns.Add($template, $true) }

    $document = $word.Documents.Open($inputPath, $false, $true, $false)
    $document.Activate()
    $word.ActiveWindow.View.Type = 3
    if ($document.ProtectionType -ne -1) { throw 'Remove document protection before inserting equations.' }
    if ($document.TrackRevisions) { throw 'Turn off Track Changes and resolve revisions before inserting equations.' }

    $ranges = [Collections.Generic.List[object]]::new()
    foreach ($item in $items) {
        $range = Find-Anchor $document $item
        if ($range.Fields.Count -gt 0 -or $range.InlineShapes.Count -gt 0) {
            throw 'An anchor or bookmark must not contain an existing field or OLE object.'
        }
        $paragraphText = ([string]$range.Paragraphs.Item(1).Range.Text).Trim([char[]]" `t`r`n")
        $anchorText = ([string]$range.Text).Trim([char[]]" `t`r`n")
        if ($item.mode -ne 'inline') {
            if ($range.Information(12)) { throw 'Display equations inside tables are not supported by this version.' }
            if ($paragraphText -ne $anchorText) {
                throw 'A display or right-numbered locator must occupy its own paragraph.'
            }
        }
        foreach ($prior in $ranges) {
            if ($range.Start -lt $prior.End -and $range.End -gt $prior.Start) { throw 'Equation locators overlap.' }
        }
        $ranges.Add($range)
    }

    if ($numberedCount -gt 0) {
        if (-not ('WordComBridge' -as [type])) { Add-Type -Path (Join-Path $PSScriptRoot 'WordComBridge.cs') }
        $hadDeferredUpdates = [WordComBridge]::HasCustomProperty($document, 'MTDeferFieldUpdate')
        if ($hadDeferredUpdates) {
            $originalDeferredUpdates = [WordComBridge]::GetCustomProperty($document, 'MTDeferFieldUpdate')
        }
        # MathType's insertion macros otherwise update every field in every story range.
        [WordComBridge]::SetCustomProperty($document, 'MTDeferFieldUpdate', '1')
        $hasSection = $false
        foreach ($field in $document.Fields) {
            if ($field.Code.Text -match 'MTEditEquationSection') { $hasSection = $true; break }
        }
        if (-not $hasSection) {
            $document.Range(0, 0).Select()
            Invoke-MathTypeMacro $word 'MTCommand_InsertNextChapter'
        }
    }

    for ($index = 0; $index -lt $items.Count; $index++) {
        $item = $items[$index]
        $range = Find-Anchor $document $item
        $location = if ($null -ne $item.PSObject.Properties['anchor']) { [string]$item.anchor } else { [string]$item.bookmark }
        Write-Host "Inserting $($index + 1)/$($items.Count): $location [$($item.mode)]"

        $latex = Get-NormalizedLatex ([string]$item.latex)
        $tex = if ($item.mode -eq 'inline') { '$' + $latex + '$' } else { '\[' + $latex + '\]' }
        $start = [int]$range.Start
        $beforeCount = [int]$document.InlineShapes.Count
        $precedingShapeCount = 0
        foreach ($existingShape in $document.InlineShapes) {
            if ([int]$existingShape.Range.Start -lt $start) { $precedingShapeCount++ }
        }
        $standaloneInline = $false
        if ($item.mode -eq 'inline') {
            $paragraphText = ([string]$range.Paragraphs.Item(1).Range.Text).Trim([char[]]" `t`r`n")
            $anchorText = ([string]$range.Text).Trim([char[]]" `t`r`n")
            $standaloneInline = ($paragraphText -eq $anchorText)
        }
        if ($standaloneInline) {
            # Surround the selected TeX temporarily so MathType does not offer an interactive
            # "convert to display" dialog for an inline object in an otherwise empty paragraph.
            $range.Text = 'x' + $tex + 'x'
            $texRange = $document.Range($start + 1, $start + 1 + $tex.Length)
        }
        else {
            $range.Text = $tex
            $texRange = $document.Range($start, $start + $tex.Length)
        }
        $texRange.Select()
        Invoke-MathTypeMacro $word 'MTCommand_TeXToggle'

        if ([int]$document.InlineShapes.Count -ne $beforeCount + 1) {
            throw "MathType did not replace '$location' with exactly one equation object. Check its supported TeX syntax."
        }
        $shape = Find-InsertedEquation $document $start ($precedingShapeCount + 1)
        if ($standaloneInline) {
            $rightSentinel = $document.Range($shape.Range.End, $shape.Range.End + 1)
            $leftSentinel = $document.Range($shape.Range.Start - 1, $shape.Range.Start)
            if ($rightSentinel.Text -ne 'x' -or $leftSentinel.Text -ne 'x') {
                throw 'Could not remove temporary inline-layout sentinels.'
            }
            $null = $rightSentinel.Delete()
            $null = $leftSentinel.Delete()
        }
        if ($shape.Width -le 0 -or $shape.Height -le 0) { throw 'MathType returned an equation with empty dimensions.' }

        $numberFields = 0
        if ($item.mode -eq 'right-numbered') {
            $tail = $shape.Range.Duplicate
            $tail.Collapse(0)
            $tail.Select()
            $word.Selection.TypeText("`t")
            Invoke-MathTypeMacro $word 'MTCommand_InsertEqnNum'
            $paragraph = $shape.Range.Paragraphs.Item(1).Range
            $codes = @($paragraph.Fields | ForEach-Object { $_.Code.Text })
            if (-not ($codes -match 'MACROBUTTON MTPlaceRef') -or -not ($codes -match 'SEQ MTEqn')) {
                throw 'MathType native equation-number fields were not created.'
            }
            $null = $paragraph.Fields.Update()
            $null = $paragraph.Fields.Update()
            $numberFields = $paragraph.Fields.Count
        }

        if ($null -ne $item.PSObject.Properties['bookmark']) {
            $null = $document.Bookmarks.Add([string]$item.bookmark, $shape.Range)
        }
        $completed.Add([pscustomobject]@{
            location = $location
            mode = [string]$item.mode
            progId = [string]$shape.OLEFormat.ProgID
            widthPoints = [Math]::Round([double]$shape.Width, 2)
            heightPoints = [Math]::Round([double]$shape.Height, 2)
            nativeNumberFields = $numberFields
        })
    }

    if ($numberedCount -gt 0) {
        if ($hadDeferredUpdates) {
            [WordComBridge]::SetCustomProperty($document, 'MTDeferFieldUpdate', [string]$originalDeferredUpdates)
        }
        else {
            [WordComBridge]::DeleteCustomProperty($document, 'MTDeferFieldUpdate')
        }
    }
    $document.Repaginate()
    $null = [IO.Directory]::CreateDirectory($outputDirectory)
    if ($pdfPath) { $null = [IO.Directory]::CreateDirectory((Split-Path $pdfPath)) }
    $document.SaveAs2($tempDocx, 16)

    if ($pdfPath) {
        $document.ExportAsFixedFormat($tempPdf, 17)
    }

    [pscustomobject]@{
        input = $inputPath
        output = $outputPath
        pdf = $pdfPath
        equations = $completed
        equationRepresentation = 'MathType Equation.DSMT4 OLE'
        numbering = 'MathType native MTPlaceRef and MTEqn fields'
        unrelatedFieldUpdatesSuppressed = ($numberedCount -gt 0)
        visualReview = 'pending'
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $tempReport -Encoding UTF8

    # Close the staged DOCX before publishing it. The original remains untouched on any failure.
    $document.Close(0)
    $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    $document = $null
    $published = [Collections.Generic.List[string]]::new()
    try {
        if ($pdfPath) { Move-Item -LiteralPath $tempPdf -Destination $pdfPath; $published.Add($pdfPath) }
        Move-Item -LiteralPath $tempReport -Destination $reportPath
        $published.Add($reportPath)
        Move-Item -LiteralPath $tempDocx -Destination $outputPath
        $published.Add($outputPath)
    }
    catch {
        foreach ($path in $published) {
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
        }
        throw
    }
    Write-Host "Saved: $outputPath"
}
catch {
    Write-Host $_.ScriptStackTrace
    throw
}
finally {
    if ($document) {
        $document.Close(0)
        $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($loadedAddIn) { $loadedAddIn.Installed = $false }
    if ($word) {
        $word.NormalTemplate.Saved = $true
        $word.Quit(0)
        $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    foreach ($path in @($tempDocx, $tempReport, $tempPdf)) {
        if ($path -and (Test-Path -LiteralPath $path)) { Remove-Item -LiteralPath $path -Force }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
