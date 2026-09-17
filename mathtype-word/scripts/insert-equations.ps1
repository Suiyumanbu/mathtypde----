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

function Get-ItemOperation($Item) {
    if ($null -ne $Item.PSObject.Properties['operation']) { return [string]$Item.operation }
    return 'insert-latex'
}

function Get-ItemLocation($Item) {
    if ($null -ne $Item.PSObject.Properties['anchor']) { return [string]$Item.anchor }
    if ($null -ne $Item.PSObject.Properties['bookmark']) { return [string]$Item.bookmark }
    return "MathType equation #$([int]$Item.equationIndex)"
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

function Get-FieldCodes($Range) {
    $codes = [Collections.Generic.List[string]]::new()
    $fieldCount = [int]$Range.Fields.Count
    for ($fieldIndex = 1; $fieldIndex -le $fieldCount; $fieldIndex++) {
        $codes.Add([string]$Range.Fields.Item($fieldIndex).Code.Text)
    }
    return $codes.ToArray()
}

function Add-NativeEquationNumber($Word, $Shape) {
    $paragraph = $Shape.Range.Paragraphs.Item(1).Range
    $existingCodes = @(Get-FieldCodes $paragraph)
    if (($existingCodes -match 'MACROBUTTON MTPlaceRef') -or ($existingCodes -match 'SEQ MTEqn')) {
        throw 'The target equation already has a MathType native number.'
    }

    $tail = $Shape.Range.Duplicate
    $tail.Collapse(0)
    $tail.Select()
    $Word.Selection.TypeText("`t")
    Invoke-MathTypeMacro $Word 'MTCommand_InsertEqnNum'

    $codes = @(Get-FieldCodes $paragraph)
    if (-not ($codes -match 'MACROBUTTON MTPlaceRef') -or -not ($codes -match 'SEQ MTEqn')) {
        throw 'MathType native equation-number fields were not created.'
    }
    return [int]$paragraph.Fields.Count
}

function Update-NativeEquationNumbers($Document, [int]$Start) {
    $searchRange = $Document.Range($Start, $Document.Content.End)
    $numberFields = [Collections.Generic.List[object]]::new()
    $paragraphStarts = [Collections.Generic.HashSet[int]]::new()
    $fieldCount = [int]$searchRange.Fields.Count
    for ($fieldIndex = 1; $fieldIndex -le $fieldCount; $fieldIndex++) {
        $field = $searchRange.Fields.Item($fieldIndex)
        if ([string]$field.Code.Text -match 'MACROBUTTON MTPlaceRef') {
            $null = $paragraphStarts.Add([int]$field.Code.Paragraphs.Item(1).Range.Start)
            $numberFields.Add($field)
        }
    }
    foreach ($numberField in $numberFields) {
        # MTPlaceRef contains nested sequence fields. Update the inner sequence
        # once before its enclosing display, without touching unrelated fields.
        # Range.Fields may include an adjacent field at a boundary. Update the
        # actual native field object and its own nested collection instead.
        $null = $numberField.Code.Fields.Update()
        $null = $numberField.Update()
    }
    return $paragraphStarts.Count
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
    $itemAllowed = @('anchor', 'bookmark', 'equationIndex', 'latex', 'mode', 'operation')
    foreach ($property in $item.PSObject.Properties.Name) {
        if ($property -notin $itemAllowed) { throw "Unknown equation property: $property" }
    }
    $operation = Get-ItemOperation $item
    if ($operation -notin @('insert-latex', 'number-existing')) {
        throw "Unsupported equation operation: $operation"
    }
    $hasAnchor = $null -ne $item.PSObject.Properties['anchor']
    $hasBookmark = $null -ne $item.PSObject.Properties['bookmark']
    $hasEquationIndex = $null -ne $item.PSObject.Properties['equationIndex']

    if ($operation -eq 'insert-latex') {
        if ($hasAnchor -eq $hasBookmark -or $hasEquationIndex) {
            throw 'A LaTeX insertion needs exactly one locator: anchor or bookmark.'
        }
        if ($null -eq $item.PSObject.Properties['mode'] -or $item.mode -notin @('inline', 'display', 'right-numbered')) {
            throw 'mode must be inline, display, or right-numbered.'
        }
        if ($null -eq $item.PSObject.Properties['latex'] -or -not ($item.latex -is [string])) {
            throw 'Each LaTeX insertion needs a LaTeX string.'
        }
        $latex = Get-NormalizedLatex ([string]$item.latex)
        if ([string]::IsNullOrWhiteSpace($latex)) { throw 'LaTeX cannot be empty.' }
        if ($latex -match '\\(documentclass|usepackage|newcommand|renewcommand|def|input|include)\b') {
            throw 'Pass a self-contained equation body, without a LaTeX document or custom macro definitions.'
        }
    }
    else {
        if ($hasAnchor -or ($hasBookmark -eq $hasEquationIndex)) {
            throw 'number-existing needs exactly one locator: bookmark or equationIndex.'
        }
        if ($null -ne $item.PSObject.Properties['latex'] -or $null -ne $item.PSObject.Properties['mode']) {
            throw 'number-existing does not accept latex or mode.'
        }
        if ($hasEquationIndex) {
            $parsedIndex = 0
            if (-not [int]::TryParse([string]$item.equationIndex, [ref]$parsedIndex) -or $parsedIndex -lt 1) {
                throw 'equationIndex must be a positive integer.'
            }
        }
    }

    $locator = Get-ItemLocation $item
    if ([string]::IsNullOrWhiteSpace($locator) -or $locator -match '[\r\n\^]') {
        throw 'Locator is empty or contains unsupported characters.'
    }
    $locatorKey = if ($hasAnchor) { "anchor:$locator" } elseif ($hasBookmark) { "bookmark:$locator" } else { "equationIndex:$parsedIndex" }
    if ($seen.ContainsKey($locatorKey)) { throw "Duplicate locator: $locator" }
    $seen[$locatorKey] = $true
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
$numberedCount = @($items | Where-Object {
    $operation = Get-ItemOperation $_
    if ($operation -eq 'number-existing') { return $true }
    return [string]$_.mode -eq 'right-numbered'
}).Count
$hadDeferredUpdates = $false
$originalDeferredUpdates = $null
$minimumNumberStart = $null
$updatedNumberParagraphs = 0
$timings = [ordered]@{}
$totalTimer = [Diagnostics.Stopwatch]::StartNew()
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
    $phaseTimer = [Diagnostics.Stopwatch]::StartNew()
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.Options.SaveNormalPrompt = $false
    $word.Options.UpdateLinksAtOpen = $false
    $alreadyLoaded = @($word.Templates | Where-Object { $_.FullName -eq $template }).Count -gt 0
    if (-not $alreadyLoaded) { $loadedAddIn = $word.AddIns.Add($template, $true) }
    $phaseTimer.Stop()
    $timings.wordStartupAndAddInMilliseconds = [Math]::Round($phaseTimer.Elapsed.TotalMilliseconds, 1)

    $phaseTimer.Restart()
    $document = $word.Documents.Open($inputPath, $false, $true, $false)
    $document.Activate()
    $word.ActiveWindow.View.Type = 3
    if ($document.ProtectionType -ne -1) { throw 'Remove document protection before inserting equations.' }
    if ($document.TrackRevisions) { throw 'Turn off Track Changes and resolve revisions before inserting equations.' }
    $phaseTimer.Stop()
    $timings.openDocumentMilliseconds = [Math]::Round($phaseTimer.Elapsed.TotalMilliseconds, 1)

    if ($numberedCount -gt 0) {
        $phaseTimer.Restart()
        if (-not ('WordComBridge' -as [type])) { Add-Type -Path (Join-Path $PSScriptRoot 'WordComBridge.cs') }
        $hadDeferredUpdates = [WordComBridge]::TryGetCustomProperty(
            $document,
            'MTDeferFieldUpdate',
            [ref]$originalDeferredUpdates
        )
        # MathType's insertion macros otherwise update every field in every story range.
        [WordComBridge]::SetCustomProperty($document, 'MTDeferFieldUpdate', '1')
        $sectionValue = $null
        $hasSection = [WordComBridge]::TryGetCustomProperty(
            $document,
            'MTEquationSection',
            [ref]$sectionValue
        )
        if (-not $hasSection) {
            for ($fieldIndex = 1; $fieldIndex -le $document.Fields.Count; $fieldIndex++) {
                if ($document.Fields.Item($fieldIndex).Code.Text -match 'MTEditEquationSection') {
                    $hasSection = $true
                    break
                }
            }
        }
        if (-not $hasSection) {
            $document.Range(0, 0).Select()
            Invoke-MathTypeMacro $word 'MTCommand_InsertNextChapter'
        }
        $phaseTimer.Stop()
        $timings.numberingSetupMilliseconds = [Math]::Round($phaseTimer.Elapsed.TotalMilliseconds, 1)
    }
    else {
        $timings.numberingSetupMilliseconds = 0
    }

    $phaseTimer.Restart()
    $mathTypeShapes = [Collections.Generic.List[object]]::new()
    $initialInlineShapeStarts = [Collections.Generic.List[int]]::new()
    $needsInitialShapePositions = @($items | Where-Object {
        (Get-ItemOperation $_) -eq 'insert-latex'
    }).Count -gt 0
    $needsEquationIndex = @($items | Where-Object {
        (Get-ItemOperation $_) -eq 'number-existing' -and
        $null -ne $_.PSObject.Properties['equationIndex']
    }).Count -gt 0
    if ($needsInitialShapePositions -or $needsEquationIndex) {
        $initialShapeCount = [int]$document.InlineShapes.Count
        for ($shapeIndex = 1; $shapeIndex -le $initialShapeCount; $shapeIndex++) {
            $candidate = $document.InlineShapes.Item($shapeIndex)
            if ($needsInitialShapePositions) {
                $initialInlineShapeStarts.Add([int]$candidate.Range.Start)
            }
            if ($needsEquationIndex) {
                $candidateProgId = $null
                try { $candidateProgId = [string]$candidate.OLEFormat.ProgID } catch { continue }
                if ($candidateProgId -eq 'Equation.DSMT4') {
                    $mathTypeShapes.Add($candidate)
                }
            }
        }
    }

    $plan = [Collections.Generic.List[object]]::new()
    foreach ($item in $items) {
        $operation = Get-ItemOperation $item
        $location = Get-ItemLocation $item
        $existingShape = $null
        if ($operation -eq 'number-existing') {
            if ($null -ne $item.PSObject.Properties['bookmark']) {
                $range = Find-Anchor $document $item
                if ($range.InlineShapes.Count -ne 1) {
                    throw "Bookmark '$location' must contain exactly one existing MathType equation."
                }
                $existingShape = $range.InlineShapes.Item(1)
            }
            else {
                $equationIndex = [int]$item.equationIndex
                if ($equationIndex -gt $mathTypeShapes.Count) {
                    throw "MathType equation index $equationIndex is out of range; found $($mathTypeShapes.Count)."
                }
                $existingShape = $mathTypeShapes[$equationIndex - 1]
                $range = $existingShape.Range.Duplicate
            }
            $existingProgId = $null
            try { $existingProgId = [string]$existingShape.OLEFormat.ProgID } catch {}
            if ($existingProgId -ne 'Equation.DSMT4') {
                throw "The target '$location' is not an Equation.DSMT4 object."
            }
            $paragraph = $existingShape.Range.Paragraphs.Item(1).Range
            if ($paragraph.Information(12)) { throw 'Right-numbered equations inside tables are not supported by this version.' }
            if ($paragraph.InlineShapes.Count -ne 1) {
                throw 'An existing equation must be the only inline object in its paragraph.'
            }
            $codes = @(Get-FieldCodes $paragraph)
            if (($codes -match 'MACROBUTTON MTPlaceRef') -or ($codes -match 'SEQ MTEqn')) {
                throw "The target '$location' is already right-numbered."
            }
            $ordinaryText = ([string]$paragraph.Text).Replace([string][char]1, '').Trim([char[]]" `t`r`n")
            if (-not [string]::IsNullOrWhiteSpace($ordinaryText)) {
                throw 'An existing equation must occupy its own paragraph before it can be numbered.'
            }
            $mode = 'right-numbered'
        }
        else {
            $range = Find-Anchor $document $item
            if ($range.Fields.Count -gt 0 -or $range.InlineShapes.Count -gt 0) {
                throw 'An insertion anchor or bookmark must not contain an existing field or OLE object.'
            }
            $paragraphText = ([string]$range.Paragraphs.Item(1).Range.Text).Trim([char[]]" `t`r`n")
            $anchorText = ([string]$range.Text).Trim([char[]]" `t`r`n")
            $mode = [string]$item.mode
            if ($mode -ne 'inline') {
                if ($range.Information(12)) { throw 'Display equations inside tables are not supported by this version.' }
                if ($paragraphText -ne $anchorText) {
                    throw 'A display or right-numbered locator must occupy its own paragraph.'
                }
            }
        }
        $plan.Add([pscustomobject]@{
            Item = $item
            Operation = $operation
            Location = $location
            Mode = $mode
            Range = $range
            ExistingShape = $existingShape
            InitialShapesBefore = 0
            Start = [int]$range.Start
            End = [int]$range.End
        })
    }
    $orderedPlan = @($plan | Sort-Object Start, End)
    # Both lists are in document order. Advance once instead of scanning all cached
    # positions for each locator; no per-equation full-document COM traversal.
    $initialShapeCursor = 0
    foreach ($target in $orderedPlan) {
        while ($initialShapeCursor -lt $initialInlineShapeStarts.Count -and
               $initialInlineShapeStarts[$initialShapeCursor] -lt $target.Start) {
            $initialShapeCursor++
        }
        $target.InitialShapesBefore = $initialShapeCursor
    }
    for ($planIndex = 1; $planIndex -lt $orderedPlan.Count; $planIndex++) {
        $prior = $orderedPlan[$planIndex - 1]
        $current = $orderedPlan[$planIndex]
        $overlap = $current.Start -lt $prior.End -and $current.End -gt $prior.Start
        $sameRange = $current.Start -eq $prior.Start -and $current.End -eq $prior.End
        if ($overlap -or $sameRange) { throw 'Equation locators overlap or resolve to the same range.' }
    }
    $phaseTimer.Stop()
    $timings.locateAndValidateMilliseconds = [Math]::Round($phaseTimer.Elapsed.TotalMilliseconds, 1)

    $insertedShapeCount = 0
    for ($index = 0; $index -lt $orderedPlan.Count; $index++) {
        $target = $orderedPlan[$index]
        $item = $target.Item
        $location = [string]$target.Location
        $mode = [string]$target.Mode
        $operation = [string]$target.Operation
        $itemTimer = [Diagnostics.Stopwatch]::StartNew()
        $conversionMilliseconds = 0
        $numberingMilliseconds = 0
        Write-Host "Processing $($index + 1)/$($orderedPlan.Count): $location [$operation -> $mode]"

        if ($operation -eq 'insert-latex') {
            $conversionTimer = [Diagnostics.Stopwatch]::StartNew()
            $range = $target.Range.Duplicate
            $latex = Get-NormalizedLatex ([string]$item.latex)
            $tex = if ($mode -eq 'inline') { '$' + $latex + '$' } else { '\[' + $latex + '\]' }
            $start = [int]$range.Start
            $beforeCount = [int]$document.InlineShapes.Count
            $standaloneInline = $false
            if ($mode -eq 'inline') {
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
            $expectedShapeIndex = [int]$target.InitialShapesBefore + $insertedShapeCount + 1
            $shape = Find-InsertedEquation $document $start $expectedShapeIndex
            $insertedShapeCount++
            if ($standaloneInline) {
                $rightSentinel = $document.Range($shape.Range.End, $shape.Range.End + 1)
                $leftSentinel = $document.Range($shape.Range.Start - 1, $shape.Range.Start)
                if ($rightSentinel.Text -ne 'x' -or $leftSentinel.Text -ne 'x') {
                    throw 'Could not remove temporary inline-layout sentinels.'
                }
                $null = $rightSentinel.Delete()
                $null = $leftSentinel.Delete()
            }
            $conversionTimer.Stop()
            $conversionMilliseconds = [Math]::Round($conversionTimer.Elapsed.TotalMilliseconds, 1)
        }
        else {
            $shape = $target.ExistingShape
        }

        $progId = [string]$shape.OLEFormat.ProgID
        $width = [double]$shape.Width
        $height = [double]$shape.Height
        if ($progId -ne 'Equation.DSMT4' -or $width -le 0 -or $height -le 0) {
            throw 'MathType returned an invalid or empty equation object.'
        }

        $numberFields = 0
        if ($mode -eq 'right-numbered') {
            $numberingTimer = [Diagnostics.Stopwatch]::StartNew()
            $numberFields = Add-NativeEquationNumber $word $shape
            $numberStart = [int]$shape.Range.Start
            if ($null -eq $minimumNumberStart -or $numberStart -lt $minimumNumberStart) {
                $minimumNumberStart = $numberStart
            }
            $numberingTimer.Stop()
            $numberingMilliseconds = [Math]::Round($numberingTimer.Elapsed.TotalMilliseconds, 1)
        }

        if ($null -ne $item.PSObject.Properties['bookmark']) {
            $null = $document.Bookmarks.Add([string]$item.bookmark, $shape.Range)
        }
        $itemTimer.Stop()
        $completed.Add([pscustomobject]@{
            location = $location
            operation = $operation
            source = if ($operation -eq 'number-existing') { 'existing' } else { 'latex' }
            mode = $mode
            progId = $progId
            widthPoints = [Math]::Round($width, 2)
            heightPoints = [Math]::Round($height, 2)
            nativeNumberFields = $numberFields
            conversionMilliseconds = $conversionMilliseconds
            numberingMilliseconds = $numberingMilliseconds
            durationMilliseconds = [Math]::Round($itemTimer.Elapsed.TotalMilliseconds, 1)
        })
    }

    if ($null -ne $minimumNumberStart) {
        $phaseTimer.Restart()
        $updatedNumberParagraphs = Update-NativeEquationNumbers $document $minimumNumberStart
        $phaseTimer.Stop()
        $timings.numberFieldUpdateMilliseconds = [Math]::Round($phaseTimer.Elapsed.TotalMilliseconds, 1)
    }
    else {
        $timings.numberFieldUpdateMilliseconds = 0
    }

    if ($numberedCount -gt 0) {
        if ($hadDeferredUpdates) {
            [WordComBridge]::SetCustomProperty($document, 'MTDeferFieldUpdate', [string]$originalDeferredUpdates)
        }
        else {
            [WordComBridge]::DeleteCustomProperty($document, 'MTDeferFieldUpdate')
        }
    }
    $null = [IO.Directory]::CreateDirectory($outputDirectory)
    if ($pdfPath) { $null = [IO.Directory]::CreateDirectory((Split-Path $pdfPath)) }

    $phaseTimer.Restart()
    $document.SaveAs2($tempDocx, 16)
    $phaseTimer.Stop()
    $timings.saveDocxMilliseconds = [Math]::Round($phaseTimer.Elapsed.TotalMilliseconds, 1)

    if ($pdfPath) {
        $phaseTimer.Restart()
        $document.ExportAsFixedFormat($tempPdf, 17)
        $phaseTimer.Stop()
        $timings.exportPdfMilliseconds = [Math]::Round($phaseTimer.Elapsed.TotalMilliseconds, 1)
    }
    else {
        $timings.exportPdfMilliseconds = 0
    }
    $totalTimer.Stop()
    $timings.processingMilliseconds = [Math]::Round($totalTimer.Elapsed.TotalMilliseconds, 1)

    [pscustomobject]@{
        input = $inputPath
        output = $outputPath
        pdf = $pdfPath
        equations = $completed
        equationRepresentation = 'MathType Equation.DSMT4 OLE'
        numbering = 'MathType native MTPlaceRef and MTEqn fields'
        unrelatedFieldUpdatesSuppressed = ($numberedCount -gt 0)
        nativeNumberParagraphsUpdated = $updatedNumberParagraphs
        timings = [pscustomobject]$timings
        visualReview = if ($pdfPath) { 'pending' } else { 'not-requested' }
    } | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $tempReport -Encoding UTF8

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
