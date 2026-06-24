[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputPath,
    [string]$OutputPath,
    [ValidateRange(-1, 100000)]
    [int]$ExpectedInline = -1,
    [ValidateRange(-1, 100000)]
    [int]$ExpectedDisplay = -1,
    [string]$MacroName = "MTCommand_OnTexToggle",
    [ValidateRange(0, 3000)]
    [int]$DelayMilliseconds = 100,
    [switch]$AddEquationNumbers,
    [switch]$SkipTextIntegrityCheck,
    [switch]$Force
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$wdDoNotSaveChanges = 0
$wdFieldEmpty = -1
$wdCollapseEnd = 0
$wdCollapseStart = 1
$wdAlignParagraphLeft = 0
$wdAlignParagraphCenter = 1
$wdAlignTabCenter = 1
$wdAlignTabRight = 2
$wdTabLeaderSpaces = 0
$wdLineSpaceAtLeast = 3
$wdWithInTable = 12
$rpcRejected = @(-2147418111, -2147417846)
$script:NumericFallbackConversions = 0
$script:PrimeFallbackConversions = 0
$script:CommaListFallbackConversions = 0
$script:TableInlineDialogWatchers = 0
$script:DisplayParagraphsReformatted = 0
$script:DisplayEmptyParagraphsRemoved = 0

function Write-UserActionError {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Reason,
        [string[]]$Actions = @()
    )
    Write-Host ""
    Write-Host "================ MathType conversion stopped ================" -ForegroundColor Red
    Write-Host $Title -ForegroundColor Red
    Write-Host "Reason : $Reason" -ForegroundColor Yellow
    if ($Actions.Count -gt 0) {
        Write-Host "Action :" -ForegroundColor Cyan
        foreach ($action in $Actions) {
            Write-Host "  - $action" -ForegroundColor Cyan
        }
    }
    Write-Host "==============================================================" -ForegroundColor Red
    Write-Host ""
}

$displayRegex = [regex]::new('\\\[(?<content>[^\r\a]+?)\\\]', [System.Text.RegularExpressions.RegexOptions]::Singleline)
$inlineRegex = [regex]::new('(?<!\$)\$(?!\$)(?<content>[^\r\a$]+?)(?<!\$)\$(?!\$)', [System.Text.RegularExpressions.RegexOptions]::Singleline)

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Operation,
        [int]$Retries = 12,
        [int]$Delay = 350
    )
    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        try {
            return & $Operation
        } catch [System.Runtime.InteropServices.COMException] {
            if ($rpcRejected -notcontains $_.Exception.HResult -or $attempt -eq $Retries) {
                throw
            }
            Start-Sleep -Milliseconds $Delay
        }
    }
}

function Get-RunningWord {
    try {
        $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    } catch {
        Write-UserActionError `
            -Title "Microsoft Word is not open or is not visible to this PowerShell session." `
            -Reason "The converter attaches to an already-running Word COM instance; it does not start Word by itself." `
            -Actions @(
                "Open Microsoft Word manually.",
                "Create or keep one blank Word document open.",
                "Run PowerShell and Word at the same privilege level; do not run one as Administrator and the other normally.",
                "Verify MathType Toggle TeX works manually, then rerun this script."
            )
        throw "No running Word instance was found. Open Word manually, verify MathType Toggle TeX works, leave Word open, and rerun."
    }
    if ($word.Name -notlike "*Microsoft Word*" -or $word.Path -match "WPS|Kingsoft") {
        Write-UserActionError `
            -Title "The active Word.Application is not Microsoft Word." `
            -Reason "Detected $($word.Name) at $($word.Path)." `
            -Actions @("Close WPS/Kingsoft Word-compatible editors.", "Open Microsoft Word and rerun.")
        throw "The running Word.Application is not Microsoft Word: $($word.Name) at $($word.Path)"
    }
    $word.Visible = $true
    return $word
}

function Invoke-MathTypeToggle {
    param([Parameter(Mandatory = $true)]$Word)
    try {
        $null = $Word.GetType().InvokeMember(
            "Run",
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null,
            $Word,
            @($MacroName, $null)
        )
    } catch {
        $message = $_.Exception.Message
        if ($_.Exception.InnerException) {
            $message = $_.Exception.InnerException.Message
        }
        Write-UserActionError `
            -Title "MathType Toggle TeX could not run." `
            -Reason $message `
            -Actions @(
                "Confirm MathType 6 is licensed and activated in this Windows account.",
                "In Word, manually select a test formula such as `$x^2`$ and run MathType Toggle TeX.",
                "If Word shows a MathType license/demo dialog, close it and fix activation before rerunning.",
                "Keep Word open and do not interact with it while conversion is running."
            )
        throw "MathType Toggle TeX failed: $message"
    }
}

function Test-MathTypeInlineShape {
    param($Shape)
    try {
        return $Shape.OLEFormat.ProgID -eq "Equation.DSMT4"
    } catch {
        return $false
    }
}

function Get-MathTypeInlineShapes {
    param([Parameter(Mandatory = $true)]$Document)
    $result = [System.Collections.Generic.List[object]]::new()
    for ($i = 1; $i -le $Document.InlineShapes.Count; $i++) {
        $shape = $Document.InlineShapes.Item($i)
        if (Test-MathTypeInlineShape -Shape $shape) {
            $result.Add($shape)
        }
    }
    return @($result)
}

function Get-MathTypeObjectCount {
    param([Parameter(Mandatory = $true)]$Document)
    $count = 0
    for ($i = 1; $i -le $Document.InlineShapes.Count; $i++) {
        $shape = $Document.InlineShapes.Item($i)
        try {
            if (Test-MathTypeInlineShape -Shape $shape) { $count++ }
        } finally {
            Release-ComObject $shape
        }
    }
    return $count
}

function Clear-TemporaryDocument {
    param([Parameter(Mandatory = $true)]$Document)
    $range = $Document.Content
    $range.Delete() | Out-Null
    Release-ComObject $range
}

function Put-FormulaInTemporaryDocument {
    param(
        [Parameter(Mandatory = $true)]$Word,
        [Parameter(Mandatory = $true)]$TempDocument,
        [Parameter(Mandatory = $true)][string]$Literal
    )
    Clear-TemporaryDocument -Document $TempDocument
    $TempDocument.Activate()
    $range = $TempDocument.Range(0, 0)
    $range.Text = $Literal
    $range.SetRange(0, $Literal.Length)
    $range.Select()
    Release-ComObject $range

    Invoke-WithRetry -Operation { Invoke-MathTypeToggle -Word $Word } | Out-Null
    if ($DelayMilliseconds -gt 0) {
        Start-Sleep -Milliseconds $DelayMilliseconds
    }

    $shapes = @(Get-MathTypeInlineShapes -Document $TempDocument)
    if ($shapes.Count -ne 1) {
        throw "MathType produced $($shapes.Count) Equation.DSMT4 objects for formula: $Literal"
    }
    return $shapes[0]
}

function ConvertTo-AsciiPrimeText {
    param([Parameter(Mandatory = $true)][string]$Text)
    # Avoid non-ASCII literals in regex source. Windows PowerShell 5 may read a
    # UTF-8 script through the active ANSI code page when the file has no BOM.
    return $Text.Replace([char]0x2018, [char]0x27).Replace([char]0x2019, [char]0x27).Replace([char]0x2032, [char]0x27)
}

function Test-FormulaRangeInTable {
    param([Parameter(Mandatory = $true)]$FormulaRange)
    try { return [bool]$FormulaRange.Information($wdWithInTable) } catch { return $false }
}

function Start-InlineStyleDialogWatcher {
    $watcherCode = @'
$ErrorActionPreference = "SilentlyContinue"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$deadline = [DateTime]::UtcNow.AddSeconds(20)
$root = [System.Windows.Automation.AutomationElement]::RootElement
$buttonCondition = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
    [System.Windows.Automation.ControlType]::Button
)
while ([DateTime]::UtcNow -lt $deadline) {
    $windows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($window in $windows) {
        $buttons = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $buttonCondition)
        foreach ($button in $buttons) {
            $name = $button.Current.Name
            if ($name -match 'Create\s+Inline\s+Style\s+Equation|Inline\s+Style\s+Equation') {
                $pattern = $button.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                ([System.Windows.Automation.InvokePattern]$pattern).Invoke()
                exit 0
            }
        }
    }
    Start-Sleep -Milliseconds 100
}
exit 2
'@
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($watcherCode))
    $powerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $process = Start-Process -FilePath $powerShellExe -ArgumentList @(
        "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-EncodedCommand", $encoded
    ) -PassThru -WindowStyle Hidden
    Start-Sleep -Milliseconds 150
    return $process
}

function Stop-InlineStyleDialogWatcher {
    param($Process)
    if ($null -eq $Process) { return }
    try {
        if (-not $Process.HasExited) {
            $Process.Kill()
            [void]$Process.WaitForExit(1000)
        }
    } catch {
    } finally {
        $Process.Dispose()
    }
}

function Invoke-MathTypeToggleForFormula {
    param(
        [Parameter(Mandatory = $true)]$Word,
        [bool]$WatchInlineStyleDialog
    )
    $watcher = $null
    try {
        if ($WatchInlineStyleDialog) {
            $watcher = Start-InlineStyleDialogWatcher
            $script:TableInlineDialogWatchers++
        }
        Invoke-WithRetry -Operation { Invoke-MathTypeToggle -Word $Word } | Out-Null
    } finally {
        Stop-InlineStyleDialogWatcher -Process $watcher
    }
}

function Get-FormulaRetryLiterals {
    param([Parameter(Mandatory = $true)]$Formula)

    if ([string]$Formula.Kind -ne "inline") { return @() }
    $literal = [string]$Formula.Literal
    if ($literal.Length -lt 2 -or -not ($literal.StartsWith('$') -and $literal.EndsWith('$'))) {
        return @()
    }
    $content = $literal.Substring(1, $literal.Length - 2).Trim()
    $number = '[+-]?(?:\d+(?:\.\d*)?|\.\d+)'
    $interval = [regex]::Match($content, "^\s*([\[\(])\s*($number)\s*,\s*($number)\s*([\]\)])\s*$")
    if ($interval.Success) {
        $left = $interval.Groups[1].Value
        $first = $interval.Groups[2].Value
        $second = $interval.Groups[3].Value
        $right = $interval.Groups[4].Value
        return @(
            ('$' + '\left' + $left + $first + ',\,' + $second + '\right' + $right + '$'),
            ('$' + '\mathrm{' + $left + $first + ',\,' + $second + $right + '}' + '$')
        )
    }
    if ([regex]::IsMatch($content, "^\s*$number\s*$")) {
        return @(('$' + '\mathrm{' + $content + '}' + '$'))
    }
    $primeSafeContent = ConvertTo-AsciiPrimeText -Text $content
    $primeMatch = [regex]::Match($primeSafeContent, "^(?<base>.+?)(?<primes>'+)$")
    if ($primeMatch.Success) {
        $base = $primeMatch.Groups['base'].Value.TrimEnd()
        $primeCount = $primeMatch.Groups['primes'].Value.Length
        if ($base.Length -gt 0 -and $primeCount -gt 0) {
            $primeCommands = ((1..$primeCount) | ForEach-Object { '\prime' }) -join ''
            return @(('$' + $base + '^{' + $primeCommands + '}' + '$'))
        }
    }
    $commaListPattern = '^[A-Za-z][A-Za-z0-9_{}]*(?:\s*,\s*[A-Za-z][A-Za-z0-9_{}]*)+$'
    if ([regex]::IsMatch($content, $commaListPattern)) {
        $commaSafe = [regex]::Replace($content, '\s*,\s*', ',\,')
        return @(('$' + $commaSafe + '$'))
    }
    return @()
}

function Get-ParagraphStyleName {
    param([Parameter(Mandatory = $true)]$Paragraph)
    try { return [string]$Paragraph.Style.NameLocal } catch { try { return [string]$Paragraph.Style } catch { return "" } }
}

function Test-DisplayEquationParagraph {
    param([Parameter(Mandatory = $true)]$Paragraph)
    $styleName = Get-ParagraphStyleName -Paragraph $Paragraph
    if ($styleName -like "MTDisplay*") { return $true }
    return $false
}

function Test-EmptyParagraph {
    param([Parameter(Mandatory = $true)]$Paragraph)
    try {
        if ($Paragraph.Range.InlineShapes.Count -ne 0) { return $false }
        if ($Paragraph.Range.Fields.Count -ne 0) { return $false }
        $text = [string]$Paragraph.Range.Text
        return [regex]::IsMatch($text, '^[\s\r\n\a]*$')
    } catch {
        return $false
    }
}

function Get-MathTypeShapeInParagraph {
    param([Parameter(Mandatory = $true)]$Paragraph)
    $found = @()
    for ($i = 1; $i -le $Paragraph.Range.InlineShapes.Count; $i++) {
        $shape = $Paragraph.Range.InlineShapes.Item($i)
        if (Test-MathTypeInlineShape -Shape $shape) {
            $found += $shape
        } else {
            Release-ComObject $shape
        }
    }
    if ($found.Count -eq 1) { return $found[0] }
    foreach ($shape in $found) { Release-ComObject $shape }
    return $null
}

function Repair-DisplayEquationLayout {
    param([Parameter(Mandatory = $true)]$Document)
    $reformatted = 0
    $removed = 0
    for ($i = $Document.Paragraphs.Count; $i -ge 1; $i--) {
        $paragraph = $Document.Paragraphs.Item($i)
        $shape = $null
        try {
            if (-not (Test-DisplayEquationParagraph -Paragraph $paragraph)) { continue }
            $shape = Get-MathTypeShapeInParagraph -Paragraph $paragraph
            if ($null -ne $shape) {
                $prefix = $Document.Range($paragraph.Range.Start, $shape.Range.Start)
                try {
                    if (-not [string]::IsNullOrWhiteSpace([string]$prefix.Text)) {
                        $prefix.Text = ""
                    }
                } finally {
                    Release-ComObject $prefix
                }
                $suffixEnd = [int]$paragraph.Range.End - 1
                if ($shape.Range.End -le $suffixEnd) {
                    $suffix = $Document.Range($shape.Range.End, $suffixEnd)
                    try {
                        if (-not [string]::IsNullOrWhiteSpace([string]$suffix.Text)) {
                            $suffix.Text = ""
                        }
                    } finally {
                        Release-ComObject $suffix
                    }
                }
            }

            $paragraph.Format.Alignment = $wdAlignParagraphCenter
            $paragraph.Format.LeftIndent = 0
            $paragraph.Format.RightIndent = 0
            $paragraph.Format.FirstLineIndent = 0
            $paragraph.Format.TabStops.ClearAll()
            $paragraph.Format.LineSpacingRule = $wdLineSpaceAtLeast
            $paragraph.Format.LineSpacing = 30
            $paragraph.Format.SpaceBefore = 3
            $paragraph.Format.SpaceAfter = 3
            $paragraph.Format.KeepTogether = $true
            $reformatted++

            if ($i -lt $Document.Paragraphs.Count) {
                $next = $Document.Paragraphs.Item($i + 1)
                try {
                    if (Test-EmptyParagraph -Paragraph $next) {
                        $next.Range.Delete() | Out-Null
                        $removed++
                    }
                } finally {
                    Release-ComObject $next
                }
            }
        } finally {
            Release-ComObject $shape
            Release-ComObject $paragraph
        }
    }
    $script:DisplayParagraphsReformatted += $reformatted
    $script:DisplayEmptyParagraphsRemoved += $removed
}

function Convert-OneFormula {
    param(
        [Parameter(Mandatory = $true)]$Word,
        [Parameter(Mandatory = $true)]$TargetDocument,
        [Parameter(Mandatory = $true)]$Formula
    )
    $stage = "create target range"
    $targetRange = $null
    $watchInlineStyleDialog = $false
    try {
        $contentStart = [int]$TargetDocument.Content.Start
        $contentEnd = [int]$TargetDocument.Content.End
        $targetStart = [int]$Formula.Start
        $targetEnd = [int]($Formula.Start + $Formula.Length)
        if ($targetStart -lt $contentStart -or $targetEnd -gt $contentEnd -or $targetStart -ge $targetEnd) {
            throw "Target range [$targetStart,$targetEnd) is outside document content [$contentStart,$contentEnd)."
        }
        $targetRange = $TargetDocument.Range($targetStart, $targetEnd)

        $actualText = [string]$targetRange.Text
        if ($actualText -cne [string]$Formula.Literal) {
            throw "Target text does not match the expected formula. Expected=$($Formula.Literal); Actual=$actualText"
        }

        $watchInlineStyleDialog = ([string]$Formula.Kind -eq "inline") -and (Test-FormulaRangeInTable -FormulaRange $targetRange)

        # MathType 6 officially accepts $...$ as inline Texvc and \[...\] as
        # display Texvc. Convert the selected range in the target document itself;
        # do not create or copy an OLE object across documents.
        $beforeObjects = Get-MathTypeObjectCount -Document $TargetDocument
        $stage = "select target formula"
        $TargetDocument.Activate()
        $targetRange.Select()

        $stage = "run MathType Toggle TeX on target selection"
        Invoke-MathTypeToggleForFormula -Word $Word -WatchInlineStyleDialog $watchInlineStyleDialog
        if ($DelayMilliseconds -gt 0) {
            Start-Sleep -Milliseconds $DelayMilliseconds
        }

        $stage = "validate MathType object count increment"
        $afterObjects = Get-MathTypeObjectCount -Document $TargetDocument
        if ($afterObjects -ne ($beforeObjects + 1)) {
            $retryLiterals = @(Get-FormulaRetryLiterals -Formula $Formula)
            if ($retryLiterals.Count -eq 0) {
                throw "Expected MathType object count to increase from $beforeObjects to $($beforeObjects + 1), but found $afterObjects."
            }

            Release-ComObject $targetRange
            $targetRange = $null
            $currentLiteral = [string]$Formula.Literal
            $convertedByFallback = $false
            foreach ($retryLiteral in $retryLiterals) {
                $stage = "prepare MathType compatibility fallback"
                $retryRange = $TargetDocument.Range($targetStart, $targetStart + $currentLiteral.Length)
                try {
                    if ([string]$retryRange.Text -cne $currentLiteral) {
                        throw "Formula text changed after failed Toggle TeX; refusing automatic rewrite. Current=$($retryRange.Text)"
                    }
                    $retryRange.Text = $retryLiteral
                    $retryRange.SetRange($targetStart, $targetStart + $retryLiteral.Length)
                    Write-Host "Retrying formula as: $retryLiteral" -ForegroundColor Yellow

                    $stage = "run MathType Toggle TeX on compatibility fallback"
                    $TargetDocument.Activate()
                    $retryRange.Select()
                    Invoke-MathTypeToggleForFormula -Word $Word -WatchInlineStyleDialog $watchInlineStyleDialog
                    if ($DelayMilliseconds -gt 0) {
                        Start-Sleep -Milliseconds $DelayMilliseconds
                    }
                    $afterObjects = Get-MathTypeObjectCount -Document $TargetDocument
                    if ($afterObjects -eq ($beforeObjects + 1)) {
                        $convertedByFallback = $true
                        $originalContent = ([string]$Formula.Literal).Substring(1, ([string]$Formula.Literal).Length - 2).Trim()
                        $primeSafeOriginal = ConvertTo-AsciiPrimeText -Text $originalContent
                        if ([regex]::IsMatch($primeSafeOriginal, "'+$")) {
                            $script:PrimeFallbackConversions++
                        } elseif ([regex]::IsMatch($originalContent, '^[A-Za-z][A-Za-z0-9_{}]*(?:\s*,\s*[A-Za-z][A-Za-z0-9_{}]*)+$')) {
                            $script:CommaListFallbackConversions++
                        } else {
                            $script:NumericFallbackConversions++
                        }
                        break
                    }
                    $currentLiteral = $retryLiteral
                } finally {
                    Release-ComObject $retryRange
                }
            }
            if (-not $convertedByFallback) {
                $stage = "restore formula after failed compatibility fallbacks"
                $restoreRange = $TargetDocument.Range($targetStart, $targetStart + $currentLiteral.Length)
                try {
                    if ([string]$restoreRange.Text -ceq $currentLiteral) {
                        $restoreRange.Text = [string]$Formula.Literal
                    }
                } finally {
                    Release-ComObject $restoreRange
                }
                throw "MathType ignored the original formula and all safe compatibility fallback forms."
            }
        }
    } catch {
        $hresult = ""
        if ($_.Exception -is [System.Runtime.InteropServices.COMException]) {
            $hresult = " HResult=0x{0:X8}." -f ($_.Exception.HResult -band 0xffffffffL)
        }
        $sample = [string]$Formula.Literal
        if ($sample.Length -gt 160) { $sample = $sample.Substring(0, 157) + "..." }
        throw "Formula conversion failed during '$stage'.$hresult Kind=$($Formula.Kind); Start=$($Formula.Start); Length=$($Formula.Length); Formula=$sample; Error=$($_.Exception.Message)"
    } finally {
        Release-ComObject $targetRange
    }
}

function Normalize-SourceText {
    param([Parameter(Mandatory = $true)][string]$Text)
    $normalized = $displayRegex.Replace($Text, "[MATH]")
    $normalized = $inlineRegex.Replace($normalized, "[MATH]")
    $normalized = $normalized.Replace([string][char]1, "[MATH]")
    # Toggle TeX may add paragraph marks around display equations. Ignore only
    # whitespace so that non-formula Chinese text is still compared exactly.
    return [regex]::Replace($normalized, '[\s\u00A0]+', '')
}

function Normalize-ConvertedText {
    param([Parameter(Mandatory = $true)][string]$Text)
    $normalized = $Text.Replace([string][char]1, "[MATH]")
    return [regex]::Replace($normalized, '[\s\u00A0]+', '')
}

function New-TextOffsetToWordPositionMap {
    param(
        [Parameter(Mandatory = $true)]$Range,
        [Parameter(Mandatory = $true)][string]$Text
    )
    $map = [int[]]::new($Text.Length + 1)
    $offset = 0
    $map[0] = [int]$Range.Start
    $characters = $Range.Characters
    try {
        for ($i = 1; $i -le $characters.Count; $i++) {
            $character = $characters.Item($i)
            try {
                $characterText = [string]$character.Text
                for ($j = 0; $j -lt $characterText.Length; $j++) {
                    if (($offset + $j) -le $Text.Length) {
                        $map[$offset + $j] = [int]$character.Start
                    }
                }
                $offset += $characterText.Length
                if ($offset -le $Text.Length) {
                    $map[$offset] = [int]$character.End
                }
            } finally {
                Release-ComObject $character
            }
        }
    } finally {
        Release-ComObject $characters
    }
    if ($offset -ne $Text.Length) {
        throw "Unable to map paragraph text to Word positions. TextLength=$($Text.Length), CharacterTextLength=$offset, Range=[$($Range.Start),$($Range.End))."
    }
    return $map
}

function Get-FormulaWordRanges {
    param([Parameter(Mandatory = $true)]$Document)
    $result = [System.Collections.Generic.List[object]]::new()
    for ($paragraphIndex = 1; $paragraphIndex -le $Document.Paragraphs.Count; $paragraphIndex++) {
        $paragraph = $Document.Paragraphs.Item($paragraphIndex)
        $range = $null
        try {
            $range = $paragraph.Range.Duplicate
            $text = [string]$range.Text
            $displayMatches = @($displayRegex.Matches($text))
            $inlineMatches = @($inlineRegex.Matches($text))
            if (($displayMatches.Count + $inlineMatches.Count) -eq 0) { continue }

            $map = $null
            if ($text.Length -ne ([int]$range.End - [int]$range.Start)) {
                $map = New-TextOffsetToWordPositionMap -Range $range -Text $text
            }

            foreach ($entry in @(
                @($inlineMatches | ForEach-Object { [pscustomobject]@{ Match = $_; Kind = 'inline' } }) +
                @($displayMatches | ForEach-Object { [pscustomobject]@{ Match = $_; Kind = 'display' } })
            )) {
                $match = $entry.Match
                if ($null -eq $map) {
                    $start = [int]$range.Start + [int]$match.Index
                    $end = $start + [int]$match.Length
                } else {
                    $start = $map[$match.Index]
                    $end = $map[$match.Index + $match.Length]
                }
                $result.Add([pscustomobject]@{
                    Start = [int]$start
                    Length = [int]($end - $start)
                    Kind = [string]$entry.Kind
                    Literal = [string]$match.Value
                    Paragraph = $paragraphIndex
                })
            }
        } finally {
            Release-ComObject $range
            Release-ComObject $paragraph
        }
    }
    return @($result)
}

function Get-SequenceFieldCount {
    param([Parameter(Mandatory = $true)]$Document)
    $count = 0
    for ($i = 1; $i -le $Document.Fields.Count; $i++) {
        $field = $Document.Fields.Item($i)
        if ($field.Code.Text -match 'SEQ\s+Equation') {
            $count++
        }
    }
    return $count
}

function Add-DisplayEquationNumbers {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)][int]$ExpectedCount
    )
    $paragraphs = [System.Collections.Generic.List[object]]::new()
    for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        $styleName = ""
        try { $styleName = [string]$paragraph.Style.NameLocal } catch { try { $styleName = [string]$paragraph.Style } catch {} }
        if ($styleName -eq "MTDisplayEquation") {
            $mathShapes = 0
            for ($j = 1; $j -le $paragraph.Range.InlineShapes.Count; $j++) {
                if (Test-MathTypeInlineShape -Shape $paragraph.Range.InlineShapes.Item($j)) {
                    $mathShapes++
                }
            }
            if ($mathShapes -eq 1) {
                $paragraphs.Add($paragraph)
            }
        }
    }
    if ($paragraphs.Count -ne $ExpectedCount) {
        throw "Expected $ExpectedCount display-equation paragraphs, found $($paragraphs.Count)."
    }

    $section = $Document.Sections.Item(1)
    $pageSetup = $section.PageSetup
    $usableWidth = $pageSetup.PageWidth - $pageSetup.LeftMargin - $pageSetup.RightMargin

    foreach ($paragraph in $paragraphs) {
        $paragraph.Format.Alignment = $wdAlignParagraphLeft
        $paragraph.Format.LineSpacingRule = $wdLineSpaceAtLeast
        $paragraph.Format.LineSpacing = 30
        $paragraph.Format.TabStops.ClearAll()
        $null = $paragraph.Format.TabStops.Add($usableWidth / 2, $wdAlignTabCenter, $wdTabLeaderSpaces)
        $null = $paragraph.Format.TabStops.Add($usableWidth, $wdAlignTabRight, $wdTabLeaderSpaces)

        $shape = $null
        for ($j = 1; $j -le $paragraph.Range.InlineShapes.Count; $j++) {
            $candidate = $paragraph.Range.InlineShapes.Item($j)
            if (Test-MathTypeInlineShape -Shape $candidate) {
                $shape = $candidate
                break
            }
        }
        if ($null -eq $shape) {
            throw "Display paragraph lost its MathType object before numbering."
        }

        $prefix = $Document.Range($paragraph.Range.Start, $shape.Range.Start)
        $prefix.Text = "`t"
        Release-ComObject $prefix
        $shape = $paragraph.Range.InlineShapes.Item(1)
        $suffix = $Document.Range($shape.Range.End, $paragraph.Range.End - 1)
        $suffix.Text = ""
        Release-ComObject $suffix

        $insert = $Document.Range($shape.Range.End, $shape.Range.End)
        $insert.Text = "`t("
        $insert.Collapse($wdCollapseEnd)
        $field = $Document.Fields.Add($insert, $wdFieldEmpty, "SEQ Equation \* ARABIC", $true)
        $field.Update() | Out-Null
        $tail = $Document.Range($field.Result.End, $field.Result.End)
        $tail.Text = ")"

        Release-ComObject $tail
        Release-ComObject $field
        Release-ComObject $insert
        Release-ComObject $shape
    }
    $Document.Fields.Update() | Out-Null
}

$InputPath = [System.IO.Path]::GetFullPath($InputPath)
if (-not (Test-Path -LiteralPath $InputPath)) {
    throw "Input file not found: $InputPath"
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $inputDirectory = [System.IO.Path]::GetDirectoryName($InputPath)
    $inputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $inputExtension = [System.IO.Path]::GetExtension($InputPath)
    $OutputPath = Join-Path $inputDirectory ($inputBaseName + '_mathtype' + $inputExtension)
}
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
if ($InputPath -eq $OutputPath) {
    throw "InputPath and OutputPath must be different."
}
if (Test-Path -LiteralPath $OutputPath) {
    if (-not $Force) {
        throw "Output already exists: $OutputPath. Use -Force to replace it."
    }
}
if ($env:USERNAME -eq "codexsandboxonline") {
    throw "Run this script in your normal Windows PowerShell session, not the Codex terminal."
}

$outputDir = Split-Path -Parent $OutputPath
if ($outputDir) { [void][System.IO.Directory]::CreateDirectory($outputDir) }
Copy-Item -LiteralPath $InputPath -Destination $OutputPath -Force

$word = $null
$document = $null
$tempDocument = $null
$saved = $false

try {
    $word = Get-RunningWord
    Write-Host "Attached to: $($word.Name) $($word.Version)" -ForegroundColor Cyan

    $tempDocument = $word.Documents.Add()
    Write-Host "Preflight: converting inline and display formula smoke tests" -ForegroundColor Cyan
    $smokeShape = Put-FormulaInTemporaryDocument -Word $word -TempDocument $tempDocument -Literal '$x^2$'
    Release-ComObject $smokeShape
    Clear-TemporaryDocument -Document $tempDocument
    $smokeDisplayShape = Put-FormulaInTemporaryDocument -Word $word -TempDocument $tempDocument -Literal '\[x^2\]'
    Release-ComObject $smokeDisplayShape
    Clear-TemporaryDocument -Document $tempDocument

    $document = Invoke-WithRetry -Operation {
        $word.Documents.Open($OutputPath, $false, $false, $false)
    }
    $document.Activate()

    $originalParagraphs = $document.Paragraphs.Count
    $originalTables = $document.Tables.Count
    $initialMathTypeObjects = Get-MathTypeObjectCount -Document $document
    $originalText = [string]$document.Content.Text
    $sourceFingerprint = Normalize-SourceText -Text $originalText

    $displayMatches = @($displayRegex.Matches($originalText))
    $inlineMatches = @($inlineRegex.Matches($originalText))
    if ($ExpectedInline -ge 0 -and $inlineMatches.Count -ne $ExpectedInline) {
        throw "Formula count mismatch before conversion. Inline=$($inlineMatches.Count)/$ExpectedInline, Display=$($displayMatches.Count)/$ExpectedDisplay"
    }
    if ($ExpectedDisplay -ge 0 -and $displayMatches.Count -ne $ExpectedDisplay) {
        throw "Formula count mismatch before conversion. Inline=$($inlineMatches.Count)/$ExpectedInline, Display=$($displayMatches.Count)/$ExpectedDisplay"
    }
    if ($ExpectedInline -lt 0) { $ExpectedInline = $inlineMatches.Count }
    if ($ExpectedDisplay -lt 0) { $ExpectedDisplay = $displayMatches.Count }
    if (($ExpectedInline + $ExpectedDisplay) -le 0) {
        throw 'No formulas were found. Required delimiters are $...$ for inline formulas and \[...\] for display formulas.'
    }
    Write-Host ("Detected formulas: {0} inline, {1} display" -f $ExpectedInline, $ExpectedDisplay) -ForegroundColor Green

    Write-Host "Resolving LaTeX ranges in Word-native coordinates" -ForegroundColor Cyan
    $formulas = @(Get-FormulaWordRanges -Document $document)
    $resolvedInline = @($formulas | Where-Object Kind -eq 'inline').Count
    $resolvedDisplay = @($formulas | Where-Object Kind -eq 'display').Count
    if ($resolvedInline -ne $ExpectedInline -or $resolvedDisplay -ne $ExpectedDisplay) {
        throw "Formula count mismatch after resolving Word ranges. Inline=$resolvedInline/$ExpectedInline, Display=$resolvedDisplay/$ExpectedDisplay"
    }
    $ordered = @($formulas | Sort-Object Start -Descending)

    Write-Host "Converting $($ordered.Count) formulas individually. Do not interact with Word." -ForegroundColor Cyan
    for ($index = 0; $index -lt $ordered.Count; $index++) {
        $formula = $ordered[$index]
        $done = $index + 1
        if ($done -eq 1 -or $done % 10 -eq 0) {
            Write-Host ("Starting {0}/{1}: {2}, Start={3}, Length={4}" -f $done, $ordered.Count, $formula.Kind, $formula.Start, $formula.Length)
        }
        Convert-OneFormula -Word $word -TargetDocument $document -Formula $formula
        if ($done -eq 1 -or $done % 10 -eq 0 -or $done -eq $ordered.Count) {
            Write-Host ("Converted {0}/{1}" -f $done, $ordered.Count)
        }
    }

    Write-Host "Repairing display-equation paragraph layout" -ForegroundColor Cyan
    Repair-DisplayEquationLayout -Document $document

    $mathTypeObjects = @(Get-MathTypeInlineShapes -Document $document).Count
    $convertedText = [string]$document.Content.Text
    $remainingInline = $inlineRegex.Matches($convertedText).Count
    $remainingDisplay = $displayRegex.Matches($convertedText).Count
    $expectedFinalMathTypeObjects = $initialMathTypeObjects + $ExpectedInline + $ExpectedDisplay
    if ($mathTypeObjects -ne $expectedFinalMathTypeObjects) {
        throw "Expected $expectedFinalMathTypeObjects MathType objects, found $mathTypeObjects."
    }
    if ($remainingInline -ne 0 -or $remainingDisplay -ne 0) {
        throw "Residual LaTeX remains. Inline=$remainingInline, Display=$remainingDisplay"
    }
    if ($document.OMaths.Count -ne 0) {
        throw "Word-native OMath objects were found: $($document.OMaths.Count)"
    }

    # Preserve a recoverable output as soon as the core formula audit passes.
    # Later layout/text-integrity checks can still reject the result, but they no
    # longer discard a successful 227-formula conversion.
    $document.Save()
    $saved = $true
    Write-Host "Core formula conversion checkpoint saved" -ForegroundColor Green

    $finalParagraphs = $document.Paragraphs.Count
    $paragraphDelta = $finalParagraphs - $originalParagraphs
    if ($document.Tables.Count -ne $originalTables) {
        throw "Table structure changed during conversion. Tables $originalTables->$($document.Tables.Count). Output checkpoint retained: $OutputPath"
    }
    if ($paragraphDelta -lt 0 -or $paragraphDelta -gt $ExpectedDisplay) {
        throw "Unexpected paragraph change during conversion. Paragraphs $originalParagraphs->$finalParagraphs (delta=$paragraphDelta; allowed=0..$ExpectedDisplay). Output checkpoint retained: $OutputPath"
    }
    if ($paragraphDelta -gt 0) {
        Write-Host ("MathType added {0} display-equation paragraph boundaries (allowed up to {1})." -f $paragraphDelta, $ExpectedDisplay) -ForegroundColor Yellow
    }
    if (-not $SkipTextIntegrityCheck) {
        $convertedFingerprint = Normalize-ConvertedText -Text $convertedText
        if ($sourceFingerprint -cne $convertedFingerprint) {
            throw "Normalized non-formula text changed during conversion. Output checkpoint retained for inspection: $OutputPath"
        }
    }

    if ($AddEquationNumbers) {
        Write-Host "Adding $ExpectedDisplay right-aligned equation numbers" -ForegroundColor Cyan
        Add-DisplayEquationNumbers -Document $document -ExpectedCount $ExpectedDisplay
    }

    $sequenceFields = Get-SequenceFieldCount -Document $document
    if ($AddEquationNumbers -and $sequenceFields -ne $ExpectedDisplay) {
        throw "Expected $ExpectedDisplay SEQ Equation fields, found $sequenceFields."
    }

    $document.Save()
    $saved = $true
    Write-Host "Success" -ForegroundColor Green
    [pscustomobject]@{
        InputPath = $InputPath
        OutputPath = $OutputPath
        MathTypeObjects = $mathTypeObjects
        InitialMathTypeObjects = $initialMathTypeObjects
        DetectedInlineLatex = $ExpectedInline
        DetectedDisplayLatex = $ExpectedDisplay
        NumericFallbackConversions = $script:NumericFallbackConversions
        PrimeFallbackConversions = $script:PrimeFallbackConversions
        CommaListFallbackConversions = $script:CommaListFallbackConversions
        TableInlineDialogWatchers = $script:TableInlineDialogWatchers
        DisplayParagraphsReformatted = $script:DisplayParagraphsReformatted
        DisplayEmptyParagraphsRemoved = $script:DisplayEmptyParagraphsRemoved
        RemainingInlineLatex = $remainingInline
        RemainingDisplayLatex = $remainingDisplay
        WordOMathObjects = $document.OMaths.Count
        EquationNumberFields = $sequenceFields
        Paragraphs = $document.Paragraphs.Count
        Tables = $document.Tables.Count
        TextIntegrityChecked = (-not $SkipTextIntegrityCheck)
    }
} catch {
    $conversionError = $_
    Write-UserActionError `
        -Title "DOCX LaTeX to MathType conversion failed." `
        -Reason $conversionError.Exception.Message `
        -Actions @(
            "InputPath : $InputPath",
            "OutputPath: $OutputPath",
            "If partial progress was saved, inspect the output DOCX or resume using it as the next -InputPath.",
            "For delimiter issues, ensure inline formulas are `$...`$ and display formulas are \[...\]; do not use `$\$...`$\$.",
            "For Word/MathType issues, leave Word open, verify Toggle TeX manually, and rerun."
        )
    if ($null -ne $document) {
        try {
            $document.Save()
            $saved = $true
            Write-Warning "Conversion stopped, but partial progress was saved: $OutputPath"
            Write-Warning "Resume by using this partial DOCX as the next -InputPath and a new -OutputPath."
        } catch {
            Write-Warning "Conversion failed and the partial checkpoint could not be saved: $($_.Exception.Message)"
        }
    }
    throw $conversionError
} finally {
    if ($null -ne $document) {
        try { $document.Close($wdDoNotSaveChanges) } catch {}
        Release-ComObject $document
    }
    if ($null -ne $tempDocument) {
        try { $tempDocument.Close($wdDoNotSaveChanges) } catch {}
        Release-ComObject $tempDocument
    }
    if ($null -ne $word) {
        try { $word.Visible = $true } catch {}
        Release-ComObject $word
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
