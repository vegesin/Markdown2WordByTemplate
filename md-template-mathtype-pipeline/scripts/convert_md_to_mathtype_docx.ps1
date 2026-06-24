[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [Alias("Markdown", "Md")]
    [string]$MarkdownPath,

    [Parameter(Mandatory = $false)]
    [Alias("Template")]
    [string]$TemplatePath,

    [Parameter(Mandatory = $true)]
    [Alias("Output")]
    [string]$OutputPath,

    [string]$Stage1Path,
    [string]$VisioDirectory,
    [string]$MdSkillDirectory = "$env:USERPROFILE\.codex\skills\md-to-template-docx",
    [string]$MathTypeSkillDirectory = "$env:USERPROFILE\.codex\skills\word-mathtype-selective",
    [string]$PythonPath = "$env:USERPROFILE\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe",
    [string]$PandocPath = "$env:LOCALAPPDATA\Pandoc\pandoc.exe",
    [switch]$AddEquationNumbers,
    [switch]$SkipTextIntegrityCheck,
    [switch]$Force
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Resolve-InputFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Label
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label file not found: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Resolve-InputDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Label
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "$Label directory not found: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Normalize-DocxPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    if ([System.IO.Path]::GetExtension($Path) -ne ".docx") {
        $Path += ".docx"
    }
    return [System.IO.Path]::GetFullPath($Path)
}

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "== $Message ==" -ForegroundColor Cyan
}

$MarkdownPath = Resolve-InputFile -Path $MarkdownPath -Label "Markdown"
if (-not [string]::IsNullOrWhiteSpace($TemplatePath)) {
    $TemplatePath = Resolve-InputFile -Path $TemplatePath -Label "Word template"
}
$PythonPath = Resolve-InputFile -Path $PythonPath -Label "Python"
$PandocPath = Resolve-InputFile -Path $PandocPath -Label "Pandoc"
$MdSkillDirectory = Resolve-InputDirectory -Path $MdSkillDirectory -Label "md-to-template-docx skill"
$MathTypeSkillDirectory = Resolve-InputDirectory -Path $MathTypeSkillDirectory -Label "word-mathtype-selective skill"

$stage1Runner = Resolve-InputFile -Path (Join-Path $MdSkillDirectory "scripts\run_md_to_template_docx.ps1") -Label "Stage-1 runner"
$mathTypeRunner = Resolve-InputFile -Path (Join-Path $MathTypeSkillDirectory "scripts\convert_word_latex_to_mathtype.ps1") -Label "MathType converter"

$OutputPath = Normalize-DocxPath -Path $OutputPath
$outputDirectory = [System.IO.Path]::GetDirectoryName($OutputPath)
if ($outputDirectory) {
    [void][System.IO.Directory]::CreateDirectory($outputDirectory)
}
if ([string]::IsNullOrWhiteSpace($Stage1Path)) {
    $Stage1Path = Join-Path $outputDirectory ([System.IO.Path]::GetFileNameWithoutExtension($OutputPath) + "_stage1_latex.docx")
} else {
    $Stage1Path = Normalize-DocxPath -Path $Stage1Path
}
if ([string]::IsNullOrWhiteSpace($VisioDirectory)) {
    $VisioDirectory = Join-Path $outputDirectory ([System.IO.Path]::GetFileNameWithoutExtension($OutputPath) + "_visio_sources")
} else {
    $VisioDirectory = [System.IO.Path]::GetFullPath($VisioDirectory)
}

Write-Host "Markdown template MathType pipeline" -ForegroundColor Green
Write-Host "Markdown      : $MarkdownPath"
Write-Host "Template      : $(if ($TemplatePath) { $TemplatePath } else { 'bundled default template' })"
Write-Host "Stage-1 DOCX  : $Stage1Path"
Write-Host "Final DOCX    : $OutputPath"
Write-Host "Visio sources : $VisioDirectory"

Write-Step "Stage 1/2: Markdown to template DOCX"
$stage1Args = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", $stage1Runner,
    "-MarkdownPath", $MarkdownPath,
    "-OutputPath", $Stage1Path,
    "-SkillDirectory", $MdSkillDirectory,
    "-PythonPath", $PythonPath,
    "-PandocPath", $PandocPath,
    "-VisioDirectory", $VisioDirectory
)
if ($TemplatePath) {
    $stage1Args += @("-TemplatePath", $TemplatePath)
}
& powershell @stage1Args
if ($LASTEXITCODE -ne 0) {
    throw "Stage 1 failed with exit code $LASTEXITCODE."
}
if (-not (Test-Path -LiteralPath $Stage1Path -PathType Leaf)) {
    throw "Stage 1 finished but did not create: $Stage1Path"
}

Write-Step "Stage 2/2: Word LaTeX to MathType"
$stage2Args = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", $mathTypeRunner,
    "-InputPath", $Stage1Path,
    "-OutputPath", $OutputPath
)
if ($Force) { $stage2Args += "-Force" }
if ($AddEquationNumbers) { $stage2Args += "-AddEquationNumbers" }
if ($SkipTextIntegrityCheck) { $stage2Args += "-SkipTextIntegrityCheck" }

& powershell @stage2Args
if ($LASTEXITCODE -ne 0) {
    throw "Stage 2 failed with exit code $LASTEXITCODE. Stage-1 DOCX retained: $Stage1Path"
}
if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
    throw "Stage 2 finished but did not create: $OutputPath"
}

Write-Host ""
Write-Host "Pipeline success" -ForegroundColor Green
[pscustomobject]@{
    MarkdownPath = $MarkdownPath
    TemplatePath = if ($TemplatePath) { $TemplatePath } else { "bundled default template" }
    Stage1Path = $Stage1Path
    OutputPath = $OutputPath
    VisioSources = $VisioDirectory
}
