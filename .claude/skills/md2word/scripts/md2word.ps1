[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Input,

    [string]$TemplateMht = "",

    [string]$TemplateDocx = "",

    [string]$TemplateSpec = "",

    [string]$TemplateOut = "",

    [string]$TemplateReport = "",

    [string]$OutputMht = "",

    [string]$OutputDocx = "",

    [string]$ProcessedDocx = "",

    [string]$MacroName = "",

    [switch]$Visible,

    [int]$TimeoutSeconds = 180
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-AbsolutePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathValue,

        [switch]$AllowMissing
    )

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        throw "Path must not be empty"
    }

    if ($AllowMissing) {
        return [System.IO.Path]::GetFullPath($PathValue)
    }

    return (Resolve-Path -LiteralPath $PathValue).Path
}

function Ensure-ParentDirectory {
    param([Parameter(Mandatory = $true)][string]$PathValue)

    $parent = [System.IO.Path]::GetDirectoryName($PathValue)
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        [void][System.IO.Directory]::CreateDirectory($parent)
    }
}

function Get-UnescapedSingleDollarCount {
    param([Parameter(Mandatory = $true)][string]$Line)

    $count = 0
    for ($i = 0; $i -lt $Line.Length; $i++) {
        if ($Line[$i] -ne '$') {
            continue
        }

        if ($i -gt 0 -and $Line[$i - 1] -eq '\') {
            continue
        }

        $prevIsDollar = $i -gt 0 -and $Line[$i - 1] -eq '$'
        $nextIsDollar = $i + 1 -lt $Line.Length -and $Line[$i + 1] -eq '$'
        if ($prevIsDollar -or $nextIsDollar) {
            continue
        }

        $count++
    }

    return $count
}

function Test-MarkdownFormulaDelimiters {
    param([Parameter(Mandatory = $true)][string]$PathValue)

    $lines = Get-Content -LiteralPath $PathValue -Encoding UTF8
    $inCodeFence = $false
    $inDisplayMath = $false
    $displayMathStartLine = 0

    for ($lineNumber = 1; $lineNumber -le $lines.Count; $lineNumber++) {
        $line = [string]$lines[$lineNumber - 1]
        $trimmed = $line.Trim()

        if ($trimmed.StartsWith("```")) {
            $inCodeFence = -not $inCodeFence
            continue
        }

        if ($inCodeFence) {
            continue
        }

        if ($trimmed -eq '$$') {
            if (-not $inDisplayMath) {
                $inDisplayMath = $true
                $displayMathStartLine = $lineNumber
            } else {
                $inDisplayMath = $false
                $displayMathStartLine = 0
            }
            continue
        }

        if ($inDisplayMath) {
            continue
        }

        $singleDollarCount = Get-UnescapedSingleDollarCount -Line $line
        if (($singleDollarCount % 2) -ne 0) {
            throw "Unmatched inline formula delimiter `$ found at line $lineNumber."
        }
    }

    if ($inCodeFence) {
        throw "Unclosed code fence detected in markdown input."
    }

    if ($inDisplayMath) {
        throw "Unclosed display formula block starting at line $displayMathStartLine."
    }
}

function Get-PythonCommand {
    if (-not [string]::IsNullOrWhiteSpace($env:PYTHON_EXE)) {
        $cmd = Get-Command $env:PYTHON_EXE -ErrorAction SilentlyContinue
        if ($cmd) {
            return @($cmd.Source)
        }
    }

    $candidates = @(
        @("py", "-3"),
        @("python"),
        @("python3")
    )

    foreach ($candidate in $candidates) {
        $cmd = Get-Command $candidate[0] -ErrorAction SilentlyContinue
        if ($cmd) {
            if ($candidate.Count -gt 1) {
                return @($cmd.Source) + $candidate[1..($candidate.Count - 1)]
            }
            return @($cmd.Source)
        }
    }

    throw "Python 3 not found. Please install Python and make `py -3` or `python` available."
}

function Get-PowerShellExe {
    $processPath = (Get-Process -Id $PID).Path
    if ($processPath -and (Test-Path -LiteralPath $processPath)) {
        return $processPath
    }

    foreach ($name in @("powershell.exe", "pwsh.exe")) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd) {
            return $cmd.Source
        }
    }

    throw "Cannot locate a usable PowerShell executable."
}

function Invoke-CommandArray {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Command,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $exe = $Command[0]
    $args = @()
    if ($Command.Count -gt 1) {
        $args = $Command[1..($Command.Count - 1)]
    }

    $output = & $exe @args 2>&1
    $exitCode = $LASTEXITCODE

    if ($output) {
        $output | ForEach-Object { Write-Output $_ }
    }

    if ($exitCode -ne 0) {
        throw "$Label failed with exit code $exitCode."
    }
}

function Invoke-PowerShellFileWithTimeout {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string[]]$ArgumentList,

        [Parameter(Mandatory = $true)]
        [string]$Label,

        [int]$TimeoutSeconds = 180
    )

    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()
    $psExe = Get-PowerShellExe

    try {
        $args = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $FilePath) + $ArgumentList
        $process = Start-Process `
            -FilePath $psExe `
            -ArgumentList $args `
            -PassThru `
            -WindowStyle Hidden `
            -RedirectStandardOutput $stdoutPath `
            -RedirectStandardError $stderrPath

        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            try {
                $process.Kill()
            } catch {
            }
            throw "$Label timed out after $TimeoutSeconds seconds."
        }

        $stdout = (Get-Content -LiteralPath $stdoutPath -Raw).Trim()
        $stderr = (Get-Content -LiteralPath $stderrPath -Raw).Trim()

        if ($stdout) {
            $stdout -split "`r?`n" | ForEach-Object { Write-Output $_ }
        }
        if ($stderr) {
            $stderr -split "`r?`n" | ForEach-Object { Write-Output $_ }
        }

        if ($process.ExitCode -ne 0) {
            throw "$Label failed with exit code $($process.ExitCode)."
        }
    } finally {
        Remove-Item -LiteralPath $stdoutPath, $stderrPath -ErrorAction SilentlyContinue
    }
}

function Test-ValidDocx {
    param([Parameter(Mandatory = $true)][string]$PathValue)

    if (-not (Test-Path -LiteralPath $PathValue)) {
        return $false
    }

    $fileInfo = Get-Item -LiteralPath $PathValue
    if ($fileInfo.Length -le 0) {
        return $false
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($PathValue)
        $names = @{}
        foreach ($entry in $zip.Entries) {
            $names[$entry.FullName] = $true
        }
        return $names.ContainsKey("[Content_Types].xml") -and $names.ContainsKey("word/document.xml")
    } catch {
        return $false
    } finally {
        if ($zip -ne $null) {
            $zip.Dispose()
        }
    }
}

$templateSourceCount = 0
if ($TemplateMht) { $templateSourceCount++ }
if ($TemplateDocx) { $templateSourceCount++ }
if ($TemplateSpec) { $templateSourceCount++ }

if ($templateSourceCount -gt 1) {
    throw "Only one of -TemplateMht / -TemplateDocx / -TemplateSpec can be used."
}

if ($TemplateOut -and $TemplateMht) {
    throw "-TemplateOut is only meaningful with -TemplateDocx or -TemplateSpec."
}

$pythonCommand = Get-PythonCommand
$skillRoot = Split-Path -Parent $PSScriptRoot

$inputAbs = Resolve-AbsolutePath -PathValue $Input
Test-MarkdownFormulaDelimiters -PathValue $inputAbs
$inputDir = Split-Path -Parent $inputAbs
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputAbs)
$baseNoExt = Join-Path $inputDir $baseName

$outputMhtAbs = if ($OutputMht) { Resolve-AbsolutePath -PathValue $OutputMht -AllowMissing } else { "$baseNoExt.mht" }
$outputDocxAbs = if ($OutputDocx) { Resolve-AbsolutePath -PathValue $OutputDocx -AllowMissing } else { "$baseNoExt.docx" }
$processedDocxAbs = if ($ProcessedDocx) { Resolve-AbsolutePath -PathValue $ProcessedDocx -AllowMissing } else { "$baseNoExt.wordmath.docx" }
$templateReportAbs = if ($TemplateReport) { Resolve-AbsolutePath -PathValue $TemplateReport -AllowMissing } else { "" }

Ensure-ParentDirectory -PathValue $outputMhtAbs
Ensure-ParentDirectory -PathValue $outputDocxAbs
Ensure-ParentDirectory -PathValue $processedDocxAbs
if ($templateReportAbs) {
    Ensure-ParentDirectory -PathValue $templateReportAbs
}

$stageRoot = Join-Path ([System.IO.Path]::GetTempPath()) "repo2patent-md2word"
$stageId = "{0:yyyyMMdd-HHmmss}-{1}" -f (Get-Date), $PID
$stageDir = Join-Path $stageRoot $stageId
[void][System.IO.Directory]::CreateDirectory($stageDir)

$resolvedTemplateAbs = ""

if ($TemplateMht) {
    $resolvedTemplateAbs = Resolve-AbsolutePath -PathValue $TemplateMht
} elseif ($TemplateDocx) {
    $templateDocxAbs = Resolve-AbsolutePath -PathValue $TemplateDocx
    $stageTemplateMht = Join-Path $stageDir "template-normalized.mht"
    $targetTemplateOut = if ($TemplateOut) { Resolve-AbsolutePath -PathValue $TemplateOut -AllowMissing } else { $stageTemplateMht }
    Ensure-ParentDirectory -PathValue $targetTemplateOut

    $docxArgs = @("-InputDocx", $templateDocxAbs, "-OutputMht", $targetTemplateOut)
    if ($Visible.IsPresent) {
        $docxArgs += "-Visible"
    }

    Invoke-PowerShellFileWithTimeout `
        -FilePath (Join-Path $PSScriptRoot "word_template_to_mht.ps1") `
        -ArgumentList $docxArgs `
        -Label "Template DOCX to MHT" `
        -TimeoutSeconds $TimeoutSeconds

    if (-not (Test-Path -LiteralPath $targetTemplateOut)) {
        throw "Template DOCX normalization did not produce MHT: $targetTemplateOut"
    }
    $resolvedTemplateAbs = $targetTemplateOut
} elseif ($TemplateSpec) {
    $templateSpecAbs = Resolve-AbsolutePath -PathValue $TemplateSpec
    $targetTemplateOut = if ($TemplateOut) {
        Resolve-AbsolutePath -PathValue $TemplateOut -AllowMissing
    } else {
        [System.IO.Path]::ChangeExtension($templateSpecAbs, ".generated.mht")
    }
    Ensure-ParentDirectory -PathValue $targetTemplateOut

    Invoke-CommandArray `
        -Command ($pythonCommand + @(
            (Join-Path $PSScriptRoot "generate_template_mht.py"),
            "--spec", $templateSpecAbs,
            "--output", $targetTemplateOut
        )) `
        -Label "Generate template MHT"

    $resolvedTemplateAbs = $targetTemplateOut
} else {
    $resolvedTemplateAbs = Join-Path $skillRoot "assets/专利交底书模板.mht"
}

if (-not (Test-Path -LiteralPath $resolvedTemplateAbs)) {
    throw "Resolved template not found: $resolvedTemplateAbs"
}

$renderArgs = @(
    (Join-Path $PSScriptRoot "render_mht.py"),
    "--input", $inputAbs,
    "--template", $resolvedTemplateAbs,
    "--output", $outputMhtAbs
)
if ($templateReportAbs) {
    $renderArgs += @("--report", $templateReportAbs)
}

Invoke-CommandArray -Command ($pythonCommand + $renderArgs) -Label "Render MHT"

$stageDocx = Join-Path $stageDir "raw.docx"
$stageProcessedDocx = Join-Path $stageDir "processed.docx"
$pipelineArgs = @(
    "-InputMht", $outputMhtAbs,
    "-OutputDocx", $stageDocx,
    "-ProcessedDocx", $stageProcessedDocx
)
if ($MacroName) {
    $pipelineArgs += @("-MacroName", $MacroName)
}
if ($Visible.IsPresent) {
    $pipelineArgs += "-Visible"
}

Invoke-PowerShellFileWithTimeout `
    -FilePath (Join-Path $PSScriptRoot "word_mht_pipeline.ps1") `
    -ArgumentList $pipelineArgs `
    -Label "Word MHT pipeline" `
    -TimeoutSeconds $TimeoutSeconds

if (-not (Test-ValidDocx -PathValue $stageDocx)) {
    throw "Invalid raw DOCX generated by Word pipeline: $stageDocx"
}

if (-not (Test-ValidDocx -PathValue $stageProcessedDocx)) {
    throw "Invalid processed DOCX generated by Word pipeline: $stageProcessedDocx"
}

Copy-Item -LiteralPath $stageDocx -Destination $outputDocxAbs -Force
Copy-Item -LiteralPath $stageProcessedDocx -Destination $processedDocxAbs -Force

Write-Output "[ok] template: $resolvedTemplateAbs"
if ($templateReportAbs) {
    Write-Output "[ok] template report: $templateReportAbs"
}
Write-Output "[ok] rendered mht: $outputMhtAbs"
Write-Output "[ok] raw docx: $outputDocxAbs"
Write-Output "[ok] wordmath docx: $processedDocxAbs"
