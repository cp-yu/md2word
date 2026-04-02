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

    [string]$StylePreset = "",

    [switch]$Visible,

    [int]$TimeoutSeconds = 180
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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

$pythonCommand = Get-PythonCommand
$scriptPath = Join-Path $PSScriptRoot "md2word.py"
$pythonArgs = @()
if ($pythonCommand.Count -gt 1) {
    $pythonArgs = $pythonCommand[1..($pythonCommand.Count - 1)]
}
$args = @(
    $scriptPath,
    "--input", $Input
)

if ($TemplateMht) {
    $args += @("--template-mht", $TemplateMht)
}
if ($TemplateDocx) {
    $args += @("--template-docx", $TemplateDocx)
}
if ($TemplateSpec) {
    $args += @("--template-spec", $TemplateSpec)
}
if ($TemplateOut) {
    $args += @("--template-out", $TemplateOut)
}
if ($TemplateReport) {
    $args += @("--template-report", $TemplateReport)
}
if ($OutputMht) {
    $args += @("--output-mht", $OutputMht)
}
if ($OutputDocx) {
    $args += @("--output-docx", $OutputDocx)
}
if ($ProcessedDocx) {
    $args += @("--processed-docx", $ProcessedDocx)
}
if ($MacroName) {
    $args += @("--macro-name", $MacroName)
}
if ($StylePreset) {
    $args += @("--style-preset", $StylePreset)
}
if ($Visible.IsPresent) {
    $args += "--visible"
}
$args += @("--timeout-seconds", $TimeoutSeconds.ToString())

& $pythonCommand[0] @pythonArgs @args
exit $LASTEXITCODE
