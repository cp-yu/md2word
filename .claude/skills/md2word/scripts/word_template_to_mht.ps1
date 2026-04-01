[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $true)]
    [string]$OutputMht,

    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$wdDoNotSaveChanges = 0
$wdFormatWebArchive = 9

function Resolve-WindowsPath {
    param([string]$PathValue)

    if ($PathValue -match '^[A-Za-z]:\\') {
        return $PathValue
    }

    if ($PathValue -match '^\\\\') {
        return $PathValue
    }

    $normalized = $PathValue -replace '\\', '/'
    if ($normalized -match '^/mnt/([a-zA-Z])/(.*)$') {
        $drive = $matches[1].ToUpper()
        $rest = $matches[2] -replace '/', '\'
        return "${drive}:\$rest"
    }

    throw "Unsupported path for Windows conversion: $PathValue"
}

function Release-ComObject {
    param([object]$ComObject)

    if ($null -eq $ComObject) {
        return
    }

    try {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    } catch {
    }
}

$word = $null
$document = $null

try {
    $InputDocx = Resolve-WindowsPath -PathValue $InputDocx
    $OutputMht = Resolve-WindowsPath -PathValue $OutputMht

    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$Visible
    $word.DisplayAlerts = 0

    $document = $word.Documents.Open($InputDocx, $false, $true)
    $document.SaveAs([ref]$OutputMht, [ref]$wdFormatWebArchive)
} finally {
    if ($document -ne $null) {
        try {
            $document.Close([ref]$wdDoNotSaveChanges)
        } catch {
        }
    }
    if ($word -ne $null) {
        try {
            $word.Quit()
        } catch {
        }
    }

    Release-ComObject -ComObject $document
    Release-ComObject -ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
