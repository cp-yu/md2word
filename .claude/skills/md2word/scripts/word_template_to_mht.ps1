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

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir "word_common.ps1")

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
    Close-WordSession -Document $document -Word $word -SaveChanges $wdDoNotSaveChanges
}
