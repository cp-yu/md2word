[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputMht,

    [Parameter(Mandatory = $true)]
    [string]$OutputDocx,

    [Parameter(Mandatory = $true)]
    [string]$ProcessedDocx,

    [string]$MacroName = "",

    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$wdFormatDocumentDefault = 16
$wdCollapseEnd = 0

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir "word_common.ps1")

function Normalize-LatexBlockText {
    param([string]$Text)

    $normalized = $Text -replace "`r", " " -replace "`n", " "
    $normalized = [regex]::Replace($normalized, "\s+", " ").Trim()
    $normalized = $normalized -replace "\\varnothing", "\emptyset"
    # Word 的 LaTeX/BuildUp 不吃 \begin{matrix}...\end{matrix}，要先压成 \matrix{...}
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{matrix\}\s*(.*?)\s*\\end\{matrix\}",
        {
            param($match)
            $inner = [regex]::Replace($match.Groups[1].Value, "\s+", " ").Trim()
            return "\matrix{$inner}"
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{pmatrix\}\s*(.*?)\s*\\end\{pmatrix\}",
        {
            param($match)
            $inner = [regex]::Replace($match.Groups[1].Value, "\s+", " ").Trim()
            return "(\matrix{$inner})"
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{bmatrix\}\s*(.*?)\s*\\end\{bmatrix\}",
        {
            param($match)
            $inner = [regex]::Replace($match.Groups[1].Value, "\s+", " ").Trim()
            return "[\matrix{$inner}]"
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{Bmatrix\}\s*(.*?)\s*\\end\{Bmatrix\}",
        {
            param($match)
            $inner = [regex]::Replace($match.Groups[1].Value, "\s+", " ").Trim()
            return "\{\matrix{$inner}\}"
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{vmatrix\}\s*(.*?)\s*\\end\{vmatrix\}",
        {
            param($match)
            $inner = [regex]::Replace($match.Groups[1].Value, "\s+", " ").Trim()
            return "|\matrix{$inner}|"
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{Vmatrix\}\s*(.*?)\s*\\end\{Vmatrix\}",
        {
            param($match)
            $inner = [regex]::Replace($match.Groups[1].Value, "\s+", " ").Trim()
            return "\|\matrix{$inner}\|"
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{cases\}\s*(.*?)\s*\\end\{cases\}",
        {
            param($match)
            $inner = [regex]::Replace($match.Groups[1].Value, "\s+", " ").Trim()
            return "\left\{\matrix{$inner}\right."
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    return $normalized
}

function Convert-SimpleAccent {
    param(
        [string]$Text,
        [string]$Command,
        [string]$Accent
    )

    return [regex]::Replace(
        $Text,
        "\\$Command\{(\\[A-Za-z]+|[A-Za-z])\}",
        {
            param($match)
            $value = $match.Groups[1].Value
            switch ($value) {
                '\tau' { return ('{0}{1}' -f ([string][char]0x03C4), $Accent) }
                '\mu' { return ('{0}{1}' -f ([string][char]0x03BC), $Accent) }
                '\sigma' { return ('{0}{1}' -f ([string][char]0x03C3), $Accent) }
                '\theta' { return ('{0}{1}' -f ([string][char]0x03B8), $Accent) }
                '\delta' { return ('{0}{1}' -f ([string][char]0x03B4), $Accent) }
                default {
                    if ($value.StartsWith('\')) {
                        return $match.Value
                    }
                    return "$value$Accent"
                }
            }
        }
    )
}

function Normalize-LatexInlineText {
    param([string]$Text)

    $normalized = $Text.Trim()
    $normalized = $normalized -replace "\\varnothing", "\emptyset"
    $normalized = $normalized -replace "\\left", ""
    $normalized = $normalized -replace "\\right", ""
    $normalized = $normalized -replace "\\mathrm\{bird\}", "bird"
    $normalized = [regex]::Replace($normalized, '\\mathrm\{([^{}]+)\}', '$1')
    $normalized = [regex]::Replace($normalized, '\\text\{([^{}]+)\}', '$1')
    $normalized = Convert-SimpleAccent -Text $normalized -Command "hat" -Accent ([string][char]0x0302)
    $normalized = Convert-SimpleAccent -Text $normalized -Command "tilde" -Accent ([string][char]0x0303)
    $normalized = Convert-SimpleAccent -Text $normalized -Command "bar" -Accent ([string][char]0x0304)
    $normalized = $normalized -replace '\\tau', ([string][char]0x03C4)
    $normalized = $normalized -replace '\\mu', ([string][char]0x03BC)
    $normalized = $normalized -replace '\\sigma', ([string][char]0x03C3)
    $normalized = $normalized -replace '\\theta', ([string][char]0x03B8)
    $normalized = $normalized -replace '\\delta', ([string][char]0x03B4)
    $normalized = $normalized -replace '\\Omega', ([string][char]0x03A9)
    $normalized = $normalized -replace '\\emptyset', ([string][char]0x2205)
    $normalized = $normalized -replace '\\lfloor', ([string][char]0x230A)
    $normalized = $normalized -replace '\\rfloor', ([string][char]0x230B)
    $normalized = $normalized -replace '\\lceil', ([string][char]0x2308)
    $normalized = $normalized -replace '\\rceil', ([string][char]0x2309)
    $normalized = $normalized -replace '\\min', 'min'
    $normalized = $normalized -replace '\\max', 'max'
    $normalized = [regex]::Replace($normalized, '_\{([^{}]+)\}', '_($1)')
    $normalized = [regex]::Replace($normalized, '\^\{([^{}]+)\}', '^($1)')
    $normalized = [regex]::Replace($normalized, "\s+", " ").Trim()
    return $normalized
}

function Fallback-InlineFormulaText {
    param([string]$Text)

    $fallback = Normalize-LatexInlineText -Text $Text
    $fallback = $fallback -replace "\\", ""
    $fallback = [regex]::Replace($fallback, "\s+", " ").Trim()
    return $fallback
}

function Convert-RangeToWordMath {
    param(
        [object]$Document,
        [object]$Range,
        [string]$CleanText
    )

    $Range.Text = $CleanText
    [void]$Document.OMaths.Add($Range)
    if ($Range.OMaths.Count -gt 0) {
        $math = $Range.OMaths.Item(1)
        try {
            $math.BuildUp()
        } catch [System.NotImplementedException] {
            Write-Output ("[warn] BuildUp skipped for formula: " + $CleanText)
        } catch {
            Write-Output ("[warn] BuildUp failed for formula: " + $CleanText)
        }
    }
}

function Get-ParagraphContentRange {
    param([object]$Paragraph)

    $range = $Paragraph.Range.Duplicate
    if ($range.End -gt $range.Start) {
        $range.End = $range.End - 1
    }
    return $range
}

function Convert-LatexTextToWordMath {
    param([object]$Document)

    $displayPattern = [regex]::new('^\s*\$\$(.+?)\$\$\s*$', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $inlinePattern = [regex]::new('(?<!\$)\$([^\r\n\$]+?)\$(?!\$)')

    for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        $paragraphRange = Get-ParagraphContentRange -Paragraph $paragraph
        $paragraphText = [string]$paragraphRange.Text
        if ([string]::IsNullOrWhiteSpace($paragraphText)) {
            continue
        }

        $displayMatch = $displayPattern.Match($paragraphText)
        if ($displayMatch.Success) {
            $cleanText = Normalize-LatexBlockText -Text $displayMatch.Groups[1].Value
            if (-not [string]::IsNullOrWhiteSpace($cleanText)) {
                Convert-RangeToWordMath -Document $Document -Range $paragraphRange -CleanText $cleanText
            }
            continue
        }

        $inlineMatches = $inlinePattern.Matches($paragraphText)
        for ($j = $inlineMatches.Count - 1; $j -ge 0; $j--) {
            $match = $inlineMatches.Item($j)
            $formulaRange = $Document.Range(
                $paragraphRange.Start + $match.Index,
                $paragraphRange.Start + $match.Index + $match.Length
            )
            $originalText = $match.Groups[1].Value
            $cleanText = Normalize-LatexInlineText -Text $originalText
            if ([string]::IsNullOrWhiteSpace($cleanText)) {
                continue
            }
            if ($cleanText.Contains('\')) {
                $formulaRange.Text = Fallback-InlineFormulaText -Text $originalText
                continue
            }
            Convert-RangeToWordMath -Document $Document -Range $formulaRange -CleanText $cleanText
        }
    }
}

$word = $null
$document = $null

try {
    $InputMht = Resolve-WindowsPath -PathValue $InputMht
    $OutputDocx = Resolve-WindowsPath -PathValue $OutputDocx
    $ProcessedDocx = Resolve-WindowsPath -PathValue $ProcessedDocx

    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$Visible
    $word.DisplayAlerts = 0

    $document = $word.Documents.Open($InputMht)
    $document.SaveAs([ref]$OutputDocx, [ref]$wdFormatDocumentDefault)

    if ([string]::IsNullOrWhiteSpace($MacroName)) {
        Convert-LatexTextToWordMath -Document $document
    } else {
        $word.Run($MacroName)
    }

    $document.SaveAs([ref]$ProcessedDocx, [ref]$wdFormatDocumentDefault)
    Write-Output "[ok] raw_docx=$OutputDocx"
    Write-Output "[ok] processed_docx=$ProcessedDocx"
} catch {
    Write-Output ("[error] " + $_.Exception.GetType().FullName + ": " + $_.Exception.Message)
    if ($_.InvocationInfo) {
        if ($_.InvocationInfo.ScriptLineNumber) {
            Write-Output ("[error] line=" + $_.InvocationInfo.ScriptLineNumber)
        }
        if ($_.InvocationInfo.Line) {
            Write-Output ("[error] code=" + $_.InvocationInfo.Line.Trim())
        }
        if ($_.InvocationInfo.PositionMessage) {
            Write-Output ("[error] " + $_.InvocationInfo.PositionMessage)
        }
    }
    exit 1
} finally {
    Close-WordSession -Document $document -Word $word -SaveChanges $false
}
