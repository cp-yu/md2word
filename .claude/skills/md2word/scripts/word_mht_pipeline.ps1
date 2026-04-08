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

function Normalize-MatrixInnerText {
    param([string]$Text)

    $normalized = [regex]::Replace($Text, "\s+", " ").Trim()
    # Word 线性格式里，矩阵/分段公式的行分隔符使用 @，不是 LaTeX 的 \\
    $normalized = [regex]::Replace($normalized, '\s*\\\\\s*', '@')
    $normalized = [regex]::Replace($normalized, '\s*&\s*', '&')
    return $normalized
}

function Normalize-LatexBlockText {
    param([string]$Text)

    $normalized = $Text -replace "`r", " " -replace "`n", " "
    $normalized = [regex]::Replace($normalized, "\s+", " ").Trim()
    # Word 的 LaTeX/BuildUp 不吃 \begin{matrix}...\end{matrix}，要先压成 \matrix{...}
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{matrix\}\s*(.*?)\s*\\end\{matrix\}",
        {
            param($match)
            $inner = Normalize-MatrixInnerText -Text $match.Groups[1].Value
            return "\matrix{$inner}"
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{pmatrix\}\s*(.*?)\s*\\end\{pmatrix\}",
        {
            param($match)
            $inner = Normalize-MatrixInnerText -Text $match.Groups[1].Value
            return "(\matrix{$inner})"
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{bmatrix\}\s*(.*?)\s*\\end\{bmatrix\}",
        {
            param($match)
            $inner = Normalize-MatrixInnerText -Text $match.Groups[1].Value
            return "[\matrix{$inner}]"
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{Bmatrix\}\s*(.*?)\s*\\end\{Bmatrix\}",
        {
            param($match)
            $inner = Normalize-MatrixInnerText -Text $match.Groups[1].Value
            return "\{\matrix{$inner}\}"
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{vmatrix\}\s*(.*?)\s*\\end\{vmatrix\}",
        {
            param($match)
            $inner = Normalize-MatrixInnerText -Text $match.Groups[1].Value
            return "|\matrix{$inner}|"
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{Vmatrix\}\s*(.*?)\s*\\end\{Vmatrix\}",
        {
            param($match)
            $inner = Normalize-MatrixInnerText -Text $match.Groups[1].Value
            return "\|\matrix{$inner}\|"
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $normalized = [regex]::Replace(
        $normalized,
        "\\begin\{cases\}\s*(.*?)\s*\\end\{cases\}",
        {
            param($match)
            $inner = Normalize-MatrixInnerText -Text $match.Groups[1].Value
            return "\left\{\matrix{$inner}\right."
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $normalized = Normalize-CommonLatexText -Text $normalized
    $normalized = [regex]::Replace($normalized, "\s+", " ").Trim()
    return $normalized
}

function Normalize-CommonLatexText {
    param([string]$Text)

    $normalized = $Text.Trim()
    $normalized = $normalized.Replace('\qquad', ' ')
    $normalized = $normalized.Replace('\quad', ' ')
    $normalized = $normalized.Replace('\,', ' ')
    $normalized = $normalized.Replace('\:', ' ')
    $normalized = $normalized.Replace('\;', ' ')
    $normalized = $normalized.Replace('\!', '')
    $normalized = $normalized.Replace('\ ', ' ')
    $normalized = $normalized.Replace('\left\|', ([string][char]0x2016))
    $normalized = $normalized.Replace('\right\|', ([string][char]0x2016))
    $normalized = $normalized.Replace('\left|', '|')
    $normalized = $normalized.Replace('\right|', '|')
    $normalized = $normalized.Replace('\left(', '(')
    $normalized = $normalized.Replace('\right)', ')')
    $normalized = $normalized.Replace('\left[', '[')
    $normalized = $normalized.Replace('\right]', ']')
    $normalized = $normalized -replace "\\varnothing", "\emptyset"
    $normalized = $normalized -replace "\\mathrm\{bird\}", "bird"
    $normalized = [regex]::Replace($normalized, '\\mathrm\{([^{}]+)\}', '$1')
    $normalized = [regex]::Replace($normalized, '\\mathbf\{([^{}]+)\}', '$1')
    $normalized = [regex]::Replace($normalized, '\\mathit\{([^{}]+)\}', '$1')
    $normalized = [regex]::Replace($normalized, '\\text\{([^{}]+)\}', '$1')

    $commandMap = @(
        @('\rightarrow', ([string][char]0x2192)),
        @('\leftarrow', ([string][char]0x2190)),
        @('\varepsilon', ([string][char]0x03F5)),
        @('\varphi', ([string][char]0x03D5)),
        @('\epsilon', ([string][char]0x03B5)),
        @('\alpha', ([string][char]0x03B1)),
        @('\beta', ([string][char]0x03B2)),
        @('\gamma', ([string][char]0x03B3)),
        @('\delta', ([string][char]0x03B4)),
        @('\eta', ([string][char]0x03B7)),
        @('\theta', ([string][char]0x03B8)),
        @('\lambda', ([string][char]0x03BB)),
        @('\mu', ([string][char]0x03BC)),
        @('\pi', ([string][char]0x03C0)),
        @('\sigma', ([string][char]0x03C3)),
        @('\tau', ([string][char]0x03C4)),
        @('\phi', ([string][char]0x03C6)),
        @('\omega', ([string][char]0x03C9)),
        @('\Omega', ([string][char]0x03A9)),
        @('\emptyset', ([string][char]0x2205)),
        @('\infty', ([string][char]0x221E)),
        @('\neq', ([string][char]0x2260)),
        @('\leq', ([string][char]0x2264)),
        @('\geq', ([string][char]0x2265)),
        @('\le', ([string][char]0x2264)),
        @('\ge', ([string][char]0x2265)),
        @('\pm', ([string][char]0x00B1)),
        @('\times', ([string][char]0x00D7)),
        @('\cdot', ([string][char]0x00B7)),
        @('\circ', ([string][char]0x2218)),
        @('\to', ([string][char]0x2192)),
        @('\lfloor', ([string][char]0x230A)),
        @('\rfloor', ([string][char]0x230B)),
        @('\lceil', ([string][char]0x2308)),
        @('\rceil', ([string][char]0x2309)),
        @('\min', 'min'),
        @('\max', 'max')
    )

    foreach ($entry in $commandMap) {
        $pattern = [regex]::Escape($entry[0]) + '(?![A-Za-z])'
        $replacement = [string]$entry[1]
        $normalized = [regex]::Replace(
            $normalized,
            $pattern,
            [System.Text.RegularExpressions.MatchEvaluator]{
                param($match)
                return $replacement
            }
        )
    }

    return $normalized
}

function Has-UnsupportedLatexCommand {
    param([string]$Text)

    $allowedCommands = @('\bar', '\hat', '\tilde')
    $matches = [regex]::Matches($Text, '\\[A-Za-z]+')
    foreach ($match in $matches) {
        if ($allowedCommands -notcontains $match.Value) {
            return $true
        }
    }
    return $false
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

    $normalized = Normalize-CommonLatexText -Text $Text
    $normalized = $normalized -replace "\\left", ""
    $normalized = $normalized -replace "\\right", ""
    $normalized = Convert-SimpleAccent -Text $normalized -Command "hat" -Accent ([string][char]0x0302)
    $normalized = Convert-SimpleAccent -Text $normalized -Command "tilde" -Accent ([string][char]0x0303)
    $normalized = Convert-SimpleAccent -Text $normalized -Command "bar" -Accent ([string][char]0x0304)
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
            if (Has-UnsupportedLatexCommand -Text $cleanText) {
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
