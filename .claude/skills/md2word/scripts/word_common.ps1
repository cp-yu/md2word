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

function Close-WordSession {
    param(
        [object]$Document,
        [object]$Word,
        $SaveChanges = $false
    )

    if ($null -ne $Document) {
        try {
            $Document.Close([ref]$SaveChanges)
        } catch {
        }
    }
    if ($null -ne $Word) {
        try {
            $Word.Quit()
        } catch {
        }
    }

    Release-ComObject -ComObject $Document
    Release-ComObject -ComObject $Word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
