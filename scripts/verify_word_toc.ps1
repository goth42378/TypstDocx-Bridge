param([Parameter(Mandatory = $true)][string]$Docx)
$ErrorActionPreference = 'Stop'
$path = (Resolve-Path -LiteralPath $Docx).Path
function Convert-PageNumber([string]$text) {
    $value = $text.Trim()
    if ($value -match '^\d+$') { return [int]$value }
    if ($value -notmatch '^[IVXLCDM]+$') { throw "Unrecognized page number: $value" }
    $numerals = @{ I=1; V=5; X=10; L=50; C=100; D=500; M=1000 }
    $total = 0
    $last = 0
    for ($index = $value.Length - 1; $index -ge 0; $index--) {
        $number = $numerals[[string]$value[$index]]
        if ($number -lt $last) { $total -= $number } else { $total += $number; $last = $number }
    }
    return $total
}
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($path, $false, $true)
    if ($document.TablesOfContents.Count -ne 1) { throw 'Expected exactly one table of contents' }
    $toc = $document.TablesOfContents.Item(1)
    $checked = 0
    $mismatches = @()
    foreach ($field in $toc.Range.Fields) {
        if ($field.Code.Text -match 'PAGEREF\s+([^\s\\]+)') {
            $bookmark = $Matches[1]
            if (-not $document.Bookmarks.Exists($bookmark)) { throw "TOC bookmark missing: $bookmark" }
            $shown = Convert-PageNumber $field.Result.Text
            # wdActiveEndAdjustedPageNumber = 1: respect section page-number restarts.
            $actual = [int]$document.Bookmarks.Item($bookmark).Range.Information(1)
            if ($shown -ne $actual) { $mismatches += "$bookmark cached=$shown actual=$actual" }
            $checked++
        }
    }
    if ($checked -eq 0) { throw 'TOC has no PAGEREF fields' }
    $expected = 0
    foreach ($paragraph in $document.Paragraphs) {
        if ([int]$paragraph.OutlineLevel -le 3) { $expected++ }
    }
    if ($checked -ne $expected) { throw "TOC entry count $checked does not match $expected outline headings" }
    if ($mismatches.Count) { throw "TOC page mismatches: $($mismatches -join '; ')" }
    Write-Output "TOC entries checked: $checked / $expected headings; page mismatches: 0"
}
finally {
    if ($null -ne $document) {
        $document.Close($false)
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
    }
    if ($null -ne $word) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    }
}
