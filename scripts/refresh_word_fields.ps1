param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [string]$Pdf
)
$ErrorActionPreference = 'Stop'
$inputPath = (Resolve-Path -LiteralPath $Docx).Path
$pdfPath = $null
if ($Pdf) {
    $pdfPath = [System.IO.Path]::GetFullPath($Pdf)
    if (Test-Path -LiteralPath $pdfPath) { throw "PDF output already exists: $pdfPath" }
    if (-not (Test-Path -LiteralPath (Split-Path -Parent $pdfPath))) { throw 'PDF output directory does not exist' }
}
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($inputPath, $false, $false)
    # Two passes let forward REF fields pick up targets updated later in pass one.
    foreach ($pass in 1..2) {
        $document.Fields.Update() | Out-Null
        foreach ($section in $document.Sections) {
            foreach ($header in $section.Headers) {
                if ($header.Exists) { $header.Range.Fields.Update() | Out-Null }
            }
            foreach ($footer in $section.Footers) {
                if ($footer.Exists) { $footer.Range.Fields.Update() | Out-Null }
            }
        }
        $document.Repaginate()
    }
    $document.Save()
    if ($pdfPath) { $document.ExportAsFixedFormat($pdfPath, 17) }
    Write-Output "Updated fields: $inputPath"
    if ($pdfPath) { Write-Output "QA PDF: $pdfPath" }
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
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
