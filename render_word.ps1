param(
    [string]$inputDoc = 'D:\project\innovation\proposal_hamrah_entekhab.docx',
    [string]$outputPdf = 'D:\project\innovation\proposal_render\proposal_hamrah_entekhab.pdf'
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Path (Split-Path -Parent $outputPdf) -Force | Out-Null
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($inputDoc, $false, $true)
    $doc.Fields.Update() | Out-Null
    $doc.SaveAs([ref]$outputPdf, [ref]17)
    $doc.Close($false)
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
}
Write-Output $outputPdf
