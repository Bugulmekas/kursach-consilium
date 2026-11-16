param(
    [Parameter(Mandatory = $true)][string]$Path,
    [switch]$Visible,
    [switch]$KeepOpen
)
# Update all fields + TOC twice, repaginate, save, and optionally close Word.
# Usage:
#   powershell -File update_fields.ps1 -Path "D:\work\final.docx"
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$word = New-Object -ComObject Word.Application
$word.Visible = [bool]$Visible
$doc = $word.Documents.Open($Path)

$null = $doc.Fields.Update()
$doc.Repaginate()
if ($doc.TablesOfContents.Count -gt 0) {
    $doc.TablesOfContents.Item(1).Update()
    $doc.Repaginate()
    $doc.TablesOfContents.Item(1).Update()
    $doc.Repaginate()
}
$null = $doc.Fields.Update()
$doc.Repaginate()

Write-Output ('sections: {0}, pages: {1}' -f $doc.Sections.Count, $doc.ComputeStatistics(2))
$doc.Save()

if (-not $KeepOpen) {
    $doc.Close($false)
    $word.Quit($false)
    Write-Output 'saved and closed'
} else {
    $doc.Activate()
    Write-Output 'saved, Word left open'
}
