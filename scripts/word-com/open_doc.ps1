param(
    [Parameter(Mandatory = $true)][string]$Path,
    [switch]$Update,
    [int]$Zoom = 100,
    [switch]$Hidden
)
# Open a .docx in Word, optionally update fields/TOC twice and save.
# Usage:
#   powershell -File open_doc.ps1 -Path "D:\work\thesis.docx" -Update -Zoom 70
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$word = New-Object -ComObject Word.Application
$word.Visible = -not $Hidden
$doc = $word.Documents.Open($Path)
Start-Sleep -Seconds 2

if ($Update) {
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
    $doc.Save()
    Write-Output 'fields updated and saved'
}

$doc.Activate()
$word.ActiveWindow.ActivePane.View.Zoom.Percentage = $Zoom
Write-Output ('pages: {0}' -f $doc.ComputeStatistics(2))
