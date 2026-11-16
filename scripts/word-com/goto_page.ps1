param(
    [int]$Page = 1,
    [string]$DocumentName = "",
    [int]$Zoom = 70
)
# Go to a page in an already open Word document.
# Usage:
#   powershell -File goto_page.ps1 -Page 2
#   powershell -File goto_page.ps1 -DocumentName "thesis" -Page 3 -Zoom 80
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
$word.Visible = $true
$word.Activate()

$doc = $null
if ($DocumentName) {
    for ($i = 1; $i -le $word.Documents.Count; $i++) {
        $d = $word.Documents.Item($i)
        Write-Output ('doc: ' + $d.Name)
        if ($d.Name -like "*$DocumentName*") { $doc = $d }
    }
} else {
    $doc = $word.ActiveDocument
}

if ($doc -eq $null) { Write-Output 'document not found'; exit 1 }

$doc.Activate()
Start-Sleep -Milliseconds 500
$word.ActiveWindow.ActivePane.View.Zoom.Percentage = $Zoom
$null = $word.Selection.GoTo(1, 1, $Page)
Start-Sleep -Milliseconds 1200
Write-Output ('pages: {0}; at page {1}' -f $doc.ComputeStatistics(2), $word.Selection.Information(3))
