param([int]$Zoom = 60)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
$word.ActiveWindow.ActivePane.View.Zoom.Percentage = $Zoom
Write-Output 'zoomed'
