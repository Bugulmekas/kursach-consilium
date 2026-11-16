param()
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
for ($i = $word.Documents.Count; $i -ge 1; $i--) {
  $d = $word.Documents.Item($i)
  Write-Output ('closing: ' + $d.Name + ' saved=' + $d.Saved)
  $d.Close(0)
}
$word.Quit()
Start-Sleep -Seconds 3
Write-Output 'closed all'
