$r = Import-Csv .\AllUsers_Tasks_v5.csv
Write-Host ('Total rows: ' + $r.Count) -ForegroundColor Cyan
Write-Host ''
Write-Host 'By Origin:' -ForegroundColor Cyan
$r | Group-Object Origin | Format-Table Count, Name -AutoSize | Out-String | Write-Host
Write-Host 'By SurfaceType:' -ForegroundColor Cyan
$r | Group-Object SurfaceType | Format-Table Count, Name -AutoSize | Out-String | Write-Host
Write-Host 'DOCSCAN rows:' -ForegroundColor Cyan
$r | Where-Object { $_.Origin -eq 'DOCSCAN' } | Format-Table AssigneeEmail, CreatorEmail, Title, SurfaceType, SourceRef, Notes -AutoSize -Wrap | Out-String -Width 300 | Write-Host
Write-Host 'Shared rows:' -ForegroundColor Cyan
$r | Where-Object { $_.Shared -eq 'True' } | Format-Table AssigneeEmail, CreatorEmail, Title, SurfaceType, Origin -AutoSize -Wrap | Out-String -Width 300 | Write-Host
Write-Host '---- Summary ----' -ForegroundColor Cyan
if (Test-Path .\AllUsers_Tasks_v5_Summary.csv) { Import-Csv .\AllUsers_Tasks_v5_Summary.csv | Format-Table -AutoSize -Wrap | Out-String -Width 300 | Write-Host } else { Write-Host 'No summary file.' }
