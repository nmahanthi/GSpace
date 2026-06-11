$r = Import-Csv .\AllUsers_Tasks_v6.csv
Write-Host ('Total rows: ' + $r.Count) -ForegroundColor Cyan
Write-Host ''
Write-Host 'By Origin:' -ForegroundColor Cyan
$r | Group-Object Origin | Format-Table Count, Name -AutoSize | Out-String | Write-Host
Write-Host 'DOCSCAN row (full):' -ForegroundColor Cyan
$r | Where-Object { $_.Origin -eq 'DOCSCAN' } | Format-List AssigneeEmail, TaskListId, TaskListTitle, TaskId, Title, SurfaceType, SourceRef, CreatorEmail, Shared, Origin
Write-Host 'All shared rows (tasklist cols):' -ForegroundColor Cyan
$r | Where-Object { $_.Shared -eq 'True' } | Format-Table AssigneeEmail, CreatorEmail, TaskListId, TaskListTitle, Title, SurfaceType, Origin -AutoSize -Wrap | Out-String -Width 300 | Write-Host
