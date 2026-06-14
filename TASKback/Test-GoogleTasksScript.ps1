<#
.SYNOPSIS
  Self-contained test harness for Get-GoogleTasksWithCreator.ps1.
  Builds a fake GAM shim that returns canned output, runs the script
  against it, and asserts the resulting CSV.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$script = Join-Path $root 'Get-GoogleTasksWithCreator.ps1'
$workDir = Join-Path ([IO.Path]::GetTempPath()) ("gamtest_" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

$fakeGamPs1 = Join-Path $workDir 'fake-gam.ps1'
$fakeGamCmd = Join-Path $workDir 'gam.cmd'
$outputCsv = Join-Path $workDir 'out.csv'

# ---- Build canned JSON rows for 'print tasks formatjson' ----
$taskDoc = [ordered]@{
    kind = 'tasks#task'; id = 't-doc-1'; title = 'Review Q4 plan'
    status = 'needsAction'; updated = '2026-01-10T10:00:00Z'
    webViewLink = 'https://tasks.google.com/task/doc-1'
    assignmentInfo = [ordered]@{
        linkToTask        = 'https://docs.google.com/document/d/DOC123/edit?disco=abc'
        surfaceType       = 'DOCUMENT'
        driveResourceInfo = [ordered]@{ driveFileId = 'DOC123' }
    }
} | ConvertTo-Json -Compress -Depth 6

$taskSpace = [ordered]@{
    kind = 'tasks#task'; id = 't-sp-1'; title = 'Follow up in chat'
    status = 'needsAction'; updated = '2026-01-11T09:00:00Z'
    webViewLink = 'https://tasks.google.com/task/sp-1'
    assignmentInfo = [ordered]@{
        linkToTask  = 'https://chat.google.com/room/SPACEXYZ/msg'
        surfaceType = 'SPACE'
        spaceInfo   = [ordered]@{ space = 'spaces/SPACEXYZ' }
    }
} | ConvertTo-Json -Compress -Depth 6

$taskSelf = [ordered]@{
    kind = 'tasks#task'; id = 't-self-1'; title = 'Buy milk'
    status = 'needsAction'; updated = '2026-01-12T08:00:00Z'
    webViewLink = 'https://tasks.google.com/task/self-1'
} | ConvertTo-Json -Compress -Depth 6

# GAM formatjson CSV: quote fields containing commas with ", and double internal "
function ConvertTo-CsvField([string]$s) { '"' + ($s -replace '"', '""') + '"' }
$taskCsv = @()
$taskCsv += 'User,tasklistId,id,taskId,title,JSON'
$taskCsv += ('alice@test.local,list-1,t-doc-1,list-1/t-doc-1,Review Q4 plan,{0}' -f (ConvertTo-CsvField $taskDoc))
$taskCsv += ('alice@test.local,list-1,t-sp-1,list-1/t-sp-1,Follow up in chat,{0}' -f (ConvertTo-CsvField $taskSpace))
$taskCsv += ('alice@test.local,list-1,t-self-1,list-1/t-self-1,Buy milk,{0}' -f (ConvertTo-CsvField $taskSelf))
$taskCsvPath = Join-Path $workDir 'tasks.csv'
$taskCsv -join "`r`n" | Out-File -FilePath $taskCsvPath -Encoding UTF8

$tasklistsCsv = @()
$tasklistsCsv += 'User,id,title,updated'
$tasklistsCsv += 'alice@test.local,list-1,My Tasks,2026-01-10T10:00:00Z'
$tasklistsCsvPath = Join-Path $workDir 'tasklists.csv'
$tasklistsCsv -join "`r`n" | Out-File -FilePath $tasklistsCsvPath -Encoding UTF8

# Canned filelist (alice owns one Doc: DOC123 already has an API task, DOC999 only has a comment).
$filelistCsv = @()
$filelistCsv += 'Owner,id,name,mimeType'
$filelistCsv += 'alice@test.local,DOC123,Q4 Plan.gdoc,application/vnd.google-apps.document'
$filelistCsv += 'alice@test.local,DOC999,Roadmap.gdoc,application/vnd.google-apps.document'
$filelistCsvPath = Join-Path $workDir 'filelist.csv'
$filelistCsv -join "`r`n" | Out-File -FilePath $filelistCsvPath -Encoding UTF8

# Two filecomments files keyed per file ID; one has an assignment, one is empty.
$commentsHeader = 'User,fileId,commentId,replyId,anchor,assigneeEmailAddress,author.displayName,author.me,content,createdTime,deleted,modifiedTime,quotedFileContent.mimeType,quotedFileContent.value,resolved'
$commentsDoc123 = @($commentsHeader) -join "`r`n"
$commentsDoc123Path = Join-Path $workDir 'filecomments_DOC123.csv'
$commentsDoc123 | Out-File -FilePath $commentsDoc123Path -Encoding UTF8

$commentsDoc999 = @()
$commentsDoc999 += $commentsHeader
$commentsDoc999 += 'alice@test.local,DOC999,cmt-1,,kix.abc,dave@test.local,Alice Author,True,@dave@test.local please review,2026-02-01T10:00:00Z,False,2026-02-01T10:00:00Z,text/html,Roadmap Q2 goals,False'
$commentsDoc999Path = Join-Path $workDir 'filecomments_DOC999.csv'
$commentsDoc999 -join "`r`n" | Out-File -FilePath $commentsDoc999Path -Encoding UTF8

$fileInfoTxt = @"
id: DOC123
name: Q4 Plan.gdoc
owners:
  emailAddress: docowner@test.local
  displayName: Doc Owner
"@
$fileInfoPath = Join-Path $workDir 'fileinfo.txt'
$fileInfoTxt | Out-File -FilePath $fileInfoPath -Encoding UTF8

$spaceInfoTxt = "name: spaces/SPACEXYZ`r`ndisplayName: Project Alpha Space"
$spaceInfoPath = Join-Path $workDir 'spaceinfo.txt'
$spaceInfoTxt | Out-File -FilePath $spaceInfoPath -Encoding UTF8

# Chat messages for the space: one matching "Created a task for @..." from bob,
# timestamped 1s after the SPACE task's updated value (2026-01-11T09:00:00Z),
# and one unrelated "Completed a task" noise event.
$chatCsv = @()
$chatCsv += 'User,space.name,name,thread.name,argumentText,createTime,sender.email,sender.type'
$chatCsv += 'alice@test.local,spaces/SPACEXYZ,spaces/SPACEXYZ/messages/m1,spaces/SPACEXYZ/threads/t1,Created a task for @Alice (via Tasks),2026-01-11T09:00:01Z,bob@test.local,HUMAN'
$chatCsv += 'alice@test.local,spaces/SPACEXYZ,spaces/SPACEXYZ/messages/m2,spaces/SPACEXYZ/threads/t2,Completed a task (via Tasks),2026-02-01T10:00:00Z,bob@test.local,HUMAN'
$chatCsv += 'alice@test.local,spaces/SPACEXYZ,spaces/SPACEXYZ/messages/m3,spaces/SPACEXYZ/threads/t3,Created a task (via Tasks),2026-03-05T14:00:00Z,carol@test.local,HUMAN'
$chatCsvPath = Join-Path $workDir 'chatmessages.csv'
$chatCsv -join "`r`n" | Out-File -FilePath $chatCsvPath -Encoding UTF8

# Canned chatspaces enumeration and members output for tenant scan tests.
$spacesCsv = @()
$spacesCsv += 'User,name,displayName,membershipCount.joinedDirectHumanUserCount'
$spacesCsv += 'alice@test.local,spaces/SPACEXYZ,Project Alpha Space,2'
$spacesCsvPath = Join-Path $workDir 'chatspaces.csv'
$spacesCsv -join "`r`n" | Out-File -FilePath $spacesCsvPath -Encoding UTF8

$membersCsv = @()
$membersCsv += 'User,space.name,name,member.email,member.type,role,state'
$membersCsv += 'alice@test.local,spaces/SPACEXYZ,spaces/SPACEXYZ/members/1,alice@test.local,HUMAN,ROLE_MEMBER,JOINED'
$membersCsvPath = Join-Path $workDir 'chatmembers.csv'
$membersCsv -join "`r`n" | Out-File -FilePath $membersCsvPath -Encoding UTF8

# ---- Fake GAM shim ----
$shim = @'
param()
$a = $args
if ($a.Count -ge 4 -and $a[0] -eq 'user' -and $a[2] -eq 'print' -and $a[3] -eq 'tasks') {
    Get-Content -Raw -LiteralPath "__TASKS__"; exit 0
}
if ($a.Count -ge 4 -and $a[0] -eq 'user' -and $a[2] -eq 'print' -and $a[3] -eq 'tasklists') {
    Get-Content -Raw -LiteralPath "__TASKLISTS__"; exit 0
}
if ($a.Count -ge 4 -and $a[0] -eq 'file' -and $a[2] -eq 'print' -and $a[3] -eq 'tasks') {
    Get-Content -Raw -LiteralPath "__TASKS__"; exit 0
}
if ($a.Count -ge 4 -and $a[0] -eq 'file' -and $a[2] -eq 'print' -and $a[3] -eq 'tasklists') {
    Get-Content -Raw -LiteralPath "__TASKLISTS__"; exit 0
}
if ($a.Count -ge 4 -and $a[0] -eq 'user' -and $a[2] -eq 'print' -and $a[3] -eq 'filelist') {
    Get-Content -Raw -LiteralPath "__FILELIST__"; exit 0
}
if ($a.Count -ge 5 -and $a[0] -eq 'user' -and $a[2] -eq 'print' -and $a[3] -eq 'filecomments') {
    $fid = $a[4]
    $p = "__COMMENTS_DIR__\filecomments_$fid.csv"
    if (Test-Path $p) { Get-Content -Raw -LiteralPath $p; exit 0 }
    # Empty but valid header so ConvertFrom-GamCsv returns 0 rows.
    "User,fileId,commentId,assigneeEmailAddress,author.displayName,author.me,content,createdTime,deleted,modifiedTime,resolved"; exit 0
}
if ($a.Count -ge 4 -and $a[0] -eq 'user' -and $a[2] -eq 'show' -and $a[3] -eq 'fileinfo') {
    Get-Content -Raw -LiteralPath "__FILEINFO__"; exit 0
}
if ($a.Count -ge 4 -and $a[0] -eq 'user' -and $a[2] -eq 'info' -and $a[3] -eq 'chatspace') {
    Get-Content -Raw -LiteralPath "__SPACEINFO__"; exit 0
}
if ($a.Count -ge 4 -and $a[0] -eq 'user' -and $a[2] -eq 'print' -and $a[3] -eq 'chatmessages') {
    Get-Content -Raw -LiteralPath "__CHATMSGS__"; exit 0
}
if ($a.Count -ge 4 -and $a[0] -eq 'user' -and $a[2] -eq 'print' -and $a[3] -eq 'chatspaces') {
    Get-Content -Raw -LiteralPath "__CHATSPACES__"; exit 0
}
if ($a.Count -ge 4 -and $a[0] -eq 'user' -and $a[2] -eq 'print' -and $a[3] -eq 'chatmembers') {
    Get-Content -Raw -LiteralPath "__CHATMEMBERS__"; exit 0
}
Write-Error ("fake-gam: unhandled args: " + ($a -join ' '))
exit 1
'@
$shim = $shim.Replace('__TASKS__', $taskCsvPath).Replace('__TASKLISTS__', $tasklistsCsvPath).Replace('__FILEINFO__', $fileInfoPath).Replace('__SPACEINFO__', $spaceInfoPath).Replace('__CHATMSGS__', $chatCsvPath).Replace('__CHATSPACES__', $spacesCsvPath).Replace('__CHATMEMBERS__', $membersCsvPath).Replace('__FILELIST__', $filelistCsvPath).Replace('__COMMENTS_DIR__', $workDir)
$shim | Out-File -FilePath $fakeGamPs1 -Encoding UTF8

# .cmd wrapper so the script can invoke "gam" as an executable
$cmd = "@echo off`r`npowershell -NoProfile -ExecutionPolicy Bypass -File `"$fakeGamPs1`" %*"
$cmd | Out-File -FilePath $fakeGamCmd -Encoding ASCII

Write-Host "Workdir: $workDir" -ForegroundColor DarkGray
Write-Host "Running script against fake GAM..." -ForegroundColor Cyan

$summaryCsv = Join-Path $workDir 'summary.csv'
& $script -Users 'alice@test.local' -OutputCsv $outputCsv -SummaryCsv $summaryCsv -GamPath $fakeGamCmd

if (-not (Test-Path $outputCsv)) { throw "Output CSV was not produced." }
$rows = Import-Csv $outputCsv
Write-Host ("Rows produced: {0}" -f $rows.Count) -ForegroundColor DarkGray
$rows | Format-Table AssigneeEmail, TaskId, SurfaceType, CreatorEmail, Shared, Origin, SourceRef, Notes -AutoSize | Out-String | Write-Host

$fail = @()
function Assert-Row($row, $expected, $label) {
    foreach ($k in $expected.Keys) {
        if ([string]$row.$k -ne [string]$expected[$k]) {
            $script:fail += "[{0}] {1}: expected '{2}', got '{3}'" -f $label, $k, $expected[$k], $row.$k
        }
    }
}
# Expect: 3 rows from main pipeline + 1 orphaned thread (t3) + 1 Docs-comment row (DOC999).
if ($rows.Count -ne 5) { $fail += "Expected 5 rows, got $($rows.Count)" }
$byId = @{}; foreach ($r in $rows) { $byId[$r.TaskId] = $r }

Assert-Row $byId['t-doc-1']  @{ AssigneeEmail = 'alice@test.local'; SurfaceType = 'DOCUMENT'; CreatorEmail = 'docowner@test.local'; CreatorName = 'Doc Owner'; SourceRef = 'DOC123'; Origin = 'GAM'; Shared = 'True'; TaskListId = 'list-1'; TaskListTitle = 'My Tasks' } 'doc'
Assert-Row $byId['t-sp-1']   @{ AssigneeEmail = 'alice@test.local'; SurfaceType = 'SPACE'; CreatorEmail = 'bob@test.local'; SourceRef = 'spaces/SPACEXYZ'; Origin = 'GAM'; Shared = 'True'; TaskListId = 'list-1'; TaskListTitle = 'My Tasks' } 'space'
Assert-Row $byId['t-self-1'] @{ AssigneeEmail = 'alice@test.local'; SurfaceType = 'SELF'; CreatorEmail = 'alice@test.local'; Origin = 'GAM'; Shared = 'False'; TaskListId = 'list-1'; TaskListTitle = 'My Tasks' } 'self'

$orphan = $rows | Where-Object { $_.Origin -eq 'SPACESCAN' -and $_.CreatorEmail -eq 'carol@test.local' } | Select-Object -First 1
if (-not $orphan) { $fail += "[orphan] No SPACESCAN row for carol's unassigned task" }
else {
    Assert-Row $orphan @{ SurfaceType = 'SPACE'; SourceRef = 'spaces/SPACEXYZ'; AssigneeEmail = ''; Status = 'unassigned'; Shared = 'False' } 'orphan'
    if ($orphan.Notes -notmatch 'Orphaned') { $fail += "[orphan] Notes missing 'Orphaned' marker" }
}

$docScan = $rows | Where-Object { $_.Origin -eq 'DOCSCAN' } | Select-Object -First 1
if (-not $docScan) { $fail += "[docscan] No DOCSCAN row emitted for DOC999 comment" }
else {
    Assert-Row $docScan @{ SurfaceType = 'DOCUMENT'; SourceRef = 'DOC999'; AssigneeEmail = 'dave@test.local'; CreatorEmail = 'alice@test.local'; Shared = 'True'; Status = 'needsAction'; TaskListTitle = '(Google Docs comment)' } 'docscan'
    if ($docScan.Notes -notmatch 'Roadmap') { $fail += "[docscan] Notes missing doc name" }
}

$docScanDup = $rows | Where-Object { $_.Origin -eq 'DOCSCAN' -and $_.SourceRef -eq 'DOC123' }
if ($docScanDup) { $fail += "[docscan] DOC123 should NOT produce a DOCSCAN row (already in API results)" }

if ($byId['t-sp-1'].Notes -notmatch 'Project Alpha Space') { $fail += "[space] Notes missing space name" }
if ($byId['t-sp-1'].Notes -notmatch 'correlation') { $fail += "[space] Notes missing correlation marker" }
if ($byId['t-doc-1'].Notes -notmatch 'Q4 Plan') { $fail += "[doc] Notes missing doc name" }

if (-not (Test-Path $summaryCsv)) { $fail += "Summary CSV was not produced." }
else {
    $summary = Import-Csv $summaryCsv
    $bobRow = $summary | Where-Object { $_.CreatorEmail -eq 'bob@test.local' } | Select-Object -First 1
    if (-not $bobRow) { $fail += "[summary] No row for bob@test.local" }
    elseif ([int]$bobRow.SharedCount -lt 1) { $fail += "[summary] bob SharedCount should be >=1" }
}

if ($fail.Count -gt 0) {
    Write-Host ""; Write-Host "FAIL:" -ForegroundColor Red
    $fail | ForEach-Object { Write-Host ("  - " + $_) -ForegroundColor Red }
    Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}
Write-Host "All assertions passed." -ForegroundColor Green
Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
