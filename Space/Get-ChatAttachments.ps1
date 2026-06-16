# =============================================================================
# Get-ChatAttachments.ps1
# Audits Google Chat attachment count + size per Space / Group Chat
#
# Phase 1  Pull all Spaces (asadmin) -> deduplicate -> pick one member runner
# Phase 2  Per-space: gam print chatmessages fields attachment -> parse rows
# Phase 3a Drive-source attachments -> gam show fileinfo -> real byte count
# Phase 3b Uploaded/inline attachments -> HEAD on downloadUri -> Content-Length
#
# Outputs
#   CHAT_ATTACH_DETAIL_<ts>.csv      one row per attachment (all spaces)
#   CHAT_ATTACH_PER_SPACE_<ts>.csv   aggregated totals per space
#   CHAT_ATTACH_TOP_LARGEST_<ts>.csv top 100 largest attachments
# =============================================================================

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
$AdminEmail          = "admin-narendra@rocheua.com"
$GamPath             = Join-Path (Split-Path $PSScriptRoot) "gam7\gam.exe"
$OutputDir           = $PSScriptRoot
$Timestamp           = Get-Date -Format "yyyyMMdd_HHmmss"
$TimeWindowDays      = 90          # how far back to scan messages
$LargeFileMB         = 10         # flag threshold for "large" attachments
$IncludeGroupChats   = $true      # include GROUP_CHAT spaces (not only SPACE)
$HeadTimeoutSec      = 10         # timeout per HEAD request for upload sizing
$MaxHeadRequests     = 500        # cap to avoid rate-limit on large tenants

# ── EXISTING SPACE LIST (reuse if fresh, else re-pull) ────────────────────────
# chat_spaces_target.csv produced by the BOT/CHAT run already has User + name
# + spaceType which is exactly what we need.
$ExistingSpaceCsv    = Join-Path (Split-Path $PSScriptRoot) `
    "gam_exports\Bot&Chat App related_usingGAM\RUN_BOTCHAT_20260615_154743Z\chat_spaces_target.csv"

# ── OUTPUT PATHS ──────────────────────────────────────────────────────────────
$TmpMsgCsv           = Join-Path $OutputDir "tmp_chatmsgs_$Timestamp.csv"
$DetailCsv           = Join-Path $OutputDir "CHAT_ATTACH_DETAIL_$Timestamp.csv"
$PerSpaceCsv         = Join-Path $OutputDir "CHAT_ATTACH_PER_SPACE_$Timestamp.csv"
$TopCsv              = Join-Path $OutputDir "CHAT_ATTACH_TOP_LARGEST_$Timestamp.csv"

Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host "  Google Chat Attachment Audit (GAM7 Hybrid)" -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host "Admin       : $AdminEmail"
Write-Host "Window      : Last $TimeWindowDays days"
Write-Host "Large flag  : > $LargeFileMB MB"
Write-Host "Timestamp   : $Timestamp"
Write-Host ""

# ── HELPER: wait for a file to appear ─────────────────────────────────────────
function Wait-ForFile($path, $timeoutSec = 300) {
    $elapsed = 0
    Write-Host "   Waiting for GAM output" -NoNewline
    while (-not (Test-Path $path) -or (Get-Item $path).Length -eq 0) {
        if ($elapsed -ge $timeoutSec) {
            Write-Host "`n   ERROR: timed out after $timeoutSec s" -ForegroundColor Red; exit 1
        }
        Start-Sleep -Seconds 3; $elapsed += 3; Write-Host "." -NoNewline
    }
    Write-Host " OK" -ForegroundColor Green
}

# ── HELPER: safe column reader ─────────────────────────────────────────────────
function Get-Col($row, [string[]]$names) {
    foreach ($n in $names) {
        $v = $row.PSObject.Properties[$n]
        if ($v -and "$($v.Value)".Trim() -ne "") { return "$($v.Value)".Trim() }
    }
    return ""
}

# =============================================================================
# PHASE 1 – Build unique space list with a runner user per space
# =============================================================================
Write-Host "[1/3] Building space list..." -ForegroundColor Yellow

$spaceRows = $null
if (Test-Path $ExistingSpaceCsv) {
    $spaceRows = Import-Csv $ExistingSpaceCsv
    Write-Host "   Loaded existing space CSV: $($spaceRows.Count) rows"
} else {
    Write-Host "   Existing CSV not found - pulling fresh from GAM..." -ForegroundColor DarkYellow
    $tmpSpaces = Join-Path $OutputDir "tmp_spaces_$Timestamp.csv"
    & $GamPath redirect csv $tmpSpaces user $AdminEmail print chatspaces asadmin `
        fields "name,displayname,spacetype,membershipcount,createtime,lastactivetime"
    Wait-ForFile $tmpSpaces
    $spaceRows = Import-Csv $tmpSpaces
    Remove-Item $tmpSpaces -Force -ErrorAction SilentlyContinue
    Write-Host "   Pulled $($spaceRows.Count) space rows"
}

# Deduplicate: one runner per unique space name
$allowedTypes = @("SPACE")
if ($IncludeGroupChats) { $allowedTypes += "GROUP_CHAT" }

$spacemap = @{}   # key = spaces/XXX -> @{ Runner; DisplayName; SpaceType; MemberCount; SpaceUri }
foreach ($r in $spaceRows) {
    $sname = (Get-Col $r "name").Trim()
    $stype = (Get-Col $r "spaceType","spacetype").Trim().ToUpper()
    if (-not $sname -or $spacemap.ContainsKey($sname)) { continue }
    if ($stype -notin $allowedTypes) { continue }
    $runner = Get-Col $r "User","TargetUser"
    if (-not $runner) { $runner = $AdminEmail }
    $spacemap[$sname] = [ordered]@{
        Runner      = $runner
        DisplayName = Get-Col $r "displayName","displayname"
        SpaceType   = $stype
        MemberCount = Get-Col $r "membershipCount.joinedDirectHumanUserCount","membershipcount"
        SpaceUri    = Get-Col $r "spaceUri","spaceuri"
        CreateTime  = Get-Col $r "createTime","createtime"
        LastActive  = Get-Col $r "lastActiveTime","lastactivetime"
    }
}
$uniqueSpaces = $spacemap.Keys | Sort-Object
Write-Host "   Unique spaces to scan : $($uniqueSpaces.Count)" -ForegroundColor Green

# =============================================================================
# PHASE 2 – Per-space: pull messages, parse attachment columns
# =============================================================================
Write-Host ""
Write-Host "[2/3] Pulling messages with attachments per space..." -ForegroundColor Yellow

# Time filter: ISO-8601, UTC
$filterTime  = (Get-Date).AddDays(-$TimeWindowDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$msgFilter   = "createTime > `"$filterTime`""

$detailRows  = [System.Collections.Generic.List[PSCustomObject]]::new()
$driveIds    = [System.Collections.Generic.List[PSCustomObject]]::new()   # for Phase 3a
$uploadedRows= [System.Collections.Generic.List[PSCustomObject]]::new()   # for Phase 3b

$spaceIdx = 0
foreach ($sname in $uniqueSpaces) {
    $spaceIdx++
    $meta   = $spacemap[$sname]
    $runner = $meta.Runner
    $disp   = if ($meta.DisplayName) { $meta.DisplayName } else { $sname }

    Write-Host ("  [{0,3}/{1}] {2} ({3}) as {4}" -f `
        $spaceIdx, $uniqueSpaces.Count, $disp, $meta.SpaceType, $runner) -ForegroundColor DarkCyan

    # Remove stale temp file
    if (Test-Path $TmpMsgCsv) { Remove-Item $TmpMsgCsv -Force }

    # GAM: print chatmessages — no formatjson so GAM flattens attachment.N.* columns
    $gamArgs = @(
        "redirect", "csv", $TmpMsgCsv,
        "user",     $runner,
        "print",    "chatmessages", $sname,
        "filter",   $msgFilter,
        "fields",   "name,createtime,sender,attachment,attachedgifs"
    )
    & $GamPath @gamArgs 2>$null

    if (-not (Test-Path $TmpMsgCsv) -or (Get-Item $TmpMsgCsv).Length -lt 10) {
        Write-Host "      (no messages or access denied)" -ForegroundColor DarkGray
        continue
    }

    $msgs = Import-Csv $TmpMsgCsv
    if (-not $msgs -or $msgs.Count -eq 0) { continue }

    # Find all attachment index slots present in this batch (0,1,2,...)
    $sampleCols = $msgs[0].PSObject.Properties.Name
    $attachIdxs = $sampleCols |
        Where-Object { $_ -match '^attachment\.(\d+)\.' } |
        ForEach-Object { [int]($_ -replace '^attachment\.(\d+)\..*','$1') } |
        Sort-Object -Unique

    foreach ($msg in $msgs) {
        $msgName    = Get-Col $msg "name"
        $createTime = Get-Col $msg "createTime","createtime"
        $sender     = Get-Col $msg "sender.name","sender.email","sender"
        $senderDisp = Get-Col $msg "sender.displayName","sender.displayname"

        # ── Attached Gifs (count only, no byte size available) ───────────────
        $gifCols = $msg.PSObject.Properties.Name | Where-Object { $_ -match '^attachedGifs\.' }
        $gifCount = ($gifCols | ForEach-Object { $_ -replace '^attachedGifs\.(\d+)\..*','$1' } |
                     Sort-Object -Unique).Count

        if ($gifCount -gt 0) {
            $detailRows.Add([PSCustomObject]@{
                SpaceID        = $sname
                SpaceName      = $disp
                SpaceType      = $meta.SpaceType
                MemberCount    = $meta.MemberCount
                MessageID      = $msgName
                CreateTime     = $createTime
                Sender         = $senderDisp
                SenderID       = $sender
                AttachIndex    = "gifs"
                AttachName     = "Animated GIF(s)"
                ContentType    = "image/gif"
                Source         = "GIPHY_TENOR"
                IsInlineImage  = $true
                DriveFileID    = ""
                DownloadUri    = ""
                Bytes          = 0
                SizeMB         = 0
                IsLarge        = $false
                Note           = "$gifCount gif(s) - URL only, no size"
            })
        }

        # ── File attachments ─────────────────────────────────────────────────
        foreach ($idx in $attachIdxs) {
            $prefix      = "attachment.$idx"
            $aName       = Get-Col $msg "$prefix.name"
            if (-not $aName) { continue }   # slot empty for this message

            $contentName = Get-Col $msg "$prefix.contentName"
            $contentType = Get-Col $msg "$prefix.contentType"
            $source      = Get-Col $msg "$prefix.source"
            $driveId     = Get-Col $msg "$prefix.driveDataRef.driveFileId"
            $dlUri       = Get-Col $msg "$prefix.downloadUri"
            $isImage     = $contentType -match "^image/"

            $detRow = [PSCustomObject]@{
                SpaceID        = $sname
                SpaceName      = $disp
                SpaceType      = $meta.SpaceType
                MemberCount    = $meta.MemberCount
                MessageID      = $msgName
                CreateTime     = $createTime
                Sender         = $senderDisp
                SenderID       = $sender
                AttachIndex    = $idx
                AttachName     = $contentName
                ContentType    = $contentType
                Source         = $source
                IsInlineImage  = $isImage
                DriveFileID    = $driveId
                DownloadUri    = $dlUri
                Bytes          = 0           # filled in Phase 3
                SizeMB         = 0
                IsLarge        = $false
                Note           = ""
            }
            $detailRows.Add($detRow)

            if ($source -eq "DRIVE_FILE" -and $driveId) {
                $driveIds.Add([PSCustomObject]@{ DriveFileID = $driveId; Row = $detRow })
            } elseif ($source -eq "UPLOADED_CONTENT" -and $dlUri) {
                $uploadedRows.Add([PSCustomObject]@{ Uri = $dlUri; Row = $detRow })
            }
        }
    }
    Remove-Item $TmpMsgCsv -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "   Total attachment detail rows : $($detailRows.Count)" -ForegroundColor Green
Write-Host "   Drive-source attachments     : $($driveIds.Count)"
Write-Host "   Uploaded/inline attachments  : $($uploadedRows.Count)"

# =============================================================================
# PHASE 3a – Enrich Drive-source attachments with real byte size
#            gam user admin show fileinfo <id> fields id,name,size,mimetype
# =============================================================================
Write-Host ""
Write-Host "[3a/3] Enriching Drive attachment sizes..." -ForegroundColor Yellow

if ($driveIds.Count -gt 0) {
    # De-duplicate Drive IDs (same file may be shared in multiple spaces)
    $uniqueDriveIds = $driveIds | Select-Object -ExpandProperty DriveFileID -Unique
    Write-Host "   Unique Drive file IDs to query : $($uniqueDriveIds.Count)"

    $driveSizeMap = @{}   # fileId -> bytes

    $batchSize = 50
    $batches   = [Math]::Ceiling($uniqueDriveIds.Count / $batchSize)

    for ($b = 0; $b -lt $batches; $b++) {
        $slice = $uniqueDriveIds | Select-Object -Skip ($b * $batchSize) -First $batchSize
        Write-Host ("   Batch {0}/{1} ({2} IDs)" -f ($b+1), $batches, $slice.Count) -ForegroundColor DarkGray

        $tmpDriveCsv = Join-Path $OutputDir "tmp_drive_$Timestamp`_$b.csv"

        # Build a parentId query using id: clauses joined by OR
        # GAM filelist query supports: id = 'X' or id = 'Y' ...
        $idQuery = ($slice | ForEach-Object { "id = '$_'" }) -join " or "

        $gamDriveArgs = @(
            "redirect", "csv", $tmpDriveCsv,
            "user",     $AdminEmail,
            "print",    "filelist",
            "query",    $idQuery,
            "fields",   "id,name,size,mimetype,owners",
            "allfields", "false"
        )
        & $GamPath @gamDriveArgs 2>$null

        if (Test-Path $tmpDriveCsv) {
            $driveRows = Import-Csv $tmpDriveCsv
            foreach ($dr in $driveRows) {
                $fid   = Get-Col $dr "id"
                $bytes = [long]0
                $raw   = Get-Col $dr "size"
                if ($raw -match '^\d+$') { $bytes = [long]$raw }
                if ($fid) { $driveSizeMap[$fid] = $bytes }
            }
            Remove-Item $tmpDriveCsv -Force -ErrorAction SilentlyContinue
        }
    }

    # Write sizes back to detail rows
    $notFound = 0
    foreach ($entry in $driveIds) {
        $fid = $entry.DriveFileID
        if ($driveSizeMap.ContainsKey($fid)) {
            $bytes = $driveSizeMap[$fid]
            $entry.Row.Bytes  = $bytes
            $entry.Row.SizeMB = [Math]::Round($bytes / 1MB, 2)
            $entry.Row.IsLarge= ($bytes -gt ($LargeFileMB * 1MB))
            if (-not $driveSizeMap[$fid]) { $entry.Row.Note = "size=0 - may be Google Workspace native file" }
        } else {
            $notFound++
            $entry.Row.Note = "Drive file not accessible by admin"
        }
    }

    Write-Host "   Drive sizes resolved : $($driveSizeMap.Count) / $($uniqueDriveIds.Count)" -ForegroundColor Green
    if ($notFound -gt 0) {
        Write-Host "   Not accessible (Shared Drive / permission) : $notFound" -ForegroundColor DarkYellow
    }
} else {
    Write-Host "   No Drive attachments found - skipping." -ForegroundColor DarkGray
}

# =============================================================================
# PHASE 3b – Size uploaded/inline attachments via HTTP HEAD -> Content-Length
# =============================================================================
Write-Host ""
Write-Host "[3b/3] Sizing uploaded/inline attachments via HEAD requests..." -ForegroundColor Yellow

$headDone = 0; $headFailed = 0; $headSkipped = 0

if ($uploadedRows.Count -gt 0) {
    $toHead = $uploadedRows
    if ($toHead.Count -gt $MaxHeadRequests) {
        Write-Host ("   Capping HEAD requests at {0} (of {1} total) to avoid rate-limit" `
            -f $MaxHeadRequests, $toHead.Count) -ForegroundColor DarkYellow
        $toHead = $toHead | Select-Object -First $MaxHeadRequests
        $headSkipped = $uploadedRows.Count - $MaxHeadRequests
    }

    # Use a single HttpClient for efficiency (faster than Invoke-WebRequest per call)
    Add-Type -AssemblyName System.Net.Http
    $handler = [System.Net.Http.HttpClientHandler]::new()
    $handler.AllowAutoRedirect = $true
    $client  = [System.Net.Http.HttpClient]::new($handler)
    $client.Timeout = [TimeSpan]::FromSeconds($HeadTimeoutSec)

    $hi = 0
    foreach ($entry in $toHead) {
        $hi++
        if ($hi % 50 -eq 0) {
            Write-Host ("   HEAD {0}/{1}..." -f $hi, $toHead.Count) -ForegroundColor DarkGray
        }
        try {
            $req = [System.Net.Http.HttpRequestMessage]::new(
                        [System.Net.Http.HttpMethod]::Head, $entry.Uri)
            $resp = $client.SendAsync($req).GetAwaiter().GetResult()
            $cl   = $resp.Content.Headers.ContentLength
            if ($cl -and $cl -gt 0) {
                $bytes = [long]$cl
                $entry.Row.Bytes   = $bytes
                $entry.Row.SizeMB  = [Math]::Round($bytes / 1MB, 2)
                $entry.Row.IsLarge = ($bytes -gt ($LargeFileMB * 1MB))
                $headDone++
            } else {
                $entry.Row.Note = "HEAD: no Content-Length header"
                $headFailed++
            }
        } catch {
            $entry.Row.Note = "HEAD failed: $($_.Exception.Message -replace '`n',' ')"
            $headFailed++
        }
    }
    $client.Dispose()

    Write-Host "   HEAD succeeded : $headDone"   -ForegroundColor Green
    Write-Host "   HEAD failed    : $headFailed"  -ForegroundColor $(if($headFailed -gt 0){"DarkYellow"}else{"Green"})
    if ($headSkipped -gt 0) {
        Write-Host "   HEAD skipped (cap) : $headSkipped - re-run with higher MaxHeadRequests" -ForegroundColor DarkYellow
    }
} else {
    Write-Host "   No uploaded attachments found - skipping." -ForegroundColor DarkGray
}

# =============================================================================
# OUTPUT – Export CSVs
# =============================================================================
Write-Host ""
Write-Host "Exporting reports..." -ForegroundColor Yellow

# 1. Full detail (one row per attachment)
$detailRows | Sort-Object Bytes -Descending |
    Export-Csv $DetailCsv -NoTypeInformation
Write-Host "   Detail CSV       -> $DetailCsv"

# 2. Per-Space aggregated summary
$perSpaceRows = $detailRows |
    Where-Object { $_.Source -ne "GIPHY_TENOR" } |   # exclude pure-gif rows from byte totals
    Group-Object -Property SpaceID |
    ForEach-Object {
        $grp   = $_.Group
        $first = $grp[0]
        $totalBytes    = ($grp | Measure-Object Bytes -Sum).Sum
        $driveBytes    = ($grp | Where-Object { $_.Source -eq "DRIVE_FILE"       } | Measure-Object Bytes -Sum).Sum
        $uploadBytes   = ($grp | Where-Object { $_.Source -eq "UPLOADED_CONTENT" } | Measure-Object Bytes -Sum).Sum
        $driveCount    = ($grp | Where-Object { $_.Source -eq "DRIVE_FILE" }).Count
        $uploadCount   = ($grp | Where-Object { $_.Source -eq "UPLOADED_CONTENT" }).Count
        $imageCount    = ($grp | Where-Object { $_.IsInlineImage -eq $true -and $_.Source -eq "UPLOADED_CONTENT" }).Count
        $gifCount      = ($detailRows | Where-Object { $_.SpaceID -eq $first.SpaceID -and $_.Source -eq "GIPHY_TENOR" } |
                          Select-Object -ExpandProperty Note | ForEach-Object {
                              if ($_ -match "^(\d+) gif") { [int]$Matches[1] } else { 0 }
                          } | Measure-Object -Sum).Sum
        $largeCount    = ($grp | Where-Object { $_.IsLarge -eq $true }).Count
        $largest       = $grp | Sort-Object Bytes -Descending | Select-Object -First 1

        [PSCustomObject]@{
            SpaceID                  = $first.SpaceID
            SpaceName                = $first.SpaceName
            SpaceType                = $first.SpaceType
            MemberCount              = $first.MemberCount
            SpaceUri                 = $spacemap[$first.SpaceID].SpaceUri
            TotalAttachmentCount     = $grp.Count
            Drive_AttachCount        = $driveCount
            Uploaded_AttachCount     = $uploadCount
            InlineImage_Count        = $imageCount
            AnimatedGif_Count        = [int]$gifCount
            TotalBytes               = $totalBytes
            TotalSizeMB              = [Math]::Round($totalBytes / 1MB, 2)
            Drive_TotalBytes         = $driveBytes
            Drive_TotalMB            = [Math]::Round($driveBytes / 1MB, 2)
            Uploaded_TotalBytes      = $uploadBytes
            Uploaded_TotalMB         = [Math]::Round($uploadBytes / 1MB, 2)
            LargeAttachments_Count   = $largeCount
            Largest_AttachName       = $largest.AttachName
            Largest_Bytes            = $largest.Bytes
            Largest_SizeMB           = $largest.SizeMB
            Largest_Source           = $largest.Source
            Largest_MessageID        = $largest.MessageID
        }
    } | Sort-Object TotalBytes -Descending

$perSpaceRows | Export-Csv $PerSpaceCsv -NoTypeInformation
Write-Host "   Per-Space CSV    -> $PerSpaceCsv"

# 3. Top 100 largest individual attachments
$detailRows |
    Where-Object { $_.Bytes -gt 0 } |
    Sort-Object Bytes -Descending |
    Select-Object -First 100 |
    Export-Csv $TopCsv -NoTypeInformation
Write-Host "   Top Largest CSV  -> $TopCsv"

# =============================================================================
# CONSOLE SUMMARY
# =============================================================================
$grandTotalBytes  = ($detailRows | Measure-Object Bytes -Sum).Sum
$grandTotalCount  = ($detailRows | Where-Object { $_.Source -ne "GIPHY_TENOR" }).Count
$driveTotal       = ($detailRows | Where-Object { $_.Source -eq "DRIVE_FILE"       } | Measure-Object Bytes -Sum).Sum
$uploadTotal      = ($detailRows | Where-Object { $_.Source -eq "UPLOADED_CONTENT" } | Measure-Object Bytes -Sum).Sum
$largeTotal       = ($detailRows | Where-Object { $_.IsLarge -eq $true }).Count
$spacesWithAttach = ($detailRows | Select-Object -ExpandProperty SpaceID -Unique).Count
$top5             = $detailRows | Where-Object { $_.Bytes -gt 0 } | Sort-Object Bytes -Descending | Select-Object -First 5

Write-Host ""
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host "              ATTACHMENT AUDIT COMPLETE" -ForegroundColor Green
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host ("Spaces scanned          : {0}"   -f $uniqueSpaces.Count)
Write-Host ("Spaces with attachments : {0}"   -f $spacesWithAttach)
Write-Host ("Total attachments found : {0}"   -f $grandTotalCount)
Write-Host ("  Drive-source          : {0}"   -f ($driveIds.Count))
Write-Host ("  Uploaded/inline       : {0}"   -f ($uploadedRows.Count))
Write-Host ("  Flagged large (>{0} MB): {1}"  -f $LargeFileMB, $largeTotal) `
    -ForegroundColor $(if($largeTotal -gt 0){"Yellow"}else{"Green"})
Write-Host ""
Write-Host ("Total size (Drive)      : {0:N2} MB"  -f ($driveTotal  / 1MB))
Write-Host ("Total size (Uploaded)   : {0:N2} MB"  -f ($uploadTotal / 1MB))
Write-Host ("Grand total sized       : {0:N2} MB"  -f ($grandTotalBytes / 1MB))
Write-Host ""
Write-Host "TOP 5 LARGEST ATTACHMENTS:" -ForegroundColor Cyan
$rank = 1
foreach ($t in $top5) {
    Write-Host ("  #{0}  {1:N1} MB  [{2}]  {3}  in {4}" -f `
        $rank++,
        ($t.Bytes / 1MB),
        $t.Source,
        $(if($t.AttachName){"'$($t.AttachName)'"}else{"(unnamed)"}),
        $t.SpaceName
    )
}
Write-Host ""
Write-Host "Reports:" -ForegroundColor Cyan
Write-Host "  Detail     -> $DetailCsv"
Write-Host "  Per-Space  -> $PerSpaceCsv"
Write-Host "  Top 100    -> $TopCsv"
Write-Host ("=" * 60) -ForegroundColor Cyan
