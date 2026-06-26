#Requires -Version 5.1
<#
.SYNOPSIS
    SharePoint Online Migration Assessment Report
.DESCRIPTION
    Comprehensive SPO migration assessment: site collections, lists/libraries,
    permissions, large files, classic workflows, InfoPath forms, and migration
    risk scoring.  Outputs a multi-sheet Excel workbook and an HTML dashboard.
.PARAMETER AdminUrl
    SPO admin center URL, e.g. https://contoso-admin.sharepoint.com
.PARAMETER OutputPath
    Output folder (created if absent). Defaults to script directory + timestamp.
.PARAMETER ClientId / ClientSecret / TenantId
    App-only credentials. If omitted, interactive browser login is used.
.PARAMETER LargeFileSizeMB
    Files >= this size (MB) are flagged as large. Default: 500.
.PARAMETER LargeListThreshold
    Lists with >= this many items are flagged. Default: 5000.
.PARAMETER SkipPersonalSites
    Skip OneDrive personal sites. Default: $true.
.PARAMETER SkipSystemSites
    Skip system/app-catalog sites. Default: $true.
.PARAMETER CheckLargeFiles
    Scan document libraries for large files (adds time on big tenants).
.PARAMETER MaxSites
    Cap on sites processed; 0 = all. Default: 0.
.PARAMETER ExportCSV
    Also write raw CSV files alongside Excel and HTML.
.EXAMPLE
    .\SPO-Migration-Assessment.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -ExportCSV
.EXAMPLE
    .\SPO-Migration-Assessment.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" `
        -ClientId "xxx" -ClientSecret "yyy" -TenantId "zzz" -CheckLargeFiles -MaxSites 50
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$AdminUrl,
    [string]$OutputPath         = (Join-Path $PSScriptRoot "SPO_Assessment_$(Get-Date -f 'yyyyMMdd_HHmmss')"),
    [string]$ClientId           = "",
    [string]$ClientSecret       = "",
    [string]$TenantId           = "",
    [int]   $LargeFileSizeMB    = 500,
    [int]   $LargeListThreshold = 5000,
    [bool]  $SkipPersonalSites  = $true,
    [bool]  $SkipSystemSites    = $true,
    [switch]$CheckLargeFiles,
    [int]   $MaxSites           = 0,
    [switch]$ExportCSV
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$script:HasErrors      = $false

#region HELPERS
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts    = Get-Date -Format "HH:mm:ss"
    $color = switch ($Level) { "SUCCESS"{"Green"} "WARN"{"Yellow"} "ERROR"{"Red"} default{"Cyan"} }
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $color
    if ($Level -eq "ERROR") { $script:HasErrors = $true }
}

function Invoke-SafePnP {
    param([scriptblock]$Block, [string]$Ctx = "")
    try   { & $Block 2>$null }
    catch {
        $msg = ($_.Exception.Message -split "`n")[0].Trim()
        Write-Log "Skipped$(if($Ctx){ " [$Ctx]" }): $msg" "WARN"
        return $null
    }
}
#endregion

#region MODULE CHECK
function Assert-Modules {
    $req  = @("PnP.PowerShell","ImportExcel")
    $miss = @($req | Where-Object { -not (Get-Module -ListAvailable -Name $_) })
    if ($miss) {
        $cmds = ($miss | ForEach-Object { "  Install-Module $_ -Scope CurrentUser -Force" }) -join "`n"
        Write-Log "Missing modules. Run the following, then re-execute this script:`n$cmds" "ERROR"
        exit 1
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
    Import-Module ImportExcel    -ErrorAction Stop
    Write-Log "Modules ready: PnP.PowerShell, ImportExcel." "SUCCESS"
}
#endregion

#region CONNECTION
function Connect-Admin {
    Write-Log "Connecting to SPO Admin: $AdminUrl" "INFO"
    if ($ClientId -and $ClientSecret -and $TenantId) {
        Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -ClientSecret $ClientSecret -Tenant $TenantId
    } else {
        Connect-PnPOnline -Url $AdminUrl -Interactive
    }
    Write-Log "Connected to admin center." "SUCCESS"
}

function Connect-Site { param([string]$Url)
    if ($ClientId -and $ClientSecret -and $TenantId) {
        Connect-PnPOnline -Url $Url -ClientId $ClientId -ClientSecret $ClientSecret -Tenant $TenantId -ErrorAction Stop
    } else {
        Connect-PnPOnline -Url $Url -Interactive -ErrorAction Stop
    }
}
#endregion

#region SITE COLLECTION ENUMERATION
function Get-AllSiteCollections {
    Write-Log "=== Enumerating Site Collections ===" "INFO"
    Connect-Admin
    $all = @(Invoke-SafePnP { Get-PnPTenantSite -Detailed } "TenantSite")
    if (-not $all -or $all.Count -eq 0) { Write-Log "No sites returned. Verify SharePoint Admin role." "ERROR"; return @() }
    if ($SkipPersonalSites) { $all = @($all | Where-Object { $_.Url -notlike "*/personal/*" }) }
    if ($SkipSystemSites) {
        $all = @($all | Where-Object {
            $_.Url      -notlike "*/appcatalog/*"      -and
            $_.Url      -notlike "*/contentTypeHub*"   -and
            $_.Template -notlike "SRCHCEN*"            -and
            $_.Template -notlike "POINTPUBLISHINGHUB*" -and
            $_.Template -notlike "RedirectSite*"
        })
    }
    if ($MaxSites -gt 0) { $all = $all | Select-Object -First $MaxSites }
    Write-Log "Found $($all.Count) site collections to process." "INFO"
    return $all
}
#endregion

#region SITE DETAILS
function Get-SiteDetails {
    param($Sites)
    Write-Log "=== Collecting Site Details ===" "INFO"
    $siteResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    $libResults  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $fileResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i = 0

    foreach ($site in $Sites) {
        $i++
        Write-Progress -Activity "Analyzing Sites" `
                       -Status "$i / $($Sites.Count) : $($site.Url)" `
                       -PercentComplete (($i / $Sites.Count) * 100)

        $connected    = $false
        $subsiteCount = 0; $wfCount = 0; $infoPathCount = 0
        $uniquePerms  = 0; $totalItems = 0; $largeFilesCount = 0
        $listCount    = 0

        try {
            Connect-Site -Url $site.Url
            $connected = $true
        } catch {
            Write-Log "Cannot connect to $($site.Url): $(($_.Exception.Message -split '`n')[0])" "WARN"
        }

        if ($connected) {
            # Subsites
            $subsites     = @(Invoke-SafePnP { Get-PnPSubWeb -Recurse } $site.Url)
            $subsiteCount = $subsites.Count

            # Lists & Libraries
            $lists = @(Invoke-SafePnP {
                Get-PnPList -Includes HasUniqueRoleAssignments,WorkflowAssociations,ContentTypes,ItemCount,RootFolder,EnableVersioning,MajorVersionLimit,BaseType,BaseTemplate,Hidden
            } $site.Url)

            foreach ($list in $lists) {
                if ($list.Hidden) { continue }
                $listCount++
                $isLib      = $list.BaseType -eq "DocumentLibrary"
                $isBig      = $list.ItemCount -ge $LargeListThreshold
                $hasWF      = $list.WorkflowAssociations.Count -gt 0
                $isInfoPath = @($list.ContentTypes | Where-Object { $_.Name -like "*InfoPath*" -or $_.Name -like "*XmlDocument*" }).Count -gt 0

                if ($isInfoPath)                  { $infoPathCount++ }
                if ($hasWF)                       { $wfCount += $list.WorkflowAssociations.Count }
                if ($list.HasUniqueRoleAssignments){ $uniquePerms++ }
                $totalItems += $list.ItemCount

                # Large file detection (optional, document libraries only)
                if ($CheckLargeFiles -and $isLib) {
                    $bigFiles = @(Invoke-SafePnP {
                        Get-PnPListItem -List $list -PageSize 500 `
                            -Fields "FileLeafRef","File_x0020_Size","Modified","FileDirRef","File_x0020_Type" |
                        Where-Object {
                            $_.FileSystemObjectType -eq "File" -and
                            ([int64]$_.FieldValues["File_x0020_Size"]) -ge ($LargeFileSizeMB * 1MB)
                        }
                    } $list.Title)
                    foreach ($f in $bigFiles) {
                        $largeFilesCount++
                        $fileResults.Add([PSCustomObject]@{
                            SiteUrl  = $site.Url
                            Library  = $list.Title
                            FileName = $f.FieldValues["FileLeafRef"]
                            SizeMB   = [math]::Round(([int64]$f.FieldValues["File_x0020_Size"]) / 1MB, 2)
                            FileType = $f.FieldValues["File_x0020_Type"]
                            Modified = if ($f.FieldValues["Modified"]) { ([datetime]$f.FieldValues["Modified"]).ToString("yyyy-MM-dd") } else { "" }
                            FilePath = $f.FieldValues["FileDirRef"]
                        })
                    }
                }

                $libResults.Add([PSCustomObject]@{
                    SiteUrl        = $site.Url
                    SiteTitle      = $site.Title
                    ListTitle      = $list.Title
                    ListType       = $list.BaseType
                    BaseTemplate   = $list.BaseTemplate
                    ItemCount      = $list.ItemCount
                    IsLargeList    = $isBig
                    HasUniquePerms = $list.HasUniqueRoleAssignments
                    HasWorkflows   = $hasWF
                    WorkflowCount  = $list.WorkflowAssociations.Count
                    HasInfoPath    = $isInfoPath
                    Versioning     = $list.EnableVersioning
                    MajorVersions  = $list.MajorVersionLimit
                })
            }
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }

        # Migration risk score
        $score  = 0
        $score += if ($wfCount        -gt  0)    { 3 } else { 0 }
        $score += if ($infoPathCount  -gt  0)    { 2 } else { 0 }
        $score += if ($largeFilesCount -gt  5)   { 2 } else { 0 }
        $score += if ($uniquePerms    -gt 10)    { 2 } else { 0 }
        $score += if ($totalItems     -gt 100000){ 3 } else { 0 }
        $score += if ([double]$site.StorageUsageCurrent -gt 51200) { 1 } else { 0 } # >50 GB (MB units)
        $risk   = if ($score -ge 5) { "High" } elseif ($score -ge 2) { "Medium" } else { "Low" }

        $siteResults.Add([PSCustomObject]@{
            SiteUrl           = $site.Url
            Title             = $site.Title
            Template          = $site.Template
            Owner             = $site.Owner
            StorageUsedGB     = [math]::Round([double]$site.StorageUsageCurrent / 1024, 3)
            StorageQuotaGB    = [math]::Round([double]$site.StorageMaximumLevel  / 1024, 3)
            StorageUsedPct    = if ([double]$site.StorageMaximumLevel -gt 0) {
                                    [math]::Round(([double]$site.StorageUsageCurrent / [double]$site.StorageMaximumLevel) * 100, 1)
                                } else { 0 }
            LastModified      = if ($site.LastContentModifiedDate) { $site.LastContentModifiedDate.ToString("yyyy-MM-dd") } else { "N/A" }
            Created           = if ($site.Created) { ([datetime]$site.Created).ToString("yyyy-MM-dd") } else { "N/A" }
            SharingCapability = $site.SharingCapability
            LockState         = $site.LockState
            IsHubSite         = $site.IsHubSite
            HubSiteId         = $site.HubSiteId
            SubsiteCount      = $subsiteCount
            ListsLibraries    = $listCount
            TotalItems        = $totalItems
            ClassicWorkflows  = $wfCount
            InfoPathForms     = $infoPathCount
            UniquePermObjects = $uniquePerms
            LargeFilesCount   = $largeFilesCount
            MigrationRisk     = $risk
            RiskScore         = $score
            Connected         = $connected
        })
    }
    Write-Progress -Activity "Analyzing Sites" -Completed
    Write-Log "Site detail collection complete: $($siteResults.Count) sites." "SUCCESS"
    return @{ Sites = $siteResults; Libraries = $libResults; LargeFiles = $fileResults }
}
#endregion

#region EXCEL EXPORT
function Export-ToExcel {
    param([string]$OutputFile, $SitesData, $LibrariesData, $LargeFilesData)
    Write-Log "Generating Excel workbook..." "INFO"

    # --- Migration Summary sheet ---
    $highRisk  = @($SitesData | Where-Object { $_.MigrationRisk -eq "High"   }).Count
    $medRisk   = @($SitesData | Where-Object { $_.MigrationRisk -eq "Medium" }).Count
    $lowRisk   = @($SitesData | Where-Object { $_.MigrationRisk -eq "Low"    }).Count
    $summary   = [ordered]@{
        "Total Site Collections"      = $SitesData.Count
        "High Risk Sites"             = $highRisk
        "Medium Risk Sites"           = $medRisk
        "Low Risk Sites"              = $lowRisk
        "Total Storage Used (GB)"     = [math]::Round([double](($SitesData | Measure-Object StorageUsedGB -Sum).Sum), 2)
        "Sites with Classic Workflows"= @($SitesData | Where-Object { $_.ClassicWorkflows  -gt 0 }).Count
        "Sites with InfoPath Forms"   = @($SitesData | Where-Object { $_.InfoPathForms      -gt 0 }).Count
        "Sites with External Sharing" = @($SitesData | Where-Object { $_.SharingCapability -ne "Disabled" }).Count
        "Total Lists & Libraries"     = $LibrariesData.Count
        "Large Libraries (>=$LargeListThreshold items)" = @($LibrariesData | Where-Object { $_.IsLargeList }).Count
        "Large Files Flagged (>=$LargeFileSizeMB MB)"   = $LargeFilesData.Count
        "Total Items Across All Sites"= [int](($SitesData | Measure-Object TotalItems -Sum).Sum)
        "Sites with Subsites"         = @($SitesData | Where-Object { $_.SubsiteCount -gt 0 }).Count
    }
    $summaryData = $summary.Keys | ForEach-Object { [PSCustomObject]@{ Metric = $_; Value = $summary[$_] } }

    # --- Build workbook ---
    $xl = $summaryData | Export-Excel -Path $OutputFile -WorksheetName "Migration Summary" `
        -TableName "Summary" -TableStyle Medium2 -AutoSize -FreezeTopRow -PassThru

    # Sites sheet
    $SitesData | Export-Excel -ExcelPackage $xl -WorksheetName "Site Collections" `
        -TableName "Sites" -TableStyle Medium6 -AutoSize -FreezeTopRow

    # Conditional color on MigrationRisk column in Sites sheet
    $ws       = $xl.Workbook.Worksheets["Site Collections"]
    $colNames = ($SitesData | Select-Object -First 1).PSObject.Properties.Name
    $riskIdx  = [array]::IndexOf($colNames, "MigrationRisk") + 1
    if ($riskIdx -gt 0 -and $SitesData.Count -gt 0) {
        $endRow = $SitesData.Count + 1
        Add-ConditionalFormatting -WorkSheet $ws -Address "${riskIdx}:${riskIdx}" `
            -RuleType Equal -ConditionValue '"High"'   -BackgroundColor ([System.Drawing.Color]::FromArgb(209,52,56))  -ForegroundColor ([System.Drawing.Color]::White)
        Add-ConditionalFormatting -WorkSheet $ws -Address "${riskIdx}:${riskIdx}" `
            -RuleType Equal -ConditionValue '"Medium"' -BackgroundColor ([System.Drawing.Color]::FromArgb(255,140,0))  -ForegroundColor ([System.Drawing.Color]::Black)
        Add-ConditionalFormatting -WorkSheet $ws -Address "${riskIdx}:${riskIdx}" `
            -RuleType Equal -ConditionValue '"Low"'    -BackgroundColor ([System.Drawing.Color]::FromArgb(16,124,16))  -ForegroundColor ([System.Drawing.Color]::White)
    }

    # Libraries sheet
    if ($LibrariesData.Count -gt 0) {
        $LibrariesData | Export-Excel -ExcelPackage $xl -WorksheetName "Lists & Libraries" `
            -TableName "Libraries" -TableStyle Medium6 -AutoSize -FreezeTopRow
    }

    # Large Files sheet
    if ($LargeFilesData.Count -gt 0) {
        $LargeFilesData | Export-Excel -ExcelPackage $xl -WorksheetName "Large Files" `
            -TableName "LargeFiles" -TableStyle Medium9 -AutoSize -FreezeTopRow
    }

    # Migration Waves sheet — suggested grouping by risk
    $waveFlat = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($risk in @("Low","Medium","High")) {
        $waveLabel = switch ($risk) { "Low"{"Wave 1 - Low Risk"} "Medium"{"Wave 2 - Medium Risk"} "High"{"Wave 3 - High Risk"} }
        $matched   = @($SitesData | Where-Object { $_.MigrationRisk -eq $risk })
        foreach ($s in $matched) {
            $waveFlat.Add([PSCustomObject]@{
                MigrationWave    = $waveLabel
                SiteUrl          = $s.SiteUrl
                Title            = $s.Title
                StorageUsedGB    = $s.StorageUsedGB
                TotalItems       = $s.TotalItems
                ClassicWorkflows = $s.ClassicWorkflows
                InfoPathForms    = $s.InfoPathForms
                SharingCapability= $s.SharingCapability
                MigrationRisk    = $s.MigrationRisk
            })
        }
    }
    if ($waveFlat.Count -gt 0) {
        $waveFlat | Export-Excel -ExcelPackage $xl -WorksheetName "Migration Waves" `
            -TableName "Waves" -TableStyle Medium4 -AutoSize -FreezeTopRow
    }

    Close-ExcelPackage $xl
    Write-Log "Excel workbook saved: $OutputFile" "SUCCESS"
}
#endregion

#region HTML EXPORT
function Export-ToHTML {
    param([string]$OutputFile, $SitesData, $LibrariesData, $LargeFilesData,
          [string]$TenantUrl, [datetime]$RunDate)
    Write-Log "Generating HTML report..." "INFO"

    $totalSites = $SitesData.Count
    $totalGB    = [math]::Round([double](($SitesData | Measure-Object StorageUsedGB -Sum).Sum), 2)
    $highRisk   = @($SitesData | Where-Object { $_.MigrationRisk -eq "High"   }).Count
    $medRisk    = @($SitesData | Where-Object { $_.MigrationRisk -eq "Medium" }).Count
    $lowRisk    = @($SitesData | Where-Object { $_.MigrationRisk -eq "Low"    }).Count
    $wfSites    = @($SitesData | Where-Object { $_.ClassicWorkflows  -gt 0 }).Count
    $ipSites    = @($SitesData | Where-Object { $_.InfoPathForms      -gt 0 }).Count
    $extSharing = @($SitesData | Where-Object { $_.SharingCapability -ne "Disabled" }).Count
    $totalLibs  = $LibrariesData.Count
    $totalFiles = $LargeFilesData.Count
    $reportDate = $RunDate.ToString("dddd, dd MMM yyyy HH:mm")

    # --- Site rows ---
    $siteRows = [System.Text.StringBuilder]::new()
    foreach ($s in $SitesData) {
        $rc = switch ($s.MigrationRisk) { "High"{"risk-high"} "Medium"{"risk-med"} default{"risk-low"} }
        $wf = if ($s.ClassicWorkflows -gt 0) { "<span class='flag'>$($s.ClassicWorkflows)</span>" } else { "0" }
        $ip = if ($s.InfoPathForms    -gt 0) { "<span class='flag'>$($s.InfoPathForms)</span>"    } else { "0" }
        [void]$siteRows.Append("<tr>")
        [void]$siteRows.Append("<td><a href='$($s.SiteUrl)' target='_blank' title='$($s.SiteUrl)'>$($s.Title)</a></td>")
        [void]$siteRows.Append("<td>$($s.Template)</td><td>$($s.Owner)</td>")
        [void]$siteRows.Append("<td>$($s.StorageUsedGB)</td><td>$($s.StorageUsedPct)%</td>")
        [void]$siteRows.Append("<td>$($s.TotalItems)</td><td>$($s.ListsLibraries)</td><td>$($s.SubsiteCount)</td>")
        [void]$siteRows.Append("<td>$wf</td><td>$ip</td><td>$($s.UniquePermObjects)</td>")
        [void]$siteRows.Append("<td>$($s.SharingCapability)</td><td>$($s.LockState)</td>")
        [void]$siteRows.Append("<td><span class='$rc'>$($s.MigrationRisk)</span></td>")
        [void]$siteRows.Append("<td>$($s.LastModified)</td></tr>")
    }

    # --- Library rows (cap at 2000 for HTML perf) ---
    $libRows = [System.Text.StringBuilder]::new()
    foreach ($l in ($LibrariesData | Select-Object -First 2000)) {
        $bc = if ($l.IsLargeList)    { "flag" } else { "" }
        $wc = if ($l.HasWorkflows)   { "flag" } else { "" }
        $ic = if ($l.HasInfoPath)    { "flag" } else { "" }
        [void]$libRows.Append("<tr>")
        [void]$libRows.Append("<td title='$($l.SiteUrl)'>$($l.SiteTitle)</td>")
        [void]$libRows.Append("<td>$($l.ListTitle)</td><td>$($l.ListType)</td>")
        [void]$libRows.Append("<td class='$bc'>$($l.ItemCount)</td>")
        [void]$libRows.Append("<td class='$wc'>$(if($l.HasWorkflows){'Yes (' + $l.WorkflowCount + ')'}else{'No'})</td>")
        [void]$libRows.Append("<td class='$ic'>$(if($l.HasInfoPath){'Yes'}else{'No'})</td>")
        [void]$libRows.Append("<td>$(if($l.HasUniquePerms){'Yes'}else{'No'})</td>")
        [void]$libRows.Append("<td>$(if($l.Versioning){"Yes ($($l.MajorVersions) maj)"}else{'No'})</td></tr>")
    }

    # --- Large file rows ---
    $fileRows = [System.Text.StringBuilder]::new()
    foreach ($f in ($LargeFilesData | Select-Object -First 1000)) {
        [void]$fileRows.Append("<tr>")
        [void]$fileRows.Append("<td title='$($f.SiteUrl)'>$($f.SiteUrl.Split('/')[-1])</td>")
        [void]$fileRows.Append("<td>$($f.Library)</td><td title='$($f.FilePath)'>$($f.FileName)</td>")
        [void]$fileRows.Append("<td>$($f.SizeMB)</td><td>$($f.FileType)</td><td>$($f.Modified)</td></tr>")
    }

    # Build HTML using StringBuilder — avoids here-string delimiter issues in PS 5.1
    $h = [System.Text.StringBuilder]::new(65536)

    # CSS block (no variable expansion needed — use plain string concat)
    $css  = ':root{--bg:#f0f4f8;--card:#fff;--acc:#0078d4;--txt:#1a1a2e;--sub:#6b7280;--bdr:#e2e8f0;'
    $css += '--high:#d13438;--med:#ff8c00;--low:#107c10;}'
    $css += '*{box-sizing:border-box;margin:0;padding:0;}'
    $css += "body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--txt);font-size:14px;}"
    $css += 'header{background:linear-gradient(135deg,#0078d4,#00b4d8);color:#fff;padding:28px 40px;}'
    $css += 'header h1{font-size:1.7rem;font-weight:700;}header p{opacity:.9;font-size:.87rem;margin-top:6px;}'
    $css += '.summary{display:grid;grid-template-columns:repeat(auto-fill,minmax(155px,1fr));gap:12px;padding:20px 40px;}'
    $css += '.card{background:var(--card);border-radius:10px;padding:16px 18px;box-shadow:0 1px 4px rgba(0,0,0,.08);border-left:4px solid var(--acc);}'
    $css += '.card .val{font-size:1.85rem;font-weight:700;color:var(--acc);}'
    $css += '.card .lbl{color:var(--sub);font-size:.7rem;text-transform:uppercase;letter-spacing:.5px;margin-top:4px;}'
    $css += '.card.high{border-color:var(--high);}.card.high .val{color:var(--high);}'
    $css += '.card.med{border-color:var(--med);}.card.med .val{color:var(--med);}'
    $css += '.card.good{border-color:var(--low);}.card.good .val{color:var(--low);}'
    $css += 'nav{display:flex;padding:0 40px;background:#fff;border-bottom:1px solid var(--bdr);}'
    $css += 'nav button{background:none;border:none;padding:13px 20px;cursor:pointer;font-size:.9rem;color:var(--sub);border-bottom:3px solid transparent;transition:.2s;font-weight:500;}'
    $css += 'nav button.active{color:var(--acc);border-bottom-color:var(--acc);font-weight:600;}'
    $css += '.tab{display:none;padding:20px 40px;}.tab.active{display:block;}'
    $css += '.section-hdr{font-size:1rem;font-weight:600;margin-bottom:10px;}'
    $css += '.search-bar{margin-bottom:10px;}'
    $css += '.search-bar input{padding:7px 12px;border:1px solid var(--bdr);border-radius:8px;width:340px;font-size:.87rem;outline:none;}'
    $css += '.search-bar input:focus{border-color:var(--acc);}'
    $css += '.tbl-wrap{overflow-x:auto;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.07);}'
    $css += 'table{width:100%;border-collapse:collapse;background:#fff;}'
    $css += 'th{background:var(--acc);color:#fff;padding:10px 12px;text-align:left;font-size:.77rem;white-space:nowrap;font-weight:600;}'
    $css += 'td{padding:8px 12px;border-bottom:1px solid var(--bdr);font-size:.81rem;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}'
    $css += 'tr:nth-child(even) td{background:#f8fafc;}tr:hover td{background:#ebf5ff;}'
    $css += 'td a{color:var(--acc);text-decoration:none;}td a:hover{text-decoration:underline;}'
    $css += '.risk-high{background:var(--high);color:#fff;border-radius:12px;padding:2px 9px;font-size:.74rem;font-weight:700;}'
    $css += '.risk-med{background:var(--med);color:#000;border-radius:12px;padding:2px 9px;font-size:.74rem;font-weight:700;}'
    $css += '.risk-low{background:var(--low);color:#fff;border-radius:12px;padding:2px 9px;font-size:.74rem;font-weight:700;}'
    $css += '.flag{color:var(--high);font-weight:700;}'
    $css += 'footer{text-align:center;padding:16px;color:var(--sub);font-size:.75rem;border-top:1px solid var(--bdr);margin-top:20px;}'

    # JS (no variable expansion — plain string)
    $js  = 'function showTab(id,btn){'
    $js += "document.querySelectorAll('.tab').forEach(function(t){t.classList.remove('active');});"
    $js += "document.querySelectorAll('nav button').forEach(function(b){b.classList.remove('active');});"
    $js += "document.getElementById(id).classList.add('active');btn.classList.add('active');}"
    $js += 'function flt(inp,tbl){'
    $js += 'var q=document.getElementById(inp).value.toLowerCase();'
    $js += 'var rows=document.getElementById(tbl).getElementsByTagName("tr");'
    $js += 'for(var i=1;i<rows.length;i++){rows[i].style.display=rows[i].innerText.toLowerCase().indexOf(q)>-1?"":"none";}}'

    # Assemble HTML
    [void]$h.Append('<!DOCTYPE html><html lang="en"><head>')
    [void]$h.Append('<meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>')
    [void]$h.Append('<title>SPO Migration Assessment</title>')
    [void]$h.Append("<style>$css</style></head><body>")
    [void]$h.Append('<header>')
    [void]$h.Append('<h1>&#128202; SharePoint Online - Migration Assessment</h1>')
    [void]$h.Append("<p>Tenant: <strong>$TenantUrl</strong> &nbsp;&bull;&nbsp; Report Date: <strong>$reportDate</strong></p>")
    [void]$h.Append('</header><div class="summary">')
    [void]$h.Append("<div class='card'><div class='val'>$totalSites</div><div class='lbl'>Total Sites</div></div>")
    [void]$h.Append("<div class='card'><div class='val'>$totalGB GB</div><div class='lbl'>Total Storage</div></div>")
    [void]$h.Append("<div class='card high'><div class='val'>$highRisk</div><div class='lbl'>High Risk Sites</div></div>")
    [void]$h.Append("<div class='card med'><div class='val'>$medRisk</div><div class='lbl'>Medium Risk Sites</div></div>")
    [void]$h.Append("<div class='card good'><div class='val'>$lowRisk</div><div class='lbl'>Low Risk Sites</div></div>")
    [void]$h.Append("<div class='card high'><div class='val'>$wfSites</div><div class='lbl'>Sites w/ Classic Workflows</div></div>")
    [void]$h.Append("<div class='card med'><div class='val'>$ipSites</div><div class='lbl'>Sites w/ InfoPath</div></div>")
    [void]$h.Append("<div class='card'><div class='val'>$extSharing</div><div class='lbl'>External Sharing On</div></div>")
    [void]$h.Append("<div class='card'><div class='val'>$totalLibs</div><div class='lbl'>Lists &amp; Libraries</div></div>")
    [void]$h.Append("<div class='card high'><div class='val'>$totalFiles</div><div class='lbl'>Large Files Flagged</div></div>")
    [void]$h.Append('</div><nav>')
    [void]$h.Append("<button class='active' onclick='showTab(`"sites`",this)'>&#127760; Sites ($totalSites)</button>")
    [void]$h.Append("<button onclick='showTab(`"libs`",this)'>&#128196; Lists &amp; Libraries ($totalLibs)</button>")
    [void]$h.Append("<button onclick='showTab(`"files`",this)'>&#128190; Large Files ($totalFiles)</button>")
    [void]$h.Append('</nav>')
    [void]$h.Append('<div id="sites" class="tab active">')
    [void]$h.Append('<p class="section-hdr">Site Collections - Risk &amp; Migration Readiness</p>')
    [void]$h.Append('<div class="search-bar"><input id="srchS" onkeyup="flt(''srchS'',''tblS'')" placeholder="Search sites..."/></div>')
    [void]$h.Append('<div class="tbl-wrap"><table id="tblS"><thead><tr>')
    [void]$h.Append('<th>Title</th><th>Template</th><th>Owner</th><th>Storage(GB)</th><th>Used%</th>')
    [void]$h.Append('<th>Items</th><th>Lists/Libs</th><th>Subsites</th><th>Workflows</th><th>InfoPath</th>')
    [void]$h.Append('<th>Unique Perms</th><th>Ext Sharing</th><th>Lock</th><th>Risk</th><th>Last Modified</th>')
    [void]$h.Append("</tr></thead><tbody>$($siteRows.ToString())</tbody></table></div></div>")
    [void]$h.Append('<div id="libs" class="tab">')
    [void]$h.Append('<p class="section-hdr">Lists &amp; Libraries - Detail (first 2000; see Excel for full data)</p>')
    [void]$h.Append('<div class="search-bar"><input id="srchL" onkeyup="flt(''srchL'',''tblL'')" placeholder="Search lists..."/></div>')
    [void]$h.Append('<div class="tbl-wrap"><table id="tblL"><thead><tr>')
    [void]$h.Append('<th>Site</th><th>List/Library</th><th>Type</th><th>Items</th>')
    [void]$h.Append('<th>Workflows</th><th>InfoPath</th><th>Unique Perms</th><th>Versioning</th>')
    [void]$h.Append("</tr></thead><tbody>$($libRows.ToString())</tbody></table></div></div>")
    [void]$h.Append('<div id="files" class="tab">')
    [void]$h.Append("<p class='section-hdr'>Large Files (>= $LargeFileSizeMB MB) - Requires Attention Before Migration</p>")
    [void]$h.Append('<div class="search-bar"><input id="srchF" onkeyup="flt(''srchF'',''tblF'')" placeholder="Search files..."/></div>')
    [void]$h.Append('<div class="tbl-wrap"><table id="tblF"><thead><tr>')
    [void]$h.Append('<th>Site</th><th>Library</th><th>File Name</th><th>Size (MB)</th><th>Type</th><th>Modified</th>')
    [void]$h.Append("</tr></thead><tbody>$($fileRows.ToString())</tbody></table></div></div>")
    [void]$h.Append("<footer>SPO Migration Assessment - $reportDate - SPO-Migration-Assessment.ps1</footer>")
    [void]$h.Append("<script>$js</script>")
    [void]$h.Append('</body></html>')

    $h.ToString() | Out-File -FilePath $OutputFile -Encoding UTF8 -Force
    Write-Log "HTML report saved: $OutputFile" "SUCCESS"
}
#endregion

#region MAIN
function Main {
    Write-Host "`n================================================================" -ForegroundColor Cyan
    Write-Host "   SharePoint Online - Migration Assessment" -ForegroundColor Cyan
    Write-Host "================================================================`n" -ForegroundColor Cyan

    Assert-Modules

    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Log "Output folder: $OutputPath" "INFO"

    # Collect data
    $sites = Get-AllSiteCollections
    if (-not $sites -or $sites.Count -eq 0) { Write-Log "No sites to process. Exiting." "ERROR"; return }

    $data = Get-SiteDetails -Sites $sites

    # Optional CSV export
    if ($ExportCSV) {
        $data.Sites      | Export-Csv -Path (Join-Path $OutputPath "SPO_Sites.csv")      -NoTypeInformation -Encoding UTF8
        $data.Libraries  | Export-Csv -Path (Join-Path $OutputPath "SPO_Libraries.csv")  -NoTypeInformation -Encoding UTF8
        if ($data.LargeFiles.Count -gt 0) {
            $data.LargeFiles | Export-Csv -Path (Join-Path $OutputPath "SPO_LargeFiles.csv") -NoTypeInformation -Encoding UTF8
        }
        Write-Log "CSV files written to: $OutputPath" "SUCCESS"
    }

    # Excel report
    $excelFile = Join-Path $OutputPath "SPO_Migration_Assessment.xlsx"
    Export-ToExcel -OutputFile     $excelFile `
                   -SitesData      $data.Sites `
                   -LibrariesData  $data.Libraries `
                   -LargeFilesData $data.LargeFiles

    # HTML report
    $htmlFile = Join-Path $OutputPath "SPO_Migration_Assessment.html"
    Export-ToHTML -OutputFile     $htmlFile `
                  -SitesData      $data.Sites `
                  -LibrariesData  $data.Libraries `
                  -LargeFilesData $data.LargeFiles `
                  -TenantUrl      $AdminUrl `
                  -RunDate        (Get-Date)

    Write-Host "`n================================================================" -ForegroundColor Cyan
    Write-Log "Assessment complete!" "SUCCESS"
    Write-Log "  Excel : $excelFile" "SUCCESS"
    Write-Log "  HTML  : $htmlFile"  "SUCCESS"
    Write-Host "================================================================`n" -ForegroundColor Cyan

    if ($script:HasErrors) { Write-Log "Some sites had errors/warnings. Review WARN messages above." "WARN" }

    # Open HTML in default browser
    Start-Process $htmlFile
}

Main
#endregion
