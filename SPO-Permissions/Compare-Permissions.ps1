<#
.SYNOPSIS
    Compares GSite permissions against SPO permissions and outputs differences.
.DESCRIPTION
    Loads the GSite permissions CSV and SPO permissions CSV, resolves the
    GSite->SPO site mapping from the SPOSiteUrl column already embedded in the
    GSite CSV (written there by Export-GSitePermissions.ps1 from InputSites.csv),
    translates Google roles to SPO permission levels via Config.psd1, then
    produces a differences CSV with three categories:
      - MISSING   : permission exists in GSite but is absent from SPO
      - EXTRA     : permission exists in SPO but was not in GSite (potential over-grant)
      - MISMATCH  : principal exists in both but has a different permission level

    No separate SiteMappings file is required - the mapping is read directly
    from the SPOSiteUrl column in the GSite permissions CSV.
.PARAMETER ConfigPath
    Path to Config.psd1.  Defaults to .\Config.psd1
.PARAMETER GSitePermissionsFile
    Path to the GSite permissions CSV produced by Export-GSitePermissions.ps1.
    If omitted the newest file matching the pattern in Output.Directory is used.
.PARAMETER SPOPermissionsFile
    Path to the SPO permissions CSV produced by Export-SPOPermissions.ps1.
    If omitted the newest matching file is used.
.PARAMETER OutputFile
    Overrides the output CSV path from Config.psd1.
.EXAMPLE
    .\Compare-Permissions.ps1
    .\Compare-Permissions.ps1 -GSitePermissionsFile .\Output\GSite_20240601.csv `
                               -SPOPermissionsFile  .\Output\SPO_20240601.csv
#>
[CmdletBinding()]
param(
    [string]$ConfigPath            = ".\Config.psd1",
    [string]$GSitePermissionsFile  = "",
    [string]$SPOPermissionsFile    = "",
    [string]$OutputFile            = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Load config
# ---------------------------------------------------------------------------
$cfg  = Import-PowerShellDataFile -Path $ConfigPath
$date = Get-Date -Format "yyyyMMdd_HHmmss"
$dir  = $cfg.Output.Directory
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

Write-Host "`n=== Compare-Permissions.ps1 ===" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# Resolve input files (pick newest if not specified)
# ---------------------------------------------------------------------------
function Resolve-LatestCsv { param([string]$Pattern)
    Get-ChildItem -Path $dir -Filter $Pattern | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
}

if (-not $GSitePermissionsFile) {
    $GSitePermissionsFile = Resolve-LatestCsv "$($cfg.Output.GSitePermissionsFile)*.csv"
}
if (-not $SPOPermissionsFile) {
    $SPOPermissionsFile = Resolve-LatestCsv "$($cfg.Output.SPOPermissionsFile)*.csv"
}
if (-not $OutputFile) {
    $OutputFile = Join-Path $dir "$($cfg.Output.DifferencesFile)_$date.csv"
}

Write-Host "GSite source : $GSitePermissionsFile"
Write-Host "SPO source   : $SPOPermissionsFile"

$permMap = $cfg.PermissionMapping   # Google role -> SPO level

# ---------------------------------------------------------------------------
# Load CSVs
# ---------------------------------------------------------------------------
$gRecords = Import-Csv $GSitePermissionsFile
$sRecords = Import-Csv $SPOPermissionsFile

# ---------------------------------------------------------------------------
# Build GSite -> SPO mapping from SPOSiteUrl column embedded in the GSite CSV.
# Export-GSitePermissions.ps1 copies SPOSiteUrl from InputSites.csv into every row.
# Also support fallback to InputSites.csv (Config.InputSitesFile) if the column
# is absent (e.g. the CSV was produced by an older version of the script).
# ---------------------------------------------------------------------------
$mappings = @{}

$hasSPOColumn = ($gRecords | Select-Object -First 1).PSObject.Properties.Name -contains 'SPOSiteUrl'
if ($hasSPOColumn) {
    $gRecords | Where-Object { $_.SPOSiteUrl } |
        ForEach-Object { $mappings[$_.SiteUrl] = $_.SPOSiteUrl }
    Write-Host "Site mappings read from GSite CSV SPOSiteUrl column: $($mappings.Count) unique site(s)."
} elseif ($cfg.ContainsKey('InputSitesFile') -and (Test-Path $cfg.InputSitesFile)) {
    Import-Csv $cfg.InputSitesFile |
        Where-Object { $_.SiteUrl -and $_.SPOSiteUrl } |
        ForEach-Object { $mappings[$_.SiteUrl] = $_.SPOSiteUrl }
    Write-Host "Site mappings loaded from InputSites.csv: $($mappings.Count) entries."
} else {
    Write-Warning "No SPOSiteUrl column in GSite CSV and InputSites.csv not found. Comparison may produce no results."
}

# ---------------------------------------------------------------------------
# Build a lookup for SPO keyed by (SPOSiteUrl | NormalizedEmail | PermissionLevel)
# Key = "<siteUrl>|<email>"  -> PermissionLevel
# ---------------------------------------------------------------------------
$spoLookup = @{}
foreach ($s in $sRecords) {
    if ($s.IsInherited -eq "True") { continue }
    $email = ($s.PrincipalLogin -replace "i:0#.f|membership|", "" -replace "c:0t.c|tenant|", "").ToLower().Trim()
    $key   = "$($s.SiteCollectionUrl.ToLower().TrimEnd('/'))|$email"
    if (-not $spoLookup.ContainsKey($key)) { $spoLookup[$key] = [System.Collections.Generic.List[string]]::new() }
    $spoLookup[$key].Add($s.PermissionLevel)
}

# Also build a set of all SPO (siteUrl|email) pairs for EXTRA detection
$allSpoKeys = $spoLookup.Keys | ForEach-Object { $_ }

$diffs = [System.Collections.Generic.List[object]]::new()
$matchedSpoKeys = [System.Collections.Generic.HashSet[string]]::new()

# ---------------------------------------------------------------------------
# Compare: for each GSite permission -> find matching SPO entry
# ---------------------------------------------------------------------------
foreach ($g in $gRecords) {
    if ($g.PrincipalType -eq "anyone") { continue }   # skip public/anyone

    # SPOSiteUrl comes directly from the GSite CSV row (preferred) or from mapping table
    $spoSiteUrl = if ($g.PSObject.Properties['SPOSiteUrl'] -and $g.SPOSiteUrl) {
                      $g.SPOSiteUrl
                  } elseif ($mappings.ContainsKey($g.SiteUrl)) {
                      $mappings[$g.SiteUrl]
                  } else { $null }

    if (-not $spoSiteUrl) {
        Write-Verbose "No SPO mapping for GSite: $($g.SiteUrl) - skipping."
        continue
    }

    $email       = $g.PrincipalEmail.ToLower().Trim()
    $expectedSPO = if ($permMap.ContainsKey($g.GoogleRole)) { $permMap[$g.GoogleRole] } else { "Read" }
    $key         = "$($spoSiteUrl.ToLower().TrimEnd('/'))|$email"

    if ($spoLookup.ContainsKey($key)) {
        $matchedSpoKeys.Add($key) | Out-Null
        if (-not ($spoLookup[$key] -contains $expectedSPO)) {
            $diffs.Add([PSCustomObject]@{
                DifferenceType    = "MISMATCH"
                GSiteSiteUrl      = $g.SiteUrl
                SPOSiteUrl        = $spoSiteUrl
                PrincipalEmail    = $email
                PrincipalName     = $g.PrincipalName
                GSiteRole         = $g.GoogleRole
                ExpectedSPOLevel  = $expectedSPO
                ActualSPOLevel    = ($spoLookup[$key] -join "; ")
                Notes             = "Permission level differs from expected."
            })
        }
    } else {
        $diffs.Add([PSCustomObject]@{
            DifferenceType    = "MISSING"
            GSiteSiteUrl      = $g.SiteUrl
            SPOSiteUrl        = $spoSiteUrl
            PrincipalEmail    = $email
            PrincipalName     = $g.PrincipalName
            GSiteRole         = $g.GoogleRole
            ExpectedSPOLevel  = $expectedSPO
            ActualSPOLevel    = ""
            Notes             = "User has access in GSite but no matching permission in SPO."
        })
    }
}

# ---------------------------------------------------------------------------
# EXTRA permissions in SPO (not mapped from any GSite)
# ---------------------------------------------------------------------------
foreach ($key in $allSpoKeys) {
    if (-not $matchedSpoKeys.Contains($key)) {
        $parts = $key -split '\|', 2
        $diffs.Add([PSCustomObject]@{
            DifferenceType    = "EXTRA"
            GSiteSiteUrl      = ""
            SPOSiteUrl        = $parts[0]
            PrincipalEmail    = $parts[1]
            PrincipalName     = ""
            GSiteRole         = ""
            ExpectedSPOLevel  = ""
            ActualSPOLevel    = ($spoLookup[$key] -join "; ")
            Notes             = "Permission exists in SPO but has no corresponding GSite entry."
        })
    }
}

$diffs | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

$missing  = ($diffs | Where-Object DifferenceType -eq "MISSING").Count
$extra    = ($diffs | Where-Object DifferenceType -eq "EXTRA").Count
$mismatch = ($diffs | Where-Object DifferenceType -eq "MISMATCH").Count

Write-Host "`nComparison complete:" -ForegroundColor Cyan
Write-Host "  MISSING  (in GSite, absent in SPO)  : $missing"  -ForegroundColor Red
Write-Host "  EXTRA    (in SPO, no GSite entry)   : $extra"    -ForegroundColor Yellow
Write-Host "  MISMATCH (wrong permission level)   : $mismatch" -ForegroundColor Magenta
Write-Host "Differences written -> $OutputFile`n" -ForegroundColor Green
return $OutputFile
