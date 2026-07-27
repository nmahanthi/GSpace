<#
.SYNOPSIS
    Deduplicate and score the standalone Shared Drive GSites export
    (output of 01d_run_gam_exports_shareddrive.cmd).

.DESCRIPTION
    1. Deduplicates GSites_SharedDrive_Inventory.csv (by site id) and
       GSites_SharedDrive_Permissions.csv (by site id + permission id).
       Not usually needed for the "select teamdriveid" export pattern (each
       site is looked up directly, not iterated per-member), but kept as a
       safety net in case the same Shared Drive ID appears more than once in
       the input CSV, or a site sits in a nested/shared subfolder reachable
       from more than one row.
    2. Calls 05_score_sites.ps1 against the Shared Drive files, writing
       GSites_SharedDrive_Complexity_Report.csv.

    No page/embed/external-domain crawl data is included (this standalone
    utility does not run the Playwright/API crawler), so StructurePoints and
    EmbedPoints will be 0 for every site - only SecurityPoints (from
    permissions) contribute to the score. This is expected.

.PARAMETER OutputDir
    Folder containing the Shared Drive export CSVs. Defaults to .\output.

.PARAMETER PrimaryDomain
    Your primary domain, used to distinguish internal vs. external grantees
    in the permissions-based security score.

.EXAMPLE
    .\05b_score_shareddrive_sites.ps1 -PrimaryDomain "rocheua.com"
#>

param(
    [string]$OutputDir = "$PSScriptRoot\output",
    [string]$PrimaryDomain = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info { param([string]$Message) Write-Host "[INFO] $Message" -ForegroundColor Yellow }
function Write-Success { param([string]$Message) Write-Host "[OK] $Message" -ForegroundColor Green }

function Get-SafeProperty {
    param(
        [Parameter(Mandatory = $true)][psobject]$InputObject,
        [Parameter(Mandatory = $true)][string[]]$PropertyNames
    )
    foreach ($prop in $PropertyNames) {
        $p = $InputObject.psobject.Properties[$prop]
        if ($null -ne $p -and -not [string]::IsNullOrWhiteSpace($p.Value)) {
            return [string]$p.Value
        }
    }
    return $null
}

$inventoryFile = 'GSites_SharedDrive_Inventory.csv'
$permissionsFile = 'GSites_SharedDrive_Permissions.csv'
$reportFile = 'GSites_SharedDrive_Complexity_Report.csv'

$inventoryPath = Join-Path $OutputDir $inventoryFile
if (-not (Test-Path $inventoryPath)) {
    throw "Inventory file not found: $inventoryPath. Run 01d_run_gam_exports_shareddrive.cmd first."
}

# --- Step 1: Dedup ---
Write-Info "Deduplicating Shared Drive GAM exports..."

$rows = @(Import-Csv $inventoryPath)
if ($rows.Count -gt 0) {
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $deduped = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $rows) {
        $id = Get-SafeProperty -InputObject $row -PropertyNames @('id')
        if ([string]::IsNullOrWhiteSpace($id)) { $deduped.Add($row) | Out-Null; continue }
        if ($seen.Contains($id)) { continue }
        $deduped.Add($row) | Out-Null
        $seen.Add($id) | Out-Null
    }
    if ($deduped.Count -lt $rows.Count) {
        $deduped | Export-Csv -NoTypeInformation -Path $inventoryPath
        Write-Success "Deduplicated $inventoryFile : $($rows.Count) rows -> $($deduped.Count) unique site(s)"
    }
    else {
        Write-Info "  $inventoryFile already unique ($($rows.Count) rows)"
    }
}

$permsPath = Join-Path $OutputDir $permissionsFile
if (Test-Path $permsPath) {
    $rows = @(Import-Csv $permsPath)
    if ($rows.Count -gt 0) {
        $seen = [System.Collections.Generic.HashSet[string]]::new()
        $deduped = [System.Collections.Generic.List[object]]::new()
        foreach ($row in $rows) {
            $id = Get-SafeProperty -InputObject $row -PropertyNames @('id')
            if ([string]::IsNullOrWhiteSpace($id)) { $deduped.Add($row) | Out-Null; continue }
            $permId = Get-SafeProperty -InputObject $row -PropertyNames @('permission.id')
            $key = "$id|$permId"
            if ($seen.Contains($key)) { continue }
            $deduped.Add($row) | Out-Null
            $seen.Add($key) | Out-Null
        }
        if ($deduped.Count -lt $rows.Count) {
            $deduped | Export-Csv -NoTypeInformation -Path $permsPath
            Write-Success "Deduplicated $permissionsFile : $($rows.Count) rows -> $($deduped.Count) unique permission row(s)"
        }
        else {
            Write-Info "  $permissionsFile already unique ($($rows.Count) rows)"
        }
    }
}

# --- Step 2: Score ---
Write-Info "Scoring Shared Drive sites (no crawl data - structure/embed points will be 0)..."

$scoreScript = Join-Path $PSScriptRoot '05_score_sites.ps1'
& $scoreScript -OutputDir $OutputDir -PrimaryDomain $PrimaryDomain `
    -InventoryFile $inventoryFile -PermissionsFile $permissionsFile `
    -PagesFile 'GSites_SharedDrive_Pages.csv' -EmbedsFile 'GSites_SharedDrive_Embeds.csv' `
    -ExternalDomainsFile 'GSites_SharedDrive_ExternalDomains.csv' -ReportFile $reportFile

Write-Success "Shared Drive scoring completed: $reportFile"
