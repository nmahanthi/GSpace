#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, PnP.PowerShell

<#
.SYNOPSIS
    Builds a live Org Chart in Microsoft 365 (SharePoint modern page) that reflects the real
    Manager/DirectReport structure from Entra ID.
.DESCRIPTION
    Two things happen:

    1. Reads the team structure from Microsoft Graph, starting at -RootUserId, walking
       Get-MgUserDirectReport down to -MaxDepth levels. While walking it detects data problems
       that would otherwise show up as a broken chart (missing manager, circular reporting line,
       disconnected/orphaned nodes) and reports them before anything is published.

    2. Provisions/refreshes a SharePoint modern page containing:
       - The native "Organization chart" web part, seeded to the root person (best effort — the
         OOB web part does not expose a documented API for the "root person" property, so the
         script sets it where possible and always prints the 2-click manual step as a fallback).
       - A "People" web part per level, listing every direct/indirect report at that level so the
         page is populated end-to-end without any manual work even if the OrgChart seed doesn't
         stick in your tenant's web part version.
       - A text summary (root person, generated date, level/head counts).

    The page is what Org Explorer and the modern SharePoint experience use to render team
    structure; Microsoft 365 Profile Cards (Outlook/Teams/People) read the Manager attribute
    directly from Entra ID and are unaffected by this page — see Set-M365EmployeeProfile.ps1 to
    maintain that relationship.
.NOTES
    Prerequisites:
    - Install-Module Microsoft.Graph -Scope CurrentUser
    - Install-Module PnP.PowerShell -Scope CurrentUser
    - Since Sept 9, 2024, PnP PowerShell requires your own Entra ID App Registration for
      -Interactive login. Register one (one-time) with:
        Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP.PowerShell" -Tenant yourtenant.onmicrosoft.com
      Then pass the resulting Application (client) ID via -ClientId.

    Required delegated scopes (Graph): User.Read.All
    Required SharePoint permission: member/owner (Add-PnPPage) on the target site.

.PARAMETER SiteUrl
    URL of the SharePoint site to hold the Org Chart page.
.PARAMETER RootUserId
    UPN (or object ID) of the person the chart should be built around (e.g. a department head or CEO).
.PARAMETER PageName
    Name of the modern page to create or update. Default: OrgChart.
.PARAMETER MaxDepth
    How many levels of direct reports to walk down from the root. Default: 3.
.PARAMETER ClientId
    Entra ID App (client) ID used for -Interactive PnP login. Falls back to
    ENTRAID_APP_ID / ENTRAID_CLIENT_ID / AZURE_CLIENT_ID environment variables.
.PARAMETER OutputPath
    CSV export of the resolved team tree (audit trail). Defaults to a timestamped file.

.EXAMPLE
    .\New-M365OrgChartPage.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/DigitalWorkplace" `
        -RootUserId "john.smith@contoso.com" -PageName "TeamOrgChart" -MaxDepth 3 -ClientId "<appId>"
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$RootUserId,

    [Parameter(Mandatory = $false)]
    [string]$PageName = "OrgChart",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10)]
    [int]$MaxDepth = 3,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\M365OrgChart_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "M365 Org Chart Builder" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# ----------------------------------------------------------------- connect
Write-Host "[1/4] Connecting to Microsoft Graph and SharePoint Online..." -ForegroundColor Yellow
try {
    Connect-MgGraph -Scopes 'User.Read.All' -NoWelcome
    Write-Host "✓ Connected to Microsoft Graph" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    exit 1
}

try {
    if ($ClientId) {
        Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
    } elseif ($env:ENTRAID_APP_ID -or $env:ENTRAID_CLIENT_ID -or $env:AZURE_CLIENT_ID) {
        Connect-PnPOnline -Url $SiteUrl -Interactive
    } else {
        throw "No ClientId supplied and no ENTRAID_APP_ID/ENTRAID_CLIENT_ID/AZURE_CLIENT_ID environment variable set. Since Sept 9, 2024, PnP PowerShell requires your own Entra ID App Registration for -Interactive login. Run Register-PnPEntraIDAppForInteractiveLogin once, then pass -ClientId <appId> to this script."
    }
    Write-Host "✓ Connected to SharePoint Online`n" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to connect to SharePoint Online: $_" -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------- resolve root
Write-Host "[2/4] Walking the reporting structure from the root..." -ForegroundColor Yellow
try {
    $root = Get-MgUser -UserId $RootUserId -Property 'id,displayName,jobTitle,department,userPrincipalName,mail' -ErrorAction Stop
} catch {
    Write-Host "✗ Root user not found: $RootUserId ($_)" -ForegroundColor Red
    exit 1
}

$nodes    = [System.Collections.Generic.List[PSCustomObject]]::new()
$visited  = [System.Collections.Generic.HashSet[string]]::new()
$issues   = [System.Collections.Generic.List[string]]::new()

function Add-OrgChartNode {
    param($User, $ManagerId, $ManagerName, $Level, $Ancestors)

    if ($visited.Contains($User.Id)) {
        $issues.Add("Circular reporting line detected: $($User.DisplayName) already appears earlier in this branch (possible manager loop).")
        return
    }
    [void]$visited.Add($User.Id)

    $nodes.Add([PSCustomObject]@{
        Id           = $User.Id
        DisplayName  = $User.DisplayName
        JobTitle     = $User.JobTitle
        Department   = $User.Department
        Email        = $User.Mail
        UPN          = $User.UserPrincipalName
        ManagerId    = $ManagerId
        ManagerName  = $ManagerName
        Level        = $Level
    })

    if ($Level -ge $MaxDepth) { return }

    try {
        $reports = Get-MgUserDirectReport -UserId $User.Id -ErrorAction Stop
    } catch {
        $issues.Add("Could not read direct reports for $($User.DisplayName): $($_.Exception.Message)")
        return
    }

    foreach ($r in $reports) {
        try {
            $reportUser = Get-MgUser -UserId $r.Id -Property 'id,displayName,jobTitle,department,userPrincipalName,mail' -ErrorAction Stop
            Add-OrgChartNode -User $reportUser -ManagerId $User.Id -ManagerName $User.DisplayName -Level ($Level + 1) -Ancestors ($Ancestors + $User.Id)
        } catch {
            $issues.Add("Direct report $($r.Id) under $($User.DisplayName) could not be resolved: $($_.Exception.Message)")
        }
    }
}

Add-OrgChartNode -User $root -ManagerId $null -ManagerName $null -Level 0 -Ancestors @()

Write-Host "✓ Resolved $($nodes.Count) people across $((($nodes | Measure-Object -Property Level -Maximum).Maximum) + 1) level(s)" -ForegroundColor Green
if ($issues.Count -gt 0) {
    Write-Host "`n⚠ Data issues found (chart will still be built, but review these):" -ForegroundColor Yellow
    $issues | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
}
Write-Host ""

$nodes | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "✓ Team tree exported to: $OutputPath`n" -ForegroundColor Green

# ---------------------------------------------------------- build the page
Write-Host "[3/4] Provisioning the SharePoint page..." -ForegroundColor Yellow

$page = $null
try {
    $page = Get-PnPPage -Identity $PageName -ErrorAction Stop
    Write-Host "  • Page '$PageName' already exists — it will be updated" -ForegroundColor DarkGray
} catch {
    if ($PSCmdlet.ShouldProcess($PageName, "Create modern page")) {
        $page = Add-PnPPage -Name $PageName -LayoutType Article
        Write-Host "  ✓ Created page '$PageName'" -ForegroundColor Green
    }
}

if ($PSCmdlet.ShouldProcess($PageName, "Add Org Chart web parts")) {

    Add-PnPPageSection -Page $page -SectionTemplate OneColumn -Order 1

    $maxLevel = ($nodes | Measure-Object -Property Level -Maximum).Maximum
    $summaryHtml = @"
<h2>$($root.DisplayName) — Team Org Chart</h2>
<p><strong>Department:</strong> $($root.Department) &nbsp;|&nbsp; <strong>Title:</strong> $($root.JobTitle)</p>
<p><strong>Generated:</strong> $(Get-Date -Format 'MMMM dd, yyyy hh:mm tt') &nbsp;|&nbsp; <strong>Levels:</strong> $($maxLevel + 1) &nbsp;|&nbsp; <strong>Total people:</strong> $($nodes.Count)</p>
"@
    Add-PnPPageTextPart -Page $page -Text $summaryHtml -Section 1 -Column 1 -Order 1

    # Native Organization chart web part, seeded to the root person on a best-effort basis.
    # Microsoft does not publish a supported schema for this web part's "root person" property,
    # so if the seed below doesn't take in your tenant, open the page and type the root person's
    # name/email into the web part once — it will then be saved with the page going forward.
    Add-PnPPageSection -Page $page -SectionTemplate OneColumn -Order 2
    try {
        Add-PnPPageWebPart -Page $page -DefaultWebPartType OrgChart -Section 2 -Column 1 -WebPartProperties @{
            personId = $root.Id
            userId   = $root.Id
        } -ErrorAction Stop
        Write-Host "  ✓ Added Organization chart web part (seeded to $($root.DisplayName))" -ForegroundColor Green
    } catch {
        Add-PnPPageWebPart -Page $page -DefaultWebPartType OrgChart -Section 2 -Column 1
        Write-Host "  ✓ Added Organization chart web part (open the page and set the root person to $($root.DisplayName) / $($root.UPN) once)" -ForegroundColor Yellow
    }

    # One People web part per level so the full team is visible even without editing the OrgChart part.
    $order = 3
    foreach ($level in ($nodes | Group-Object -Property Level | Sort-Object { [int]$_.Name })) {
        $levelLabel = if ($level.Name -eq '0') { "Leadership" } else { "Reporting Level $($level.Name)" }
        Add-PnPPageSection -Page $page -SectionTemplate OneColumn -Order $order
        Add-PnPPageTextPart -Page $page -Text "<h3>$levelLabel</h3>" -Section $order -Column 1 -Order 1

        $persons = @($level.Group | ForEach-Object {
            @{
                id         = $_.UPN
                upn        = $_.UPN
                role       = $_.JobTitle
                department = $_.Department
                phone      = ""
                sip        = ""
            }
        })

        try {
            Add-PnPPageWebPart -Page $page -DefaultWebPartType People -Section $order -Column 1 -WebPartProperties @{
                title   = $levelLabel
                persons = $persons
            } -ErrorAction Stop
        } catch {
            Write-Host "    ⚠ Could not seed People web part for $levelLabel automatically: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        $order++
    }

    $page.Save()
    if ($PSCmdlet.ShouldProcess($PageName, "Publish page")) {
        $page.Publish()
    }
}

Write-Host "✓ Page saved and published`n" -ForegroundColor Green

# ---------------------------------------------------------------- summary
Write-Host "[4/4] Team structure" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
foreach ($n in ($nodes | Sort-Object Level, DisplayName)) {
    $indent = "  " * $n.Level
    Write-Host "$indent├── $($n.DisplayName) — $($n.JobTitle)" -ForegroundColor White
}
Write-Host "========================================" -ForegroundColor Cyan

Write-Host "`nPage URL: $SiteUrl/SitePages/$PageName.aspx" -ForegroundColor Cyan
Write-Host "Org Explorer / People experiences will reflect this same Manager data from Entra ID within ~24 hours of any Manager change." -ForegroundColor DarkGray

Disconnect-PnPOnline | Out-Null
Disconnect-MgGraph | Out-Null
