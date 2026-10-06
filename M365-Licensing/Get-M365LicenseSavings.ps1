<#
.SYNOPSIS
    Read-only Microsoft 365 licence waste and savings report for ONE tenant (HTML + CSV + JSON).
.DESCRIPTION
    Finds money being wasted on licences and estimates what you could save:
      1. Unassigned seats      - purchased but not assigned to anyone
      2. Disabled accounts     - sign-in blocked but still holding paid licences
      3. Inactive users        - no successful sign-in for -InactiveDays (or never signed in)
      4. Redundant licences    - a user holds two SKUs where one already includes all of the other's service plans
      5. Email-only users      - Exchange activity but no Teams/SharePoint/OneDrive activity (downgrade candidates)
    Each user is counted under ONE category only (in the order above) so savings are not double counted.
.NOTES
    Requires: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
    Graph permissions (read-only): User.Read.All, Organization.Read.All, AuditLog.Read.All, Reports.Read.All
      (application permissions for unattended use; delegated scopes are requested when signing in interactively).
    Microsoft has no price API. Prices come from prices.csv (SkuPartNumber,FriendlyName,PriceUSD,PriceINR; per user per month).
    The shipped prices are PLACEHOLDERS: replace them with the customer's real contract prices.
    Sign-in data (signInActivity) needs Entra ID P1 or higher; if unavailable, inactive-user checks are skipped.
    Usage reports show real user names only if the tenant setting "Display concealed user, group and site names" is off;
    otherwise email-only detection is skipped.
.EXAMPLE
    ./Get-M365LicenseSavings.ps1 -TenantId <guid> -Open            # USD and INR reports
    ./Get-M365LicenseSavings.ps1 -TenantId <guid> -Currencies INR  # INR only
    ./Get-M365LicenseSavings.ps1 -TenantId <guid> -ClientId <guid> -SecretEnvVar M365_SECRET_CONTOSO
#>
[CmdletBinding()]
param(
    [string]$CustomerName,   # optional: defaults to the tenant's own name
    [string]$TenantId,
    [string]$ClientId,
    [string]$SecretEnvVar,
    [string]$PricesCsv = (Join-Path $PSScriptRoot 'prices.csv'),
    [ValidateSet('USD','INR')] [string[]]$Currencies = @('USD','INR'),   # one report per currency
    [int]$InactiveDays = 90,
    [string]$EmailOnlyTargetSku = 'EXCHANGESTANDARD',
    [string]$PartnerTermsCsv = (Join-Path $PSScriptRoot 'partner-terms.csv'),
    [switch]$SkipPartnerReport,
    [string]$OutputDir = (Join-Path $PSScriptRoot 'output'),
    [switch]$Open
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
function Format-Money([double]$v) {
    if ($Currency -eq 'INR') {   # Indian digit grouping: 12,34,567
        $n = [string][math]::Round([math]::Abs($v)); $sign = if ($v -lt 0) { '-' } else { '' }
        if ($n.Length -gt 3) { $n = ($n.Substring(0, $n.Length - 3) -replace '\B(?=(\d{2})+$)', ',') + ',' + $n.Substring($n.Length - 3) }
        return "$sign$sym$n"
    }
    if ($v -lt 0) { return ('-{0}{1:N0}' -f $sym, [math]::Abs($v)) }
    '{0}{1:N0}' -f $sym, $v
}

# ---- Connect ---------------------------------------------------------------
if ($ClientId -and $SecretEnvVar -and $TenantId) {
    $secret = [Environment]::GetEnvironmentVariable($SecretEnvVar)
    if (-not $secret) { throw "Env var '$SecretEnvVar' is empty." }
    $cred = [pscredential]::new($ClientId, (ConvertTo-SecureString $secret -AsPlainText -Force))
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $cred -NoWelcome
} else {
    $scopes = 'User.Read.All','Organization.Read.All','AuditLog.Read.All','Reports.Read.All'
    if ($TenantId) { Connect-MgGraph -TenantId $TenantId -Scopes $scopes -NoWelcome } else { Connect-MgGraph -Scopes $scopes -NoWelcome }
}

if (-not $CustomerName) {
    $CustomerName = try { (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization').value[0].displayName } catch { $null }
    if (-not $CustomerName) { $CustomerName = if ($TenantId) { $TenantId } else { 'Tenant' } }
}
$safe = ($CustomerName -replace '[^\w\-]', '_')

function Get-GraphAll([string]$Uri) {
    $items = @()
    while ($Uri) {
        $r = Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers @{ ConsistencyLevel = 'eventual' }
        $items += $r.value
        $Uri = $r.'@odata.nextLink'
    }
    $items
}

function Get-SkuName($sku) { if ($friendly[$sku]) { $friendly[$sku] } else { $sku } }

# ---- Collect ---------------------------------------------------------------
Write-Verbose 'Reading subscribed SKUs'
$skus = @(Get-GraphAll 'https://graph.microsoft.com/v1.0/subscribedSkus') | Where-Object { $_.capabilityStatus -eq 'Enabled' }
$skuById = @{}; foreach ($s in $skus) { $skuById[$s.skuId] = $s }
$planSets = @{}
foreach ($s in $skus) { $planSets[$s.skuId] = [Collections.Generic.HashSet[string]]::new([string[]]@($s.servicePlans | ForEach-Object { $_.servicePlanId })) }

$baseWarnings = @()
$warnings = $baseWarnings
Write-Verbose 'Reading users'
$sel = 'id,displayName,userPrincipalName,accountEnabled,userType,assignedLicenses,createdDateTime'
$haveSignIn = $true
try {
    $users = @(Get-GraphAll "https://graph.microsoft.com/v1.0/users?`$select=$sel,signInActivity&`$top=120")
} catch {
    $haveSignIn = $false
    $warnings += 'Sign-in activity unavailable (needs Entra ID P1+ and AuditLog.Read.All): inactive-user check skipped.'
    $users = @(Get-GraphAll "https://graph.microsoft.com/v1.0/users?`$select=$sel&`$top=999")
}
$licensedUsers = @($users | Where-Object { @($_.assignedLicenses).Count -gt 0 })

# Usage report (for email-only detection)
$usage = @{}
try {
    $tmp = [IO.Path]::GetTempFileName()
    Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D90')" -OutputFilePath $tmp | Out-Null
    foreach ($r in Import-Csv $tmp) { if ($r.'Is Deleted' -ne 'True') { $usage[$r.'User Principal Name'.ToLower()] = $r } }
    Remove-Item $tmp -ErrorAction SilentlyContinue
    $matched = @($licensedUsers | Where-Object { $usage.ContainsKey($_.userPrincipalName.ToLower()) }).Count
    if ($licensedUsers.Count -gt 0 -and $matched -lt ($licensedUsers.Count * 0.5)) {
        $usage = @{}
        $warnings += 'Usage report names look concealed: email-only detection skipped.'
    }
} catch { $warnings += "Usage report unavailable: $($_.Exception.Message)" }

$baseWarnings = @($warnings)
$allOutputs = @()
foreach ($Currency in $Currencies) {
$sym = if ($Currency -eq 'INR') { '₹' } else { '$' }
$warnings = @($baseWarnings)
$price = @{}; $friendly = @{}
foreach ($p in Import-Csv $PricesCsv) {
    $price[$p.SkuPartNumber] = [double]$p."Price$Currency"
    $friendly[$p.SkuPartNumber] = $p.FriendlyName
}
# ---- Per-SKU summary (spend and unassigned seats) -------------------------
$skuRows = foreach ($s in $skus) {
    $name = $s.skuPartNumber
    $enabled = [int]$s.prepaidUnits.enabled
    $assigned = [int]$s.consumedUnits
    $unassigned = [math]::Max(0, $enabled - $assigned)
    $pp = if ($price.ContainsKey($name)) { $price[$name] } else { $null }
    if ($null -eq $pp) { $warnings += "No price for $name in prices.csv: excluded from cost figures." }
    [pscustomobject]@{
        Sku = $name; Name = (Get-SkuName $name); Purchased = $enabled; Assigned = $assigned; Unassigned = $unassigned
        Price = $pp; MonthlySpend = $(if ($pp) { $enabled * $pp } else { 0 })
        UnassignedWaste = $(if ($pp) { $unassigned * $pp } else { 0 })
    }
}
$skuRows = @($skuRows | Where-Object { $_.Purchased -gt 0 })

# ---- Per-user findings (one category per user) ----------------------------
$cutoff = (Get-Date).AddDays(-$InactiveDays)
function Get-PaidLicenses($u) {
    @($u.assignedLicenses | ForEach-Object { $skuById[$_.skuId] } | Where-Object { $_ -and $price[$_.skuPartNumber] -gt 0 })
}
function Get-LastSignIn($u) {
    $a = $u.signInActivity; if (-not $a) { return $null }
    $d = $a.lastSuccessfulSignInDateTime; if (-not $d) { $d = $a.lastSignInDateTime }
    if ($d) { [datetime]$d } else { $null }
}
$findings = New-Object System.Collections.Generic.List[object]
function Add-Finding($u, $cat, $skuNames, $monthly, $action) {
    $findings.Add([pscustomobject]@{
        Category = $cat; User = $u.displayName; UPN = $u.userPrincipalName; Licences = ($skuNames -join '; ')
        Confidence = $(switch -Wildcard ($cat) { 'Disabled*' {'High'} 'Redundant*' {'High'} 'Inactive*' {'Medium'} default {'Low'} })
        MonthlySaving = [math]::Round($monthly, 2); Action = $action })
}

foreach ($u in $licensedUsers) {
    $paid = @(Get-PaidLicenses $u)
    if ($paid.Count -eq 0) { continue }
    $paidNames = @($paid | ForEach-Object { Get-SkuName $_.skuPartNumber })
    $total = ($paid | ForEach-Object { $price[$_.skuPartNumber] } | Measure-Object -Sum).Sum

    if ($u.accountEnabled -eq $false) {
        Add-Finding $u 'Disabled account with licence' $paidNames $total 'Remove licences (convert mailbox to shared or archive first if data is needed).'
        continue
    }
    if ($haveSignIn) {
        $last = Get-LastSignIn $u
        $created = if ($u.createdDateTime) { [datetime]$u.createdDateTime } else { $null }
        if ($u.userType -ne 'Guest' -and (($last -and $last -lt $cutoff) -or (-not $last -and $created -and $created -lt $cutoff))) {
            $when = if ($last) { "last sign-in $($last.ToString('yyyy-MM-dd'))" } else { 'never signed in' }
            Add-Finding $u "Inactive $InactiveDays+ days" $paidNames $total "Confirm with manager, then remove licences ($when)."
            continue
        }
    }
    # Redundant: SKU B is fully contained in SKU A held by the same user
    $redundant = @()
    foreach ($b in $paid) {
        foreach ($a in $paid) {
            if ($a.skuId -eq $b.skuId) { continue }
            $pa = $planSets[$a.skuId]; $pb = $planSets[$b.skuId]
            if ($pb.Count -eq 0 -or -not $pb.IsSubsetOf($pa)) { continue }
            $strict = $pa.Count -gt $pb.Count
            $tieCheaper = ($pa.Count -eq $pb.Count) -and ($price[$b.skuPartNumber] -le $price[$a.skuPartNumber]) -and ($b.skuId -lt $a.skuId -or $price[$b.skuPartNumber] -lt $price[$a.skuPartNumber])
            if ($strict -or $tieCheaper) { $redundant += $b; break }
        }
    }
    if ($redundant.Count) {
        $redundant = @($redundant | Sort-Object skuId -Unique)
        $save = ($redundant | ForEach-Object { $price[$_.skuPartNumber] } | Measure-Object -Sum).Sum
        Add-Finding $u 'Redundant licence' @($redundant | ForEach-Object { Get-SkuName $_.skuPartNumber }) $save 'Remove the licence; another assigned SKU already includes all its features.'
        continue
    }
    # Email-only: Exchange used, nothing else, in the last 90 days
    $r = $usage[$u.userPrincipalName.ToLower()]
    if ($r -and $price.ContainsKey($EmailOnlyTargetSku) -and $u.userType -ne 'Guest') {
        $isRecent = { param($d) $d -and ([datetime]$d -ge $cutoff) }
        $ex = & $isRecent $r.'Exchange Last Activity Date'
        $other = (& $isRecent $r.'Teams Last Activity Date') -or (& $isRecent $r.'SharePoint Last Activity Date') -or (& $isRecent $r.'OneDrive Last Activity Date')
        $save = $total - $price[$EmailOnlyTargetSku]
        if ($ex -and -not $other -and $save -gt 0 -and $paid.skuPartNumber -notcontains $EmailOnlyTargetSku) {
            Add-Finding $u 'Email-only user (downgrade)' $paidNames $save "Review desktop app use, then move to $(Get-SkuName $EmailOnlyTargetSku)."
        }
    }
}

# ---- Totals ----------------------------------------------------------------
$monthlySpend = ($skuRows | Measure-Object MonthlySpend -Sum).Sum
$unassignedWaste = ($skuRows | Measure-Object UnassignedWaste -Sum).Sum
$byCat = @($findings | Group-Object Category | ForEach-Object {
    [pscustomobject]@{ Category = $_.Name; Users = $_.Count; Monthly = ($_.Group | Measure-Object MonthlySaving -Sum).Sum } } | Sort-Object Monthly -Descending)
$userWaste = ($findings | Measure-Object MonthlySaving -Sum).Sum
$totalMonthly = $unassignedWaste + $userWaste
# Conservative = only high-confidence findings (disabled accounts, redundant licences) plus unassigned seats
$conservativeMonthly = $unassignedWaste + (($findings | Where-Object Confidence -eq 'High' | Measure-Object MonthlySaving -Sum).Sum)
$unusedSkus = @($skuRows | Where-Object { $_.Assigned -eq 0 -and $_.Purchased -gt 0 -and $_.Price -gt 0 })
$pct = if ($monthlySpend) { [math]::Round(100 * $totalMonthly / $monthlySpend, 1) } else { 0 }

$tips = @()
if ($unusedSkus.Count) { $tips += "<b>Cancel unused subscriptions.</b> No users are assigned: $(($unusedSkus | ForEach-Object { $_.Name }) -join ', ')." }
$tips += "<b>Cut seats only at renewal.</b> Annual-commitment subscriptions usually cannot be reduced mid-term; diarise renewal dates and right-size then."
$tips += "<b>Automate leavers.</b> Use group-based licensing and remove licences in your offboarding checklist so disabled accounts stop costing money."
$tips += "<b>Check billing frequency.</b> Paying monthly on an annual term, or monthly terms, often costs more than paying annually; ask your partner for the difference."
$tips += "<b>Review premium tiers.</b> For users who do not use E5-only security, compliance or voice features, E3 plus targeted add-ons is often cheaper. Business Premium suits organisations of up to 300 users."
$tips += "<b>Remove unused add-ons</b> (Visio, Project, Power BI Pro, Defender/Purview add-ons) held by people who do not use them."
$tips += "<b>Re-run quarterly.</b> This script is read-only and safe to schedule."

# ---- Output files ----------------------------------------------------------
$csvPath = Join-Path $OutputDir "$safe-licensing-actions-$Currency.csv"
$findings | Sort-Object MonthlySaving -Descending | Export-Csv $csvPath -NoTypeInformation -Encoding UTF8
$jsonPath = Join-Path $OutputDir "$safe-licensing-$Currency.json"
[ordered]@{
    CustomerName = $CustomerName; Currency = $Currency; GeneratedOn = (Get-Date -Format 'yyyy-MM-dd'); InactiveDays = $InactiveDays
    MonthlySpend = $monthlySpend; ConservativeAnnualSaving = $conservativeMonthly * 12; MonthlyWaste = $totalMonthly; AnnualSaving = $totalMonthly * 12; WastePercent = $pct
    UnassignedSeatsWaste = $unassignedWaste; Skus = $skuRows; Categories = $byCat; Warnings = @($warnings | Select-Object -Unique)
} | ConvertTo-Json -Depth 5 | Set-Content $jsonPath -Encoding UTF8

$enc = { param($t) [Net.WebUtility]::HtmlEncode([string]$t) }
$skuHtml = ($skuRows | Sort-Object MonthlySpend -Descending | ForEach-Object {
    $p = if ($null -ne $_.Price) { Format-Money $_.Price } else { 'n/a' }
    "<tr><td>$(& $enc $_.Name)</td><td>$($_.Purchased)</td><td>$($_.Assigned)</td><td>$($_.Unassigned)</td><td>$p</td><td>$(Format-Money $_.MonthlySpend)</td><td class='bad'>$(Format-Money $_.UnassignedWaste)</td></tr>" }) -join "`n"
$maxCat = [math]::Max(1, [double](@($byCat | Measure-Object Monthly -Maximum).Maximum))
$catHtml = ($byCat | ForEach-Object {
    $w = [int](100 * $_.Monthly / $maxCat)
    "<tr><td>$(& $enc $_.Category)</td><td>$($_.Users)</td><td>$(Format-Money $_.Monthly)</td><td><div class='bar' style='width:$w%'></div></td></tr>" }) -join "`n"
if ($unassignedWaste -gt 0) { $uw = [int](100 * $unassignedWaste / [math]::Max($maxCat, $unassignedWaste)); $catHtml = "<tr><td>Unassigned seats</td><td>-</td><td>$(Format-Money $unassignedWaste)</td><td><div class='bar' style='width:$uw%'></div></td></tr>`n" + $catHtml }
$rowsHtml = ($findings | Sort-Object MonthlySaving -Descending | Select-Object -First 300 | ForEach-Object {
    "<tr><td>$(& $enc $_.Category)</td><td>$(& $enc $_.User)<br><small>$(& $enc $_.UPN)</small></td><td>$(& $enc $_.Licences)</td><td>$(Format-Money $_.MonthlySaving)</td><td>$($_.Confidence)</td><td>$(& $enc $_.Action)</td></tr>" }) -join "`n"
$warnHtml = if ($warnings) { "<div class='warn'>" + ((@($warnings | Select-Object -Unique) | ForEach-Object { "<div>$(& $enc $_)</div>" }) -join '') + '</div>' } else { '' }
$tipsHtml = ($tips | ForEach-Object { "<li>$_</li>" }) -join "`n"

$html = @"
<!DOCTYPE html><html><head><meta charset='utf-8'><title>$(& $enc $CustomerName) - Microsoft 365 licence savings</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:2em auto;max-width:1100px;color:#222}
.cards{display:flex;gap:1em;flex-wrap:wrap}.card{flex:1;min-width:200px;border:1px solid #ddd;border-radius:8px;padding:1em}
.card .n{font-size:1.8em;font-weight:bold}.bad{color:#c0392b;font-weight:bold}.good{color:#27ae60}
table{border-collapse:collapse;width:100%;margin:1em 0}th,td{border:1px solid #ddd;padding:6px;text-align:left;vertical-align:top}
th{background:#f4f4f4;cursor:pointer}.bar{height:14px;background:#e67e22;border-radius:3px}
.warn{background:#fff6e0;border:1px solid #f0c36d;padding:.6em 1em;margin:1em 0;border-radius:6px}small{color:#777}
</style></head><body>
<h1>Microsoft 365 licence savings - $(& $enc $CustomerName)</h1>
<p>Generated $(Get-Date -Format 'yyyy-MM-dd'). Amounts are in $Currency. Prices come from prices.csv (PriceUSD / PriceINR) and are estimates: confirm against the customer's agreement.</p>
$warnHtml
<div class='cards'>
<div class='card'>Current licence spend<div class='n'>$(Format-Money $monthlySpend)/mo</div>$(Format-Money ($monthlySpend*12))/yr</div>
<div class='card'>Identified waste<div class='n bad'>$(Format-Money $totalMonthly)/mo</div>$pct% of spend</div>
<div class='card'>Potential annual saving<div class='n good'>$(Format-Money ($totalMonthly*12))</div>$($findings.Count) user actions</div>
<div class='card'>Conservative annual saving<div class='n good'>$(Format-Money ($conservativeMonthly*12))</div>high-confidence items only</div>
</div>
<h2>Where the waste is</h2>
<table><tr><th>Category</th><th>Users</th><th>Saving / month</th><th></th></tr>$catHtml</table>
<h2>Licences purchased</h2>
<table><tr><th>Licence</th><th>Purchased</th><th>Assigned</th><th>Unassigned</th><th>Price / user / mo</th><th>Monthly spend</th><th>Unassigned cost / mo</th></tr>$skuHtml</table>
<h2>How to save more</h2><ol>$tipsHtml</ol>
<h2>Action list (top 300 by saving; full list in the CSV)</h2>
<table id='t'><tr><th>Category</th><th>User</th><th>Licences</th><th>Saving / mo</th><th>Confidence</th><th>Recommended action</th></tr>$rowsHtml</table>
<p><small>Each user is counted once (disabled, then inactive, then redundant, then email-only). Unassigned seats can usually only be removed at renewal. Review every action with the customer before removing licences.</small></p>
</body></html>
"@
$htmlPath = Join-Path $OutputDir "$safe-licensing-$Currency.html"
$html | Set-Content $htmlPath -Encoding UTF8
# ---- Partner profit view (INTERNAL - do not send to the customer) ----------
if (-not $SkipPartnerReport) {
    $t = $null
    if (Test-Path $PartnerTermsCsv) { $t = Import-Csv $PartnerTermsCsv | Where-Object Currency -eq $Currency | Select-Object -First 1 }
    function Get-Term([string]$Name, [double]$Default) {
        if ($t -and $t.PSObject.Properties[$Name] -and "$($t.$Name)" -ne '') { [double]$t.$Name } else { $Default }
    }
    $gs = (Get-Term 'GainSharePercent' 20) / 100;  $fixedFee = Get-Term 'FixedAuditFee' 0
    $bill = Get-Term 'BillableHourlyRate' 0;       $cost = Get-Term 'InternalHourlyCost' 0
    $csp = (Get-Term 'CspMarginPercent' 0) / 100
    $bkPrice = Get-Term 'BackupCustomerPricePerUser' 0; $bkCost = Get-Term 'BackupVendorCostPerUser' 0
    $setupH = Get-Term 'SetupHours' 1; $meetH = Get-Term 'ReviewMeetingHours' 1; $tool = Get-Term 'MonthlyToolingCost' 0
    $renewMin = Get-Term 'RenewalMinutesPerSku' 30
    $minutes = @{ 'Disabled' = Get-Term 'MinutesDisabled' 10; 'Inactive' = Get-Term 'MinutesInactive' 20; 'Redundant' = Get-Term 'MinutesRedundant' 5; 'Email-only' = Get-Term 'MinutesEmailOnly' 30 }
    function Get-ActionMinutes($list) {
        $sum = 0.0
        foreach ($f in $list) { foreach ($k in $minutes.Keys) { if ($f.Category -like "$k*") { $sum += $minutes[$k]; break } } }
        $sum
    }
    $users = $licensedUsers.Count
    $highF = @($findings | Where-Object Confidence -eq 'High')
    $renewH = (@($skuRows | Where-Object { $_.Unassigned -gt 0 }).Count * $renewMin) / 60
    $cleanHcons = ((Get-ActionMinutes $highF) / 60) + $renewH
    $cleanHfull = ((Get-ActionMinutes $findings) / 60) + $renewH
    $hoursCons = $setupH + $meetH + $cleanHcons; $hoursFull = $setupH + $meetH + $cleanHfull
    $annualFull = $totalMonthly * 12; $annualCons = $conservativeMonthly * 12
    $residual = [math]::Max(0, ($monthlySpend - $totalMonthly) * 12)

    $S = @{}   # scenario figures: Cons / Full
    foreach ($sc in 'Cons','Full') {
        $ann = if ($sc -eq 'Cons') { $annualCons } else { $annualFull }
        $ch  = if ($sc -eq 'Cons') { $cleanHcons } else { $cleanHfull }
        $hrs = if ($sc -eq 'Cons') { $hoursCons } else { $hoursFull }
        $rev = [ordered]@{
            'Gain-share fee'                = $gs * $ann
            'Fixed audit fee'               = $fixedFee
            'Billable clean-up work'        = $ch * $bill
            'Licence resale margin (CSP)'   = $csp * $residual
            'Backup resale (customer pays)' = $users * $bkPrice * 12
        }
        $costs = [ordered]@{
            'Backup vendor cost'            = $users * $bkCost * 12
            'Your time (cost basis)'        = $hrs * $cost
            'Tooling'                       = $tool * 12
        }
        $tr = ($rev.Values | Measure-Object -Sum).Sum; $tc = ($costs.Values | Measure-Object -Sum).Sum
        $feesOneOff = $rev['Gain-share fee'] + $rev['Fixed audit fee'] + $rev['Billable clean-up work']
        $after = ($monthlySpend - ($ann / 12)) * 12
        $S[$sc] = @{ Rev = $rev; Costs = $costs; TotRev = $tr; TotCost = $tc; Net = $tr - $tc; Hours = $hrs
                     Hourly = $(if ($hrs -gt 0) { ($tr - $tc + $hrs * $cost) / $hrs } else { 0 })
                     Saving = $ann; Fees = $feesOneOff
                     CustBefore = $monthlySpend * 12; CustAfter = $after + $rev['Backup resale (customer pays)'] + $feesOneOff }
    }
    function Row($label, $c, $f, $cls = '') { "<tr class='$cls'><td>$(& $enc $label)</td><td>$(Format-Money $c)</td><td>$(Format-Money $f)</td></tr>" }
    $revHtml  = (($S.Cons.Rev.Keys  | ForEach-Object { Row $_ $S.Cons.Rev[$_]  $S.Full.Rev[$_] }) -join "`n")
    $costHtml = (($S.Cons.Costs.Keys | ForEach-Object { Row $_ $S.Cons.Costs[$_] $S.Full.Costs[$_] }) -join "`n")
    $custHtml = (Row 'Licences now (per year)' $S.Cons.CustBefore $S.Full.CustBefore) + (Row 'Licences after savings' ($S.Cons.CustBefore - $S.Cons.Saving) ($S.Full.CustBefore - $S.Full.Saving)) +
                (Row 'Backup (per year)' $S.Cons.Rev['Backup resale (customer pays)'] $S.Full.Rev['Backup resale (customer pays)']) +
                (Row 'Your one-off fees' $S.Cons.Fees $S.Full.Fees) +
                (Row 'Customer year-one total vs today' ($S.Cons.CustAfter - $S.Cons.CustBefore) ($S.Full.CustAfter - $S.Full.CustBefore) 'tot')
    $custDelta = $S.Cons.CustAfter - $S.Cons.CustBefore
    $custNote = if ($custDelta -le 0) { "Even with backup and your fees, the customer pays <b>$(Format-Money (-$custDelta)) less</b> than today in year one (conservative). Lead with this." }
                else { "With backup and your fees the customer pays <b>$(Format-Money $custDelta) more</b> than today in year one (conservative); the savings do not fully fund the backup. Position backup as risk protection, or lower the fee." }
    $pitch = "We reviewed your Microsoft 365 licences: about $(Format-Money $annualCons) a year of high-confidence savings ($(Format-Money $annualFull) if everything is accepted). " +
             $(if ($custDelta -le 0) { "Including backup and our fee you would still spend $(Format-Money (-$custDelta)) less than today in year one." } else { "That covers most of the cost of backup and our fee." })
    $tc = $S.Cons; $tf = $S.Full
    $phtml = @"
<!DOCTYPE html><html><head><meta charset='utf-8'><title>INTERNAL - partner profit - $(& $enc $CustomerName)</title>
<style>body{font-family:Segoe UI,Arial,sans-serif;margin:2em auto;max-width:1000px}table{border-collapse:collapse;width:100%;margin:1em 0}th,td{border:1px solid #ddd;padding:6px;text-align:left}th{background:#f4f4f4}tr.tot td{font-weight:bold;background:#f9f9f9}.bad{color:#c0392b;font-weight:bold}.good{color:#27ae60;font-weight:bold}.warn{background:#fff6e0;border:1px solid #f0c36d;padding:.6em 1em;border-radius:6px}</style></head><body>
<div class='warn'><b>INTERNAL - not for the customer.</b> Uses the assumptions in partner-terms.csv ($Currency row); the figures are only as good as those inputs. Conservative = high-confidence items only (disabled accounts, redundant licences, unassigned seats).</div>
<h1>Partner profit - $(& $enc $CustomerName) ($Currency, year one)</h1>
<h2>Revenue</h2><table><tr><th></th><th>Conservative</th><th>Full</th></tr>$revHtml$(Row 'Total revenue' $tc.TotRev $tf.TotRev 'tot')</table>
<h2>Your costs</h2><table><tr><th></th><th>Conservative</th><th>Full</th></tr>$costHtml$(Row 'Total costs' $tc.TotCost $tf.TotCost 'tot')</table>
<h2>Net profit</h2><table><tr><th></th><th>Conservative</th><th>Full</th></tr>
$(Row 'Net profit' $tc.Net $tf.Net 'tot')
<tr><td>Hours of work (setup, meeting, clean-up, renewal)</td><td>$([math]::Round($tc.Hours,1)) h</td><td>$([math]::Round($tf.Hours,1)) h</td></tr>
<tr><td>Effective earnings per hour (before your time cost)</td><td>$(Format-Money $tc.Hourly)</td><td>$(Format-Money $tf.Hourly)</td></tr></table>
<h2>What the customer actually pays (year one)</h2><table><tr><th></th><th>Conservative</th><th>Full</th></tr>$custHtml</table>
<p>$custNote</p>
<h2>Suggested pitch</h2><p>$(& $enc $pitch)</p>
<h2>Notes</h2><ul>
<li>Licence resale margin only applies if you are the customer's CSP (Microsoft's reseller programme); otherwise leave CspMarginPercent at 0.</li>
<li>Backup lines assume the customer buys backup for every licensed user; remove them from partner-terms.csv (price 0) if not selling it.</li>
<li>Savings from unassigned seats usually arrive only at renewal, so year-one cash savings can be lower than shown.</li>
<li>Next conversations: security and Entra clean-up, a quarterly licence-review retainer, and a renewal right-sizing reminder 90 days before renewal.</li></ul>
</body></html>
"@
    $pPath = Join-Path $OutputDir "$safe-partner-profit-$Currency.html"
    $phtml | Set-Content $pPath -Encoding UTF8
    Write-Output "Partner view ($Currency, INTERNAL): $pPath"
}
Write-Output "Report ($Currency): $htmlPath"
Write-Output "Actions ($Currency): $csvPath"
if ($Open) { try { Invoke-Item $htmlPath } catch { Write-Warning 'Could not open browser automatically.' } }
}
Disconnect-MgGraph | Out-Null
