#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

<#
.SYNOPSIS
    Populates Microsoft 365 / Entra ID employee profile attributes and maintains the Manager relationship.
.DESCRIPTION
    Writes the identity/profile attributes that surface in Org Explorer, Microsoft 365 Profile Cards,
    Outlook, Teams and People Search:

        Display Name -> displayName
        Job Title    -> jobTitle
        Department   -> department
        Manager      -> manager relationship (manager/$ref)
        Company      -> companyName
        Office       -> officeLocation
        Business Ph. -> businessPhones
        Mobile       -> mobilePhone
        Email        -> mail / userPrincipalName (read-only here, used as the key)
        Employee ID  -> employeeId
        Employee Type-> employeeType
        Hire Date    -> employeeHireDate

    The Manager is a navigation property, not a writable attribute, so it is maintained with
    Set-MgUserManagerByRef (and removed with Remove-MgUserManagerByRef when -Manager is "NONE").

    Works for a single user (parameters) or in bulk (-CsvPath). Supports -WhatIf / -Confirm.
.NOTES
    Prerequisites:
    - Install-Module Microsoft.Graph -Scope CurrentUser
    - Role: User Administrator (or Global Administrator)

    Required delegated scopes:
    - User.ReadWrite.All
    - Directory.ReadWrite.All   (required to write the manager reference)

    Propagation to the experiences:
    - Entra ID / Graph            : immediate
    - Outlook & Teams profile card: usually < 1 hour
    - Org Explorer / People Search: up to 24 hours (search index + org chart rebuild)

.PARAMETER UserPrincipalName
    UPN of the user to update (single-user mode).
.PARAMETER Manager
    UPN of the manager. Pass "NONE" to clear an existing manager.
.PARAMETER CsvPath
    CSV file for bulk mode. Column headers must match the parameter names.
.PARAMETER SampleCsv
    Writes a template CSV to -CsvPath and exits.
.PARAMETER OutputPath
    Result/verification CSV. Defaults to a timestamped file in the current folder.

.EXAMPLE
    .\Set-M365EmployeeProfile.ps1 -UserPrincipalName narendra@contoso.com `
        -DisplayName "Narendra Mahanthi" -JobTitle "M365 Solution Architect" `
        -Department "Digital Workplace" -Manager "john.smith@contoso.com" `
        -Company "Contoso" -OfficeLocation "Hyderabad" `
        -BusinessPhone "+91 40 1234 5678" -MobilePhone "+91 98765 43210" `
        -EmployeeId "EMP001245" -EmployeeType "Employee"

.EXAMPLE
    .\Set-M365EmployeeProfile.ps1 -SampleCsv -CsvPath .\employees.csv
    .\Set-M365EmployeeProfile.ps1 -CsvPath .\employees.csv -WhatIf
#>

[CmdletBinding(DefaultParameterSetName = 'Single', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Single')]
    [string]$UserPrincipalName,

    [Parameter(ParameterSetName = 'Single')]
    [string]$DisplayName,

    [Parameter(ParameterSetName = 'Single')]
    [string]$GivenName,

    [Parameter(ParameterSetName = 'Single')]
    [string]$Surname,

    [Parameter(ParameterSetName = 'Single')]
    [string]$JobTitle,

    [Parameter(ParameterSetName = 'Single')]
    [string]$Department,

    [Parameter(ParameterSetName = 'Single')]
    [string]$Manager,

    [Parameter(ParameterSetName = 'Single')]
    [string]$Company,

    [Parameter(ParameterSetName = 'Single')]
    [string]$OfficeLocation,

    [Parameter(ParameterSetName = 'Single')]
    [string]$BusinessPhone,

    [Parameter(ParameterSetName = 'Single')]
    [string]$MobilePhone,

    [Parameter(ParameterSetName = 'Single')]
    [string]$EmployeeId,

    [Parameter(ParameterSetName = 'Single')]
    [string]$EmployeeType,

    [Parameter(ParameterSetName = 'Single')]
    [string]$EmployeeHireDate,

    [Parameter(ParameterSetName = 'Single')]
    [string]$City,

    [Parameter(ParameterSetName = 'Single')]
    [string]$State,

    [Parameter(ParameterSetName = 'Single')]
    [string]$Country,

    [Parameter(ParameterSetName = 'Single')]
    [string]$UsageLocation,

    [Parameter(Mandatory = $true, ParameterSetName = 'Bulk')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Sample')]
    [string]$CsvPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'Sample')]
    [switch]$SampleCsv,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\M365EmployeeProfile_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

# Profile attribute -> Graph property map (Manager and phones are handled separately)
$PropertyMap = [ordered]@{
    DisplayName      = 'displayName'
    GivenName        = 'givenName'
    Surname          = 'surname'
    JobTitle         = 'jobTitle'
    Department       = 'department'
    Company          = 'companyName'
    OfficeLocation   = 'officeLocation'
    MobilePhone      = 'mobilePhone'
    EmployeeId       = 'employeeId'
    EmployeeType     = 'employeeType'
    EmployeeHireDate = 'employeeHireDate'
    City             = 'city'
    State            = 'state'
    Country          = 'country'
    UsageLocation    = 'usageLocation'
}

# ---------------------------------------------------------------- sample CSV
if ($SampleCsv) {
    $sample = [PSCustomObject]@{
        UserPrincipalName = 'narendra@contoso.com'
        DisplayName       = 'Narendra Mahanthi'
        GivenName         = 'Narendra'
        Surname           = 'Mahanthi'
        JobTitle          = 'M365 Solution Architect'
        Department        = 'Digital Workplace'
        Manager           = 'john.smith@contoso.com'
        Company           = 'Contoso'
        OfficeLocation    = 'Hyderabad'
        BusinessPhone     = '+91 40 1234 5678'
        MobilePhone       = '+91 98765 43210'
        EmployeeId        = 'EMP001245'
        EmployeeType      = 'Employee'
        EmployeeHireDate  = '2021-06-01'
        City              = 'Hyderabad'
        State             = 'Telangana'
        Country           = 'India'
        UsageLocation     = 'IN'
    }
    $sample | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "✓ Sample CSV written to $CsvPath" -ForegroundColor Green
    return
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "M365 Employee Profile & Manager Update" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# ----------------------------------------------------------------- connect
Write-Host "[1/3] Connecting to Microsoft Graph..." -ForegroundColor Yellow
try {
    $connectParams = @{
        Scopes    = @('User.ReadWrite.All', 'Directory.ReadWrite.All')
        NoWelcome = $true
    }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $ctx = Get-MgContext
    Write-Host "✓ Connected to tenant $($ctx.TenantId) as $($ctx.Account)`n" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    exit 1
}

# ------------------------------------------------------------- input rows
Write-Host "[2/3] Loading input..." -ForegroundColor Yellow
if ($PSCmdlet.ParameterSetName -eq 'Bulk') {
    if (-not (Test-Path -Path $CsvPath)) {
        Write-Host "✗ CSV not found: $CsvPath" -ForegroundColor Red
        exit 1
    }
    $rows = @(Import-Csv -Path $CsvPath)
} else {
    $row = [ordered]@{ UserPrincipalName = $UserPrincipalName; Manager = $Manager; BusinessPhone = $BusinessPhone }
    foreach ($key in $PropertyMap.Keys) { $row[$key] = (Get-Variable -Name $key -ValueOnly -ErrorAction SilentlyContinue) }
    $rows = @([PSCustomObject]$row)
}
Write-Host "✓ $($rows.Count) user(s) to process`n" -ForegroundColor Green

# --------------------------------------------------------------- process
Write-Host "[3/3] Applying profile attributes..." -ForegroundColor Yellow
$results = @()

foreach ($entry in $rows) {

    $upn = ($entry.UserPrincipalName).Trim()
    if ([string]::IsNullOrWhiteSpace($upn)) { continue }

    Write-Host "`n→ $upn" -ForegroundColor Cyan

    $status       = 'Updated'
    $errorMessage = ''
    $managerName  = ''

    try {
        $user = Get-MgUser -UserId $upn -Property 'id,displayName,userPrincipalName,mail' -ErrorAction Stop

        # Build the patch body from populated columns only (blank = leave untouched)
        $body = @{}
        foreach ($key in $PropertyMap.Keys) {
            $value = $null
            if ($entry.PSObject.Properties[$key]) { $value = [string]$entry.$key }
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                if ($PropertyMap[$key] -eq 'employeeHireDate') {
                    $body[$PropertyMap[$key]] = ([datetime]::Parse($value)).ToUniversalTime()
                } else {
                    $body[$PropertyMap[$key]] = $value.Trim()
                }
            }
        }

        # businessPhones is a collection in Graph
        if ($entry.PSObject.Properties['BusinessPhone'] -and -not [string]::IsNullOrWhiteSpace($entry.BusinessPhone)) {
            $body['businessPhones'] = @($entry.BusinessPhone.Trim())
        }

        if ($body.Count -gt 0) {
            if ($PSCmdlet.ShouldProcess($upn, "Update $($body.Count) profile attribute(s)")) {
                Update-MgUser -UserId $user.Id -BodyParameter $body -ErrorAction Stop
                Write-Host "  ✓ Updated: $($body.Keys -join ', ')" -ForegroundColor Green
            }
        } else {
            $status = 'NoChange'
            Write-Host "  • No profile attributes supplied" -ForegroundColor DarkGray
        }

        # ---------------------------------------------------- manager link
        $managerValue = $null
        if ($entry.PSObject.Properties['Manager']) { $managerValue = [string]$entry.Manager }

        if (-not [string]::IsNullOrWhiteSpace($managerValue)) {
            $managerValue = $managerValue.Trim()

            if ($managerValue -eq 'NONE') {
                if ($PSCmdlet.ShouldProcess($upn, 'Remove manager')) {
                    try {
                        Remove-MgUserManagerByRef -UserId $user.Id -ErrorAction Stop
                        Write-Host "  ✓ Manager removed" -ForegroundColor Green
                    } catch {
                        if ($_.Exception.Message -match 'does not exist|Resource.*not found') {
                            Write-Host "  • No manager assigned" -ForegroundColor DarkGray
                        } else { throw }
                    }
                }
            }
            elseif ($managerValue -eq $upn) {
                throw "A user cannot be their own manager."
            }
            else {
                $mgr = Get-MgUser -UserId $managerValue -Property 'id,displayName,userPrincipalName' -ErrorAction Stop
                $managerName = $mgr.DisplayName

                # Guard against a direct reporting loop (manager already reports to this user)
                $mgrOfMgr = Get-MgUserManager -UserId $mgr.Id -ErrorAction SilentlyContinue
                if ($mgrOfMgr -and $mgrOfMgr.Id -eq $user.Id) {
                    throw "Circular reporting line: $($mgr.UserPrincipalName) already reports to $upn."
                }

                if ($PSCmdlet.ShouldProcess($upn, "Set manager to $($mgr.UserPrincipalName)")) {
                    Set-MgUserManagerByRef -UserId $user.Id -BodyParameter @{
                        '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($mgr.Id)"
                    } -ErrorAction Stop
                    Write-Host "  ✓ Manager set: $($mgr.DisplayName) <$($mgr.UserPrincipalName)>" -ForegroundColor Green
                }
                $status = 'Updated'
            }
        }
    }
    catch {
        $status       = 'Failed'
        $errorMessage = $_.Exception.Message
        Write-Host "  ✗ $errorMessage" -ForegroundColor Red
    }

    # -------------------------------------------------------- verification
    $verified = $null
    if ($status -ne 'Failed' -and -not $WhatIfPreference) {
        try {
            $verified = Get-MgUser -UserId $upn -Property ('id,displayName,givenName,surname,jobTitle,department,' +
                'companyName,officeLocation,businessPhones,mobilePhone,mail,userPrincipalName,' +
                'employeeId,employeeType,employeeHireDate,city,state,country,usageLocation') -ErrorAction Stop
            $currentManager = Get-MgUserManager -UserId $verified.Id -ErrorAction SilentlyContinue
            if ($currentManager) {
                $managerName = $currentManager.AdditionalProperties['displayName']
            }
        } catch {
            $errorMessage = "Verification failed: $($_.Exception.Message)"
        }
    }

    $results += [PSCustomObject]@{
        UserPrincipalName = $upn
        DisplayName       = $verified.DisplayName
        JobTitle          = $verified.JobTitle
        Department        = $verified.Department
        Manager           = $managerName
        Company           = $verified.CompanyName
        OfficeLocation    = $verified.OfficeLocation
        BusinessPhone     = ($verified.BusinessPhones -join '; ')
        MobilePhone       = $verified.MobilePhone
        Email             = $verified.Mail
        EmployeeId        = $verified.EmployeeId
        EmployeeType      = $verified.EmployeeType
        EmployeeHireDate  = $(if ($verified -and $verified.EmployeeHireDate) { $verified.EmployeeHireDate.ToString('yyyy-MM-dd') })
        City              = $verified.City
        Country           = $verified.Country
        Status            = $status
        Error             = $errorMessage
    }
}

# ---------------------------------------------------------------- summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "PROFILE CARD / ORG EXPLORER VIEW" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

foreach ($r in $results) {
    Write-Host "─ $($r.DisplayName)" -ForegroundColor White
    Write-Host "  ├── Job Title       : $($r.JobTitle)"
    Write-Host "  ├── Department      : $($r.Department)"
    Write-Host "  ├── Manager         : $($r.Manager)"
    Write-Host "  ├── Office Location : $($r.OfficeLocation)"
    Write-Host "  ├── Email           : $($r.Email)"
    Write-Host "  ├── Business Phone  : $($r.BusinessPhone)"
    Write-Host "  ├── Mobile          : $($r.MobilePhone)"
    Write-Host "  ├── Company         : $($r.Company)"
    Write-Host "  ├── Employee ID     : $($r.EmployeeId)"
    Write-Host "  └── Employee Type   : $($r.EmployeeType)`n"
}

$updated = @($results | Where-Object { $_.Status -eq 'Updated' }).Count
$failed  = @($results | Where-Object { $_.Status -eq 'Failed' }).Count
Write-Host "Processed: $($results.Count)  |  Updated: $updated  |  Failed: $failed" -ForegroundColor $(if ($failed) { 'Yellow' } else { 'Green' })

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "✓ Results exported to: $OutputPath" -ForegroundColor Green

Write-Host "`nPropagation: Graph/Entra immediate • Outlook & Teams profile card < 1 hr • Org Explorer / People Search up to 24 hrs." -ForegroundColor DarkGray

Disconnect-MgGraph | Out-Null
