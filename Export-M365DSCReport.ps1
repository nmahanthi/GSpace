<#
.SYNOPSIS
    Exports the Microsoft365DSC configuration of a tenant for all services
    and generates readable reports from it, using only the app registration's
    Client ID and the Tenant ID as input.

.DESCRIPTION
    Microsoft Entra ID does not allow an app to sign in with only a Client ID
    and Tenant ID - it also needs a credential (secret or certificate). To keep
    the input to just those two values, this script manages a certificate for
    you in the local user certificate store:

    * First run  - no certificate exists yet for this Client ID. The script
      creates a self-signed certificate (CN=M365DSC-<ClientId>), exports the
      public key (.cer) to the output folder, prints upload instructions and
      stops. Upload the .cer to the app registration.
    * Next runs  - the script finds that certificate automatically, signs in
      with it and runs the export for every workload, then builds the reports.

    A certificate is used (not a client secret) because Exchange Online and
    Security & Compliance in Microsoft365DSC only support certificate
    authentication for apps.

    Output (in -OutputPath):
      M365TenantConfig.ps1      the DSC configuration export
      M365DSCReport.<ext>       one report per -ReportType
      M365DSC-<ClientId>.cer    public certificate (first run only)

.PARAMETER ClientId
    Application (client) ID of the Entra ID app registration.

.PARAMETER TenantId
    Tenant ID as a GUID or as <tenant>.onmicrosoft.com. A GUID is resolved
    to the initial onmicrosoft.com domain, which Microsoft365DSC requires.

.PARAMETER Workloads
    Workloads to export. Defaults to all Microsoft 365 workloads supported by
    the installed Microsoft365DSC version (Azure, Azure DevOps, Sentinel and
    Services Hub are excluded because they are not Microsoft 365 services).

.PARAMETER Mode
    Default or Full. Full also exports high-volume resources (for example
    every user, mailbox permission and SharePoint site) and takes much longer.

.PARAMETER ReportType
    One or more of HTML, Excel, CSV, JSON, Markdown. Excel requires Microsoft
    Excel to be installed on this machine.

.PARAMETER OutputPath
    Folder for the export and reports. Defaults to
    .\M365DSC_Export_<yyyyMMdd_HHmmss>.

.PARAMETER CertificateValidityYears
    Validity of the self-signed certificate created on first run.

.PARAMETER InstallPrerequisites
    Install the Microsoft365DSC module and its dependencies for the current
    user if they are missing.

.EXAMPLE
    # First run: creates the certificate and tells you where to upload it
    .\Export-M365DSCReport.ps1 -ClientId 11111111-2222-3333-4444-555555555555 -TenantId contoso.onmicrosoft.com -InstallPrerequisites

.EXAMPLE
    # After uploading the .cer: export all services and build HTML + CSV reports
    .\Export-M365DSCReport.ps1 -ClientId 11111111-2222-3333-4444-555555555555 -TenantId contoso.onmicrosoft.com

.EXAMPLE
    # Only Exchange and Teams, full mode, Excel report
    .\Export-M365DSCReport.ps1 -ClientId <id> -TenantId <guid> -Workloads EXO,TEAMS -Mode Full -ReportType Excel
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string[]]$Workloads,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Default', 'Full')]
    [string]$Mode = 'Default',

    [Parameter(Mandatory = $false)]
    [ValidateSet('HTML', 'Excel', 'CSV', 'JSON', 'Markdown')]
    [string[]]$ReportType = @('HTML', 'CSV'),

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath ("M365DSC_Export_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 5)]
    [int]$CertificateValidityYears = 2,

    [Parameter(Mandatory = $false)]
    [switch]$InstallPrerequisites
)

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$Text) Write-Host "`n==> $Text" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Text) Write-Host "   OK   $Text" -ForegroundColor Green }
function Write-Warn { param([string]$Text) Write-Host "   WARN $Text" -ForegroundColor Yellow }
function Write-Info { param([string]$Text) Write-Host "        $Text" }

$NonM365Workloads = @('ADO', 'AZURE', 'SENTINEL', 'SH')
$ReportExtensions = @{ HTML = 'html'; Excel = 'xlsx'; CSV = 'csv'; JSON = 'json'; Markdown = 'md' }
$CertSubject      = "CN=M365DSC-$ClientId"
$ConfigFileName   = 'M365TenantConfig.ps1'

if (-not $IsWindows -and $PSVersionTable.PSEdition -eq 'Core') {
    throw 'This script uses the Windows certificate store (Cert:\CurrentUser\My) and must run on Windows.'
}

#region Prerequisites
Write-Step 'Checking Microsoft365DSC module'
$module = Get-Module -ListAvailable -Name Microsoft365DSC | Sort-Object Version -Descending | Select-Object -First 1
if (-not $module) {
    if (-not $InstallPrerequisites) {
        throw 'Microsoft365DSC is not installed. Re-run with -InstallPrerequisites, or run: Install-Module Microsoft365DSC -Scope CurrentUser; Update-M365DSCDependencies -Scope CurrentUser'
    }
    Write-Info 'Installing Microsoft365DSC (current user scope)...'
    Install-Module -Name Microsoft365DSC -Scope CurrentUser -Force -AllowClobber
    Import-Module Microsoft365DSC -Force
    Write-Info 'Installing Microsoft365DSC dependencies (this can take several minutes)...'
    Update-M365DSCDependencies -Scope CurrentUser
    $module = Get-Module -ListAvailable -Name Microsoft365DSC | Sort-Object Version -Descending | Select-Object -First 1
}
elseif ($InstallPrerequisites) {
    Import-Module Microsoft365DSC -Force
    Write-Info 'Ensuring Microsoft365DSC dependencies are installed...'
    Update-M365DSCDependencies -Scope CurrentUser
}
Import-Module Microsoft365DSC -Force
Write-Ok "Microsoft365DSC $($module.Version)"

$supportedWorkloads = (Get-Command Export-M365DSCConfiguration).Parameters['Workloads'].Attributes |
    Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] } |
    Select-Object -ExpandProperty ValidValues

if (-not $Workloads) {
    $Workloads = $supportedWorkloads | Where-Object { $_ -notin $NonM365Workloads }
}
else {
    $unsupported = $Workloads | Where-Object { $_ -notin $supportedWorkloads }
    if ($unsupported) {
        throw "Workload(s) not supported by Microsoft365DSC $($module.Version): $($unsupported -join ', '). Supported: $($supportedWorkloads -join ', ')"
    }
}
Write-Ok "Workloads: $($Workloads -join ', ')"
#endregion

if (-not (Test-Path -Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}
$OutputPath = (Resolve-Path -Path $OutputPath).Path

#region Certificate
Write-Step "Locating certificate $CertSubject"
$cert = Get-ChildItem -Path Cert:\CurrentUser\My |
    Where-Object { $_.Subject -eq $CertSubject -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date) } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1

if (-not $cert) {
    Write-Warn 'No valid certificate found for this Client ID. Creating a new self-signed certificate.'
    $cert = New-SelfSignedCertificate -Subject $CertSubject `
        -CertStoreLocation 'Cert:\CurrentUser\My' `
        -KeyExportPolicy NonExportable `
        -KeySpec Signature `
        -KeyLength 2048 `
        -KeyAlgorithm RSA `
        -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears($CertificateValidityYears)

    $cerPath = Join-Path -Path $OutputPath -ChildPath "M365DSC-$ClientId.cer"
    Export-Certificate -Cert $cert -FilePath $cerPath -Type CERT | Out-Null

    Write-Ok "Certificate created. Thumbprint: $($cert.Thumbprint)"
    Write-Host ''
    Write-Host 'ACTION REQUIRED - upload the public certificate to the app registration:' -ForegroundColor Yellow
    Write-Info "1. Entra admin center > App registrations > (Client ID $ClientId)"
    Write-Info '2. Certificates & secrets > Certificates > Upload certificate'
    Write-Info "3. Select: $cerPath"
    Write-Info '4. Re-run this script with the same parameters.'
    return
}
Write-Ok "Thumbprint: $($cert.Thumbprint) (expires $($cert.NotAfter.ToString('yyyy-MM-dd')))"
$thumbprint = $cert.Thumbprint
#endregion

#region Tenant / sign-in validation
Write-Step 'Validating sign-in and resolving tenant domain'
try {
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $thumbprint -NoWelcome | Out-Null
    $org = Get-MgOrganization | Select-Object -First 1
}
catch {
    throw "Sign-in failed for app $ClientId in tenant $TenantId. Confirm the certificate (thumbprint $thumbprint) is uploaded to the app registration and admin consent is granted. Details: $($_.Exception.Message)"
}

$initialDomain = ($org.VerifiedDomains | Where-Object { $_.IsInitial }).Name
if ($TenantId -notmatch '\.onmicrosoft\.|\.onsovcloud\.') {
    if (-not $initialDomain) {
        throw 'Could not resolve the initial onmicrosoft.com domain. Pass -TenantId as <tenant>.onmicrosoft.com.'
    }
    $TenantId = $initialDomain
}
Disconnect-MgGraph | Out-Null
Write-Ok "Tenant: $($org.DisplayName) ($TenantId)"
#endregion

#region Export
Write-Step "Exporting configuration (Mode: $Mode). This can take a long time on large tenants."
$exportParams = @{
    Workloads             = $Workloads
    Mode                  = $Mode
    Path                  = $OutputPath
    FileName              = $ConfigFileName
    ApplicationId         = $ClientId
    TenantId              = $TenantId
    CertificateThumbprint = $thumbprint
}
Export-M365DSCConfiguration @exportParams

$configPath = Join-Path -Path $OutputPath -ChildPath $ConfigFileName
if (-not (Test-Path -Path $configPath)) {
    throw "Export finished but $configPath was not created. Check the Microsoft365DSC event log / console output above."
}
Write-Ok "Configuration: $configPath"
#endregion

#region Reports
Write-Step "Generating report(s): $($ReportType -join ', ')"
$generated = @()
foreach ($type in $ReportType) {
    $reportPath = Join-Path -Path $OutputPath -ChildPath ("M365DSCReport.{0}" -f $ReportExtensions[$type])
    try {
        New-M365DSCReportFromConfiguration -Type $type -ConfigurationPath $configPath -OutputPath $reportPath
        $generated += $reportPath
        Write-Ok "$type -> $reportPath"
    }
    catch {
        Write-Warn "$type report failed: $($_.Exception.Message)"
    }
}
#endregion

Write-Step 'Done'
Write-Info "Output folder: $OutputPath"
$generated | ForEach-Object { Write-Info "Report: $_" }
