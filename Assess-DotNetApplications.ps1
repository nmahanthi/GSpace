<#
.SYNOPSIS
    Assesses .NET applications for AWS migration/modernization readiness.

.DESCRIPTION
    Scans one or more file-system paths (source checkouts or deployed app
    folders) - and optionally the local IIS instance - for .NET applications,
    fingerprints each one (target framework, session-state model,
    authentication mode, database provider, Windows-only / COM dependencies,
    platform target), then scores "cloud-native readiness" vs "stickiness"
    and maps each app to a recommended AWS 7R migration strategy and target
    AWS compute service.

    This is a heuristic, config/source-based scan meant to accelerate a
    portfolio assessment. Always validate findings with app owners before
    finalizing a migration wave plan.

.PARAMETER Path
    One or more root folders to scan recursively for .NET applications
    (looks for *.csproj, web.config, packages.config, *.runtimeconfig.json).
    Defaults to the current directory if neither -Path nor -UseIIS is given.

.PARAMETER UseIIS
    Also enumerate applications from the local IIS instance (Windows only,
    requires the WebAdministration module) and fold their physical paths
    and app-pool settings into the scan.

.PARAMETER OutputPath
    Folder to write the CSV/JSON assessment reports to. Defaults to the
    current directory. Created if it doesn't exist.

.PARAMETER IncludeSourceScan
    Perform an additional (slower) regex scan of .cs files for local
    file-system dependencies (Server.MapPath, hardcoded drive letters) that
    typically block moving app storage to Amazon S3.

.EXAMPLE
    .\Assess-DotNetApplications.ps1 -Path C:\Source\Apps -OutputPath C:\Reports

.EXAMPLE
    .\Assess-DotNetApplications.ps1 -UseIIS -IncludeSourceScan

.NOTES
    Maps recommendations to the AWS "7 Rs" (Retire, Retain, Rehost,
    Replatform, Repurchase, Refactor, Relocate). See AWS Prescriptive
    Guidance for background. Findings are heuristic - treat the score/
    recommendation columns as a starting point for discussion, not a
    final decision.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$Path,

    [switch]$UseIIS,

    [string]$OutputPath = (Get-Location).Path,

    [switch]$IncludeSourceScan
)

$ErrorActionPreference = 'Stop'

#region Output helpers
function Write-Ok     { param([string]$Text) Write-Host "   OK   $Text" -ForegroundColor Green }
function Write-Warn   { param([string]$Text) Write-Host "   WARN $Text" -ForegroundColor Yellow }
function Write-Fail   { param([string]$Text) Write-Host "   FAIL $Text" -ForegroundColor Red }
function Write-Detail { param([string]$Text) Write-Host "        $Text" -ForegroundColor Gray }
function Write-Step   { param([string]$Text) Write-Host "-> $Text" -ForegroundColor Yellow }
#endregion

#region Known blocker patterns
# Assembly / package names that typically block porting an app to .NET 8+
# on Linux containers. Matched against <Reference Include="..."> and
# <PackageReference Include="..."> values (case-insensitive, regex).
$script:WindowsOnlyPatterns = @(
    'System\.Web(\.|$)',
    'System\.Web\.Mvc',
    'System\.DirectoryServices',
    'System\.EnterpriseServices',
    'Microsoft\.Office\.Interop',
    'System\.Windows\.Forms',
    'PresentationFramework',
    'PresentationCore',
    'System\.Management(\.|$)',
    '^Interop\.',
    '^AxInterop\.',
    'System\.Workflow'
)
#endregion

#region Discovery
function Get-FileSystemAppRoots {
    param([string[]]$Roots)

    $bag = @{}
    $excludePattern = '[\\/](bin|obj|node_modules|\.git|packages|\.vs)[\\/]'

    foreach ($root in $Roots) {
        if (-not (Test-Path -LiteralPath $root)) {
            Write-Warn "Path not found, skipping: $root"
            continue
        }

        Write-Step "Scanning $root for .NET application markers..."
        $markerFiles = Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object {
                $_.FullName -notmatch $excludePattern -and (
                    $_.Extension -eq '.csproj' -or
                    $_.Name -ieq 'web.config' -or
                    $_.Name -ieq 'packages.config' -or
                    $_.Name -like '*.runtimeconfig.json' -or
                    $_.Name -ieq 'appsettings.json'
                )
            }

        foreach ($f in $markerFiles) {
            $dir = $f.Directory.FullName
            if (-not $bag.ContainsKey($dir)) {
                $bag[$dir] = [ordered]@{
                    AppRoot          = $dir
                    Source           = 'FileSystem'
                    Csproj           = $null
                    WebConfig        = $null
                    PackagesConfig   = $null
                    RuntimeConfig    = $null
                    AppSettingsJson  = $null
                    IisNotes         = $null
                }
            }
            switch ($true) {
                { $f.Extension -eq '.csproj' }              { $bag[$dir].Csproj = $f.FullName }
                { $f.Name -ieq 'web.config' }                { $bag[$dir].WebConfig = $f.FullName }
                { $f.Name -ieq 'packages.config' }           { $bag[$dir].PackagesConfig = $f.FullName }
                { $f.Name -like '*.runtimeconfig.json' }     { $bag[$dir].RuntimeConfig = $f.FullName }
                { $f.Name -ieq 'appsettings.json' }          { $bag[$dir].AppSettingsJson = $f.FullName }
            }
        }
    }
    return $bag
}

function Get-IisAppInfo {
    if (-not (Get-Module -ListAvailable -Name WebAdministration)) {
        Write-Warn "WebAdministration module not found - skipping IIS discovery (Windows/IIS host required)."
        return @()
    }

    try {
        Import-Module WebAdministration -ErrorAction Stop
    } catch {
        Write-Warn "Could not load WebAdministration module: $($_.Exception.Message)"
        return @()
    }

    $results = @()
    try {
        foreach ($site in Get-Website) {
            $poolName = $site.applicationPool
            $pool = $null
            try { $pool = Get-ItemProperty "IIS:\AppPools\$poolName" -ErrorAction Stop } catch {}

            $results += [PSCustomObject]@{
                SiteName            = $site.name
                PhysicalPath        = [System.Environment]::ExpandEnvironmentVariables($site.physicalPath)
                AppPoolName         = $poolName
                ManagedRuntime      = if ($pool) { $pool.managedRuntimeVersion } else { 'Unknown' }
                Enable32BitAppOnWin64 = if ($pool) { $pool.enable32BitAppOnWin64 } else { 'Unknown' }
            }
        }
    } catch {
        Write-Warn "Failed to enumerate IIS sites: $($_.Exception.Message)"
    }
    return $results
}
#endregion

#region Parsers
function Get-XmlSafe {
    param([string]$LiteralPath)
    try {
        [xml]$xml = Get-Content -LiteralPath $LiteralPath -Raw
        return $xml
    } catch {
        Write-Warn "Could not parse XML file '$LiteralPath': $($_.Exception.Message)"
        return $null
    }
}

function Test-WindowsOnlyDependency {
    param([string[]]$Names)
    # NOTE: do not name this local variable $matches - it collides with the
    # automatic $Matches hashtable populated by the -match/-imatch operator.
    $matchedNames = @()
    foreach ($n in $Names) {
        if ([string]::IsNullOrWhiteSpace($n)) { continue }
        foreach ($pattern in $script:WindowsOnlyPatterns) {
            if ($n -imatch $pattern) {
                $matchedNames += $n
                break
            }
        }
    }
    return ($matchedNames | Select-Object -Unique)
}

function Get-DbProviderFromConnectionString {
    param([string]$ConnString, [string]$ProviderName)

    if ($ProviderName -imatch 'sqlclient|system\.data\.sql') { return 'SqlServer' }
    if ($ProviderName -imatch 'oracle')                       { return 'Oracle' }
    if ($ProviderName -imatch 'mysql')                        { return 'MySQL' }
    if ($ProviderName -imatch 'npgsql|postgres')              { return 'PostgreSQL' }

    if ([string]::IsNullOrWhiteSpace($ConnString)) { return $null }
    if ($ConnString -imatch 'Initial Catalog=|Data Source=.*\\|Server=.*;Database=') { return 'SqlServer' }
    if ($ConnString -imatch 'Npgsql|Host=.*Port=5432')        { return 'PostgreSQL' }
    if ($ConnString -imatch 'Uid=|MySql')                     { return 'MySQL' }
    if ($ConnString -imatch ':1521/|Data Source=.*ORCL')      { return 'Oracle' }
    if ($ConnString -imatch 'Server=|Data Source=')           { return 'SqlServer' }
    return 'Unknown'
}

function Get-CsprojInfo {
    param([string]$CsprojPath)

    $info = [ordered]@{
        IsSdkStyle          = $false
        TargetFramework     = $null
        PlatformTarget      = 'AnyCPU'
        HasComReference     = $false
        PackageNames        = @()
        ReferenceNames      = @()
    }

    $xml = Get-XmlSafe -LiteralPath $CsprojPath
    if (-not $xml) { return $info }

    $info.IsSdkStyle = [bool]$xml.Project.Sdk

    $tf = $xml.GetElementsByTagName('TargetFramework')  | Select-Object -First 1 -ExpandProperty '#text' -ErrorAction SilentlyContinue
    if (-not $tf) {
        $tf = $xml.GetElementsByTagName('TargetFrameworks') | Select-Object -First 1 -ExpandProperty '#text' -ErrorAction SilentlyContinue
        if ($tf) { $tf = ($tf -split ';')[0] }
    }
    if (-not $tf) {
        $tfv = $xml.GetElementsByTagName('TargetFrameworkVersion') | Select-Object -First 1 -ExpandProperty '#text' -ErrorAction SilentlyContinue
        if ($tfv) { $tf = "net" + ($tfv -replace '^v', '' -replace '\.', '') } # e.g. v4.8 -> net48
    }
    $info.TargetFramework = $tf

    $pt = $xml.GetElementsByTagName('PlatformTarget') | Select-Object -First 1 -ExpandProperty '#text' -ErrorAction SilentlyContinue
    if ($pt) { $info.PlatformTarget = $pt }

    $comNodes = $xml.GetElementsByTagName('COMReference')
    $info.HasComReference = ($comNodes.Count -gt 0)

    $pkgNodes = $xml.GetElementsByTagName('PackageReference')
    $info.PackageNames = @($pkgNodes | ForEach-Object { $_.Include } | Where-Object { $_ })

    $refNodes = $xml.GetElementsByTagName('Reference')
    $info.ReferenceNames = @($refNodes | ForEach-Object { $_.Include } | Where-Object { $_ })

    return $info
}

function Get-PackagesConfigInfo {
    param([string]$PackagesConfigPath)
    $names = @()
    $xml = Get-XmlSafe -LiteralPath $PackagesConfigPath
    if ($xml) {
        $names = @($xml.GetElementsByTagName('package') | ForEach-Object { $_.id } | Where-Object { $_ })
    }
    return $names
}

function Get-WebConfigInfo {
    param([string]$WebConfigPath)

    $info = [ordered]@{
        SessionStateMode      = $null
        AuthenticationMode    = $null
        CompilationTargetFwk  = $null
        HasAspNetCoreModule   = $false
        ConnectionStrings     = @()   # array of @{Name; ConnStr; ProviderName}
    }

    $xml = Get-XmlSafe -LiteralPath $WebConfigPath
    if (-not $xml) { return $info }

    $sessionNode = $xml.GetElementsByTagName('sessionState') | Select-Object -First 1
    if ($sessionNode) { $info.SessionStateMode = $sessionNode.mode }

    $authNode = $xml.GetElementsByTagName('authentication') | Select-Object -First 1
    if ($authNode) { $info.AuthenticationMode = $authNode.mode }

    $compNode = $xml.GetElementsByTagName('compilation') | Select-Object -First 1
    if ($compNode) { $info.CompilationTargetFwk = $compNode.targetFramework }

    $aspNetCoreNode = $xml.GetElementsByTagName('aspNetCore') | Select-Object -First 1
    $info.HasAspNetCoreModule = [bool]$aspNetCoreNode

    $addNodes = $xml.GetElementsByTagName('add')
    foreach ($n in $addNodes) {
        if ($n.ParentNode -and $n.ParentNode.LocalName -eq 'connectionStrings') {
            $info.ConnectionStrings += [PSCustomObject]@{
                Name         = $n.name
                ConnStr      = $n.connectionString
                ProviderName = $n.providerName
            }
        }
    }
    return $info
}

function Get-AppSettingsJsonInfo {
    param([string]$AppSettingsPath)
    $connStrings = @()
    try {
        $json = Get-Content -LiteralPath $AppSettingsPath -Raw | ConvertFrom-Json -ErrorAction Stop
        if ($json.ConnectionStrings) {
            $json.ConnectionStrings.PSObject.Properties | ForEach-Object {
                $connStrings += [PSCustomObject]@{
                    Name         = $_.Name
                    ConnStr      = $_.Value
                    ProviderName = $null
                }
            }
        }
    } catch {
        Write-Warn "Could not parse '$AppSettingsPath': $($_.Exception.Message)"
    }
    return $connStrings
}

function Test-LocalFileSystemDependency {
    param([string]$AppRoot)
    $csFiles = Get-ChildItem -LiteralPath $AppRoot -Recurse -File -Filter *.cs -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch '[\\/](bin|obj)[\\/]' }
    foreach ($f in $csFiles) {
        $content = Get-Content -LiteralPath $f.FullName -Raw -ErrorAction SilentlyContinue
        if ($content -imatch 'Server\.MapPath|HttpContext\.Current\.Server|[A-Za-z]:\\[^"''\s]+') {
            return $true
        }
    }
    return $false
}
#endregion

#region Scoring & recommendation
function Get-MigrationRecommendation {
    param(
        [string]$FrameworkFamily,
        [string]$SessionStateMode,
        [string]$AuthenticationMode,
        [bool]$HasComReference,
        [string[]]$WindowsOnlyDeps,
        [string]$PlatformTarget,
        [bool]$HasAspNetCoreModule,
        [bool]$HasLocalFsDependency
    )

    $readiness = 50
    $stickiness = 0

    if ($FrameworkFamily -eq '.NET Core/5+') { $readiness += 30 } else { $readiness -= 10 }
    if ($HasAspNetCoreModule)                { $readiness += 10 }

    if ($HasComReference)            { $readiness -= 20; $stickiness += 25 }
    if ($WindowsOnlyDeps.Count -gt 0){ $readiness -= 15; $stickiness += 20 }

    switch ($SessionStateMode) {
        'InProc'     { $readiness -= 15; $stickiness += 30 }
        'StateServer'{ $readiness += 0;  $stickiness += 10 }
        'SQLServer'  { $readiness += 5;  $stickiness += 5 }
        'Custom'     { $readiness += 10; $stickiness += 0 }
        default      { } # not found / stateless
    }

    if ($AuthenticationMode -eq 'Windows') { $readiness -= 10; $stickiness += 20 }
    if ($PlatformTarget -eq 'x86')         { $readiness -= 5;  $stickiness += 5 }
    if ($HasLocalFsDependency)             { $readiness -= 10; $stickiness += 15 }

    $readiness  = [Math]::Max(0, [Math]::Min(100, $readiness))
    $stickiness = [Math]::Max(0, [Math]::Min(100, $stickiness))

    if ($FrameworkFamily -eq '.NET Core/5+' -and -not $HasComReference -and $WindowsOnlyDeps.Count -eq 0 -and $readiness -ge 65) {
        $rec7R  = 'Replatform'
        $target = 'ECS Fargate or AWS App Runner (Linux containers)'
    }
    elseif ($FrameworkFamily -eq '.NET Core/5+') {
        $rec7R  = 'Replatform'
        $target = 'ECS/EKS on Windows containers or EC2; externalize session to ElastiCache (Redis) first'
    }
    elseif ($FrameworkFamily -eq '.NET Framework' -and ($HasComReference -or $WindowsOnlyDeps.Count -gt 0)) {
        $rec7R  = 'Rehost'
        $target = 'EC2 (Windows) via AWS Application Migration Service (MGN); revisit refactor later'
    }
    elseif ($FrameworkFamily -eq '.NET Framework') {
        $rec7R  = 'Refactor'
        $target = 'Port to .NET 8 with AWS Porting Assistant for .NET, then ECS Fargate (Linux)'
    }
    else {
        $rec7R  = 'Assess Manually'
        $target = 'Unknown - no project/config files detected, manual review required'
    }

    return [PSCustomObject]@{
        ReadinessScore   = $readiness
        StickinessScore  = $stickiness
        Recommended7R    = $rec7R
        RecommendedAwsTarget = $target
    }
}
#endregion

#region Main app assembly
function New-AppAssessment {
    param([hashtable]$Bag, [string]$IisNotes)

    $appName = Split-Path -Leaf $Bag.AppRoot
    $notes = New-Object System.Collections.Generic.List[string]
    if ($IisNotes) { $notes.Add($IisNotes) }

    $frameworkFamily = 'Unknown'
    $targetFramework = $null
    $platformTarget  = 'AnyCPU'
    $hasCom          = $false
    $winOnlyDeps     = @()
    $packageCount    = 0

    if ($Bag.Csproj) {
        $csInfo = Get-CsprojInfo -CsprojPath $Bag.Csproj
        $targetFramework = $csInfo.TargetFramework
        $platformTarget  = $csInfo.PlatformTarget
        $hasCom          = $csInfo.HasComReference
        $allNames        = @($csInfo.PackageNames) + @($csInfo.ReferenceNames)
        $packageCount    = $csInfo.PackageNames.Count

        if ($Bag.PackagesConfig) {
            $pkgNames = Get-PackagesConfigInfo -PackagesConfigPath $Bag.PackagesConfig
            $allNames += $pkgNames
            $packageCount += $pkgNames.Count
            $notes.Add('packages.config present alongside csproj (legacy NuGet style)')
        }
        $winOnlyDeps = Test-WindowsOnlyDependency -Names $allNames

        if ($targetFramework -imatch '^net(coreapp)?[5-9]|^net[1-9][0-9]?\.0') {
            $frameworkFamily = '.NET Core/5+'
        } elseif ($targetFramework -imatch '^net(standard)?4' -or $targetFramework -imatch '^netstandard') {
            $frameworkFamily = '.NET Framework'
        } elseif ($targetFramework) {
            $frameworkFamily = '.NET Core/5+'
        }
    }
    elseif ($Bag.PackagesConfig) {
        $pkgNames = Get-PackagesConfigInfo -PackagesConfigPath $Bag.PackagesConfig
        $winOnlyDeps = Test-WindowsOnlyDependency -Names $pkgNames
        $packageCount = $pkgNames.Count
        $frameworkFamily = '.NET Framework'
        $notes.Add('Legacy packages.config found (no SDK-style csproj) - implies .NET Framework project')
    }
    elseif ($Bag.RuntimeConfig) {
        $frameworkFamily = '.NET Core/5+'
        $notes.Add('Deployed binaries only (*.runtimeconfig.json found, no source csproj) - framework family inferred')
    }

    $sessionMode = $null
    $authMode    = $null
    $connStrings = @()
    $hasAspNetCoreModule = $false

    if ($Bag.WebConfig) {
        $webInfo = Get-WebConfigInfo -WebConfigPath $Bag.WebConfig
        $sessionMode = $webInfo.SessionStateMode
        $authMode    = $webInfo.AuthenticationMode
        $connStrings += $webInfo.ConnectionStrings
        $hasAspNetCoreModule = $webInfo.HasAspNetCoreModule
        if ($hasAspNetCoreModule -and $frameworkFamily -ne '.NET Core/5+') {
            $frameworkFamily = '.NET Core/5+'
            $notes.Add('web.config has <aspNetCore> handler - ASP.NET Core app hosted via IIS/ANCM')
        }
    }

    if ($Bag.AppSettingsJson) {
        $connStrings += (Get-AppSettingsJsonInfo -AppSettingsPath $Bag.AppSettingsJson)
        if ($frameworkFamily -eq 'Unknown') { $frameworkFamily = '.NET Core/5+' }
    }

    $dbProvider = 'None'
    if ($connStrings.Count -gt 0) {
        $first = $connStrings | Select-Object -First 1
        $dbProvider = Get-DbProviderFromConnectionString -ConnString $first.ConnStr -ProviderName $first.ProviderName
    }

    $hasLocalFsDependency = $false
    if ($IncludeSourceScan) {
        $hasLocalFsDependency = Test-LocalFileSystemDependency -AppRoot $Bag.AppRoot
        if ($hasLocalFsDependency) { $notes.Add('Local file-system dependency detected (Server.MapPath / hardcoded drive path) - plan S3 migration') }
    }

    $rec = Get-MigrationRecommendation -FrameworkFamily $frameworkFamily -SessionStateMode $sessionMode `
        -AuthenticationMode $authMode -HasComReference $hasCom -WindowsOnlyDeps $winOnlyDeps `
        -PlatformTarget $platformTarget -HasAspNetCoreModule $hasAspNetCoreModule -HasLocalFsDependency $hasLocalFsDependency

    if ($hasCom)              { $notes.Add('Contains COM/COM+ interop references') }
    if ($winOnlyDeps.Count -gt 0) { $notes.Add("Windows-only dependencies: $($winOnlyDeps -join ', ')") }
    if ($sessionMode -eq 'InProc') { $notes.Add('InProc session state blocks stateless horizontal scaling') }
    if ($authMode -eq 'Windows')   { $notes.Add('Windows Authentication - plan AWS Managed Microsoft AD / trust') }

    return [PSCustomObject]@{
        AppName              = $appName
        AppRoot              = $Bag.AppRoot
        DiscoverySource      = $Bag.Source
        FrameworkFamily      = $frameworkFamily
        TargetFramework      = $targetFramework
        PlatformTarget       = $platformTarget
        SessionStateMode     = $(if ($sessionMode) { $sessionMode } else { 'NotFound' })
        AuthenticationMode   = $(if ($authMode) { $authMode } else { 'NotFound' })
        DbProvider           = $dbProvider
        HasComReference      = $hasCom
        WindowsOnlyDepCount  = $winOnlyDeps.Count
        PackageCount         = $packageCount
        HasLocalFsDependency = $hasLocalFsDependency
        ReadinessScore       = $rec.ReadinessScore
        StickinessScore      = $rec.StickinessScore
        Recommended7R        = $rec.Recommended7R
        RecommendedAwsTarget = $rec.RecommendedAwsTarget
        Notes                = ($notes -join ' | ')
    }
}
#endregion

#region Main
if (-not $Path -and -not $UseIIS) {
    Write-Warn "No -Path or -UseIIS specified - defaulting to current directory."
    $Path = @((Get-Location).Path)
}

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$bags = @{}
if ($Path) {
    $fsBags = Get-FileSystemAppRoots -Roots $Path
    foreach ($k in $fsBags.Keys) { $bags[$k] = $fsBags[$k] }
}

$iisNotesByRoot = @{}
if ($UseIIS) {
    Write-Step "Enumerating local IIS sites..."
    $iisApps = Get-IisAppInfo
    foreach ($site in $iisApps) {
        if (-not $site.PhysicalPath -or -not (Test-Path -LiteralPath $site.PhysicalPath)) { continue }
        $iisBags = Get-FileSystemAppRoots -Roots @($site.PhysicalPath)
        foreach ($k in $iisBags.Keys) {
            if (-not $bags.ContainsKey($k)) { $bags[$k] = $iisBags[$k]; $bags[$k].Source = 'IIS' }
        }
        $note = "IIS site '$($site.SiteName)' (AppPool '$($site.AppPoolName)', managedRuntimeVersion='$($site.ManagedRuntime)', 32-bit=$($site.Enable32BitAppOnWin64))"
        $iisNotesByRoot[$site.PhysicalPath] = $note
    }
}

if ($bags.Count -eq 0) {
    Write-Fail "No .NET applications found (no *.csproj, web.config, packages.config, or *.runtimeconfig.json under the given path(s))."
    return
}

Write-Step "Found $($bags.Count) candidate application(s). Assessing..."
$results = New-Object System.Collections.Generic.List[object]
foreach ($key in $bags.Keys) {
    try {
        $iisNote = $iisNotesByRoot[$key]
        $results.Add((New-AppAssessment -Bag $bags[$key] -IisNotes $iisNote))
        Write-Ok "Assessed: $key"
    } catch {
        Write-Fail "Failed to assess '$key': $($_.Exception.Message)"
    }
}

$sorted = $results | Sort-Object -Property StickinessScore -Descending

Write-Host ""
Write-Host "==================== Assessment Summary ====================" -ForegroundColor Cyan
$sorted | Format-Table AppName, FrameworkFamily, TargetFramework, SessionStateMode, AuthenticationMode, DbProvider, StickinessScore, ReadinessScore, Recommended7R -AutoSize

Write-Host ""
Write-Host "---- Recommended 7R breakdown ----" -ForegroundColor Cyan
$sorted | Group-Object Recommended7R | Sort-Object Count -Descending |
    ForEach-Object { Write-Detail ("{0,-18} {1}" -f $_.Name, $_.Count) }

Write-Host ""
Write-Host "---- Portfolio stickiness indicators ----" -ForegroundColor Cyan
Write-Detail ("InProc session state apps : {0}" -f ($sorted | Where-Object { $_.SessionStateMode -eq 'InProc' }).Count)
Write-Detail ("Windows Authentication     : {0}" -f ($sorted | Where-Object { $_.AuthenticationMode -eq 'Windows' }).Count)
Write-Detail ("COM/COM+ interop           : {0}" -f ($sorted | Where-Object { $_.HasComReference }).Count)
Write-Detail ("Windows-only dependencies  : {0}" -f ($sorted | Where-Object { $_.WindowsOnlyDepCount -gt 0 }).Count)
Write-Detail ("32-bit (x86) platform      : {0}" -f ($sorted | Where-Object { $_.PlatformTarget -eq 'x86' }).Count)
Write-Detail (".NET Framework apps        : {0}" -f ($sorted | Where-Object { $_.FrameworkFamily -eq '.NET Framework' }).Count)
Write-Detail (".NET Core/5+ apps          : {0}" -f ($sorted | Where-Object { $_.FrameworkFamily -eq '.NET Core/5+' }).Count)

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvPath  = Join-Path $OutputPath "DotNet_AWS_Assessment_$timestamp.csv"
$jsonPath = Join-Path $OutputPath "DotNet_AWS_Assessment_$timestamp.json"

$sorted | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
$sorted | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

Write-Host ""
Write-Ok "Report written to: $csvPath"
Write-Ok "Report written to: $jsonPath"
#endregion
