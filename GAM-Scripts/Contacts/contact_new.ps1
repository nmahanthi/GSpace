param(
    [string]$OutputCsv = ".\ContactDelegates_Final.csv",
    [string]$WorkingFolder = ".\GAM_ContactDelegates_Work",
    [string]$GamExe = "",
    [string]$ConfigFile = "",
    [string[]]$Users = @(),
    [string]$UsersFile = ""
)

$ErrorActionPreference = "Stop"

$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

function Resolve-AbsolutePath {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $ScriptRoot $Path))
}

$WorkingFolder = Resolve-AbsolutePath -Path $WorkingFolder
$OutputCsv = Resolve-AbsolutePath -Path $OutputCsv

function Write-Log {
    param([string]$Message)
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

function Confirm-Folder {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Get-GamConfig {
    param([string]$ConfigFile)

    $candidates = New-Object System.Collections.Generic.List[string]

    if ($ConfigFile) {
        $candidates.Add((Resolve-AbsolutePath -Path $ConfigFile))
    }

    if ($PSScriptRoot) {
        $candidates.Add((Join-Path $PSScriptRoot "gam.config.json"))
        $candidates.Add((Join-Path $PSScriptRoot "contact.config.json"))
    }

    if ($env:USERPROFILE) {
        $candidates.Add((Join-Path $env:USERPROFILE ".gam-scripts.json"))
    }

    foreach ($path in ($candidates | Select-Object -Unique)) {
        if ($path -and (Test-Path -LiteralPath $path)) {
            try {
                $json = Get-Content -LiteralPath $path -Raw -Encoding UTF8 |
                ConvertFrom-Json
                Write-Log "Loaded config: $path"
                return $json
            }
            catch {
                Write-Log "WARNING: Failed to parse config '$path': $($_.Exception.Message)"
            }
        }
    }

    return $null
}

function Resolve-GamExe {
    param(
        [string]$GamExe,
        [object]$Config
    )

    $candidates = New-Object System.Collections.Generic.List[string]

    function Add-GamCandidate {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return }
        if ((Split-Path -Leaf $Path) -ieq "gam.exe") {
            $candidates.Add($Path)
        }
        else {
            $candidates.Add((Join-Path $Path "gam.exe"))
        }
    }

    if ($GamExe) { Add-GamCandidate $GamExe }

    # 1. Explicit environment overrides
    foreach ($envVar in @("GAM", "GAM_EXE", "GAM_PATH", "GAMPATH")) {
        Add-GamCandidate ([Environment]::GetEnvironmentVariable($envVar))
    }

    # 2. Value from config file (GamExe or GamPath)
    if ($Config) {
        foreach ($key in @("GamExe", "GamPath", "gamExe", "gamPath")) {
            $value = $Config.PSObject.Properties[$key]
            if ($value -and $value.Value) { Add-GamCandidate ([string]$value.Value) }
        }
    }

    # 3. Alongside this script
    if ($PSScriptRoot) {
        Add-GamCandidate (Join-Path $PSScriptRoot "gam.exe")
        Add-GamCandidate (Join-Path $PSScriptRoot "gam7\gam.exe")
        Add-GamCandidate (Join-Path $PSScriptRoot "GAMADV-XTD3\gam.exe")
    }

    # 4. Well-known install roots (drive roots, Program Files, user profile)
    $roots = @(
        "C:\", "D:\",
        $env:ProgramFiles,
        ${env:ProgramFiles(x86)},
        $env:LOCALAPPDATA,
        $env:APPDATA,
        $env:USERPROFILE,
        (Join-Path $env:USERPROFILE "Documents")
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique

    $subdirs = @("GAM7", "gam7", "GAM", "gam", "GAMADV-XTD3", "GAMADV-XTD")

    foreach ($root in $roots) {
        foreach ($sub in $subdirs) {
            Add-GamCandidate (Join-Path $root (Join-Path $sub "gam.exe"))
        }
    }

    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    # 5. PATH lookup (gam.exe or gam.bat shim)
    foreach ($name in @("gam.exe", "gam.bat", "gam.cmd", "gam")) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cmd -and $cmd.Source) { return $cmd.Source }
    }

    # 6. Last-resort recursive search of common roots (depth-limited)
    foreach ($root in $roots) {
        try {
            $found = Get-ChildItem -LiteralPath $root -Filter "gam.exe" -Recurse -Depth 4 `
                -ErrorAction SilentlyContinue -Force | Select-Object -First 1
            if ($found) { return $found.FullName }
        }
        catch { }
    }

    throw "GAM executable not found. Pass -GamExe <path>, set the GAM environment variable, or add a 'GamExe' entry to gam.config.json."
}

function Join-NativeArguments {
    param([string[]]$Arguments)

    $escaped = foreach ($arg in $Arguments) {
        if ($null -eq $arg) {
            '""'
        }
        elseif ($arg -match '[\s"]') {
            '"' + ($arg -replace '"', '\"') + '"'
        }
        else {
            $arg
        }
    }

    return ($escaped -join ' ')
}

function Invoke-Gam {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $argString = Join-NativeArguments -Arguments $Arguments
    Write-Log ("Running: " + $script:GamExeResolved + " " + $argString)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $script:GamExeResolved
    $psi.Arguments = $argString
    $psi.WorkingDirectory = $ScriptRoot
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi

    $null = $process.Start()

    # Read stderr asynchronously so its buffer never fills and causes a deadlock
    # while ReadToEnd() is blocking on stdout (classic two-stream deadlock pattern).
    $stderrTask = $process.StandardError.ReadToEndAsync()
    $stdout = $process.StandardOutput.ReadToEnd()
    $process.WaitForExit()
    $stderr = $stderrTask.GetAwaiter().GetResult()

    # Only log on failure to avoid printing every GAM progress line to the console.
    if ($process.ExitCode -ne 0) {
        if (-not [string]::IsNullOrWhiteSpace($stderr)) {
            ($stderr -split "`r?`n") | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_)) { Write-Log "STDERR: $_" }
            }
        }
        throw "GAM command failed with exit code $($process.ExitCode).`nCommand: $script:GamExeResolved $argString`nSTDERR:`n$stderr`nSTDOUT:`n$stdout"
    }

    return [PSCustomObject]@{
        ExitCode = $process.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}

function Get-FirstExistingPropertyName {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Object,
        [Parameter(Mandatory = $true)]
        [string[]]$CandidateNames
    )

    $props = $Object.PSObject.Properties.Name

    foreach ($name in $CandidateNames) {
        if ($props -contains $name) {
            return $name
        }
    }

    foreach ($candidate in $CandidateNames) {
        $match = $props | Where-Object { $_.ToLower() -eq $candidate.ToLower() } | Select-Object -First 1
        if ($match) {
            return $match
        }
    }

    return $null
}

function New-EmptyOutput {
    param([string]$Path)

    @() | Select-Object `
    @{Name = 'User Display name'; Expression = { '' } }, `
    @{Name = 'User EmailAddress'; Expression = { '' } }, `
    @{Name = 'delegateAddress'; Expression = { '' } }, `
    @{Name = 'delegate name'; Expression = { '' } }, `
    @{Name = 'count'; Expression = { 0 } } |
    Select-Object -First 0 |
    Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Get-DelegateColumns {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SampleObject,
        [Parameter(Mandatory = $true)]
        [string]$OwnerColumn
    )

    $props = $SampleObject.PSObject.Properties.Name

    $preferred = $props | Where-Object {
        $_ -ne $OwnerColumn -and (
            $_ -match '(?i)delegate' -or
            $_ -match '(?i)contactdelegate'
        )
    }

    if ($preferred -and $preferred.Count -gt 0) {
        return $preferred
    }

    return ($props | Where-Object { $_ -ne $OwnerColumn })
}

function Get-EmailsFromText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return @()
    }

    $emailMatches = [regex]::Matches(
        $Text,
        "(?i)[A-Z0-9._%+\-']+@[A-Z0-9.\-]+\.[A-Z]{2,}"
    )

    $emails = foreach ($m in $emailMatches) {
        $m.Value.Trim()
    }

    return $emails | Sort-Object -Unique
}

function Get-DelegateEmailsFromRow {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Row,
        [Parameter(Mandatory = $true)]
        [string[]]$DelegateColumns
    )

    $emails = New-Object System.Collections.Generic.List[string]

    foreach ($col in $DelegateColumns) {
        $value = [string]$Row.$col
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        $found = Get-EmailsFromText -Text $value
        foreach ($email in $found) {
            if (-not [string]::IsNullOrWhiteSpace($email)) {
                $emails.Add($email.Trim())
            }
        }
    }

    return $emails | Sort-Object -Unique
}

function Resolve-TargetUsers {
    param(
        [string[]]$Users,
        [string]$UsersFile,
        [object]$Config
    )

    $result = New-Object System.Collections.Generic.List[string]
    $seen = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    $emailPattern = "^[A-Za-z0-9._%+\-']+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$"

    $addEmail = {
        param([string]$Email)
        if ([string]::IsNullOrWhiteSpace($Email)) { return }
        $trimmed = $Email.Trim().Trim('"').Trim("'")
        if ($trimmed -match $emailPattern -and $seen.Add($trimmed)) {
            $result.Add($trimmed)
        }
    }

    foreach ($u in $Users) {
        foreach ($part in ($u -split '[,;\s]+')) { & $addEmail $part }
    }

    $fileToUse = $UsersFile
    if (-not $fileToUse -and $Config) {
        foreach ($key in @("UsersFile", "usersFile", "TargetsFile")) {
            $v = $Config.PSObject.Properties[$key]
            if ($v -and $v.Value) { $fileToUse = [string]$v.Value; break }
        }
    }

    if ($fileToUse) {
        $resolvedPath = Resolve-AbsolutePath -Path $fileToUse
        if (-not (Test-Path -LiteralPath $resolvedPath)) {
            throw "UsersFile not found: $resolvedPath"
        }

        $firstLine = Get-Content -LiteralPath $resolvedPath -TotalCount 1 -Encoding UTF8
        $looksLikeCsv = $firstLine -and ($firstLine -match '(?i)email|user|primary')

        if ($looksLikeCsv) {
            $rows = Import-Csv -LiteralPath $resolvedPath
            if ($rows -and $rows.Count -gt 0) {
                $emailCol = Get-FirstExistingPropertyName -Object $rows[0] -CandidateNames @(
                    "primaryEmail", "PrimaryEmail", "email", "Email",
                    "userEmail", "UserEmail", "user", "User"
                )
                if (-not $emailCol) {
                    $cols = ($rows[0].PSObject.Properties.Name) -join ", "
                    throw "UsersFile CSV has no email column. Columns found: $cols"
                }
                foreach ($row in $rows) { & $addEmail ([string]$row.$emailCol) }
            }
        }
        else {
            Get-Content -LiteralPath $resolvedPath -Encoding UTF8 | ForEach-Object {
                & $addEmail $_
            }
        }
    }

    if ($result.Count -eq 0 -and $Config) {
        foreach ($key in @("Users", "users", "TargetUsers")) {
            $v = $Config.PSObject.Properties[$key]
            if ($v -and $v.Value) {
                foreach ($e in $v.Value) { & $addEmail ([string]$e) }
                break
            }
        }
    }

    return , ($result.ToArray())
}

function Get-UserDisplayName {
    param([string]$Email)

    try {
        $r = Invoke-Gam -Arguments @("info", "user", $Email, "fields", "fullname")
        foreach ($line in ($r.StdOut -split "`r?`n")) {
            if ($line -match '^\s*Full Name:\s*(.+)\s*$') {
                return $Matches[1].Trim()
            }
        }
    }
    catch {
        Write-Log "WARNING: info user failed for ${Email}: $($_.Exception.Message)"
    }
    return ""
}

$script:Config = Get-GamConfig -ConfigFile $ConfigFile
$script:GamExeResolved = Resolve-GamExe -GamExe $GamExe -Config $script:Config
Write-Log "Using GAM executable: $script:GamExeResolved"

Write-Log "Validating GAM executable..."
Invoke-Gam -Arguments @("version") | Out-Null

$targetUsers = Resolve-TargetUsers -Users $Users -UsersFile $UsersFile -Config $script:Config
$isTargeted = $targetUsers.Count -gt 0

if ($isTargeted) {
    Write-Log "Targeted mode: $($targetUsers.Count) user(s) selected."
}
else {
    Write-Log "No -Users / -UsersFile provided - scanning ALL users in the tenant."
}

Confirm-Folder -Path $WorkingFolder

$usersCsv = Join-Path $WorkingFolder "Users.csv"
$delegatesCsv = Join-Path $WorkingFolder "ContactDelegates.csv"

if (Test-Path -LiteralPath $usersCsv) { Remove-Item -LiteralPath $usersCsv     -Force }
if (Test-Path -LiteralPath $delegatesCsv) { Remove-Item -LiteralPath $delegatesCsv -Force }
if (Test-Path -LiteralPath $OutputCsv) { Remove-Item -LiteralPath $OutputCsv    -Force }

# GAM's contactdelegates output does not include the owner's display name,
# so we fetch users separately to populate 'User Display name'.
if ($isTargeted) {
    Write-Log "Looking up display names for $($targetUsers.Count) target user(s)..."
    $userRows = foreach ($email in $targetUsers) {
        [PSCustomObject]@{
            primaryEmail    = $email
            'name.fullName' = (Get-UserDisplayName -Email $email)
        }
    }
    $userRows | Export-Csv -LiteralPath $usersCsv -NoTypeInformation -Encoding UTF8

    Write-Log "Exporting contact delegates for $($targetUsers.Count) target user(s)..."
    Invoke-Gam -Arguments @(
        "redirect", "csv", $delegatesCsv,
        "users", ($targetUsers -join ','),
        "print", "contactdelegates", "shownames"
    ) | Out-Null
}
else {
    Write-Log "Exporting users (for display name lookup)..."
    Invoke-Gam -Arguments @(
        "redirect", "csv", $usersCsv,
        "print", "users", "name"
    ) | Out-Null

    Write-Log "Exporting contact delegates (shownames for delegate display name)..."
    Invoke-Gam -Arguments @(
        "redirect", "csv", $delegatesCsv,
        "all", "users", "print", "contactdelegates", "shownames"
    ) | Out-Null
}

if (-not (Test-Path -LiteralPath $usersCsv)) {
    throw "Users export file was not created: $usersCsv"
}
if (-not (Test-Path -LiteralPath $delegatesCsv)) {
    throw "Contact delegates export file was not created: $delegatesCsv"
}

$users = Import-Csv -LiteralPath $usersCsv
$delegates = Import-Csv -LiteralPath $delegatesCsv

if (-not $delegates -or $delegates.Count -eq 0) {
    Write-Log "No contact delegates found. Writing empty output."
    New-EmptyOutput -Path $OutputCsv
    exit 0
}

# Detect columns — GAM7 actual column names confirmed from live output:
#   User, delegateAddress, delegateName
$ownerColumn = Get-FirstExistingPropertyName -Object $delegates[0] -CandidateNames @(
    "User", "user", "primaryEmail", "PrimaryEmail", "owner", "Owner"
)

# GAM outputs 'delegateAddress' (not 'delegateEmail') for the delegate's email.
$delegateEmailColumn = Get-FirstExistingPropertyName -Object $delegates[0] -CandidateNames @(
    "delegateAddress", "DelegateAddress", "delegateEmail", "DelegateEmail", "delegate", "email"
)

# 'shownames' adds delegateName for the delegate's display name.
$delegateNameColumn = Get-FirstExistingPropertyName -Object $delegates[0] -CandidateNames @(
    "delegateName", "DelegateName", "delegateDisplayName"
)

if (-not $ownerColumn -or -not $delegateEmailColumn) {
    $foundCols = ($delegates[0].PSObject.Properties.Name) -join ", "
    throw "Could not detect required columns. Columns found in CSV: $foundCols"
}

# Build owner display-name lookup from the separate users export.
$userEmailCol = Get-FirstExistingPropertyName -Object $users[0] -CandidateNames @(
    "primaryEmail", "PrimaryEmail", "email", "Email"
)
$userNameCol = Get-FirstExistingPropertyName -Object $users[0] -CandidateNames @(
    "name.fullName", "Name.FullName", "fullName", "FullName", "name", "Name"
)

$userLookup = @{}
foreach ($u in $users) {
    $email = [string]$u.$userEmailCol
    if ([string]::IsNullOrWhiteSpace($email)) { continue }
    $key = $email.Trim().ToLower()
    $name = if ($userNameCol) { [string]$u.$userNameCol } else { "" }
    $userLookup[$key] = if ($name) { $name } else { $email.Trim() }
}

$normalized = New-Object System.Collections.Generic.List[object]

# Use a HashSet to deduplicate: same owner + same delegate should never produce two rows.
$seen = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

foreach ($row in $delegates) {
    $ownerEmail = ([string]$row.$ownerColumn).Trim()
    if ([string]::IsNullOrWhiteSpace($ownerEmail)) { continue }

    $delEmail = ([string]$row.$delegateEmailColumn).Trim()
    if ([string]::IsNullOrWhiteSpace($delEmail)) { continue }

    # Skip duplicates — count will then always equal the number of visible rows.
    $pairKey = "$ownerEmail|$delEmail"
    if (-not $seen.Add($pairKey)) { continue }

    $delName = if ($delegateNameColumn) { ([string]$row.$delegateNameColumn).Trim() } else { "" }

    $normalized.Add([PSCustomObject]@{
            OwnerEmail    = $ownerEmail
            DelegateEmail = $delEmail
            DelegateName  = $delName
        })
}

if ($normalized.Count -eq 0) {
    Write-Log "No usable delegate relationships found after normalization. Writing empty output."
    New-EmptyOutput -Path $OutputCsv
    exit 0
}

# One row per owner: group all delegates so count always equals the number of
# semicolon-separated values visible in 'delegateAddress' and 'delegate name'.
$result = $normalized |
Group-Object -Property OwnerEmail |
ForEach-Object {
    $ownerEmail = $_.Name
    $ownerKey = $ownerEmail.ToLower()
    $sortedGroup = $_.Group | Sort-Object DelegateEmail

    [PSCustomObject]@{
        'User Display name' = $userLookup[$ownerKey]
        'User EmailAddress' = $ownerEmail
        'delegateAddress'   = ($sortedGroup | ForEach-Object { $_.DelegateEmail }) -join "; "
        'delegate name'     = ($sortedGroup | ForEach-Object { $_.DelegateName }) -join "; "
        'count'             = $_.Group.Count
    }
}

$result |
Sort-Object 'User EmailAddress' |
Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8

Write-Log "Completed successfully."
Write-Log "Output file: $OutputCsv"