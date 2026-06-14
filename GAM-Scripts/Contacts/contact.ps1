param(
    [string]$OutputCsv = ".\ContactDelegates_Final.csv",
    [string]$WorkingFolder = ".\GAM_ContactDelegates_Work",
    [string]$GamExe = ""
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

function Resolve-GamExe {
    param([string]$GamExe)

    $candidates = New-Object System.Collections.Generic.List[string]

    if ($GamExe) {
        $candidates.Add($GamExe)
    }

    if ($PSScriptRoot) {
        $candidates.Add((Join-Path $PSScriptRoot "gam.exe"))
    }

    $candidates.Add("C:\GAM7\gam.exe")
    $candidates.Add("C:\GAMADV-XTD3\gam.exe")
    $candidates.Add("C:\Users\v-nmahanthi\OneDrive - Microsoft\Documents\gam7\gam.exe")

    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    $cmd = Get-Command gam -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    throw "GAM executable not found."
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

$script:GamExeResolved = Resolve-GamExe -GamExe $GamExe
Write-Log "Using GAM executable: $script:GamExeResolved"

Write-Log "Validating GAM executable..."
Invoke-Gam -Arguments @("version") | Out-Null

Confirm-Folder -Path $WorkingFolder

$usersCsv = Join-Path $WorkingFolder "Users.csv"
$delegatesCsv = Join-Path $WorkingFolder "ContactDelegates.csv"

if (Test-Path -LiteralPath $usersCsv) { Remove-Item -LiteralPath $usersCsv     -Force }
if (Test-Path -LiteralPath $delegatesCsv) { Remove-Item -LiteralPath $delegatesCsv -Force }
if (Test-Path -LiteralPath $OutputCsv) { Remove-Item -LiteralPath $OutputCsv    -Force }

# GAM's contactdelegates output does not include the owner's display name,
# so we fetch users separately to populate 'User Display name'.
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