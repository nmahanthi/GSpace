<#
.SYNOPSIS
  Targeted Google Tasks scan — runs Get-GoogleTasksWithCreator.ps1 for a
  specific list of users instead of the whole tenant.

.DESCRIPTION
  Reads a plain-text user list (one primaryEmail per line) or accepts an
  inline comma-separated list, then delegates to Get-GoogleTasksWithCreator.ps1
  with settings tuned for small, targeted runs:

    - Tenant-wide Chat space scan is SKIPPED by default (use -ScanSpaces to
      enable it; it is expensive and only useful for full-tenant audits).
    - Doc-comment scan is ON by default for targeted runs (few users = fast).
    - Output is stamped with date/time so repeated runs do not overwrite.
    - A pre-flight summary is printed before any GAM calls are made.

.PARAMETER UsersFile
  Path to a plain-text file with one email per line (lines starting with '#'
  are ignored). Defaults to .\TargetUsers.txt in the same folder as this script.

.PARAMETER Users
  Inline comma-separated list of primaryEmail addresses. Overrides -UsersFile.

.PARAMETER OutputDir
  Folder for output CSVs. Defaults to the same folder as this script.

.PARAMETER ConfigFile
  Path to a .psd1 config file. Defaults to TaskScan.config.psd1 in the same
  folder as this script. Edit that file once to set GamPath, OutputDir, etc.
  and then run the script with no extra arguments.

.PARAMETER GamPath
  Path to gam executable. Overrides the config file value. If neither is set
  the script auto-detects GAM from PATH and common install locations.

.PARAMETER IncludeCompleted
  Include completed tasks in the output.

.PARAMETER IncludeHidden
  Include hidden tasks.

.PARAMETER IncludeDeleted
  Include deleted tasks.

.PARAMETER ScanSpaces
  Enable the tenant-wide Chat space scan (disabled by default for targeted runs).

.PARAMETER SkipDocCommentScan
  Skip the Drive comment scan (enabled by default for targeted runs).

.PARAMETER CheckpointCsv
  Optional path for an incremental checkpoint CSV updated after each user.

.EXAMPLE
  # Scan the users listed in TargetUsers.txt (default)
  .\Invoke-TargetedTaskScan.ps1

.EXAMPLE
  # Scan two specific users inline
  .\Invoke-TargetedTaskScan.ps1 -Users alice@corp.com,bob@corp.com

.EXAMPLE
  # Use a custom users file and include completed tasks
  .\Invoke-TargetedTaskScan.ps1 -UsersFile C:\lists\hr_team.txt -IncludeCompleted
#>
[CmdletBinding()]
param(
    [string]$ConfigFile = (Join-Path $PSScriptRoot 'TaskScan.config.psd1'),
    [string]$UsersFile,
    [string[]]$Users,
    [string]$OutputDir,
    [string]$GamPath = '',
    [string]$KeyFile = '',
    [switch]$IncludeCompleted,
    [switch]$IncludeHidden,
    [switch]$IncludeDeleted,
    [switch]$ScanSpaces,
    [switch]$SkipDocCommentScan,
    [string]$CheckpointCsv
)

$ErrorActionPreference = 'Stop'
$mainScript = Join-Path $PSScriptRoot 'Get-GoogleTasksWithCreator.ps1'

# ── Load config file (values fill in anything not set on the command line) ───
$cfg = @{}
if (Test-Path $ConfigFile) {
    $cfg = Import-PowerShellDataFile -LiteralPath $ConfigFile
    Write-Host ("Config loaded: {0}" -f $ConfigFile) -ForegroundColor DarkGray
}
else {
    Write-Warning "Config file not found at '$ConfigFile'. Using defaults / auto-detection."
}

# Apply config values for settings that were NOT explicitly passed
# (switches default to $false so we check the config to honour $true there)
if (-not $GamPath)       { $GamPath   = if ($cfg.GamPath  -and $cfg.GamPath  -ne '') { $cfg.GamPath  } else { '' } }
if (-not $KeyFile)       { $KeyFile   = if ($cfg.KeyFile   -and $cfg.KeyFile   -ne '') { $cfg.KeyFile   } else { '' } }
if (-not $UsersFile)     { $UsersFile = if ($cfg.UsersFile -and $cfg.UsersFile -ne '') { $cfg.UsersFile } else { Join-Path $PSScriptRoot 'TargetUsers.txt' } }
if (-not $OutputDir)     { $OutputDir = if ($cfg.OutputDir -and $cfg.OutputDir -ne '') { $cfg.OutputDir } else { $PSScriptRoot } }
if (-not $IncludeCompleted.IsPresent -and $cfg.IncludeCompleted) { $IncludeCompleted = [switch]$true }
if (-not $IncludeHidden.IsPresent    -and $cfg.IncludeHidden)    { $IncludeHidden    = [switch]$true }
if (-not $IncludeDeleted.IsPresent   -and $cfg.IncludeDeleted)   { $IncludeDeleted   = [switch]$true }
if (-not $ScanSpaces.IsPresent       -and $cfg.ScanSpaces)       { $ScanSpaces       = [switch]$true }
if (-not $SkipDocCommentScan.IsPresent -and $cfg.SkipDocCommentScan) { $SkipDocCommentScan = [switch]$true }

# ── Auto-detect GAM executable ───────────────────────────────────────────────
function Find-Gam {
    # 1. Explicit -GamPath supplied by caller
    if ($GamPath -and $GamPath -ne '') {
        if (Test-Path $GamPath) { return $GamPath }
        throw "GAM not found at the path you supplied: $GamPath"
    }

    # 2. Already on PATH (works if the user ran 'gam config' setup)
    $onPath = Get-Command gam -ErrorAction SilentlyContinue
    if ($onPath) { return $onPath.Source }

    # 3. Common fixed install locations (GAM7 installer defaults + variants)
    $candidates = @(
        'C:\GAM7\gam.exe',
        'C:\GAM\gam.exe',
        (Join-Path $env:LOCALAPPDATA  'GAM7\gam.exe'),
        (Join-Path $env:LOCALAPPDATA  'GAM\gam.exe'),
        (Join-Path $env:USERPROFILE   'GAM7\gam.exe'),
        (Join-Path $env:USERPROFILE   'GAM\gam.exe'),
        (Join-Path $env:ProgramData   'GAM7\gam.exe'),
        (Join-Path $env:ProgramData   'GAM\gam.exe'),
        (Join-Path $env:ProgramFiles  'GAM7\gam.exe'),
        (Join-Path $env:ProgramFiles  'GAM\gam.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'GAM7\gam.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'GAM\gam.exe')
    )
    foreach ($c in $candidates) {
        if ($c -and (Test-Path $c)) { return $c }
    }

    # 4. Search the GAM config directory for a saved gampath hint
    $gamCfgDir = Join-Path $env:USERPROFILE '.gam'
    $gamPathHint = Join-Path $gamCfgDir 'gampath'
    if (Test-Path $gamPathHint) {
        $hint = (Get-Content -LiteralPath $gamPathHint -TotalCount 1).Trim()
        $exe  = Join-Path $hint 'gam.exe'
        if (Test-Path $exe) { return $exe }
    }

    throw @"
Cannot find the GAM executable. Tried PATH, common install folders, and ~/.gam/gampath.
Fix options:
  A) Re-run with -GamPath:
       .\Invoke-TargetedTaskScan.ps1 -GamPath 'C:\GAM7\gam.exe'
  B) Add GAM to PATH and restart PowerShell.
  C) Install GAM7: https://github.com/GAM-team/GAM/wiki/How-to-Install-Advanced-GAM
"@
}

$resolvedGamPath = Find-Gam
Write-Host ("GAM detected : {0}" -f $resolvedGamPath) -ForegroundColor DarkGray

# ── Auto-detect oauth2service.json key file ──────────────────────────────────
function Find-GamKeyFile {
    param([string]$GamExePath)

    # 1. Explicit -KeyFile supplied by caller or config file
    if ($KeyFile -and $KeyFile -ne '') {
        if (Test-Path $KeyFile) { return $KeyFile }
        throw "oauth2service.json not found at the path you specified: $KeyFile"
    }

    # 2. GAMCFGDIR environment variable (overrides default GAM config location)
    if ($env:GAMCFGDIR) {
        $c = Join-Path $env:GAMCFGDIR 'oauth2service.json'
        if (Test-Path $c) { return $c }
    }

    # 3. Default GAM config folder: %USERPROFILE%\.gam\
    $c = Join-Path $env:USERPROFILE '.gam\oauth2service.json'
    if (Test-Path $c) { return $c }

    # 4. Same folder as gam.exe  (some installs keep config alongside the exe)
    if ($GamExePath) {
        $c = Join-Path (Split-Path $GamExePath) 'oauth2service.json'
        if (Test-Path $c) { return $c }
    }

    # 5. Common alternative GAM config locations
    $candidates = @(
        (Join-Path $env:USERPROFILE   '.gam7\oauth2service.json'),
        (Join-Path $env:LOCALAPPDATA  'GAM7\oauth2service.json'),
        (Join-Path $env:LOCALAPPDATA  'GAM\oauth2service.json'),
        (Join-Path $env:ProgramData   'GAM7\oauth2service.json'),
        (Join-Path $env:ProgramData   'GAM\oauth2service.json')
    )
    foreach ($c in $candidates) {
        if ($c -and (Test-Path $c)) { return $c }
    }

    # 6. Not found — warn but do not throw; the REST pass is optional
    Write-Warning @"
oauth2service.json not found. The REST-based assigned-task fetch (Docs/Chat tasks) will be skipped.
To fix, set KeyFile in TaskScan.config.psd1:
  KeyFile = 'C:\Users\<you>\.gam\oauth2service.json'
To locate the file, run:  gam info domain
"@
    return ''
}

$resolvedKeyFile = Find-GamKeyFile -GamExePath $resolvedGamPath
if ($resolvedKeyFile) {
    Write-Host ("Key file     : {0}" -f $resolvedKeyFile) -ForegroundColor DarkGray
}
if (-not (Test-Path $mainScript)) {
    throw "Cannot find Get-GoogleTasksWithCreator.ps1 alongside this script at: $mainScript"
}

# ── Resolve the user list ────────────────────────────────────────────────────
$resolvedUsers = @()

if ($Users -and $Users.Count -gt 0) {
    # Inline list wins over file
    $resolvedUsers = $Users | Where-Object { $_ -match '@' } | ForEach-Object { $_.Trim() }
}
elseif (Test-Path $UsersFile) {
    $resolvedUsers = Get-Content -LiteralPath $UsersFile |
        Where-Object { $_ -and $_.Trim() -ne '' -and $_.Trim() -notmatch '^#' } |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -match '@' }
}
else {
    throw "No users provided and UsersFile not found: $UsersFile`nCreate TargetUsers.txt or pass -Users alice@corp.com,bob@corp.com"
}

if ($resolvedUsers.Count -eq 0) {
    throw "User list is empty. Add at least one primaryEmail to $UsersFile or use -Users."
}

# ── Pre-flight summary ───────────────────────────────────────────────────────
Write-Host ""
Write-Host "=== Targeted Google Tasks Scan ===" -ForegroundColor Cyan
Write-Host ("  Users to scan : {0}" -f $resolvedUsers.Count) -ForegroundColor White
Write-Host ("  Doc scan      : {0}" -f $(if ($SkipDocCommentScan) { 'SKIPPED' } else { 'ON (default for targeted)' })) -ForegroundColor White
Write-Host ("  Space scan    : {0}" -f $(if ($ScanSpaces) { 'ON (tenant-wide — may be slow)' } else { 'SKIPPED (use -ScanSpaces to enable)' })) -ForegroundColor White
Write-Host ""
Write-Host "  Users:" -ForegroundColor DarkGray
$resolvedUsers | ForEach-Object { Write-Host ("    - {0}" -f $_) -ForegroundColor DarkGray }
Write-Host ""

# ── Build output paths ───────────────────────────────────────────────────────
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
$outputCsv  = Join-Path $OutputDir "Tasks_Targeted_$stamp.csv"
$summaryCsv = Join-Path $OutputDir "Tasks_Targeted_${stamp}_Summary.csv"

# ── Build argument hashtable (hashtable splatting avoids [string[]] greediness
#    in Windows PowerShell 5 that corrupts positional array splatting) ────────
$scriptArgs = @{
    Users      = $resolvedUsers          # pass the already-resolved array directly
    OutputCsv  = $outputCsv
    SummaryCsv = $summaryCsv
    GamPath    = $resolvedGamPath        # always the fully-resolved absolute path
    KeyFile    = $resolvedKeyFile        # '' if not found; helper will warn and skip REST pass
}

# Switch parameters must be passed as [switch] values in hashtable splatting
if (-not $ScanSpaces)    { $scriptArgs['SkipTenantSpaceScan'] = $true }
if ($SkipDocCommentScan) { $scriptArgs['SkipDocCommentScan']  = $true }
if ($IncludeCompleted)   { $scriptArgs['IncludeCompleted']    = $true }
if ($IncludeHidden)      { $scriptArgs['IncludeHidden']       = $true }
if ($IncludeDeleted)     { $scriptArgs['IncludeDeleted']      = $true }
if ($CheckpointCsv)      { $scriptArgs['CheckpointCsv']       = $CheckpointCsv }

# ── Run ──────────────────────────────────────────────────────────────────────
& $mainScript @scriptArgs

Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Green
Write-Host ("  Detail CSV  : {0}" -f $outputCsv)   -ForegroundColor Green
Write-Host ("  Summary CSV : {0}" -f $summaryCsv)  -ForegroundColor Green
