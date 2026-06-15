#
# TaskScan.config.psd1 — Configuration for Invoke-TargetedTaskScan.ps1
# ======================================================================
# Edit this file ONCE and then just run .\Invoke-TargetedTaskScan.ps1
# with no extra arguments.
#
# Precedence (highest wins):
#   1. Command-line parameter  (-GamPath, -OutputDir, etc.)
#   2. Value set in this file
#   3. Built-in auto-detection / default
# ======================================================================
@{
    # ------------------------------------------------------------------
    # STEP 1 — GAM executable path
    # ------------------------------------------------------------------
    # Leave blank to let the script auto-detect GAM from PATH and common
    # install folders (C:\GAM7, C:\GAM, LOCALAPPDATA\GAM7, etc.).
    # Set this only if auto-detection does not find your installation.
    #
    # Examples:
    #   GamPath = 'C:\GAM7\gam.exe'
    #   GamPath = 'E:\GAM-Roche\gam.exe'
    # ------------------------------------------------------------------
    GamPath = ''

    # ------------------------------------------------------------------
    # STEP 1b — GAM service-account key file (oauth2service.json)
    # ------------------------------------------------------------------
    # Required for the REST-based assigned-task fetch (Docs/Chat tasks).
    # Leave blank to auto-detect: the script checks the GAM config folder,
    # the folder next to gam.exe, and the %GAMCFGDIR% environment variable.
    #
    # To find your key file run:  gam info domain
    # and look for the "Service Account" or "Config File" line.
    #
    # Examples:
    #   KeyFile = 'C:\Users\admin\.gam\oauth2service.json'
    #   KeyFile = 'E:\GAM-Roche\oauth2service.json'
    # ------------------------------------------------------------------
    KeyFile = ''

    # ------------------------------------------------------------------
    # STEP 2 — User list file
    # ------------------------------------------------------------------
    # Path to the plain-text file containing one primaryEmail per line.
    # Defaults to TargetUsers.txt in the same folder as the script.
    #
    # Example:
    #   UsersFile = 'C:\Scans\hr_team.txt'
    # ------------------------------------------------------------------
    UsersFile = ''

    # ------------------------------------------------------------------
    # STEP 3 — Output folder
    # ------------------------------------------------------------------
    # Where to write the result CSVs. Defaults to the script folder.
    #
    # Example:
    #   OutputDir = 'C:\Scans\Output'
    # ------------------------------------------------------------------
    OutputDir = ''

    # ------------------------------------------------------------------
    # Scan options  (set to $true to enable)
    # ------------------------------------------------------------------

    # Include tasks the user has already completed
    IncludeCompleted = $false

    # Include tasks marked as hidden
    IncludeHidden = $false

    # Include tasks that have been deleted
    IncludeDeleted = $false

    # Enable the tenant-wide Chat space scan.
    # Disabled by default — it scans ALL spaces visible to the first user
    # and is expensive for large tenants.
    ScanSpaces = $false

    # Skip the Drive doc-comment scan (enabled by default for targeted runs).
    SkipDocCommentScan = $false
}
