# =============================================================================
# Clone-SPOSite.ps1
# Captures a SharePoint Online site template and deploys it to 10 new sites
# Requires: PnP.PowerShell module
# Install:  Install-Module PnP.PowerShell -Scope CurrentUser
# =============================================================================

#region ——— CONFIGURATION — Edit these values ———————————————————————————————

$SourceSiteUrl  = "https://yourtenant.sharepoint.com/sites/sourcesite"
$AdminUrl       = "https://yourtenant-admin.sharepoint.com"
$TenantUrl      = "https://yourtenant.sharepoint.com"

# Template output path
$TemplatePath   = "C:\SPOTemplates\SiteTemplate.xml"
$TemplateFolder = Split-Path $TemplatePath -Parent

# Define 10 target sites — customise Alias and Title as needed
$TargetSites = @(
    @{ Alias = "clonedsite1";  Title = "Cloned Site 1"  },
    @{ Alias = "clonedsite2";  Title = "Cloned Site 2"  },
    @{ Alias = "clonedsite3";  Title = "Cloned Site 3"  },
    @{ Alias = "clonedsite4";  Title = "Cloned Site 4"  },
    @{ Alias = "clonedsite5";  Title = "Cloned Site 5"  },
    @{ Alias = "clonedsite6";  Title = "Cloned Site 6"  },
    @{ Alias = "clonedsite7";  Title = "Cloned Site 7"  },
    @{ Alias = "clonedsite8";  Title = "Cloned Site 8"  },
    @{ Alias = "clonedsite9";  Title = "Cloned Site 9"  },
    @{ Alias = "clonedsite10"; Title = "Cloned Site 10" }
)

#endregion ——— END CONFIGURATION ————————————————————————————————————————————


#region ——— LOGGING SETUP ————————————————————————————————————————————————————

$LogFile = "C:\SPOTemplates\Clone-SPOSite_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","SUCCESS","WARNING","ERROR")] [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry     = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $entry

    switch ($Level) {
        "INFO"    { Write-Host $entry -ForegroundColor Cyan    }
        "SUCCESS" { Write-Host $entry -ForegroundColor Green   }
        "WARNING" { Write-Host $entry -ForegroundColor Yellow  }
        "ERROR"   { Write-Host $entry -ForegroundColor Red     }
    }
}

#endregion


#region ——— PREFLIGHT CHECKS —————————————————————————————————————————————————

Write-Log "Starting SPO Site Clone Script"

# Ensure PnP.PowerShell is installed
if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    Write-Log "PnP.PowerShell module not found. Installing..." "WARNING"
    try {
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
        Write-Log "PnP.PowerShell installed successfully." "SUCCESS"
    } catch {
        Write-Log "Failed to install PnP.PowerShell: $_" "ERROR"
        exit 1
    }
}

Import-Module PnP.PowerShell -ErrorAction Stop

# Ensure template output folder exists
if (-not (Test-Path $TemplateFolder)) {
    New-Item -ItemType Directory -Path $TemplateFolder -Force | Out-Null
    Write-Log "Created template folder: $TemplateFolder"
}

#endregion


#region ——— STEP 1: EXTRACT TEMPLATE FROM SOURCE SITE ————————————————————————

Write-Log "========== STEP 1: Extracting template from source site =========="
Write-Log "Source site: $SourceSiteUrl"

try {
    Connect-PnPOnline -Url $SourceSiteUrl -Interactive
    Write-Log "Connected to source site." "SUCCESS"
} catch {
    Write-Log "Failed to connect to source site: $_" "ERROR"
    exit 1
}

try {
    Write-Log "Extracting site template — this may take several minutes for large sites..."

    Get-PnPSiteTemplate -Out $TemplatePath `
        -Handlers All `
        -ExcludeHandlers TermGroups, SiteSecurity `
        -PersistBrandingFiles `
        -PersistPublishingFiles `
        -PersistMultiLanguageResources `
        -Force

    Write-Log "Template extracted successfully to: $TemplatePath" "SUCCESS"
} catch {
    Write-Log "Failed to extract site template: $_" "ERROR"
    exit 1
}

Disconnect-PnPOnline

#endregion


#region ——— STEP 2: CREATE 10 SITES AND APPLY TEMPLATE ———————————————————————

Write-Log "========== STEP 2: Creating and provisioning target sites =========="

try {
    Connect-PnPOnline -Url $AdminUrl -Interactive
    Write-Log "Connected to SharePoint Admin Center." "SUCCESS"
} catch {
    Write-Log "Failed to connect to Admin Center: $_" "ERROR"
    exit 1
}

# Track results for summary report
$Results = @()

foreach ($site in $TargetSites) {

    $siteUrl = "$TenantUrl/sites/$($site.Alias)"
    Write-Log "------------------------------------------------------------------"
    Write-Log "Processing: $($site.Title) — $siteUrl"

    #— Create the site ————————————————————————————————————————————————————————
    try {
        $existingsite = Get-PnPTenantSite -Url $siteUrl -ErrorAction SilentlyContinue

        if ($existingsite) {
            Write-Log "Site already exists, skipping creation: $siteUrl" "WARNING"
        } else {
            Write-Log "Creating site: $siteUrl"

            New-PnPSite -Type TeamSite `
                -Title $site.Title `
                -Alias $site.Alias `
                -Wait

            Write-Log "Site created: $siteUrl" "SUCCESS"
        }
    } catch {
        Write-Log "Failed to create site '$siteUrl': $_" "ERROR"
        $Results += [PSCustomObject]@{
            Site    = $siteUrl
            Created = $false
            Applied = $false
            Error   = $_.ToString()
        }
        # Reconnect to admin and move to next site
        Connect-PnPOnline -Url $AdminUrl -Interactive
        continue
    }

    #— Apply template to the site —————————————————————————————————————————————
    try {
        Write-Log "Connecting to new site and applying template..."

        Connect-PnPOnline -Url $siteUrl -Interactive

        Invoke-PnPSiteTemplate -Path $TemplatePath `
            -Parameters @{ "SourceSiteUrl" = $SourceSiteUrl } `
            -ClearNavigation

        Write-Log "Template applied successfully to: $siteUrl" "SUCCESS"

        $Results += [PSCustomObject]@{
            Site    = $siteUrl
            Created = $true
            Applied = $true
            Error   = ""
        }
    } catch {
        Write-Log "Failed to apply template to '$siteUrl': $_" "ERROR"
        $Results += [PSCustomObject]@{
            Site    = $siteUrl
            Created = $true
            Applied = $false
            Error   = $_.ToString()
        }
    } finally {
        Disconnect-PnPOnline
    }

    # Reconnect to admin for next iteration
    try {
        Connect-PnPOnline -Url $AdminUrl -Interactive
    } catch {
        Write-Log "Failed to reconnect to Admin Center — stopping loop." "ERROR"
        break
    }
}

Disconnect-PnPOnline

#endregion


#region ——— STEP 3: SUMMARY REPORT ——————————————————————————————————————————

Write-Log "========== SUMMARY REPORT =========="

$successCount = ($Results | Where-Object { $_.Applied -eq $true }).Count
$failCount    = ($Results | Where-Object { $_.Applied -eq $false }).Count

Write-Log "Total sites processed : $($Results.Count)"
Write-Log "Successfully applied  : $successCount" "SUCCESS"
Write-Log "Failed                : $failCount" $(if ($failCount -gt 0) { "ERROR" } else { "INFO" })

Write-Log ""
Write-Log "Site-by-site breakdown:"

foreach ($r in $Results) {
    $status = if ($r.Applied) { "SUCCESS" } else { "ERROR" }
    $msg    = if ($r.Applied) {
        "$($r.Site) — Template applied successfully"
    } else {
        "$($r.Site) — FAILED: $($r.Error)"
    }
    Write-Log $msg $status
}

Write-Log ""
Write-Log "Full log saved to: $LogFile"
Write-Log "Script completed."

#endregion
