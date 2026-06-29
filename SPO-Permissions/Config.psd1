# =============================================================================
# Config.psd1 - Centralised configuration for GSites -> SPO Permission Migration
# =============================================================================
# INSTRUCTIONS:
#   Replace every placeholder value (marked with <...>) before running any script.
#   Never commit this file with real credentials to source control.
# =============================================================================
@{
    # -------------------------------------------------------------------------
    # Google Workspace / Google Sites settings
    # -------------------------------------------------------------------------
    Google = @{
        # Path to your Google Service Account JSON key file (downloaded from GCP console)
        ServiceAccountKeyPath = ".\service-account-key.json"

        # The Google Workspace super-admin email used for domain-wide delegation
        AdminEmail            = "admin@<your-google-domain>.com"

        # Your Google Workspace primary domain (e.g. contoso.com)
        Domain                = "<your-google-domain>.com"

        # OAuth2 scopes required by the script
        Scopes = @(
            "https://www.googleapis.com/auth/drive.readonly",
            "https://www.googleapis.com/auth/admin.directory.user.readonly",
            "https://www.googleapis.com/auth/admin.directory.group.readonly"
        )
    }

    # -------------------------------------------------------------------------
    # SharePoint Online / Microsoft 365 settings
    # -------------------------------------------------------------------------
    SharePoint = @{
        # Root tenant URL  e.g. https://contoso.sharepoint.com
        TenantUrl   = "https://<tenant>.sharepoint.com"

        # SharePoint Admin Center URL  e.g. https://contoso-admin.sharepoint.com
        AdminUrl    = "https://<tenant>-admin.sharepoint.com"

        # Azure AD App Registration Client ID (must have Sites.FullControl.All)
        ClientId    = "<app-client-id>"

        # PFX certificate used for app-only auth (recommended over client secrets)
        CertificatePath     = ".\spo-app-cert.pfx"
        CertificatePassword = "<certificate-password>"

        # Microsoft 365 tenant ID (found in Azure AD -> Overview)
        TenantId = "<tenant-id>"
    }

    # -------------------------------------------------------------------------
    # Output file settings
    # -------------------------------------------------------------------------
    Output = @{
        Directory             = ".\Output"
        GSitePermissionsFile  = "GSite_Permissions"
        SPOPermissionsFile    = "SPO_Permissions"
        DifferencesFile       = "Permission_Differences"
        FixLogFile            = "Fix_Log"
    }

    # -------------------------------------------------------------------------
    # Input sites file - drives BOTH which sites to scan AND the GSite->SPO mapping.
    # Columns: SiteId, SiteUrl, SiteName, SPOSiteUrl, Notes
    #   SiteId   - Google Drive file ID (fastest lookup - preferred)
    #   SiteUrl  - Google Site URL (ID auto-extracted if present in URL)
    #   SiteName - Display name (Drive search used as fallback)
    #   SPOSiteUrl - Corresponding SharePoint Online site URL
    # At least one of SiteId / SiteUrl / SiteName must be filled per row.
    # -------------------------------------------------------------------------
    InputSitesFile = ".\InputSites.csv"

    # -------------------------------------------------------------------------
    # Permission level mapping  (Google role -> SPO permission level name)
    # Adjust the SPO side to match your tenant's custom permission levels if needed.
    # -------------------------------------------------------------------------
    PermissionMapping = @{
        "owner"         = "Full Control"
        "organizer"     = "Full Control"
        "fileOrganizer" = "Design"
        "writer"        = "Contribute"
        "commenter"     = "Read"
        "reader"        = "Read"
    }

    # -------------------------------------------------------------------------
    # Fix behaviour
    # -------------------------------------------------------------------------
    Fix = @{
        # Set to $true to also REMOVE permissions in SPO that do not exist in GSites.
        # Recommended: start with $false and review the differences CSV first.
        RemoveExtraPermissions = $false

        # Simulate all fix actions without applying changes (dry-run mode).
        WhatIf = $true
    }
}
