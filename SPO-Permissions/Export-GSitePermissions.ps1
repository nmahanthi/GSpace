<#
.SYNOPSIS
    Exports Google Sites permissions to a CSV file for a selected list of sites.
.DESCRIPTION
    Authenticates with Google using a Service Account (domain-wide delegation),
    then reads a user-supplied InputSites.csv to process ONLY the specified sites
    (instead of scanning the entire tenant of 100k+ sites).

    Each row in InputSites.csv must have at least ONE of:
      SiteId   - Google Drive file ID (fastest, most reliable)
      SiteUrl  - Google Site URL (file ID extracted automatically if embedded)
      SiteName - Display name (Drive name search used as fallback)

    Resolution priority per row:  SiteId > SiteUrl > SiteName

    The same InputSites.csv is also used by Compare-Permissions.ps1 for
    the GSiteUrl -> SPOSiteUrl mapping (SPOSiteUrl column).

.PARAMETER ConfigPath
    Path to Config.psd1.  Defaults to .\Config.psd1
.PARAMETER InputCsvFile
    Path to InputSites.csv listing the sites to scan.
    If omitted, falls back to Config.psd1 InputSitesFile value.
    If that is also missing, ALL sites in the tenant are scanned (slow for large tenants).
.PARAMETER OutputFile
    Overrides the output CSV path from Config.psd1.
.EXAMPLE
    # Scan only the 100 sites listed in InputSites.csv
    .\Export-GSitePermissions.ps1 -InputCsvFile ".\InputSites.csv"

    # Override config and output location
    .\Export-GSitePermissions.ps1 -InputCsvFile ".\MySites.csv" -OutputFile ".\Output\GSite_Perms.csv"
#>
[CmdletBinding()]
param(
    [string]$ConfigPath   = ".\Config.psd1",
    [string]$InputCsvFile = "",
    [string]$OutputFile   = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Helper: Build a Google OAuth2 access token from a Service Account key file
# Uses RSA-SHA256 JWT signed with the service account private key.
# ---------------------------------------------------------------------------
function Get-GoogleAccessToken {
    param(
        [string]$KeyFilePath,
        [string]$ImpersonateEmail,   # admin email for domain-wide delegation
        [string[]]$Scopes
    )

    $key = Get-Content $KeyFilePath -Raw | ConvertFrom-Json

    $now  = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $exp  = $now + 3600

    $header = [Convert]::ToBase64String(
        [Text.Encoding]::UTF8.GetBytes('{"alg":"RS256","typ":"JWT"}')
    ).TrimEnd('=').Replace('+','-').Replace('/','_')

    $claimSet = @{
        iss   = $key.client_email
        sub   = $ImpersonateEmail
        scope = ($Scopes -join ' ')
        aud   = "https://oauth2.googleapis.com/token"
        iat   = $now
        exp   = $exp
    } | ConvertTo-Json -Compress

    $claim = [Convert]::ToBase64String(
        [Text.Encoding]::UTF8.GetBytes($claimSet)
    ).TrimEnd('=').Replace('+','-').Replace('/','_')

    $toSign = "$header.$claim"

    # Import RSA private key (PKCS#8)
    $pkText  = $key.private_key -replace "-----BEGIN PRIVATE KEY-----","" `
                                -replace "-----END PRIVATE KEY-----",""  `
                                -replace "\n",""
    $pkBytes = [Convert]::FromBase64String($pkText)
    $rsa     = [System.Security.Cryptography.RSA]::Create()
    $rsa.ImportPkcs8PrivateKey($pkBytes, [ref]$null) | Out-Null

    $sigBytes = $rsa.SignData(
        [Text.Encoding]::UTF8.GetBytes($toSign),
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
    )
    $sig = [Convert]::ToBase64String($sigBytes).TrimEnd('=').Replace('+','-').Replace('/','_')

    $jwt = "$toSign.$sig"

    $response = Invoke-RestMethod -Method Post `
        -Uri "https://oauth2.googleapis.com/token" `
        -ContentType "application/x-www-form-urlencoded" `
        -Body "grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=$jwt"

    return $response.access_token
}

# ---------------------------------------------------------------------------
# Helper: Paginated Drive API call
# ---------------------------------------------------------------------------
function Invoke-DriveApiPagedGet {
    param([string]$Uri, [string]$Token, [string]$ItemsKey = "files")
    $results = [System.Collections.Generic.List[object]]::new()
    $nextPage = $null
    do {
        $url = if ($nextPage) { "$Uri&pageToken=$nextPage" } else { $Uri }
        $resp = Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $Token" }
        if ($resp.$ItemsKey) { $results.AddRange($resp.$ItemsKey) }
        $nextPage = $resp.nextPageToken
    } while ($nextPage)
    return $results
}

# ---------------------------------------------------------------------------
# Helper: Extract Drive file ID from a Google Sites URL
#   Handles formats:
#     https://sites.google.com/d/<fileId>/s/<name>   (New Sites - ID in URL)
#     https://sites.google.com/d/<fileId>/edit
#     https://sites.google.com/a/<domain>/<name>     (Classic/vanity - no ID)
#     https://sites.google.com/view/<name>            (Public view - no ID)
# ---------------------------------------------------------------------------
function Get-FileIdFromUrl {
    param([string]$Url)
    if ($Url -match 'sites\.google\.com/d/([a-zA-Z0-9_-]+)') {
        return $Matches[1]
    }
    return $null
}

# ---------------------------------------------------------------------------
# Helper: Resolve a single CSV row to a Drive file object.
#   Priority: SiteId > SiteUrl (ID extraction) > SiteUrl (name search) > SiteName
# ---------------------------------------------------------------------------
function Resolve-DriveFile {
    param(
        [string]$SiteId,
        [string]$SiteUrl,
        [string]$SiteName,
        [string]$Token
    )

    $fields = "id,name,webViewLink,owners"

    # --- 1. Direct file ID ---
    if ($SiteId) {
        try {
            $resp = Invoke-RestMethod `
                -Uri "https://www.googleapis.com/drive/v3/files/$($SiteId)?fields=$fields&supportsAllDrives=true" `
                -Headers @{ Authorization = "Bearer $Token" }
            return $resp
        } catch {
            Write-Warning "    SiteId '$SiteId' not found in Drive: $_"
            return $null
        }
    }

    # --- 2. Extract file ID embedded in URL ---
    if ($SiteUrl) {
        $idFromUrl = Get-FileIdFromUrl -Url $SiteUrl
        if ($idFromUrl) {
            try {
                $resp = Invoke-RestMethod `
                    -Uri "https://www.googleapis.com/drive/v3/files/$($idFromUrl)?fields=$fields&supportsAllDrives=true" `
                    -Headers @{ Authorization = "Bearer $Token" }
                return $resp
            } catch {
                Write-Warning "    ID extracted from URL ($idFromUrl) not found: $_"
            }
        }

        # --- 3. No ID in URL: derive name from URL path and search ---
        $derivedName = ($SiteUrl.TrimEnd('/') -split '/')[-1]
        if ($derivedName) { $SiteName = $derivedName }
    }

    # --- 4. Search Drive by display name ---
    if ($SiteName) {
        $q   = [Uri]::EscapeDataString("name='$SiteName' and mimeType='application/vnd.google-apps.site' and trashed=false")
        $uri = "https://www.googleapis.com/drive/v3/files?q=$q&fields=files($fields)&pageSize=10&includeItemsFromAllDrives=true&supportsAllDrives=true"
        try {
            $resp = Invoke-RestMethod -Uri $uri -Headers @{ Authorization = "Bearer $Token" }
            if ($resp.files -and $resp.files.Count -gt 0) {
                if ($resp.files.Count -gt 1) {
                    Write-Warning "    Multiple sites match name '$SiteName' - using the first result. Provide SiteId for precision."
                }
                return $resp.files[0]
            } else {
                Write-Warning "    No Drive file found with name '$SiteName'."
            }
        } catch {
            Write-Warning "    Drive name search failed for '$SiteName': $_"
        }
    }

    return $null
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host "`n=== Export-GSitePermissions.ps1 ===" -ForegroundColor Cyan

$cfg = Import-PowerShellDataFile -Path $ConfigPath
$date = Get-Date -Format "yyyyMMdd_HHmmss"

if (-not $OutputFile) {
    $dir = $cfg.Output.Directory
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $OutputFile = Join-Path $dir "$($cfg.Output.GSitePermissionsFile)_$date.csv"
}

Write-Host "Authenticating with Google..." -ForegroundColor Yellow
$token = Get-GoogleAccessToken `
    -KeyFilePath       $cfg.Google.ServiceAccountKeyPath `
    -ImpersonateEmail  $cfg.Google.AdminEmail `
    -Scopes            $cfg.Google.Scopes

# ---------------------------------------------------------------------------
# Determine which sites to process
# ---------------------------------------------------------------------------

# Resolve input CSV path: parameter > config > full scan
if (-not $InputCsvFile -and $cfg.ContainsKey('InputSitesFile')) {
    $InputCsvFile = $cfg.InputSitesFile
}

$sites       = [System.Collections.Generic.List[object]]::new()
$inputRows   = @()   # raw CSV rows kept for SPOSiteUrl lookup later

if ($InputCsvFile -and (Test-Path $InputCsvFile)) {
    $inputRows = Import-Csv $InputCsvFile
    Write-Host "Input CSV loaded: $InputCsvFile  ($($inputRows.Count) rows)" -ForegroundColor Yellow
    Write-Host "Resolving sites from CSV via Drive API..." -ForegroundColor Yellow

    $rowNum = 0
    foreach ($row in $inputRows) {
        $rowNum++
        $id   = if ($row.PSObject.Properties['SiteId'])   { $row.SiteId.Trim()   } else { "" }
        $url  = if ($row.PSObject.Properties['SiteUrl'])  { $row.SiteUrl.Trim()  } else { "" }
        $name = if ($row.PSObject.Properties['SiteName']) { $row.SiteName.Trim() } else { "" }

        if (-not $id -and -not $url -and -not $name) {
            Write-Warning "  Row $rowNum : SiteId, SiteUrl, and SiteName are all empty - skipping."
            continue
        }

        Write-Host "  [$rowNum/$($inputRows.Count)] Resolving: $(if($id){$id} elseif($url){$url} else{$name})" -ForegroundColor DarkCyan

        $driveFile = Resolve-DriveFile -SiteId $id -SiteUrl $url -SiteName $name -Token $token
        if ($driveFile) {
            # Attach the original SPOSiteUrl so it stays available when building output
            $driveFile | Add-Member -NotePropertyName InputSPOSiteUrl `
                -NotePropertyValue (if ($row.PSObject.Properties['SPOSiteUrl']) { $row.SPOSiteUrl } else { "" }) `
                -Force
            $sites.Add($driveFile)
        } else {
            Write-Warning "  Row $rowNum : Could not resolve site - it will be skipped."
        }
    }
    Write-Host "  Resolved $($sites.Count) of $($inputRows.Count) site(s)." -ForegroundColor Green

} else {
    # No CSV supplied - fall back to full-tenant scan (use with caution on large tenants)
    Write-Warning "No InputSites.csv provided. Scanning ALL sites in the tenant (may be slow)."
    $query    = [Uri]::EscapeDataString("mimeType='application/vnd.google-apps.site' and trashed=false")
    $fields   = "nextPageToken,files(id,name,webViewLink,owners)"
    $driveUri = "https://www.googleapis.com/drive/v3/files?q=$query&fields=$fields&pageSize=100&includeItemsFromAllDrives=true&supportsAllDrives=true"
    $allSites = Invoke-DriveApiPagedGet -Uri $driveUri -Token $token
    $allSites | ForEach-Object { $sites.Add($_) }
    Write-Host "  Found $($sites.Count) Google Site(s) in tenant." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Fetch permissions for each resolved site
# ---------------------------------------------------------------------------
$records = [System.Collections.Generic.List[object]]::new()
$i = 0

foreach ($site in $sites) {
    $i++
    Write-Host "  [$i/$($sites.Count)] Fetching permissions: $($site.name)" -ForegroundColor DarkCyan

    $permUri = "https://www.googleapis.com/drive/v3/files/$($site.id)/permissions" +
               "?fields=permissions(id,emailAddress,role,type,displayName,domain)&pageSize=100"

    try {
        $perms = Invoke-DriveApiPagedGet -Uri $permUri -Token $token -ItemsKey "permissions"
    } catch {
        Write-Warning "    Could not retrieve permissions for '$($site.name)': $_"
        continue
    }

    $spoUrl = if ($site.PSObject.Properties['InputSPOSiteUrl']) { $site.InputSPOSiteUrl } else { "" }

    foreach ($p in $perms) {
        $records.Add([PSCustomObject]@{
            SiteId          = $site.id
            SiteName        = $site.name
            SiteUrl         = $site.webViewLink
            SPOSiteUrl      = $spoUrl           # carried from InputSites.csv for downstream use
            PermissionId    = $p.id
            PrincipalType   = $p.type           # user | group | domain | anyone
            PrincipalEmail  = $p.emailAddress
            PrincipalName   = $p.displayName
            Domain          = $p.domain
            GoogleRole      = $p.role           # owner|organizer|fileOrganizer|writer|commenter|reader
            CapturedAt      = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        })
    }
}

$records | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
Write-Host "`nGoogle Sites permissions exported -> $OutputFile" -ForegroundColor Green
Write-Host "Total permission entries : $($records.Count)"
Write-Host "Sites processed          : $($sites.Count)`n"
return $OutputFile
