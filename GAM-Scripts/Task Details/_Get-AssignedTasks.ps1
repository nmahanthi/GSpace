<#
.SYNOPSIS
  PS 7 helper: mint user-impersonating access tokens from the GAM
  service-account key and call the Google Tasks API with showAssigned=true
  to retrieve Docs/Chat-assigned tasks that the GAM CLI cannot return.

.DESCRIPTION
  Modes:
    -User <email>       Single-user JSON array on stdout (legacy).
    -UsersFile <path>   One email per line; emits NDJSON, one line per user:
                        { "user": "<email>", "entries": [ {tasklistId,tasklistTitle,task}, ... ] }
    -Users a,b,c        Inline bulk list; same NDJSON output as -UsersFile.
  Bulk modes run -Parallel threads concurrently (default 8).

.NOTES
  Requires PowerShell 7+ (RSA.ImportFromPem + ForEach-Object -Parallel).
  The service account must have domain-wide delegation for scope
  'https://www.googleapis.com/auth/tasks.readonly'.
#>
[CmdletBinding(DefaultParameterSetName = 'Single')]
param(
    [Parameter(ParameterSetName = 'Single', Mandatory)][string]$User,
    [Parameter(ParameterSetName = 'Bulk', Mandatory)][string]$UsersFile,
    [Parameter(ParameterSetName = 'BulkInline', Mandatory)][string[]]$Users,
    [string]$KeyFile = (Join-Path $env:USERPROFILE '.gam\oauth2service.json'),
    [int]$Parallel = 8,
    [switch]$IncludeCompleted,
    [switch]$IncludeHidden,
    [switch]$IncludeDeleted
)

$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -lt 7) {
    throw "This helper requires PowerShell 7+ (current: $($PSVersionTable.PSVersion)). Invoke via 'pwsh -File'."
}
if (-not (Test-Path $KeyFile)) { throw "Service account key not found: $KeyFile" }

$key = Get-Content -Raw -LiteralPath $KeyFile | ConvertFrom-Json

# One self-contained per-user worker; invoked either directly (Single mode)
# or from inside ForEach-Object -Parallel (Bulk modes).
$perUserBlock = {
    param($Subject, $Key, $IncludeCompleted, $IncludeHidden, $IncludeDeleted)

    function ConvertTo-Base64Url([byte[]]$b) {
        [Convert]::ToBase64String($b).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }

    try {
        $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        $header = @{ alg = 'RS256'; typ = 'JWT'; kid = $Key.private_key_id } | ConvertTo-Json -Compress
        $claims = [ordered]@{
            iss   = $Key.client_email
            sub   = $Subject
            scope = 'https://www.googleapis.com/auth/tasks.readonly'
            aud   = 'https://oauth2.googleapis.com/token'
            iat   = $now
            exp   = $now + 3600
        } | ConvertTo-Json -Compress
        $h = ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($header))
        $c = ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($claims))
        $signInput = "$h.$c"
        $rsa = [System.Security.Cryptography.RSA]::Create()
        $rsa.ImportFromPem($Key.private_key)
        $sig = $rsa.SignData(
            [Text.Encoding]::UTF8.GetBytes($signInput),
            [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $jwt = "$signInput." + (ConvertTo-Base64Url $sig)

        $token = (Invoke-RestMethod -Method Post `
                -Uri 'https://oauth2.googleapis.com/token' `
                -ContentType 'application/x-www-form-urlencoded' `
                -Body @{ grant_type = 'urn:ietf:params:oauth:grant-type:jwt-bearer'; assertion = $jwt }).access_token
        $hdr = @{ Authorization = "Bearer $token" }

        $tasklists = @()
        $next = $null
        do {
            $u = if ($next) { "https://tasks.googleapis.com/tasks/v1/users/@me/lists?maxResults=100&pageToken=$next" }
            else { "https://tasks.googleapis.com/tasks/v1/users/@me/lists?maxResults=100" }
            $r = Invoke-RestMethod -Headers $hdr -Uri $u
            if ($r.items) { $tasklists += $r.items }
            $next = $r.nextPageToken
        } while ($next)

        $entries = [System.Collections.Generic.List[object]]::new()
        foreach ($tl in $tasklists) {
            $qs = 'showAssigned=true&maxResults=100'
            if ($IncludeCompleted) { $qs += '&showCompleted=true' }
            if ($IncludeHidden) { $qs += '&showHidden=true' }
            if ($IncludeDeleted) { $qs += '&showDeleted=true' }
            $base = "https://tasks.googleapis.com/tasks/v1/lists/$($tl.id)/tasks?$qs"
            $n = $null
            do {
                $u2 = if ($n) { "$base&pageToken=$n" } else { $base }
                try { $r = Invoke-RestMethod -Headers $hdr -Uri $u2 }
                catch {
                    $sc = $_.Exception.Response.StatusCode.value__
                    [Console]::Error.WriteLine("tasks.list failed for $Subject list $($tl.id) (HTTP $sc): $($_.Exception.Message)")
                    break
                }
                foreach ($t in $r.items) {
                    if ($t.assignmentInfo) {
                        $entries.Add([pscustomobject]@{
                                tasklistId    = $tl.id
                                tasklistTitle = $tl.title
                                task          = $t
                            })
                    }
                }
                $n = $r.nextPageToken
            } while ($n)
        }
        return , $entries
    }
    catch {
        [Console]::Error.WriteLine("helper failed for $Subject`: $($_.Exception.Message)")
        return , @()
    }
}

function Invoke-BulkHelper {
    param([string[]]$UserList)
    if (-not $UserList -or $UserList.Count -eq 0) { return }
    # ForEach-Object -Parallel rejects scriptblocks passed via $using:, so we
    # ship the worker as a string and rehydrate it inside the parallel scope.
    $blockText = $perUserBlock.ToString()
    $keyObj = $key
    $ic = $IncludeCompleted.IsPresent
    $ih = $IncludeHidden.IsPresent
    $id = $IncludeDeleted.IsPresent
    $UserList | ForEach-Object -Parallel {
        $u = $_
        $worker = [scriptblock]::Create($using:blockText)
        $entries = & $worker -Subject $u -Key $using:keyObj `
            -IncludeCompleted:$using:ic -IncludeHidden:$using:ih -IncludeDeleted:$using:id
        [pscustomobject]@{ user = $u; entries = $entries } | ConvertTo-Json -Depth 12 -Compress
    } -ThrottleLimit $Parallel
}

switch ($PSCmdlet.ParameterSetName) {
    'Single' {
        $entries = & $perUserBlock -Subject $User -Key $key `
            -IncludeCompleted:$IncludeCompleted -IncludeHidden:$IncludeHidden -IncludeDeleted:$IncludeDeleted
        $entries | ConvertTo-Json -Depth 12 -Compress
    }
    'Bulk' {
        if (-not (Test-Path $UsersFile)) { throw "UsersFile not found: $UsersFile" }
        $list = Get-Content -LiteralPath $UsersFile | Where-Object { $_ -match '@' } | ForEach-Object { $_.Trim() }
        Invoke-BulkHelper -UserList $list
    }
    'BulkInline' { Invoke-BulkHelper -UserList $Users }
}
