param(
    [string]$AccessToken,
    [string]$OutputDir = "$PSScriptRoot\output",
    [int]$ThrottleLimit = 10
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

if ([string]::IsNullOrWhiteSpace($AccessToken)) {
    $AccessToken = $env:GCP_ACCESS_TOKEN
}
# Trim \r\n that gcloud or PowerShell Receive-Job appends - prevents HTTP 401
$AccessToken = ($AccessToken ?? '').Trim()
if ([string]::IsNullOrWhiteSpace($AccessToken)) {
    throw 'Provide -AccessToken or set GCP_ACCESS_TOKEN.'
}

$embedsFile = Join-Path $OutputDir 'Embeds.csv'

$embedRows = @()
if (Test-Path $embedsFile) { $embedRows = Import-Csv $embedsFile }

$linkedForms = $embedRows | Where-Object { $_.ArtifactType -eq 'Form' } |
Group-Object ArtifactUrl | ForEach-Object { $_.Group[0] }

# Sheets and Scripts enrichment skipped - the Apps Script and Sheets APIs return 403
# for this account. Empty output files are written to keep downstream steps intact.
Write-Host "  [SKIP] Sheets enrichment (account lacks Sheets API access - 403)"
@() | Export-Csv -NoTypeInformation -Path (Join-Path $OutputDir 'Sheets_Enrichment.csv')

Write-Host "  Enriching $(@($linkedForms).Count) forms (parallel, throttle=$ThrottleLimit)..."
$formBag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
$linkedForms | ForEach-Object -Parallel {
    $row = $_
    $token = $using:AccessToken
    $outBag = $using:formBag
    $hdrs = @{ Authorization = "Bearer $token" }

    if ($row.ArtifactUrl -notmatch '/forms/d/([a-zA-Z0-9-_]+)') { return }
    $formId = $Matches[1]
    $uri = "https://forms.googleapis.com/v1/forms/$formId"

    $resp = $null; $delay = 1
    for ($a = 0; $a -lt 5; $a++) {
        try { $resp = Invoke-RestMethod -Method Get -Uri $uri -Headers $hdrs -ContentType 'application/json' -ErrorAction Stop; break }
        catch {
            $sc = $null; try { $sc = [int]$_.Exception.Response.StatusCode } catch {}
            if ($sc -eq 429 -or $sc -eq 503) { Start-Sleep -Seconds $delay; $delay = [Math]::Min($delay * 2, 32); continue }
            Write-Warning "Forms $formId`: $($_.Exception.Message)"; break
        }
    }
    if ($null -eq $resp) { return }

    $items = @($resp.items)
    $qCount = 0; $sCount = 0
    foreach ($item in $items) {
        if ($null -ne $item.questionItem) { $qCount++ }
        if ($null -ne $item.pageBreakItem) { $sCount++ }
    }
    $null = $outBag.Add([pscustomobject]@{
            SiteId           = $row.SiteId
            PageUrl          = $row.PageUrl
            FormId           = $resp.formId
            FormTitle        = $resp.info.title
            ItemCount        = $items.Count
            QuestionCount    = $qCount
            SectionCount     = $sCount
            LinkedSheetId    = $resp.linkedSheetId
            RevisionId       = $resp.revisionId
            ComplexityPoints = (($qCount * 2) + ($sCount * 3))
        })
} -ThrottleLimit $ThrottleLimit
if ($formBag.Count -gt 0) { $formBag | Export-Csv -NoTypeInformation -Path (Join-Path $OutputDir 'Forms_Enrichment.csv') }
else { @() | Export-Csv -NoTypeInformation -Path (Join-Path $OutputDir 'Forms_Enrichment.csv') }
Write-Host "  Forms done: $($formBag.Count) enriched"

Write-Host "  [SKIP] Scripts enrichment (account lacks Apps Script API access - 403)"
@() | Export-Csv -NoTypeInformation -Path (Join-Path $OutputDir 'Scripts_Enrichment.csv')

Write-Host 'Artifact enrichment completed.'
