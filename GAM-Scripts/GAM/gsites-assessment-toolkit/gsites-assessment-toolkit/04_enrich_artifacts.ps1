param(
    [string]$AccessToken,
    [string]$OutputDir = "$PSScriptRoot\output"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($AccessToken)) {
    $AccessToken = $env:GCP_ACCESS_TOKEN
}
if ([string]::IsNullOrWhiteSpace($AccessToken)) {
    throw 'Provide -AccessToken or set GCP_ACCESS_TOKEN.'
}

$headers = @{ Authorization = "Bearer $AccessToken" }

function Invoke-GoogleJson {
    param([string]$Uri)
    try {
        Start-Sleep -Milliseconds 150
        return Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers -ContentType 'application/json'
    }
    catch {
        Write-Warning "Failed: $Uri :: $($_.Exception.Message)"
        return $null
    }
}

function Get-SheetIdFromUrl {
    param([string]$Url)
    if ($Url -match '/spreadsheets/d/([a-zA-Z0-9-_]+)') { return $matches[1] }
    return $null
}

function Get-FormIdFromUrl {
    param([string]$Url)
    if ($Url -match '/forms/d/([a-zA-Z0-9-_]+)') { return $matches[1] }
    return $null
}

function Count-Functions {
    param([string]$Code)
    if ([string]::IsNullOrWhiteSpace($Code)) { return 0 }
    return ([regex]::Matches($Code, '(?m)^\s*function\s+[A-Za-z0-9_]+\s*\(')).Count
}

$embedsFile = Join-Path $OutputDir '08_Embeds.csv'
$sheetsFile = Join-Path $OutputDir '04_Candidate_Sheets.csv'
$formsFile = Join-Path $OutputDir '05_Candidate_Forms.csv'
$scriptsFile = Join-Path $OutputDir '06_Candidate_Scripts.csv'

$embedRows = @()
if (Test-Path $embedsFile) { $embedRows = Import-Csv $embedsFile }

$linkedSheets = $embedRows | Where-Object { $_.ArtifactType -eq 'Sheet' } |
Group-Object ArtifactUrl | ForEach-Object { $_.Group[0] }
$linkedForms = $embedRows | Where-Object { $_.ArtifactType -eq 'Form' } |
Group-Object ArtifactUrl | ForEach-Object { $_.Group[0] }

$sheetOut = New-Object System.Collections.Generic.List[object]
foreach ($row in $linkedSheets) {
    $sheetId = Get-SheetIdFromUrl $row.ArtifactUrl
    if (-not $sheetId) { continue }
    $uri = "https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=spreadsheetId,properties.title,sheets.properties,namedRanges,developerMetadata"
    $resp = Invoke-GoogleJson -Uri $uri
    if ($null -eq $resp) { continue }

    $sheetOut.Add([pscustomobject]@{
            SiteId                 = $row.SiteId
            PageUrl                = $row.PageUrl
            SpreadsheetId          = $resp.spreadsheetId
            SpreadsheetTitle       = $resp.properties.title
            SheetTabCount          = @($resp.sheets).Count
            NamedRangeCount        = @($resp.namedRanges).Count
            DeveloperMetadataCount = @($resp.developerMetadata).Count
            ComplexityPoints       = ((@($resp.sheets).Count * 2) + (@($resp.namedRanges).Count * 3) + (@($resp.developerMetadata).Count * 2))
        })
}
$sheetOut | Export-Csv -NoTypeInformation -Path (Join-Path $OutputDir '11_Sheets_Enrichment.csv')

$formOut = New-Object System.Collections.Generic.List[object]
foreach ($row in $linkedForms) {
    $formId = Get-FormIdFromUrl $row.ArtifactUrl
    if (-not $formId) { continue }
    $uri = "https://forms.googleapis.com/v1/forms/$formId"
    $resp = Invoke-GoogleJson -Uri $uri
    if ($null -eq $resp) { continue }

    $items = @($resp.items)
    $questionCount = 0
    $sectionCount = 0
    foreach ($item in $items) {
        if ($null -ne $item.questionItem) { $questionCount += 1 }
        if ($null -ne $item.pageBreakItem) { $sectionCount += 1 }
    }

    $formOut.Add([pscustomobject]@{
            SiteId           = $row.SiteId
            PageUrl          = $row.PageUrl
            FormId           = $resp.formId
            FormTitle        = $resp.info.title
            ItemCount        = $items.Count
            QuestionCount    = $questionCount
            SectionCount     = $sectionCount
            LinkedSheetId    = $resp.linkedSheetId
            RevisionId       = $resp.revisionId
            ComplexityPoints = (($questionCount * 2) + ($sectionCount * 3))
        })
}
$formOut | Export-Csv -NoTypeInformation -Path (Join-Path $OutputDir '12_Forms_Enrichment.csv')

$scriptOut = New-Object System.Collections.Generic.List[object]
if (Test-Path $scriptsFile) {
    $candidateScripts = Import-Csv $scriptsFile
    foreach ($row in $candidateScripts) {
        $scriptId = $row.id
        if ([string]::IsNullOrWhiteSpace($scriptId)) { continue }

        $meta = Invoke-GoogleJson -Uri "https://script.googleapis.com/v1/projects/$scriptId"
        $content = Invoke-GoogleJson -Uri "https://script.googleapis.com/v1/projects/$scriptId/content"
        if ($null -eq $content) { continue }

        $allFiles = @($content.files)
        $combinedCode = ($allFiles | ForEach-Object { $_.source } | Out-String)
        $functionCount = Count-Functions -Code $combinedCode
        $scopeCount = 0
        $manifestSource = ($allFiles | Where-Object { $_.name -eq 'appsscript' -or $_.type -eq 'JSON' } | Select-Object -First 1).source
        if ($manifestSource -and $manifestSource -match 'oauthScopes') {
            $scopeCount = ([regex]::Matches($manifestSource, 'https://')).Count
        }

        $scriptOut.Add([pscustomobject]@{
                ScriptId         = $scriptId
                ScriptTitle      = $meta.title
                CreateTime       = $meta.createTime
                UpdateTime       = $meta.updateTime
                ParentId         = $meta.parentId
                FileCount        = $allFiles.Count
                FunctionCount    = $functionCount
                ScopeCount       = $scopeCount
                CodeSizeChars    = ($combinedCode.Length)
                ComplexityPoints = (($allFiles.Count * 3) + ($functionCount * 2) + ($scopeCount * 2))
            })
    }
}
$scriptOut | Export-Csv -NoTypeInformation -Path (Join-Path $OutputDir '13_Scripts_Enrichment.csv')

Write-Host 'Artifact enrichment completed.'
