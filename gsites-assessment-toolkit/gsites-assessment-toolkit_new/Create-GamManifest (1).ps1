<#
.SYNOPSIS
Creates GAM Tool manifest.json files for exported CSV datasets.

.EXAMPLE
.\Create-GamManifest.ps1 `
  -RootPath "C:\exports\gam_exports" `
  -TenantName "rocheua.com" `
  -Environment "prod" `
  -RunType "assessment"

Expected folder structure:
gam_exports/{workload}/{run_id}/*.csv

Example:
gam_exports/contacts/RUN_CONTACTS_20260325_141733Z/Contacts-Delegates-Results.csv
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$RootPath,

    [Parameter(Mandatory = $true)]
    [string]$TenantName,

    [string]$Environment = "prod",

    [string]$RunType = "assessment",

    [string]$ExportSystem = "gam_tool",

    [string]$ExportVersion = "1.0",

    [string]$DeliveryFormat = "csv",

    [switch]$Overwrite
)

function Convert-ToSnakeCase {
    param([string]$Value)

    $name = [System.IO.Path]::GetFileNameWithoutExtension($Value)

    $name = $name `
        -replace '[^a-zA-Z0-9]+', '_' `
        -replace '_+', '_' `
        -replace '^_|_$', ''

    return $name.ToLower()
}

function Get-RecordCount {
    param([string]$CsvPath)

    if (-not (Test-Path $CsvPath)) {
        return 0
    }

    $lineCount = (Get-Content -Path $CsvPath | Measure-Object -Line).Lines

    if ($lineCount -le 1) {
        return 0
    }

    return ($lineCount - 1)
}

function Get-SnapshotInfoFromRunId {
    param(
        [string]$RunId,
        [string]$Workload
    )

    # Expected format:
    # RUN_CONTACTS_20260325_141733Z

    $pattern = "^RUN_([A-Z0-9]+)_(\d{8})_(\d{6})Z$"

    if ($RunId -match $pattern) {
        $workloadPart = $matches[1]
        $datePart = $matches[2]
        $timePart = $matches[3]

        $snapshotId = "$workloadPart`_$datePart`_$timePart`Z"

        $year = $datePart.Substring(0, 4)
        $month = $datePart.Substring(4, 2)
        $day = $datePart.Substring(6, 2)

        $hour = $timePart.Substring(0, 2)
        $minute = $timePart.Substring(2, 2)
        $second = $timePart.Substring(4, 2)

        $snapshotTimeUtc = "$year-$month-$day`T$hour`:$minute`:$second`Z"

        return @{
            SnapshotId      = $snapshotId
            SnapshotTimeUtc = $snapshotTimeUtc
        }
    }

    # Fallback if run_id does not match expected pattern
    return @{
        SnapshotId      = $RunId
        SnapshotTimeUtc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
}

if (-not (Test-Path $RootPath)) {
    throw "Root path does not exist: $RootPath"
}

$workloadFolders = Get-ChildItem -Path $RootPath -Directory

foreach ($workloadFolder in $workloadFolders) {

    $workloadName = $workloadFolder.Name.ToLower()
    $runFolders = Get-ChildItem -Path $workloadFolder.FullName -Directory

    foreach ($runFolder in $runFolders) {

        $runId = $runFolder.Name
        $manifestPath = Join-Path $runFolder.FullName "manifest.json"

        if ((Test-Path $manifestPath) -and (-not $Overwrite)) {
            Write-Host "Skipping existing manifest: $manifestPath"
            continue
        }

        $snapshotInfo = Get-SnapshotInfoFromRunId `
            -RunId $runId `
            -Workload $workloadName

        $csvFiles = Get-ChildItem -Path $runFolder.FullName -Filter "*.csv" -File

        $datasets = @()

        foreach ($csvFile in $csvFiles) {

            $datasetName = Convert-ToSnakeCase -Value $csvFile.Name
            $recordCount = Get-RecordCount -CsvPath $csvFile.FullName

            $datasets += [ordered]@{
                dataset_name = $datasetName
                file_name    = $csvFile.Name
                record_count = $recordCount
                schema_name  = "canonical_$datasetName"
                description  = "$datasetName dataset exported from GAM Tool"
            }
        }

        $manifest = [ordered]@{
            export_system          = $ExportSystem
            source_workload        = $workloadName
            export_version         = $ExportVersion
            snapshot_id            = $snapshotInfo.SnapshotId
            snapshot_time_utc      = $snapshotInfo.SnapshotTimeUtc
            run_type               = $RunType
            environment            = $Environment
            tenant_name            = $TenantName
            exported_by            = "${workloadName}_gam_tool"
            delivery_format        = $DeliveryFormat
            file_naming_convention = "{dataset}_{snapshot_id}_B{NNN}.csv"
            datasets               = $datasets
            ingestion_control      = [ordered]@{
                run_id_pattern         = $runId
                workload_name          = $workloadName
                snapshot_granularity   = "full"
                is_delta               = $false
                is_validation_extract  = $false
            }
        }

        $json = $manifest | ConvertTo-Json -Depth 10

        Set-Content -Path $manifestPath -Value $json -Encoding UTF8

        Write-Host "Created manifest: $manifestPath"
    }
}