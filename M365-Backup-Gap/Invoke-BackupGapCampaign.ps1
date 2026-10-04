<#
.SYNOPSIS
    Runs the assessment (where tenant details exist) and outreach for every customer in a CSV.
    Used by the scheduled GitHub Actions workflow and runnable by hand.
.EXAMPLE
    ./Invoke-BackupGapCampaign.ps1 -CustomersCsv ./customers.csv -PricePerUser 3
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$CustomersCsv,
    [decimal]$PricePerUser = 3,
    [decimal]$LossCostPerUser = 500,
    [string]$YourName = 'Your Name',
    [string]$OutputDir = (Join-Path $PSScriptRoot 'output')
)
$summary = foreach ($c in Import-Csv $CustomersCsv) {
    $status = 'outreach-only'
    if ($c.Organization -and $c.ClientId -and $c.PfxEnvVar) {
        try {
            & (Join-Path $PSScriptRoot 'Invoke-M365PurviewChecks.ps1') -CustomerName $c.CustomerName -Organization $c.Organization `
                -ClientId $c.ClientId -PfxEnvVar $c.PfxEnvVar -PfxPasswordEnvVar $c.PfxPasswordEnvVar -OutputDir $OutputDir
        } catch { Write-Warning "purview-failed: $($_.Exception.Message)" }
    }
    if ($c.TenantId -and $c.ClientId -and $c.SecretEnvVar) {
        try {
            & (Join-Path $PSScriptRoot 'Invoke-M365BackupGapAssessment.ps1') -CustomerName $c.CustomerName `
                -TenantId $c.TenantId -ClientId $c.ClientId -SecretEnvVar $c.SecretEnvVar `
                -PricePerUser $PricePerUser -LossCostPerUser $LossCostPerUser -OutputDir $OutputDir
            $status = 'assessed'
        } catch { $status = "assessment-failed: $($_.Exception.Message)"; Write-Warning $status }
    }
    [pscustomobject]@{ Customer = $c.CustomerName; Status = $status }
}
& (Join-Path $PSScriptRoot 'New-BackupGapOutreach.ps1') -CustomersCsv $CustomersCsv -PricePerUser $PricePerUser `
    -LossCostPerUser $LossCostPerUser -YourName $YourName -OutputDir $OutputDir
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$summary | Export-Csv (Join-Path $OutputDir 'campaign-summary.csv') -NoTypeInformation
$summary | Format-Table -AutoSize
