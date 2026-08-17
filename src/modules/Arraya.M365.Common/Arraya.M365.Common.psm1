$public = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public') -Filter *.ps1 -ErrorAction SilentlyContinue)
$privatePath = Join-Path $PSScriptRoot 'Private'
$private = if (Test-Path -Path $privatePath -PathType Container) {
    @(Get-ChildItem -Path $privatePath -Filter *.ps1 -ErrorAction SilentlyContinue)
} else {
    @()
}
foreach ($file in @($private) + @($public)) { . $file.FullName }
Export-ModuleMember -Function @(
    'Convert-ArrayaLegacyTenantStatsToSnapshot',
    'Convert-ArrayaObjectToArray',
    'Convert-ArrayaSnapshotToLegacyTenantStatsHash',
    'Convert-ArrayaToDate',
    'Convert-ArrayaToNumber',
    'ConvertTo-ExportFriendlyRecord',
    'ConvertTo-ExportFriendlyValue',
    'Export-ArrayaErrorReports',
    'Export-ArrayaTenantToTenantCutoverPack',
    'Export-ArrayaTenantSnapshot',
    'Export-HashTableToExcel',
    'Filter-TenantStatsHash',
    'Get-ArrayaAssessmentOutputProfilePolicy',
    'Get-ArrayaAssessmentOutputRoot',
    'Get-ArrayaBitTitanLicenseEstimate',
    'Get-ArrayaBitTitanLicenseModel',
    'Get-ArrayaCollectorCacheValue',
    'Get-ArrayaCollectorPlanStatus',
    'Get-ArrayaObjectValue',
    'Get-ArrayaTenantSnapshotMetricSet',
    'Get-ArrayaTenantSnapshotMetricSetFromContext',
    'Import-ArrayaOffice365CustomLocal',
    'Import-ArrayaTenantSnapshotContext',
    'Import-ArrayaTenantSnapshot',
    'Invoke-ArrayaCollectionStepSafe',
    'Invoke-ArrayaCollectorPlan',
    'Invoke-ArrayaGraphCollectionBatch',
    'Invoke-ArrayaGraphCollectionRequest',
    'Invoke-QuietCommand',
    'New-ArrayaAssessmentOperatorSummary',
    'New-ArrayaAssessmentContext',
    'New-ArrayaCollectorStep',
    'New-ArrayaTenantSnapshot',
    'Resolve-ArrayaSnapshotOutputContext',
    'Set-ArrayaCollectorCacheValue',
    'Test-ArrayaTenantSnapshot',
    'Update-ArrayaTenantSnapshot',
    'Write-ArrayaAssessmentArtifactManifest'
)
