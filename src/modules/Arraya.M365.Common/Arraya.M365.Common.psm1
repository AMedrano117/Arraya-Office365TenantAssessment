$public = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public') -Filter *.ps1 -ErrorAction SilentlyContinue)
$private = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter *.ps1 -ErrorAction SilentlyContinue)
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
    'Export-ArrayaTenantSnapshot',
    'Export-HashTableToExcel',
    'Filter-TenantStatsHash',
    'Get-ArrayaAssessmentOutputProfilePolicy',
    'Get-ArrayaAssessmentOutputRoot',
    'Get-ArrayaObjectValue',
    'Get-ArrayaTenantSnapshotMetricSet',
    'Import-ArrayaOffice365CustomLocal',
    'Import-ArrayaTenantSnapshotContext',
    'Import-ArrayaTenantSnapshot',
    'Invoke-ArrayaCollectionStepSafe',
    'Invoke-QuietCommand',
    'New-ArrayaAssessmentContext',
    'New-ArrayaTenantSnapshot',
    'Resolve-ArrayaSnapshotOutputContext',
    'Test-ArrayaTenantSnapshot',
    'Update-ArrayaTenantSnapshot',
    'Write-ArrayaAssessmentArtifactManifest'
)
