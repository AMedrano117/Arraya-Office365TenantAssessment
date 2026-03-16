@{
    RootModule        = 'Arraya.M365.Common.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '355a0633-5588-4c49-b021-9dd1ab4ce4ea'
    Author            = 'Arraya Solutions'
    CompanyName       = 'Arraya Solutions'
    PowerShellVersion = '7.0'
    FunctionsToExport = @(
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
        'Import-ArrayaOffice365CustomLocal',
        'Import-ArrayaTenantSnapshot',
        'Invoke-ArrayaCollectionStepSafe',
        'Invoke-QuietCommand',
        'New-ArrayaAssessmentContext',
        'New-ArrayaTenantSnapshot',
        'Test-ArrayaTenantSnapshot',
        'Update-ArrayaTenantSnapshot',
        'Write-ArrayaAssessmentArtifactManifest'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
