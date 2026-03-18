@{
    RootModule           = 'Office365Custom.psm1'
    ModuleVersion        = '1.2.1'
    GUID                 = '6affc4b2-e003-496c-85a1-2dbf043de773'
    Author               = 'Aaron Medrano'
    CompanyName          = 'Aaron Medrano'
    Copyright            = '(c) 2024 Aaron Medrano'
    Description          = 'Custom helpers for administering Microsoft 365 services and reports.'
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    FunctionsToExport    = @(
        'Capture-ErrorHelper',
        'Connect-Office365',
        'Convert-MgObjectToUri',
        'Export-MigratingDomainRecipients',
        'Get-AllOneDriveURLs',
        'Get-ExportPath',
        'Get-GraphAPIActivityReport',
        'Get-GraphData',
        'Handle-ErrorHelper',
        'Load-HashTableFromJson',
        'Manage-DomainAlias',
        'Manage-UserBatches',
        'Save-HashTableToJson',
        'Select-FromMultipleO365Matches',
        'Select-ImportedHeader',
        'Set-PrimarySMTPAddress',
        'Start-T2TDomainCutoverMigration',
        'Update-UserUPN',
        'Validate-ImportedData',
        'Write-Log',
        'Write-ProgressHelper',
        'Write-MigrationLog'
    )
    AliasesToExport      = @()
    CmdletsToExport      = @()
    VariablesToExport    = @()
    RequiredModules      = @(
        'Microsoft.Graph.Authentication',
        'ExchangeOnlineManagement',
        'ImportExcel'
    )
    PrivateData          = @{
        PSData = @{
            Tags        = @('Microsoft365', 'Graph', 'ExchangeOnline', 'SharePoint', 'Automation')
            ProjectUri  = 'https://github.com/'
            LicenseUri  = ''
            ReleaseNotes = 'Version 1.2.1 refresh with updated manifest metadata, corrected exported function list, and documentation alignment.'
        }
    }
}
