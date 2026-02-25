@{
    RootModule           = 'Office365Custom.psm1'
    ModuleVersion        = '1.2.0'
    GUID                 = '6affc4b2-e003-496c-85a1-2dbf043de773'
    Author               = 'Aaron Medrano'
    CompanyName          = 'Aaron Medrano'
    Copyright            = '(c) 2024 Aaron Medrano'
    Description          = 'Custom helpers for administering Microsoft 365 services and reports.'
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    FunctionsToExport    = @(
        'Capture-ErrorHelper',
        'Connect-MicrosoftGraph',
        'Connect-MicrosoftGraphAPI',
        'Connect-Office365Services',
        'Convert-MgObjectToUri',
        'Get-AllOneDriveURLs',
        'Get-ExportPath',
        'Get-GraphAPIActivityReport',
        'Get-GraphData',
        'Handle-ErrorHelper',
        'Load-HashTableFromJson',
        'Save-HashTableToJson',
        'Select-ImportedHeader',
        'Validate-ImportedData',
        'Write-Log',
        'Write-ProgressHelper',
        'Start-T2TDomainCutoverMigration', # Domain Cutover Functions
        'Write-MigrationLog', # Domain Cutover Functions
        'Export-MigratingDomainRecipients', # Domain Cutover Functions
        'Set-PrimarySMTPAddress', # Domain Cutover Functions
        'Update-UserUPN', # Domain Cutover Functions
        'Select-FromMultipleO365Matches', # Domain Cutover Functions
        'Manage-DomainAlias', # Domain Cutover Functions
        'Manage-UserBatches' # Domain Cutover Functions
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
            ReleaseNotes = 'Initial module manifest.'
        }
    }
}
