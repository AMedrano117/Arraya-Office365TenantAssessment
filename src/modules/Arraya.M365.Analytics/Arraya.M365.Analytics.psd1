@{
    RootModule        = 'Arraya.M365.Analytics.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '8c3f21b6-5d47-4a92-9e18-7b6c0f4a2d35'
    Author            = 'Arraya Solutions'
    CompanyName       = 'Arraya Solutions'
    Description       = 'Shared tenant analysis helpers used by both the HTML assessment report and the tenant improvement plan.'
    PowerShellVersion = '7.0'
    FunctionsToExport = @(
        'Add-ArrayaGroupLicensingEffectiveCounts'
        'Get-ArrayaConditionalAccessPolicyState'
        'Get-ArrayaDateText'
        'Get-ArrayaGroupLicenseAssignedUserCountLookup'
        'Get-ArrayaLatestDateFromRows'
        'Get-ArrayaLicenseSkuLookup'
        'Get-ArrayaLicenseSuiteFamily'
        'Get-ArrayaNormalizedPrincipalKey'
        'Get-ArrayaUserLicenseActivityState'
        'Get-ArrayaUserLicenseAssignmentRows'
        'New-ArrayaActivityLookup'
        'New-ArrayaConditionalAccessOptimizationRows'
        'New-ArrayaGroupLicensingSummaryRows'
        'New-ArrayaLicenseOptimizationCandidateRows'
        'New-ArrayaMfaMethodPostureSummaryRows'
        'New-ArrayaPrivilegedAccessRemediationSummaryRows'
        'New-ArrayaTeamsGroupsCleanupCandidateRows'
        'Resolve-ArrayaLicenseSkuName'
        'Test-ArrayaGroupManagesLicenses'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
