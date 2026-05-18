@{
    RootModule        = 'Arraya.M365.AssessmentRunner.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = 'a91132d8-b1ec-4338-94e3-c4314c3f000f'
    Author            = 'Arraya Solutions'
    CompanyName       = 'Arraya Solutions'
    Description       = 'Runner commands for Microsoft 365 tenant assessments, AD assessments, snapshot comparisons, and improvement planning.'
    PowerShellVersion = '7.0'
    FunctionsToExport = @(
        'Invoke-M365TenantAssessment',
        'Invoke-M365TenantConnectionPreflight',
        'Invoke-M365TenantDataCollection',
        'Invoke-M365TenantAssessmentExport',
        'Invoke-ADTenantAssessment',
        'Invoke-GraphActivityAssessment',
        'Invoke-M365ImprovementPlan',
        'Invoke-M365AssessmentComparison'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
