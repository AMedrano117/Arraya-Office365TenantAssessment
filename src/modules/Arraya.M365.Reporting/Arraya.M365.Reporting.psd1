@{
    RootModule        = 'Arraya.M365.Reporting.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '4ea90f4f-95e8-478f-95fe-aca01fdd8f0b'
    Author            = 'Arraya Solutions'
    CompanyName       = 'Arraya Solutions'
    PowerShellVersion = '7.0'
    FunctionsToExport = @(
        'Get-ArrayaAssessmentRecommendationText'
        'Get-ArrayaAssessmentWorksheetName'
        'Get-ArrayaEmployeeExperienceInsightsAnalysis'
        'Get-GraphUserStats'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}

