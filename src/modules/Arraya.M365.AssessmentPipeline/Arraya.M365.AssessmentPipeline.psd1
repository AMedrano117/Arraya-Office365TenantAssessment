@{
    RootModule        = 'Arraya.M365.AssessmentPipeline.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '5ab53c95-4a25-49ba-b7a1-cff22f8c5bd0'
    Author            = 'Arraya Solutions'
    CompanyName       = 'Arraya Solutions'
    Description       = 'Phase-based pipeline runtime for Microsoft 365 tenant assessments with checkpoint and resume support.'
    PowerShellVersion = '7.0'
    FunctionsToExport = @(
        'Get-ArrayaAssessmentPipelinePhase',
        'New-ArrayaAssessmentPipelineContext',
        'Invoke-ArrayaAssessmentPipeline',
        'Resume-ArrayaAssessmentPipeline'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
