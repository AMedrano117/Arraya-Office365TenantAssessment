function Get-ArrayaAssessmentOutputProfilePolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
        [string]$OutputProfile = 'SolutionsEngineer'
    )

    switch ($OutputProfile) {
        'Presales' {
            return [PSCustomObject]@{
                OutputProfile              = $OutputProfile
                ReportingMode              = 'Minimum'
                GenerateWorkbook           = $false
                GenerateTechnicalHtml      = $true
                GenerateBestPracticesHtml  = $false
                GenerateQuestionnaire      = $true
                GenerateJson               = $false
                GeneratePdf                = $false
            }
        }
        'SolutionsEngineer' {
            return [PSCustomObject]@{
                OutputProfile              = $OutputProfile
                ReportingMode              = 'Combined'
                GenerateWorkbook           = $true
                GenerateTechnicalHtml      = $true
                GenerateBestPracticesHtml  = $false
                GenerateQuestionnaire      = $false
                GenerateJson               = $false
                GeneratePdf                = $false
            }
        }
        'ExecutiveLevel' {
            return [PSCustomObject]@{
                OutputProfile              = $OutputProfile
                ReportingMode              = 'Minimum'
                GenerateWorkbook           = $false
                GenerateTechnicalHtml      = $false
                GenerateBestPracticesHtml  = $true
                GenerateQuestionnaire      = $false
                GenerateJson               = $false
                GeneratePdf                = $false
            }
        }
        'TenantToTenantMigration' {
            return [PSCustomObject]@{
                OutputProfile              = $OutputProfile
                ReportingMode              = 'Combined'
                GenerateWorkbook           = $true
                GenerateTechnicalHtml      = $true
                GenerateBestPracticesHtml  = $false
                GenerateQuestionnaire      = $true
                GenerateJson               = $false
                GeneratePdf                = $false
            }
        }
        'Geek' {
            return [PSCustomObject]@{
                OutputProfile              = $OutputProfile
                ReportingMode              = 'Geek'
                GenerateWorkbook           = $true
                GenerateTechnicalHtml      = $true
                GenerateBestPracticesHtml  = $true
                GenerateQuestionnaire      = $true
                GenerateJson               = $true
                GeneratePdf                = $false
            }
        }
        'Machine' {
            return [PSCustomObject]@{
                OutputProfile              = $OutputProfile
                ReportingMode              = 'Geek'
                GenerateWorkbook           = $false
                GenerateTechnicalHtml      = $false
                GenerateBestPracticesHtml  = $false
                GenerateQuestionnaire      = $false
                GenerateJson               = $true
                GeneratePdf                = $false
            }
        }
    }
}
