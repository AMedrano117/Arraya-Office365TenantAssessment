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
                WorkbookExportPolicy       = 'Default'
                TechnicalHtmlPolicy        = 'Default'
                GenerateMigrationPack      = $false
                CollectionScopePolicy      = 'Default'
            }
        }
        'SolutionsEngineer' {
            return [PSCustomObject]@{
                OutputProfile              = $OutputProfile
                ReportingMode              = 'Operator'
                GenerateWorkbook           = $true
                GenerateTechnicalHtml      = $true
                GenerateBestPracticesHtml  = $false
                GenerateQuestionnaire      = $false
                GenerateJson               = $true
                GeneratePdf                = $false
                WorkbookExportPolicy       = 'Default'
                TechnicalHtmlPolicy        = 'Default'
                GenerateMigrationPack      = $false
                CollectionScopePolicy      = 'Default'
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
                WorkbookExportPolicy       = 'Default'
                TechnicalHtmlPolicy        = 'Default'
                GenerateMigrationPack      = $false
                CollectionScopePolicy      = 'Default'
            }
        }
        'TenantToTenantMigration' {
            return [PSCustomObject]@{
                OutputProfile              = $OutputProfile
                ReportingMode              = 'All'
                GenerateWorkbook           = $true
                GenerateTechnicalHtml      = $true
                GenerateBestPracticesHtml  = $false
                GenerateQuestionnaire      = $false
                GenerateJson               = $true
                GeneratePdf                = $false
                WorkbookExportPolicy       = 'TenantToTenantCutover'
                TechnicalHtmlPolicy        = 'TenantToTenantCutover'
                GenerateMigrationPack      = $true
                CollectionScopePolicy      = 'TenantToTenantCutover'
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
                WorkbookExportPolicy       = 'Default'
                TechnicalHtmlPolicy        = 'Default'
                GenerateMigrationPack      = $false
                CollectionScopePolicy      = 'Default'
            }
        }
        'Machine' {
            return [PSCustomObject]@{
                OutputProfile              = $OutputProfile
                ReportingMode              = 'Automation'
                GenerateWorkbook           = $false
                GenerateTechnicalHtml      = $false
                GenerateBestPracticesHtml  = $false
                GenerateQuestionnaire      = $false
                GenerateJson               = $true
                GeneratePdf                = $false
                WorkbookExportPolicy       = 'Default'
                TechnicalHtmlPolicy        = 'Default'
                GenerateMigrationPack      = $false
                CollectionScopePolicy      = 'Default'
            }
        }
    }
}
