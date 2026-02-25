Describe 'Arraya.M365.AssessmentRunner' {
    It 'has manifest' {
        (Test-Path './src/modules/Arraya.M365.AssessmentRunner/Arraya.M365.AssessmentRunner.psd1') | Should Be $true
    }

    It 'has primary wrapper scripts' {
        (Test-Path './src/scripts/operations/Start-M365TenantAssessment.ps1') | Should Be $true
        (Test-Path './src/scripts/assessments/tenant-wide/Invoke-M365FullTenantAssessment.ps1') | Should Be $true
        (Test-Path './src/scripts/assessments/identity/Invoke-ActiveDirectoryTenantAssessment.ps1') | Should Be $true
        (Test-Path './src/scripts/assessments/tenant-wide/Invoke-M365GraphActivityReport.ps1') | Should Be $true
        (Test-Path './src/scripts/reporting/Invoke-M365TenantImprovementPlan.ps1') | Should Be $true
        (Test-Path './src/scripts/reporting/Invoke-M365TenantAssessmentComparison.ps1') | Should Be $true
    }

    It 'has migrated legacy scripts' {
        (Test-Path './src/scripts/migrated/legacy/Get-FullTenantReportDetails.ps1') | Should Be $true
        (Test-Path './src/scripts/migrated/legacy/Get-ActiveDirectoryReport.ps1') | Should Be $true
        (Test-Path './src/scripts/migrated/legacy/Get-GraphAPIActivityReport.ps1') | Should Be $true
        (Test-Path './src/scripts/migrated/legacy/New-M365TenantImprovementPlan.ps1') | Should Be $true
        (Test-Path './src/scripts/migrated/legacy/Compare-M365TenantAssessmentSnapshots.ps1') | Should Be $true
    }
}

