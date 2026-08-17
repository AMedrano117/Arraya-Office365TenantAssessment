BeforeAll {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    Import-Module (Join-Path $repoRoot 'src/modules/Arraya.M365.Common/Arraya.M365.Common.psd1') -Force
    $script:CollectorScriptPath = Join-Path $repoRoot 'src/scripts/migrated/legacy/Get-FullTenantReportDetails.ps1'
}

Describe 'Get-ArrayaCollectorPlanStatus' {
    It 'reports Passed when every step completed' {
        $status = Get-ArrayaCollectorPlanStatus -Result @(
            [pscustomobject]@{ Name = 'Tenant overview'; Status = 'Completed'; Required = $true }
            [pscustomobject]@{ Name = 'Users'; Status = 'Completed'; Required = $true }
        )

        $status.Status | Should -Be 'Passed'
        $status.SuggestedExitCode | Should -Be 0
        $status.FailedCount | Should -Be 0
    }

    It 'treats skipped steps as intentional rather than as failures' {
        $status = Get-ArrayaCollectorPlanStatus -Result @(
            [pscustomobject]@{ Name = 'Tenant overview'; Status = 'Completed'; Required = $true }
            [pscustomobject]@{ Name = 'Users'; Status = 'Skipped'; Required = $true; Message = 'Disabled by profile' }
        )

        $status.Status | Should -Be 'Passed'
        $status.SkippedCount | Should -Be 1
        $status.SuggestedExitCode | Should -Be 0
    }

    It 'reports Degraded when only optional collectors failed' {
        $status = Get-ArrayaCollectorPlanStatus -Result @(
            [pscustomobject]@{ Name = 'Tenant overview'; Status = 'Completed'; Required = $true }
            [pscustomobject]@{ Name = 'Email activity insights'; Status = 'Failed'; Required = $false; Message = 'Reports API threw' }
        )

        $status.Status | Should -Be 'Degraded'
        $status.SuggestedExitCode | Should -Be 0
        $status.OptionalFailedCount | Should -Be 1
        $status.OptionalFailedSteps | Should -Contain 'Email activity insights'
        $status.Summary | Should -Match 'Email activity insights'
    }

    It 'reports Failed with a nonzero exit code when a required collector failed' {
        $status = Get-ArrayaCollectorPlanStatus -Result @(
            [pscustomobject]@{ Name = 'Tenant overview'; Status = 'Completed'; Required = $true }
            [pscustomobject]@{ Name = 'Users'; Status = 'Failed'; Required = $true; Message = 'Graph returned 403' }
            [pscustomobject]@{ Name = 'Email activity insights'; Status = 'Failed'; Required = $false; Message = 'Reports API threw' }
        )

        $status.Status | Should -Be 'Failed'
        $status.SuggestedExitCode | Should -Be 2
        $status.RequiredFailedCount | Should -Be 1
        $status.RequiredFailedSteps | Should -Contain 'Users'
        $status.FailedCount | Should -Be 2
    }

    It 'carries the failure message through for reporting' {
        $status = Get-ArrayaCollectorPlanStatus -Result @(
            [pscustomobject]@{ Name = 'Users'; Section = 'Identity'; Workload = 'Identity'; Status = 'Failed'; Required = $true; Message = 'Graph returned 403' }
        )

        $failure = @($status.FailedSteps)[0]
        $failure.Name | Should -Be 'Users'
        $failure.Section | Should -Be 'Identity'
        $failure.Required | Should -BeTrue
        $failure.Message | Should -Be 'Graph returned 403'
    }

    It 'treats a step with no Required property as optional' {
        $status = Get-ArrayaCollectorPlanStatus -Result @(
            [pscustomobject]@{ Name = 'Legacy step'; Status = 'Failed'; Message = 'boom' }
        )

        $status.Status | Should -Be 'Degraded'
        $status.SuggestedExitCode | Should -Be 0
    }

    It 'reports Passed for an empty plan' {
        $status = Get-ArrayaCollectorPlanStatus -Result @()

        $status.Status | Should -Be 'Passed'
        $status.TotalSteps | Should -Be 0
    }
}

Describe 'Invoke-ArrayaCollectorPlan required propagation' {
    It 'carries the Required flag from the step definition into the result record' {
        $sections = @([pscustomobject]@{ Name = 'Identity' })
        $steps = @(
            New-ArrayaCollectorStep -Name 'Required step' -Section 'Identity' -Required $true -ScriptBlock { 'ok' }
            New-ArrayaCollectorStep -Name 'Optional step' -Section 'Identity' -ScriptBlock { 'ok' }
        )

        $results = Invoke-ArrayaCollectorPlan -Sections $sections -Steps $steps

        (@($results | Where-Object { $_.Name -eq 'Required step' })[0]).Required | Should -BeTrue
        (@($results | Where-Object { $_.Name -eq 'Optional step' })[0]).Required | Should -BeFalse
    }

    It 'records a required collector that throws as a Failed required step' {
        $sections = @([pscustomobject]@{ Name = 'Identity' })
        $steps = @(
            New-ArrayaCollectorStep -Name 'Users' -Section 'Identity' -Required $true -ScriptBlock { throw 'Graph returned 403' }
        )

        $results = Invoke-ArrayaCollectorPlan -Sections $sections -Steps $steps
        $status = Get-ArrayaCollectorPlanStatus -Result $results

        $status.Status | Should -Be 'Failed'
        $status.SuggestedExitCode | Should -Be 2
        $status.RequiredFailedSteps | Should -Contain 'Users'
    }
}

Describe 'Assessment collector status wiring' {
    BeforeAll {
        $script:collectorSource = Get-Content -Path $script:CollectorScriptPath -Raw
    }

    It 'captures the collector plan results instead of discarding them' {
        $script:collectorSource | Should -Not -Match 'Invoke-ProfileAwareAssessmentStep -Name \(\[string\]\$Step\.Name\).*\r?\n\s*\} \| Out-Null'
        $script:collectorSource | Should -Match '\$collectorPlanResults = @\('
    }

    It 'evaluates the plan status before exporting deliverables' {
        $statusIndex = $script:collectorSource.IndexOf('Get-ArrayaCollectorPlanStatus -Result $collectorPlanResults')
        $exportIndex = $script:collectorSource.IndexOf('### Export Reports ###')

        $statusIndex | Should -BeGreaterThan 0
        $exportIndex | Should -BeGreaterThan 0
        $statusIndex | Should -BeLessThan $exportIndex
    }

    It 'marks the load-bearing collectors as required' {
        $script:collectorSource | Should -Match "-Name 'Tenant overview'[^\r\n]*-Required \`$true"
        $script:collectorSource | Should -Match "-Name 'Users'[^\r\n]*-Required \`$true"
        $script:collectorSource | Should -Match "-Name 'Exchange mailboxes'[^\r\n]*-Required \`$true"
    }

    It 'exits nonzero only when required data failed' {
        $script:collectorSource | Should -Match 'exit \(\[int\]\$script:AssessmentRunExitCode\)'
    }
}
