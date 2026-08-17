Describe 'Legacy authentication enforcement evidence' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:collectorPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Get-FullTenantReportDetails.ps1'
        Import-Module -Name (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1') -Force -ErrorAction Stop

        $collectorSource = [System.IO.File]::ReadAllText($script:collectorPath, [System.Text.Encoding]::UTF8)
        $tokens = $null
        $parseErrors = $null
        $collectorAst = [System.Management.Automation.Language.Parser]::ParseInput(
            $collectorSource,
            [ref]$tokens,
            [ref]$parseErrors
        )
        $parseErrors | Should -BeNullOrEmpty

        $wantedFunctions = @(
            'Convert-ToAssessmentStringArray'
            'Test-AssessmentMeaningfulNestedValue'
            'Test-AssessmentConditionalAccessRequiresMfa'
            'Get-AssessmentConditionalAccessPolicyState'
            'Get-AssessmentEnabledAuthenticationMethodNames'
            'Get-AssessmentMfaEnforcementEvidence'
            'Convert-ToAssessmentBoolean'
        )
        $definitions = @($collectorAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $wantedFunctions -contains $node.Name
        }, $true))
        $definitions.Count | Should -Be $wantedFunctions.Count
        foreach ($definition in $definitions) {
            . ([scriptblock]::Create($definition.Extent.Text))
        }

        $authenticationCollectorDefinition = @($collectorAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Get-AuthenticationConfiguration'
        }, $true) | Select-Object -First 1)
        $authenticationCollectorDefinition.Count | Should -Be 1
        $script:authenticationCollectorSource = $authenticationCollectorDefinition[0].Extent.Text

        function script:New-TestConditionalAccessPolicy {
            param(
                [Parameter(Mandatory = $true)]
                [string]$State,
                [switch]$RequiresMfa
            )

            $grantControls = [System.Collections.Generic.Dictionary[string, object]]::new()
            $grantControls['builtInControls'] = if ($RequiresMfa) { @('mfa') } else { @('block') }

            $policy = [System.Collections.Generic.Dictionary[string, object]]::new()
            $policy['state'] = $State
            $policy['grantControls'] = $grantControls
            return $policy
        }

        function script:New-TestAuthenticationMethodConfiguration {
            param(
                [Parameter(Mandatory = $true)]
                [string]$State,
                [Parameter(Mandatory = $true)]
                [string]$OdataType
            )

            $method = [System.Collections.Generic.Dictionary[string, object]]::new()
            $method['state'] = $State
            $method['@odata.type'] = $OdataType
            return $method
        }
    }

    It 'keeps method availability separate from disabled and report-only MFA policies' {
        $disabledPolicy = New-TestConditionalAccessPolicy -State 'disabled' -RequiresMfa
        $reportOnlyPolicy = New-TestConditionalAccessPolicy -State 'enabledForReportingButNotEnforced' -RequiresMfa
        $emailMethod = New-TestAuthenticationMethodConfiguration `
            -State 'enabled' `
            -OdataType '#microsoft.graph.emailAuthenticationMethodConfiguration'
        $authenticationMethods = @(Get-AssessmentEnabledAuthenticationMethodNames -AuthenticationMethodConfigurations @($emailMethod))

        $evidence = Get-AssessmentMfaEnforcementEvidence `
            -AuthenticationMethods $authenticationMethods `
            -ConditionalAccessPolicies @($disabledPolicy, $reportOnlyPolicy) `
            -SecurityDefaultsEnabled $null

        $authenticationMethods | Should -Be @('emailAuthenticationMethodConfiguration')
        $evidence.AuthenticationMethodsAvailable | Should -BeTrue
        $evidence.AuthenticationMethodCount | Should -Be 1
        $evidence.EnabledMfaConditionalAccessPolicyCount | Should -Be 0
        $evidence.ReportOnlyMfaConditionalAccessPolicyCount | Should -Be 1
        $evidence.MfaConditionalAccessEnforced | Should -BeFalse
        $evidence.MfaEnforcementEvidenced | Should -BeFalse
        $evidence.MFAEnabled | Should -BeFalse
    }

    It 'treats an enabled MFA-grant Conditional Access policy as enforcement evidence' {
        $enabledPolicy = New-TestConditionalAccessPolicy -State 'enabled' -RequiresMfa

        $evidence = Get-AssessmentMfaEnforcementEvidence `
            -AuthenticationMethods @('emailAuthenticationMethodConfiguration') `
            -ConditionalAccessPolicies @($enabledPolicy)

        $evidence.EnabledMfaConditionalAccessPolicyCount | Should -Be 1
        $evidence.ReportOnlyMfaConditionalAccessPolicyCount | Should -Be 0
        $evidence.MfaConditionalAccessEnforced | Should -BeTrue
        $evidence.MfaEnforcementEvidenced | Should -BeTrue
        $evidence.MFAEnabled | Should -BeTrue
    }

    It 'retains Security Defaults as an independent enforcement source' {
        $evidence = Get-AssessmentMfaEnforcementEvidence `
            -AuthenticationMethods @('emailAuthenticationMethodConfiguration') `
            -ConditionalAccessPolicies @() `
            -SecurityDefaultsEnabled $true

        $evidence.MfaConditionalAccessEnforced | Should -BeFalse
        $evidence.SecurityDefaultsEnabled | Should -BeTrue
        $evidence.MfaEnforcementEvidenced | Should -BeTrue
        $evidence.MFAEnabled | Should -BeTrue
    }

    It 'uses the evidence helper at the collector boundary instead of method-count inference' {
        $script:authenticationCollectorSource | Should -Match 'Get-AssessmentMfaEnforcementEvidence'
        $script:authenticationCollectorSource | Should -Match 'AuthenticationMethodsAvailable'
        $script:authenticationCollectorSource | Should -Match 'MFAConditionalAccessEnforced'
        $script:authenticationCollectorSource | Should -Match 'SecurityDefaultsEnabled'
        $script:authenticationCollectorSource | Should -Not -Match 'MFAMethods\.Count\s*-gt\s*0\)\s*\{\s*\$authMethodsPolicy\.MFAEnabled\s*=\s*\$true'
    }
}
