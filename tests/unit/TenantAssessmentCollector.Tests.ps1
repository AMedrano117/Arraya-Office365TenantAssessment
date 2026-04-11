Describe 'Get-FullTenantReportDetails permission preflight' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:collectorPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Get-FullTenantReportDetails.ps1'
        $script:collectorSource = Get-Content -Raw -Path $script:collectorPath
        $script:graphDataPath = Join-Path $script:repoRoot 'src\vendor\Office365Custom\1.2.1\Public\Get-GraphData.ps1'
        $script:graphDataSource = Get-Content -Raw -Path $script:graphDataPath
    }

    It 'defines a hard permission preflight and invokes it before collection starts' {
        Test-Path $script:collectorPath | Should -BeTrue
        $script:collectorSource | Should -Match 'function Test-AssessmentPermissionPreflight'
        $script:collectorSource | Should -Match 'Permission preflight failed\. The assessment will not continue'
        $script:collectorSource | Should -Match 'Permission preflight warnings:'
        $script:collectorSource | Should -Match 'Validating required permissions and service access'
        $script:collectorSource | Should -Match 'Test-AssessmentPermissionPreflight -ConnectionResult \$connectionResult'
    }

    It 'checks the critical Graph permissions used by the collector' {
        @(
            'Organization.Read.All'
            'User.Read.All'
            'AuditLog.Read.All'
            'Group.Read.All'
            'GroupMember.Read.All'
            'RoleManagement.Read.Directory'
            'Domain.Read.All'
            'Device.Read.All'
            'Policy.Read.All'
            'CrossTenantInformation.ReadBasic.All'
            'Application.Read.All'
            'Sites.Read.All'
            'SharePointTenantSettings.Read.All'
            'OnPremDirectorySynchronization.Read.All'
            'SecurityEvents.Read.All'
            'Reports.Read.All'
            'ReportSettings.Read.All'
        ) | ForEach-Object {
            $escapedPattern = [regex]::Escape($_)
            $script:collectorSource | Should -Match $escapedPattern
        }
    }

    It 'checks Exchange and Purview access in addition to Graph' {
        $script:collectorSource | Should -Match 'Exchange mailbox read access'
        $script:collectorSource | Should -Match 'Exchange unified group read access'
        $script:collectorSource | Should -Match 'Ensure-PurviewComplianceSession'
        $script:collectorSource | Should -Match 'Get-RetentionCompliancePolicy'
        $script:collectorSource | Should -Match 'Get-DlpCompliancePolicy'
    }

    It 'treats directory synchronization feature access as a non-blocking validation warning' {
        $script:collectorSource | Should -Match 'OnPremDirectorySynchronization\.Read\.All'
        $script:collectorSource | Should -Match 'IsBlocking\s*=\s*\$false'
        $script:collectorSource | Should -Match 'Hybrid sync and password lifecycle fields may be marked as not validated in current auth mode'
    }

    It 'uses a stable preflight progress counter and suppresses inner Graph record-count progress' {
        $script:collectorSource | Should -Match '\$preflightProgressTotal\s*=\s*\$graphChecks\.Count \+ \$exchangeChecks\.Count'
        $script:collectorSource | Should -Match 'Write-ProgressHelper -Total \(\[Math\]::Max\(\$preflightProgressTotal, 1\)\) -Id \$preflightProgressId'
        $script:collectorSource | Should -Match '\$ProgressIndex\.Value\+\+'
        $script:collectorSource | Should -Match '\(\[ref\]\$preflightProgressIndex\)'
        $script:collectorSource | Should -Match 'Get-ArrayaGraphResource .* -SuppressProgress'
        $script:collectorSource | Should -Match 'Get-ArrayaGraphAdminReportSettings -Headers \$global:GraphHeaders -SuppressProgress'
        $script:graphDataSource | Should -Match '\(\?i\)\(\?:\[\?&\]\)\\\$top='
        $script:graphDataSource | Should -Match '\[switch\]\$SuppressProgress'
    }

    It 'normalizes guest role labels, auth config arrays, and cross-tenant trust parsing for collector summaries' {
        $script:collectorSource | Should -Match 'function Get-AssessmentGuestUserRoleLabel'
        $script:collectorSource | Should -Match 'function Get-AssessmentMfaMethodProfile'
        $script:collectorSource | Should -Match 'function Convert-AssessmentMfaCountMapToText'
        $script:collectorSource | Should -Match 'GuestUserRoleLabel'
        $script:collectorSource | Should -Match 'MfaEnrollmentSummary'
        $script:collectorSource | Should -Match 'MfaEnforcementSummary'
        $script:collectorSource | Should -Match 'UsersWithWeakMethodsOnly'
        $script:collectorSource | Should -Match 'EnabledPoliciesRequiringMfa'
        $script:collectorSource | Should -Match 'SSOApplications\s*=\s*@\('
        $script:collectorSource | Should -Match 'FederatedDomains\s*=\s*@\('
        $script:collectorSource | Should -Match 'function Get-CrossTenantNestedValue'
        $script:collectorSource | Should -Match 'AutomaticTrustSettings\.IsMfaAccepted'
    }

    It 'maps software one-time passcode cleanly and uses a safe camel-case fallback formatter' {
        $script:collectorSource | Should -Match 'softwareonetimepasscode'
        $script:collectorSource | Should -Match 'Software one-time passcode'
        $script:collectorSource | Should -Match '\(\?<=\[a-z\]\)\(\?=\[A-Z\]\)'
    }

    It 'uses meaningful guest scope checks and authentication-strength-aware MFA policy detection' {
        $script:collectorSource | Should -Match 'function Test-AssessmentMeaningfulNestedValue'
        $script:collectorSource | Should -Match 'function Test-AssessmentConditionalAccessRequiresMfa'
        $script:collectorSource | Should -Match 'GrantControls_AuthenticationStrength'
        $script:collectorSource | Should -Match 'UsesAuthenticationStrengthForMfa'
        $script:collectorSource | Should -Match 'RequiresMfaEnforcement'
        $script:collectorSource | Should -Match 'Test-AssessmentMeaningfulNestedValue -Value \$includeGuestsOrExternalUsers'
        $script:collectorSource | Should -Match 'Test-AssessmentMeaningfulNestedValue -Value \$excludeGuestsOrExternalUsers'
        $script:collectorSource | Should -Match 'Test-AssessmentMeaningfulNestedValue -Value \$includeGuestsOrExternalValue'
        $script:collectorSource | Should -Match 'Test-AssessmentMeaningfulNestedValue -Value \$excludeGuestsOrExternalValue'
        $script:collectorSource | Should -Match 'Where-Object \{ Test-AssessmentConditionalAccessRequiresMfa -Policy \$_ \}'
    }

    It 'stores Conditional Access detail rows with stable unique keys and aligns MFA enforcement counts to CA summary totals' {
        $script:collectorSource | Should -Match '\$policyStorageKey'
        $script:collectorSource | Should -Match '\[string\]\$policy\.Id'
        $script:collectorSource | Should -Match 'ConditionalAccessPoliciesReviewed\s*=\s*\$conditionalAccessPoliciesReviewed'
        $script:collectorSource | Should -Match 'Get-ArrayaObjectValue -Object \$conditionalAccessSummaryRecord -Names @\(''TotalPolicies''\)'
    }

    It 'persists estimated MFA enforcement user coverage from enabled Conditional Access policies' {
        $script:collectorSource | Should -Match 'Get-AssessmentMfaEnforcementCoverageSummary -EnabledMfaPolicies \$enabledMfaPolicies'
        $script:collectorSource | Should -Match 'UsersCoveredByEnabledMfaPolicies'
        $script:collectorSource | Should -Match 'UserCoveragePercent'
        $script:collectorSource | Should -Match 'MemberUserCoveragePercent'
        $script:collectorSource | Should -Match 'GuestUserCoveragePercent'
        $script:collectorSource | Should -Match 'CoverageCalculationNote'
    }

    It 'persists MFA enforcement gap users and scope review detail rows for workbook review' {
        $script:collectorSource | Should -Match 'GapUsers'
        $script:collectorSource | Should -Match 'ScopeReview'
        $script:collectorSource | Should -Match 'MfaEnforcementGapUsers'
        $script:collectorSource | Should -Match 'MfaEnforcementScopeReview'
        $script:collectorSource | Should -Match 'Excluded from enabled MFA CA policy'
        $script:collectorSource | Should -Match 'Outside enabled MFA CA include scope'
    }

    It 'maps the broader SharePoint tenant settings surface used by the sharing review' {
        @(
            'ExternalServicesEnabled'
            'EnableAzureADB2BIntegration'
            'AnyoneLinkTrackUsers'
            'NotifyOwnersWhenItemsReshared'
            'MarkNewFilesSensitiveByDefault'
            'RestrictedOneDriveLicense'
        ) | ForEach-Object {
            $escapedPattern = [regex]::Escape($_)
            $script:collectorSource | Should -Match $escapedPattern
        }
    }
}
