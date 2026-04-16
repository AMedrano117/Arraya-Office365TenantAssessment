Describe 'Get-FullTenantReportDetails permission preflight' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:collectorPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Get-FullTenantReportDetails.ps1'
        $script:collectorSource = Get-Content -Raw -Path $script:collectorPath
        $script:graphDataPath = Join-Path $script:repoRoot 'src\vendor\Office365Custom\1.2.1\Public\Get-GraphData.ps1'
        $script:graphDataSource = Get-Content -Raw -Path $script:graphDataPath
    }

    It 'defines a permission preflight with an explicit skip switch and invokes it before collection starts by default' {
        Test-Path $script:collectorPath | Should -BeTrue
        $script:collectorSource | Should -Match 'function Test-AssessmentImportedModuleMatchesManifestPath'
        $script:collectorSource | Should -Match 'function Test-AssessmentPermissionPreflight'
        $script:collectorSource | Should -Match '\[switch\]\$PreflightOnly'
        $script:collectorSource | Should -Match '\[switch\]\$SkipPermissionPreflight'
        $script:collectorSource | Should -Match 'Permission preflight failed\. The assessment will not continue'
        $script:collectorSource | Should -Match 'Permission preflight warnings:'
        $script:collectorSource | Should -Match 'Test-AssessmentPermissionPreflight -ConnectionResult \(\[pscustomobject\]\$authResult\) -Workload Graph'
        $script:collectorSource | Should -Match 'Test-AssessmentPermissionPreflight -ConnectionResult \(\[pscustomobject\]\$authResult\) -Workload ExchangeOnline'
        $script:collectorSource | Should -Match 'Test-AssessmentPermissionPreflight -ConnectionResult \(\[pscustomobject\]\$authResult\) -Workload Purview'
        $script:collectorSource | Should -Match 'Connection / Preflight'
        $script:collectorSource | Should -Match 'Write-ConnectionPreflightSummary'
        $script:collectorSource | Should -Match 'if \(\$runPreflightOnly\)\s*\{\s*return'
        $script:collectorSource | Should -Match 'Workload preflight skipped by request'
        $script:collectorSource | Should -Not -Match 'Legend: cyan=section/progress, green=completed, yellow=warnings/skips\.'
        $script:collectorSource | Should -Not -Match 'Clear-Host'
        $script:collectorSource | Should -Not -Match 'Progress view: overall step completion is shown after each major task\.'
        $script:collectorSource | Should -Match 'Join-Path -Path \$Module\.ModuleBase -ChildPath \(\[System\.IO\.Path\]::GetFileName\(\$resolvedManifestPath\)\)'
        $script:collectorSource | Should -Not -Match '\$loadedCommonModule\.Path -ne \$resolvedCommonManifestPath'
        $script:collectorSource | Should -Not -Match '\$loadedReportingModule\.Path -ne \$resolvedReportingManifestPath'
        $script:collectorSource | Should -Not -Match 'Connect-Office365 @connectOffice365Params'
    }

    It 'uses an assessment-owned auth orchestrator instead of the generic Office365 connector bootstrap' {
        $script:collectorSource | Should -Match 'function Resolve-AssessmentProfileCollectionPlan'
        $script:collectorSource | Should -Match 'function Resolve-AssessmentRequestedAuthMode'
        $script:collectorSource | Should -Match 'function Get-AssessmentGraphDelegatedScopes'
        $script:collectorSource | Should -Match 'function Resolve-AssessmentAuthWorkloadPlan'
        $script:collectorSource | Should -Match 'function Connect-AssessmentGraph'
        $script:collectorSource | Should -Match 'function Connect-AssessmentExchange'
        $script:collectorSource | Should -Match 'function Connect-AssessmentPurview'
        $script:collectorSource | Should -Match 'function Connect-AssessmentSharePoint'
        $script:collectorSource | Should -Match 'function Connect-AssessmentTeams'
        $script:collectorSource | Should -Match 'function Test-AssessmentExistingSessions'
        $script:collectorSource | Should -Match 'function Initialize-AssessmentAuthentication'
        $script:collectorSource | Should -Match 'Assessment login mode:'
        $script:collectorSource | Should -Match 'Required auth workloads:'
        $script:collectorSource | Should -Match 'Resolve-AssessmentProfileCollectionPlan `'
        $script:collectorSource | Should -Match 'Resolve-AssessmentAuthWorkloadPlan'
        $script:collectorSource | Should -Match 'Initialize-AssessmentAuthentication'
        $script:collectorSource | Should -Not -Match "'Connect-Office365',"
    }

    It 'makes auth workload planning profile-driven with explicit fallback/skipped workload states' {
        $script:collectorSource | Should -Match 'RequiredWorkloads'
        $script:collectorSource | Should -Match 'ConnectedWorkloads'
        $script:collectorSource | Should -Match 'SkippedWorkloads'
        $script:collectorSource | Should -Match 'FallbackWorkloads'
        $script:collectorSource | Should -Match 'PurviewCompliance'
        $script:collectorSource | Should -Match 'SharePointOnline'
        $script:collectorSource | Should -Match 'GraphFallback'
        $script:collectorSource | Should -Match 'SkippedByDesign'
        $script:collectorSource | Should -Match 'GraphOnly'
        $script:collectorSource | Should -Match 'OptionalModuleConnect'
        $script:collectorSource | Should -Match 'OptionalPowerShellConnect'
    }

    It 'validates only the required workload sessions when SkipAuth is used' {
        $script:collectorSource | Should -Match 'Reusing existing workload sessions for this run'
        $script:collectorSource | Should -Match 'SkipAuth was requested, but no existing Microsoft Graph session was found'
        $script:collectorSource | Should -Match 'SkipAuth was requested, but Exchange Online cmdlets are not available in the current session'
        $script:collectorSource | Should -Match 'function Test-AssessmentPurviewSessionReady'
        $script:collectorSource | Should -Match 'SkipAuth was requested, but Purview compliance session is not usable in the current session'
        $script:collectorSource | Should -Match 'Test-AssessmentExistingSessions -WorkloadPlan \$WorkloadPlan'
    }

    It 'uses plain-language governance progress labels instead of legacy Tier B terminology' {
        $script:collectorSource | Should -Match 'Exchange governance summaries'
        $script:collectorSource | Should -Match 'Operational governance summaries'
        $script:collectorSource | Should -Match 'governance findings'
        $script:collectorSource | Should -Not -Match 'Exchange governance Tier B summaries'
        $script:collectorSource | Should -Not -Match 'Operational Tier B summaries'
    }

    It 'separates connection from assessment and uses the new six-section assessment flow' {
        $script:collectorSource | Should -Match "Write-ConsoleSection -Step 'Connection' -Title 'Connection / Preflight'"
        $script:collectorSource | Should -Match "Write-ConsoleSection -Step '1/6' -Title 'Tenant Overview'"
        $script:collectorSource | Should -Match "Write-ConsoleSection -Step '2/6' -Title 'Identity'"
        $script:collectorSource | Should -Match "Write-ConsoleSection -Step '3/6' -Title 'Exchange'"
        $script:collectorSource | Should -Match "Write-ConsoleSection -Step '4/6' -Title 'Collaboration'"
        $script:collectorSource | Should -Match "Write-ConsoleSection -Step '5/6' -Title 'Endpoint'"
        $script:collectorSource | Should -Match "Write-ConsoleSection -Step '6/6' -Title 'Governance'"
        $script:collectorSource | Should -Match "Write-ConsoleSection -Step 'Export' -Title 'Exporting results'"
        $script:collectorSource | Should -Not -Match 'Consolidating Discovery Report'
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
        $script:collectorSource | Should -Match 'Exchange Online session exists, but EXO cmdlets are not visible in the collector scope'
        $script:collectorSource | Should -Match 'function Ensure-PurviewComplianceCommandAvailable'
        $script:collectorSource | Should -Match 'function Set-PurviewComplianceDiagnosticState'
        $script:collectorSource | Should -Match 'function Get-PurviewComplianceDiagnosticMessage'
        $script:collectorSource | Should -Match 'Ensure-PurviewComplianceSession'
        $script:collectorSource | Should -Match 'function Invoke-PurviewComplianceDelegatedConnect'
        $script:collectorSource | Should -Match 'Get-RetentionCompliancePolicy'
        $script:collectorSource | Should -Match 'Get-DlpCompliancePolicy'
        $script:collectorSource | Should -Match 'ExchangeOnlineManagement'
        $script:collectorSource | Should -Match "Import-Module 'ExchangeOnlineManagement'"
        $script:collectorSource | Should -Match 'Connect-IPPSSession'
        $script:collectorSource | Should -Match 'Connected to Purview compliance PowerShell using certificate authentication'
        $script:collectorSource | Should -Match 'interactive authentication with -DisableWAM'
        $script:collectorSource | Should -Match 'device code authentication'
        $script:collectorSource | Should -Match 'Client secret authentication is not supported for Purview compliance PowerShell in this workflow'
        $script:collectorSource | Should -Match 'Purview compliance PowerShell session could not be established\. Underlying error:'
        $script:collectorSource | Should -Match 'Organization used:'
        $script:collectorSource | Should -Match 'Next step:'
        $script:collectorSource | Should -Match 'Missing compliance cmdlets after connect:'
        $script:collectorSource | Should -Match 'The certificate and app registration were accepted, but this tenant did not expose the Purview retention/DLP cmdlets to that app session'
    }

    It 'treats directory synchronization feature access as a non-blocking validation warning' {
        $script:collectorSource | Should -Match 'OnPremDirectorySynchronization\.Read\.All'
        $script:collectorSource | Should -Match 'IsBlocking\s*=\s*\$false'
        $script:collectorSource | Should -Match 'Hybrid sync and password lifecycle fields may be marked as not validated in current auth mode'
    }

    It 'uses a stable preflight progress counter and suppresses inner Graph record-count progress' {
        $script:collectorSource | Should -Match '\$preflightProgressTotal\s*=\s*\$selectedGraphChecks\.Count \+ \$selectedExchangeChecks\.Count'
        $script:collectorSource | Should -Match 'Write-ProgressHelper -Total \(\[Math\]::Max\(\$preflightProgressTotal, 1\)\) -Id \$preflightProgressId'
        $script:collectorSource | Should -Match '\$ProgressIndex\.Value\+\+'
        $script:collectorSource | Should -Match '\(\[ref\]\$preflightProgressIndex\)'
        $script:collectorSource | Should -Match '\$progressActivity = "Permission preflight: \{0\}: \{1\}" -f \$Area, \$Requirement'
        $script:collectorSource | Should -Match 'Write-ProgressHelper -Total \(\[Math\]::Max\(\$preflightProgressTotal, 1\)\) -Id \$preflightProgressId -Index \$ProgressIndex\.Value -Activity \$progressActivity'
        $script:collectorSource | Should -Match '\{0\}: \{1\} successful, \{2\} remaining \(\{3\} blocking, \{4\} non-blocking\)\.'
        $script:collectorSource | Should -Match 'Get-ArrayaGraphResource .* -SuppressProgress'
        $script:collectorSource | Should -Match 'Get-ArrayaGraphResource .* -SuppressAccessDeniedWarning'
        $script:collectorSource | Should -Match 'Get-ArrayaGraphAdminReportSettings -Headers \$global:GraphHeaders -SuppressProgress'
        $script:graphDataSource | Should -Match '\(\?i\)\(\?:\[\?&\]\)\\\$top='
        $script:graphDataSource | Should -Match '\[switch\]\$SuppressProgress'
        $script:graphDataSource | Should -Match '\[switch\]\$SuppressAccessDeniedWarning'
    }

    It 'fast-passes core Graph permission checks from token claims and keeps live probes for ambiguous endpoints' {
        $script:collectorSource | Should -Match '\[bool\]\$TrustClaimPresence = \$false'
        $script:collectorSource | Should -Match 'if \(\$TrustClaimPresence -and \$claimState -eq \$true\)'
        $script:collectorSource | Should -Match 'TrustClaimPresence = \$true'
        $script:collectorSource | Should -Match "PermissionNames = @\('SharePointTenantSettings.Read.All'\)"
        $script:collectorSource | Should -Match "PermissionNames = @\('Reports.Read.All'\)"
        $script:collectorSource | Should -Match "PermissionNames = @\('ReportSettings.Read.All'\)"
        $script:collectorSource | Should -Match "PermissionNames = @\('OnPremDirectorySynchronization.Read.All'\)"
        $script:collectorSource | Should -Match 'Sites\.ReadWrite\.All'
        $script:collectorSource | Should -Match 'Application\.ReadWrite\.All'
        $script:collectorSource | Should -Match 'RoleManagement\.ReadWrite\.Directory'
        $script:collectorSource | Should -Match "PermissionNames = @\('CrossTenantInformation.ReadBasic.All'\)"
        $script:collectorSource | Should -Match "PermissionNames = @\('Application.Read.All'\)"
        $script:collectorSource | Should -Match "PermissionNames = @\('Sites.Read.All'\)"
        $script:collectorSource | Should -Match "PermissionNames = @\('Channel.ReadBasic.All'\)"
    }

    It 'forwards the access-denied warning suppression flag through the local Graph wrapper' {
        $script:collectorSource | Should -Match 'function Get-ArrayaGraphResource'
        $script:collectorSource | Should -Match '\[switch\]\$SuppressAccessDeniedWarning'
        $script:collectorSource | Should -Match '-SuppressAccessDeniedWarning:\$SuppressAccessDeniedWarning'
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

    It 'removes handled non-fatal Exchange and directory-role probe errors from the session error stack' {
        $script:collectorSource | Should -Match '\$errorCountBeforeRoleLookup = \$global:Error\.Count'
        $script:collectorSource | Should -Match 'while \(\$global:Error\.Count -gt \$errorCountBeforeRoleLookup\)'
        $script:collectorSource | Should -Match '\$errorCountBeforeInboxRuleLookup = \$global:Error\.Count'
        $script:collectorSource | Should -Match 'while \(\$global:Error\.Count -gt \$errorCountBeforeInboxRuleLookup\)'
        $script:collectorSource | Should -Match '\$global:Error\.RemoveAt\(0\)'
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
        $script:collectorSource | Should -Match 'Excluded from all enabled MFA CA policies that otherwise target the user'
        $script:collectorSource | Should -Match 'Outside enabled MFA CA include scope'
    }

    It 'persists admin-specific MFA registration and enforcement review outputs' {
        $script:collectorSource | Should -Match 'Get-AssessmentAdminMfaReview'
        $script:collectorSource | Should -Match 'AdminMfaSummary'
        $script:collectorSource | Should -Match 'AdminMfaRegistrationGaps'
        $script:collectorSource | Should -Match 'AdminMfaEnforcementGaps'
        $script:collectorSource | Should -Match 'GuestUserEnforcementState'
    }

    It 'batches enterprise application sign-in enrichment instead of making one Graph call per app' {
        $script:collectorSource | Should -Match 'function Get-AssessmentEnterpriseApplicationLatestSignInMap'
        $script:collectorSource | Should -Match 'Some enterprise application sign-in lookups exceeded the reviewed sign-in history window'
        $script:collectorSource | Should -Match 'Unable to retrieve batched enterprise application sign-in details'
        $script:collectorSource | Should -Not -Match 'Get-AssessmentEnterpriseApplicationLatestSignIn -AppId'
    }

    It 'emits compact one-line assessment step status output instead of the older redundant progress pair' {
        $script:collectorSource | Should -Not -Match 'Write-Host \("Gathering \{0\} \.\.\." -f \$Name\)'
        $script:collectorSource | Should -Not -Match 'Overall progress:'
        $script:collectorSource | Should -Match 'Write-AssessmentCollectorCompletionBanner'
        $script:collectorSource | Should -Match '\[\{0\}/\{1\} \| \{2\}%\] \{3\} - \{4\} in \{5\}\{6\}'
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
        $script:collectorSource | Should -Match 'AnonymousLinkExpirationInDays'
        $script:collectorSource | Should -Match '\[int\]::TryParse\(\$anonymousLinkExpirationText, \[ref\]\$parsedAnonymousLinkExpirationDays\)'
        $script:collectorSource | Should -Match "Get-Variable -Name connectionResult -Scope Script -ErrorAction SilentlyContinue"
        $script:collectorSource | Should -Match '\$sharePointConnectionValue = \$scriptConnectionResult\.Value\.SharePointOnline'
        $script:collectorSource | Should -Match '\[string\]::IsNullOrWhiteSpace\(\[string\]\$sharePointConnectionValue\)'
        $script:collectorSource | Should -Match '\$sharePointConnected = \(\[string\]\$sharePointConnectionValue -match ''\^\(\?i:true\|1\|yes\|connected\)\$''\)'
    }

    It 'builds a lightweight enterprise application inventory and SSO subset for non-minimum profiles' {
        $script:collectorSource | Should -Match 'function Test-AssessmentAppUsesSso'
        $script:collectorSource | Should -Match 'function Get-AssessmentEnterpriseApplicationIdentityProfile'
        $script:collectorSource | Should -Match 'function Get-AssessmentEnterpriseApplicationLatestSignIn'
        $script:collectorSource | Should -Match 'function Test-AssessmentEnterpriseApplicationValidRow'
        $script:collectorSource | Should -Match 'function Test-AssessmentEnterpriseApplicationCustomerRelevant'
        $script:collectorSource | Should -Match 'function Add-AssessmentEnterpriseApplicationRecord'
        $script:collectorSource | Should -Match 'auditLogs/signIns'
        $script:collectorSource | Should -Match 'LastSignInUserDisplayName'
        $script:collectorSource | Should -Match 'LastConditionalAccessStatus'
        $script:collectorSource | Should -Match 'LastClientAppUsed'
        $script:collectorSource | Should -Match 'PreferredSingleSignOnMode'
        $script:collectorSource | Should -Match 'SsoEnabled'
        $script:collectorSource | Should -Match 'SSOMode'
        $script:collectorSource | Should -Match 'PublisherName'
        $script:collectorSource | Should -Match 'SsoEnabledApplications'
        $script:collectorSource | Should -Match 'servicePrincipalType'
        $script:collectorSource | Should -Match 'ManagedIdentity'
        $script:collectorSource | Should -Match 'publisherName'
        $script:collectorSource | Should -Match 'Get-ArrayaGraphResource .*servicePrincipals'
        $script:collectorSource | Should -Match 'Checking Enterprise Applications for SSO'
        $script:collectorSource | Should -Match 'AuthenticationSSOApplications'
        $script:collectorSource | Should -Match '\$resolvedSsoApplicationRows'
    }

    It 'treats Teams member and guest counts as optional enrichment for app-based runs without TeamMember scopes' {
        $script:collectorSource | Should -Match '\$teamsChannelExpansionAllowed = \$true'
        $script:collectorSource | Should -Match '\$teamsChannelExpansionWarningLogged = \$false'
        $script:collectorSource | Should -Match '\$teamsMemberExpansionAllowed = \$true'
        $script:collectorSource | Should -Match '\$teamsMemberExpansionWarningLogged = \$false'
        $script:collectorSource | Should -Match 'https://graph\.microsoft\.com/v1\.0/teams/\{0\}/allChannels'
        $script:collectorSource | Should -Match 'https://graph\.microsoft\.com/v1\.0/teams/\{0\}/members'
        $script:collectorSource | Should -Match 'Get-ArrayaGraphResource'
        $script:collectorSource | Should -Match 'allChannels\?`\$select=displayName,membershipType'
        $script:collectorSource | Should -Match 'members\?`\$select=id,email,userId,roles'
        $script:collectorSource | Should -Match 'Teams channel inventory requires Channel\.ReadBasic\.All'
        $script:collectorSource | Should -Match 'Teams member and guest counts require TeamMember.Read.All or TeamMember.ReadWrite.All in app-based Graph collection'
        $script:collectorSource | Should -Match 'Continuing with team and channel inventory only'
        $script:collectorSource | Should -Match '\$channelErrorCountBeforeFetch = \$global:Error\.Count'
        $script:collectorSource | Should -Match '\$memberErrorCountBeforeFetch = \$global:Error\.Count'
        $script:collectorSource | Should -Match 'while \(\$global:Error\.Count -gt \$channelErrorCountBeforeFetch\)'
        $script:collectorSource | Should -Match 'while \(\$global:Error\.Count -gt \$memberErrorCountBeforeFetch\)'
    }

    It 'cleans up non-blocking hybrid-sync and health-command errors instead of leaving them in the session error stack' {
        $script:collectorSource | Should -Match '\$syncFeatureErrorCountBeforeFetch = \$global:Error\.Count'
        $script:collectorSource | Should -Match 'while \(\$global:Error\.Count -gt \$syncFeatureErrorCountBeforeFetch\)'
        $script:collectorSource | Should -Match 'Get-Command Get-AzureADConnectHealthSyncServices -ErrorAction Ignore'
        $script:collectorSource | Should -Match 'Get-Command Get-AzureADConnectHealthSyncErrors -ErrorAction Ignore'
        $script:collectorSource | Should -Match 'Get-Command Get-AzureADConnectHealthSyncAlert -ErrorAction Ignore'
    }

    It 'safely parses privileged admin last sign-in timestamps without null cast noise' {
        $script:collectorSource | Should -Match '\[datetime\]::TryParse\(\$lastSignInText, \[ref\]\$parsedLastSignIn\)'
    }

    It 'safely parses guest last sign-in timestamps and suppresses expected DNS / MSCommerce noise' {
        $script:collectorSource | Should -Match '\$selfServiceErrorCountBeforeLookup = \$global:Error\.Count'
        $script:collectorSource | Should -Match 'while \(\$global:Error\.Count -gt \$selfServiceErrorCountBeforeLookup\)'
        $script:collectorSource | Should -Match 'Resolve-DnsName -Name \$domainName -Server 1\.1\.1\.1 -Type A -ErrorAction Ignore'
        $script:collectorSource | Should -Match 'Resolve-DnsName -Name \("\{0\}\._domainkey\.\{1\}" -f \$selector, \$domainName\) -Server 1\.1\.1\.1 -Type CNAME -ErrorAction Ignore'
    }

    It 'guards version-specific B2B and Teams activity calls with supported command names and cleanup' {
        $script:collectorSource | Should -Match 'Get-Command -Name ''Get-MgPolicyB2BManagementPolicy'' -ErrorAction Ignore'
        $script:collectorSource | Should -Match '\$b2bPolicyErrorCountBeforeLookup = \$global:Error\.Count'
        $script:collectorSource | Should -Match 'while \(\$global:Error\.Count -gt \$b2bPolicyErrorCountBeforeLookup\)'
        $script:collectorSource | Should -Match '-ServiceName ''TeamsUser'''
    }
}
