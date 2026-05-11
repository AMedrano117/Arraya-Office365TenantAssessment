Describe 'Get-FullTenantReportDetails permission preflight' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:collectorPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Get-FullTenantReportDetails.ps1'
        $script:collectorSource = Get-Content -Raw -Path $script:collectorPath
        $script:exportPipelinePath = Join-Path $script:repoRoot 'src\scripts\reporting\Invoke-M365TenantAssessmentExportPipeline.ps1'
        $script:exportPipelineSource = Get-Content -Raw -Path $script:exportPipelinePath
        $script:graphDataPath = Join-Path $script:repoRoot 'src\vendor\Office365Custom\1.2.1\Public\Get-GraphData.ps1'
        $script:graphDataSource = Get-Content -Raw -Path $script:graphDataPath
        $script:entraGroupsPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Graph\Public\Get-EntraIDGroups.ps1'
        $script:entraGroupsSource = Get-Content -Raw -Path $script:entraGroupsPath
    }

    It 'defines a permission preflight with an explicit skip switch and invokes it before collection starts by default' {
        Test-Path $script:collectorPath | Should -BeTrue
        $script:collectorSource | Should -Match 'function Test-AssessmentImportedModuleMatchesManifestPath'
        $script:collectorSource | Should -Match 'function Test-AssessmentPermissionPreflight'
        $script:collectorSource | Should -Match 'function Invoke-AssessmentPermissionPreflightWithStatus'
        $script:collectorSource | Should -Match '\[switch\]\$PreflightOnly'
        $script:collectorSource | Should -Match '\[switch\]\$SkipPermissionPreflight'
        $script:collectorSource | Should -Match 'Permission preflight failed\. The assessment will not continue'
        $script:collectorSource | Should -Match 'Permission preflight warnings:'
        $script:collectorSource | Should -Match 'Invoke-AssessmentPermissionPreflightWithStatus -ConnectionResult \(\[pscustomobject\]\$authResult\) -Workload Graph'
        $script:collectorSource | Should -Match 'Invoke-AssessmentPermissionPreflightWithStatus -ConnectionResult \(\[pscustomobject\]\$authResult\) -Workload ExchangeOnline'
        $script:collectorSource | Should -Match 'Invoke-AssessmentPermissionPreflightWithStatus -ConnectionResult \(\[pscustomobject\]\$authResult\) -Workload Purview'
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

    It 'requires the assessment snapshot JSON before allowing the export stage to finish' {
        $script:collectorSource | Should -Match '\$requiresAssessmentSnapshotArtifact = \(-not \$effectiveSkipJsonReport\)'
        $script:collectorSource | Should -Match "'Assessment Snapshot JSON', 'JSON'"
        $script:collectorSource | Should -Match 'The assessment workbook export completed, but the required assessment snapshot JSON was not produced'
        $script:collectorSource | Should -Match 'Improve, export replay, and manifest-based follow-up cannot continue without that snapshot'
        $script:collectorSource | Should -Match '\$generatedArtifacts\.Contains\(''Manifest''\)'
        $script:collectorSource | Should -Match 'The assessment snapshot JSON was produced, but the required run manifest was not produced'
        $script:collectorSource | Should -Match 'Improve and manifest-based follow-up cannot continue without the manifest'
        $script:collectorSource | Should -Match 'if \(\$requiresAssessmentSnapshotArtifact\)'
        $script:collectorSource | Should -Match 'throw \('
    }

    It 'surfaces the underlying JSON snapshot export failure instead of only reporting a missing artifact' {
        $script:exportPipelineSource | Should -Match '\$jsonExportErrorMessage = "Unable to export Tenant Statistics JSON: \$\(\$_.Exception.Message\)"'
        $script:exportPipelineSource | Should -Match '\$generatedArtifacts\[''Assessment Snapshot JSON Error''\] = \$jsonExportErrorMessage'
        $script:exportPipelineSource | Should -Match 'throw \$jsonExportErrorMessage'
    }

    It 'threads workbook policy, migration-pack export, and the T2T cutover HTML mode through the export pipeline' {
        $script:exportPipelineSource | Should -Match 'WorkbookExportPolicy'
        $script:exportPipelineSource | Should -Match 'TechnicalHtmlPolicy'
        $script:exportPipelineSource | Should -Match 'GenerateMigrationPack'
        $script:collectorSource | Should -Match 'function Resolve-AssessmentHumanDeliverableTargetPath'
        $script:collectorSource | Should -Match 'Resolve-AssessmentHumanDeliverableTargetPath -ResolvedExportTargetPath \$resolvedExportTargetPath'
        $script:collectorSource | Should -Match 'Get-ExportPath -FileName \$defaultReportFileName -UserInputPath \$humanDeliverableTargetPath'
        $script:exportPipelineSource | Should -Match 'Export-HashTableToExcel -hashtable \$ExportTenantStatsHash -ExportDetails \$ExportDetails -WorkbookExportPolicy \$WorkbookExportPolicy'
        $script:exportPipelineSource | Should -Match 'Export-ArrayaTenantToTenantCutoverPack -TenantStatsHash \$TenantStatsHash -BaseExportPath \$ExportDetails'
        $script:exportPipelineSource | Should -Match 'New-TenantMigrationCutoverHtmlReport -TenantStatsHash \$TenantStatsHash -OutputPath \$htmlExportPath -CollectionScopePolicy ''TenantToTenantCutover'''
        $script:exportPipelineSource | Should -Match '''T2T Cutover Pack Workbook'''
    }

    It 'uses an assessment-owned auth orchestrator instead of the generic Office365 connector bootstrap' {
        $script:collectorSource | Should -Match 'function Resolve-AssessmentProfileCollectionPlan'
        $script:collectorSource | Should -Match 'function Resolve-AssessmentGraphScopePlan'
        $script:collectorSource | Should -Match 'function Resolve-AssessmentRequestedAuthMode'
        $script:collectorSource | Should -Match 'function Resolve-AssessmentClientSecretAuthContext'
        $script:collectorSource | Should -Match 'function Get-AssessmentGraphDelegatedScopes'
        $script:collectorSource | Should -Match 'function Resolve-AssessmentAuthWorkloadPlan'
        $script:collectorSource | Should -Match 'function Test-IsAssessmentInteractiveTokenAcquisitionFailure'
        $script:collectorSource | Should -Match 'function Get-AssessmentInteractiveTokenFailureOperatorMessage'
        $script:collectorSource | Should -Match 'function Write-AssessmentInteractiveAuthNotice'
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
        $script:collectorSource | Should -Match '\[pscredential\]\$ClientSecretCredential'
        $script:collectorSource | Should -Match '\[securestring\]\$ClientSecretSecure'
        $script:collectorSource | Should -Match 'Get-AssessmentGraphDelegatedScopes -WorkloadPlan \$WorkloadPlan'
        $script:collectorSource | Should -Match 'GraphScopePlan'
        $script:collectorSource | Should -Match 'CollectionScopePolicy'
        $script:collectorSource | Should -Match 'Client secret authentication requires -ClientId, or a PSCredential whose username is the app client ID\.'
        $script:collectorSource | Should -Not -Match "'Connect-Office365',"
    }

    It 'connects Exchange Online before Microsoft Graph for interactive assessment runs' {
        $script:collectorSource | Should -Match '\$connectExchangeFirst = \(\[string\]\$WorkloadPlan\.AuthenticationType -eq ''Interactive''\)'
        $script:collectorSource | Should -Match 'Interactive sign-in sequence: Exchange Online, Purview, then Microsoft Graph\.'
        $script:collectorSource | Should -Match 'if \(\$connectExchangeFirst\) \{[\s\S]*Connect-AssessmentExchange[\s\S]*Invoke-AssessmentPermissionPreflightWithStatus -ConnectionResult \(\[pscustomobject\]\$authResult\) -Workload ExchangeOnline[\s\S]*if \(\$WorkloadPlan\.Workloads\.PurviewCompliance\.Required\) \{[\s\S]*Connect-AssessmentPurview[\s\S]*Invoke-AssessmentPermissionPreflightWithStatus -ConnectionResult \(\[pscustomobject\]\$authResult\) -Workload Purview[\s\S]*Connect-AssessmentGraph[\s\S]*Invoke-AssessmentPermissionPreflightWithStatus -ConnectionResult \(\[pscustomobject\]\$authResult\) -Workload Graph'
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
        $script:collectorSource | Should -Match 'NeedsDeviceData'
        $script:collectorSource | Should -Match 'NeedsSecureScore'
        $script:collectorSource | Should -Match 'NeedsReportsData'
    }

    It 'adds a reduced-scope tenant-to-tenant collection policy and uses it to trim collectors and Graph scopes' {
        $script:collectorSource | Should -Match 'CollectionScopePolicy -eq ''TenantToTenantCutover'''
        $script:collectorSource | Should -Match 'CollectDevices'
        $script:collectorSource | Should -Match 'CollectSecuritySecureScore'
        $script:collectorSource | Should -Match 'CollectEntraGroups'
        $script:collectorSource | Should -Match 'CollectEntraGroupLicenseChecks'
        $script:collectorSource | Should -Match 'CollectSharePointAndOneDriveSites'
        $script:collectorSource | Should -Match 'BuildMigrationReadinessTables'
        $script:collectorSource | Should -Match 'Get-ArrayaCollectionDepthPolicy -ReportingMode \(\(Get-Culture\)\.TextInfo\.ToTitleCase\(\$reportingMode\)\) -CollectionScopePolicy \$effectiveCollectionScopePolicy'
        $script:collectorSource | Should -Match 'CollectMailboxDelegatePermissions'
        $script:collectorSource | Should -Match 'CollectMailboxCalendarDelegatePermissions'
        $script:collectorSource | Should -Match '\$plan\.CollectDevices = \$false'
        $script:collectorSource | Should -Match '\$plan\.CollectSecuritySecureScore = \$false'
        $script:collectorSource | Should -Match '\$plan\.CollectEntraGroups = \$false'
        $script:collectorSource | Should -Match '\$plan\.CollectAdmins = \$false'
        $script:collectorSource | Should -Match '\$plan\.CollectAuthenticationConfiguration = \$false'
        $script:collectorSource | Should -Match '\$plan\.CollectConditionalAccessPolicies = \$false'
        $script:collectorSource | Should -Match '\$plan\.CollectMfaRegistrationDetails = \$false'
        $script:collectorSource | Should -Match '\$plan\.CollectUnifiedGroups = \$false'
        $script:collectorSource | Should -Match '\$policy \| Add-Member -MemberType NoteProperty -Name CollectEntraGroupLicenseChecks -Value \$false -Force'
        $script:collectorSource | Should -Match '\$plan\.BuildExternalExposureSummaries = \$false'
        $script:collectorSource | Should -Match '\$plan\.BuildAssessmentReportTables = \$false'
        $script:collectorSource | Should -Match '\$plan\.BuildMigrationReadinessTables = \$true'
        $script:collectorSource | Should -Match 'CollectSsoApplicationDetails -Value \$false -Force'
        $script:collectorSource | Should -Match 'if \(\$needsDeviceData\)'
        $script:collectorSource | Should -Match 'if \(\$needsSecureScore\)'
        $script:collectorSource | Should -Match 'if \(\$needsReportsData\)'
    }

    It 'collects practical licensing signals for full assessment governance analysis' {
        $script:collectorSource | Should -Match '"DisplayName", "AssignedLicenses", "LicenseAssignmentStates", "UserPrincipalName"'
        $script:collectorSource | Should -Match 'DirectAssignedLicenses'
        $script:collectorSource | Should -Match 'GroupAssignedLicenses'
        $script:collectorSource | Should -Match 'LicenseAssignmentErrors'
        $script:collectorSource | Should -Match 'LicenseAssignmentStateSummary'
        $script:collectorSource | Should -Match 'LastUpdatedDateTime'

        $script:entraGroupsSource | Should -Match 'CollectEntraGroupLicenseChecks'
        $script:entraGroupsSource | Should -Match 'AssignedLicenseSkuIds'
        $script:entraGroupsSource | Should -Match 'AssignedLicenseSkuPartNumbers'
        $script:entraGroupsSource | Should -Match 'AssignedLicenseFriendlyNames'
        $script:entraGroupsSource | Should -Match 'assignedLicenses/any\(\)'
        $script:entraGroupsSource | Should -Match 'LicenseProcessingState'
        $script:entraGroupsSource | Should -Match 'Get-EntraLicenseGroupDirectMemberCount'
        $script:entraGroupsSource | Should -Match 'Direct member enumeration for license-managing group'
        $script:entraGroupsSource | Should -Match '\$assignedLicenseValue = Get-ArrayaObjectValue'
        $script:entraGroupsSource | Should -Match '\$assignedLicenseSkuIds\.Count -gt 0'
    }

    It 'builds a dedicated action-only T2T migration readiness checklist dataset with split mail-flow and mailbox-prep rows' {
        $script:collectorSource | Should -Match 'MigrationReadinessChecklist'
        $script:collectorSource | Should -Match 'Mail flow connectors'
        $script:collectorSource | Should -Match 'Remote domains allowing auto-forwarding'
        $script:collectorSource | Should -Match 'SMTP relay service accounts'
        $script:collectorSource | Should -Match 'Mailbox routing and proxy attributes'
        $script:collectorSource | Should -Match 'Mailbox delegate reapplication'
        $script:collectorSource | Should -Match 'Oversized mailbox batches'
        $script:collectorSource | Should -Match 'Oversized archives'
        $script:collectorSource | Should -Match 'Teams with shared channels'
        $script:collectorSource | Should -Match 'Collaboration sizing gaps'
        $script:collectorSource | Should -Match '\[string\]\$Status -in @\(''Ready'', ''Info''\)'
    }

    It 'validates only the required workload sessions when SkipAuth is used' {
        $script:collectorSource | Should -Match 'Reusing existing workload sessions for this run'
        $script:collectorSource | Should -Match 'SkipAuth was requested, but no existing Microsoft Graph session was found'
        $script:collectorSource | Should -Match 'SkipAuth was requested, but Exchange Online cmdlets are not available in the current session'
        $script:collectorSource | Should -Match 'function Test-AssessmentPurviewSessionReady'
        $script:collectorSource | Should -Match 'function Test-AssessmentSharePointSessionReady'
        $script:collectorSource | Should -Match 'function Test-AssessmentTeamsSessionReady'
        $script:collectorSource | Should -Match 'SkipAuth was requested, but Purview compliance session is not usable in the current session'
        $script:collectorSource | Should -Match 'Test-AssessmentExistingSessions -WorkloadPlan \$WorkloadPlan'
        $script:collectorSource | Should -Match 'SkipAuth was requested and no existing SharePoint admin session was found\. Graph fallback remains active'
        $script:collectorSource | Should -Match 'SkipAuth was requested and no existing Teams PowerShell session was found\. Graph-only Teams collection remains active'
    }

    It 'validates tenant alignment before reusing Graph and Exchange sessions and removes the public tenant lookup fallback' {
        $script:collectorSource | Should -Match 'function Get-AssessmentGraphOrganizationDetails'
        $script:collectorSource | Should -Match 'function Assert-AssessmentGraphContextMatchesTenant'
        $script:collectorSource | Should -Match 'function Resolve-AssessmentValidationInitialDomain'
        $script:collectorSource | Should -Match 'function Assert-AssessmentExchangeSessionMatchesTenant'
        $script:collectorSource | Should -Match 'Get-AcceptedDomain -Identity \$validationDomain'
        $script:collectorSource | Should -Match 'Test-AssessmentExistingSessions -WorkloadPlan \$WorkloadPlan -TenantId \$TenantId'
        $script:collectorSource | Should -Match 'Connect-AssessmentExchange -AuthenticationType \$WorkloadPlan\.AuthenticationType -TenantId \$TenantId'
        $script:collectorSource | Should -Match 'Disconnect-MgGraph -ErrorAction SilentlyContinue'
        $script:collectorSource | Should -Not -Match 'tenantinfoapp\.azurewebsites\.us'
        $script:collectorSource | Should -Not -Match 'beta/tenantRelationships/findTenantInformationByTenantId'
    }

    It 'requests SharePoint Graph scopes for interactive auth and gives cached-token guidance when Graph preflight fails' {
        $script:collectorSource | Should -Match "Get-AssessmentGraphDelegatedScopes"
        $script:collectorSource | Should -Match '\$scopes\.Add\(''Sites\.Read\.All''\)'
        $script:collectorSource | Should -Match '\$scopes\.Add\(''SharePointTenantSettings\.Read\.All''\)'
        $script:collectorSource | Should -Match 'Graph scopes requested:'
        $script:collectorSource | Should -Match 'Requested Microsoft Graph delegated scopes:'
        $script:collectorSource | Should -Match 'function Get-AssessmentGraphPreflightOperatorGuidance'
        $script:collectorSource | Should -Match 'Run Disconnect-MgGraph, then rerun the assessment'
        $script:collectorSource | Should -Match 'sign in as a Global Administrator'
        $script:collectorSource | Should -Match 'interactive preflight checks the delegated-safe /sites/root endpoint'
        $script:collectorSource | Should -Match 'app-only preflight checks /sites/getAllSites'
    }

    It 'restores Graph globals after the run and labels remaining beta fallbacks explicitly' {
        $script:collectorSource | Should -Match '\$script:AssessmentGraphRuntimeState = \[ordered\]@'
        $script:collectorSource | Should -Match 'function Restore-AssessmentGraphRuntimeState'
        $script:collectorSource | Should -Match 'finally \{\s*Restore-AssessmentGraphRuntimeState'
        $script:collectorSource | Should -Match 'Enterprise application service principal sign-in activities \(beta best-effort\)'
        $script:collectorSource | Should -Match 'beta best-effort fallback'
        $script:collectorSource | Should -Match 'Application activity state will remain not validated where this enrichment was required'
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
        $script:collectorSource | Should -Match 'function New-AssessmentCollectorSections'
        $script:collectorSource | Should -Match "\[pscustomobject\]@\{ Step = '1/6'; Name = 'Tenant Overview' \}"
        $script:collectorSource | Should -Match "\[pscustomobject\]@\{ Step = '2/6'; Name = 'Identity' \}"
        $script:collectorSource | Should -Match "\[pscustomobject\]@\{ Step = '3/6'; Name = 'Exchange' \}"
        $script:collectorSource | Should -Match "\[pscustomobject\]@\{ Step = '4/6'; Name = 'Collaboration' \}"
        $script:collectorSource | Should -Match "\[pscustomobject\]@\{ Step = '5/6'; Name = 'Endpoint' \}"
        $script:collectorSource | Should -Match "\[pscustomobject\]@\{ Step = '6/6'; Name = 'Governance' \}"
        $script:collectorSource | Should -Match 'Invoke-ArrayaCollectorPlan'
        $script:collectorSource | Should -Match 'Write-ConsoleSection -Step \(\[string\]\$Section\.Step\) -Title \(\[string\]\$Section\.Name\)'
        $script:collectorSource | Should -Match "Write-ConsoleSection -Step 'Export' -Title 'Exporting results'"
        $script:collectorSource | Should -Not -Match 'Consolidating Discovery Report'
    }

    It 'moves license collection into Tenant Overview and removes combined mailbox reporting from Exchange collection' {
        $script:collectorSource | Should -Match "New-ArrayaCollectorStep -Name 'License SKUs' -Section 'Tenant Overview'"
        $script:collectorSource | Should -Not -Match "New-ArrayaCollectorStep -Name 'License SKUs' -Section 'Identity'"
        $script:collectorSource | Should -Not -Match "Invoke-ProfileAwareAssessmentStep -Name 'Combined user/mailbox reporting'"
        $script:collectorSource | Should -Match 'function Prepare-AssessmentExportData'
        $script:collectorSource | Should -Match 'Export preparation: combined user/mailbox reporting'
    }

    It 'keeps top-level progress labels clean while surfacing key live substeps for long-running user collection' {
        $script:collectorSource | Should -Not -Match 'Write-Host "Gathering Tenant Overview Info \.\.\."'
        $script:collectorSource | Should -Match "Write-AssessmentCollectorBanner -Message 'Gathering Tenant Overview Info \.\.\.' -NoNewline"
        $script:collectorSource | Should -Match "Write-AssessmentConsoleSubstep -Message 'Users: Graph inventory retrieval and per-user enrichment'"
        $script:collectorSource | Should -Match 'Users: processed \{0\} total directory records'
        $script:collectorSource | Should -Not -Match '\$consoleProgressInterval = 500'
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
        $script:collectorSource | Should -Match 'https://graph\.microsoft\.com/v1\.0/sites/root\?\$select=id,webUrl,displayName'
        $script:collectorSource | Should -Match 'https://graph\.microsoft\.com/v1\.0/sites/getAllSites\?\$top=1'
        $script:collectorSource | Should -Match '\$isInteractiveGraphAuth'
    }

    It 'handles SharePoint getAllSites access denied with operator guidance and SPO fallback when available' {
        $script:collectorSource | Should -Match 'Microsoft Graph getAllSites requires application permissions'
        $script:collectorSource | Should -Match 'Confirm the app has Sites.Read.All application permission with admin consent'
        $script:collectorSource | Should -Match 'falling back to connected SharePoint Online PowerShell session'
        $script:collectorSource | Should -Match "SharePointCollectionSummary"
        $script:collectorSource | Should -Match 'Site usage report coverage unavailable because no SharePoint/OneDrive site inventory rows were collected'
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
        $script:collectorSource | Should -Match 'Connect-IPPSSession does not expose device code in this ExchangeOnlineManagement version'
        $script:collectorSource | Should -Match 'Purview auth: interactive sign-in hit a Windows broker / WAM token issue, retrying with -DisableWAM'
        $script:collectorSource | Should -Match 'Purview auth: interactive sign-in hit a Windows broker / WAM token issue, retrying with device code'
        $script:collectorSource | Should -Match 'Client secret authentication is not supported for Purview compliance PowerShell in this workflow'
        $script:collectorSource | Should -Match 'Purview compliance PowerShell session could not be established\. Underlying error:'
        $script:collectorSource | Should -Match 'Purview compliance PowerShell interactive sign-in hit a Windows broker / WAM token acquisition failure\. Underlying error:'
        $script:collectorSource | Should -Match 'Organization used:'
        $script:collectorSource | Should -Match 'Next step:'
        $script:collectorSource | Should -Match 'Missing compliance cmdlets after connect:'
        $script:collectorSource | Should -Match 'The certificate and app registration were accepted, but this tenant did not expose the Purview retention/DLP cmdlets to that app session'
        $script:collectorSource | Should -Match 'Exchange Administrator role'
        $script:collectorSource | Should -Match 'Connect-IPPSSession successfully in the current PowerShell session and rerun the assessment with session reuse'
        $script:collectorSource | Should -Match 'Interactive Purview sign-in hit a Windows broker / WAM token acquisition failure'
    }

    It 'turns Graph interactive broker token failures into explicit operator guidance before and after the device-code fallback' {
        $script:collectorSource | Should -Match 'Microsoft Graph interactive sign-in'
        $script:collectorSource | Should -Match 'Windows broker / WAM token acquisition failure'
        $script:collectorSource | Should -Match 'Write-AssessmentInteractiveAuthNotice -ServiceName ''Microsoft Graph'''
        $script:collectorSource | Should -Match 'browser sign-in prompt should appear'
        $script:collectorSource | Should -Match 'Graph auth: device code prompt should appear in this console'
        $script:collectorSource | Should -Match 'Retrying Microsoft Graph sign-in with device code'
        $script:collectorSource | Should -Match 'The follow-up device code sign-in did not complete'
        $script:collectorSource | Should -Match 'try again from a fresh PowerShell window'
        $script:collectorSource | Should -Match 'certificate-based authentication'
    }

    It 'makes the broader interactive connection path operator-friendly across Exchange, SharePoint, Teams, and Purview' {
        $script:collectorSource | Should -Match 'Write-AssessmentInteractiveAuthNotice -ServiceName ''Exchange Online'''
        $script:collectorSource | Should -Match 'Exchange auth: broker / WAM sign-in failed, retrying with -DisableWAM'
        $script:collectorSource | Should -Match 'Exchange auth: broker / WAM sign-in failed, retrying with device code'
        $script:collectorSource | Should -Match 'Exchange Online interactive sign-in hit a Windows broker / WAM token acquisition failure'
        $script:collectorSource | Should -Match 'SharePoint admin already connected for this session.'
        $script:collectorSource | Should -Match 'Write-AssessmentInteractiveAuthNotice -ServiceName ''SharePoint admin'''
        $script:collectorSource | Should -Match 'SharePoint admin interactive sign-in'
        $script:collectorSource | Should -Match 'Graph fallback remains active for SharePoint data in this run'
        $script:collectorSource | Should -Match 'Teams PowerShell already connected for this session.'
        $script:collectorSource | Should -Match 'Write-AssessmentInteractiveAuthNotice -ServiceName ''Teams PowerShell'''
        $script:collectorSource | Should -Match 'Teams interactive sign-in'
        $script:collectorSource | Should -Match 'Graph-only Teams collection remains active for this run'
        $script:collectorSource | Should -Match 'Write-AssessmentInteractiveAuthNotice -ServiceName ''Purview compliance PowerShell'''
        $script:collectorSource | Should -Match 'Write-AssessmentInteractiveAuthNotice -ServiceName ''Purview compliance PowerShell'' -SupportsDisableWam:\$ConnectCommandMetadata\.Parameters\.ContainsKey\(''DisableWAM''\) -SupportsDeviceCode:\$ConnectCommandMetadata\.Parameters\.ContainsKey\(''Device''\)'
        $script:collectorSource | Should -Not -Match 'Invoke-PurviewComplianceDelegatedConnect -BaseParameters \$connectParams -ConnectCommandMetadata \$connectCommand -GraphAccount \$graphAccount'
    }

    It 'suppresses noisy SharePoint and Teams module import warnings during runtime' {
        $script:collectorSource | Should -Match 'function Ensure-AssessmentSharePointModuleAvailable'
        $script:collectorSource | Should -Match 'function Ensure-AssessmentTeamsModuleAvailable'
        $script:collectorSource | Should -Match "Import-Module 'Microsoft\.Online\.SharePoint\.PowerShell' -UseWindowsPowerShell -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop"
        $script:collectorSource | Should -Match "Import-Module 'Microsoft\.Online\.SharePoint\.PowerShell' -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop"
        $script:collectorSource | Should -Match "Import-Module 'MicrosoftTeams' -UseWindowsPowerShell -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop"
        $script:collectorSource | Should -Match "Import-Module 'MicrosoftTeams' -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop"
    }

    It 'treats directory synchronization feature access as a non-blocking validation warning' {
        $script:collectorSource | Should -Match 'OnPremDirectorySynchronization\.Read\.All'
        $script:collectorSource | Should -Match 'IsBlocking\s*=\s*\$false'
        $script:collectorSource | Should -Match 'Hybrid sync and password lifecycle fields may be marked as not validated in current auth mode'
    }

    It 'uses a stable preflight progress counter and suppresses inner Graph record-count progress' {
        $script:collectorSource | Should -Match 'Checking \{0\} access\.\.\.'
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
        $script:collectorSource | Should -Match 'GuestCoveredPolicyNames'
        $script:collectorSource | Should -Match 'GuestExplicitScopePolicyNames'
        $script:collectorSource | Should -Match 'GuestMfaEnforcementSources'
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

    It 'surfaces guest MFA enforcement details in the guest access summary' {
        $script:collectorSource | Should -Match 'GuestMfaEnforcementState'
        $script:collectorSource | Should -Match 'GuestMfaEnforcementSources'
        $script:collectorSource | Should -Match 'GuestMfaConditionalAccessPolicies'
        $script:collectorSource | Should -Match 'GuestMfaExplicitScopePolicies'
    }

    It 'batches enterprise application sign-in enrichment instead of making one Graph call per app' {
        $script:collectorSource | Should -Match 'function Get-AssessmentEnterpriseApplicationLatestSignInMap'
        $script:collectorSource | Should -Match 'Some enterprise application sign-in lookups exceeded the reviewed sign-in history window'
        $script:collectorSource | Should -Match 'Unable to retrieve batched enterprise application sign-in details'
        $script:collectorSource | Should -Not -Match 'Get-AssessmentEnterpriseApplicationLatestSignIn -AppId'
    }

    It 'uses the shared run-scoped Graph request cache for high-volume Graph collection paths' {
        $script:collectorSource | Should -Match 'function Get-ArrayaGraphResource'
        $script:collectorSource | Should -Match 'Invoke-ArrayaGraphCollectionRequest `'
        $script:collectorSource | Should -Match 'Secure Score most recent'
        $script:collectorSource | Should -Match "Get-ArrayaGraphResource -Uri 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings'"
        $script:collectorSource | Should -Match 'Get-ArrayaGraphResource -Uri \$uri -Activity "Fetching MFA Registration Details"'
        $script:collectorSource | Should -Match '\$script:tenantStatsHash\[''CollectorGraphApiStats''\]'
        $script:collectorSource | Should -Match '\$script:tenantStatsHash\[''CollectorCacheStats''\]'
    }

    It 'caches Graph report CSV downloads and preserves populated runtime caches during script context sync' {
        $script:collectorSource | Should -Match 'function Export-ArrayaGraphReportCsv'
        $script:collectorSource | Should -Match 'GraphReportCsv:\$Uri'
        $script:collectorSource | Should -Match 'Get-ArrayaCollectorCacheValue -Context \$context -Key \$cacheKey'
        $script:collectorSource | Should -Match 'Set-ArrayaCollectorCacheValue -Context \$context -Key \$cacheKey -Value \$rows'
        $script:collectorSource | Should -Match '\$script:MailboxUsageGraphLookup -is \[System\.Collections\.IDictionary\] -and \$script:MailboxUsageGraphLookup\.Count -gt 0'
        $script:collectorSource | Should -Match '\$script:UnifiedGroupsInventoryCache -and @\(\$script:UnifiedGroupsInventoryCache\)\.Count -gt 0'
        $script:collectorSource | Should -Match '\$script:Office365GroupsActivityMailboxLookup\.PSObject\.Properties\[''ByGroupId''\]'
    }

    It 'emits compact one-line assessment step status output instead of the older redundant progress pair' {
        $script:collectorSource | Should -Not -Match 'Write-Host \("Gathering \{0\} \.\.\." -f \$Name\)'
        $script:collectorSource | Should -Not -Match 'Overall progress:'
        $script:collectorSource | Should -Match 'Write-AssessmentCollectorCompletionBanner'
        $script:collectorSource | Should -Match 'if \(\$script:AssessmentProgressState\) \{\s*return\s*\}'
        $script:collectorSource | Should -Match '\[\{0\}/\{1\} \| \{2\}%\] \{3\} - \{4\} in \{5\}\{6\}'
    }

    It 'shows visible substeps for long-running group, governance, and export preparation work without verbose forwarding checkpoint spam' {
        $script:collectorSource | Should -Match 'Exchange governance: shared mailbox review'
        $script:collectorSource | Should -Match 'Exchange governance: forwarding policy review'
        $script:collectorSource | Should -Match 'Exchange governance: inbox rule forwarding review'
        $script:collectorSource | Should -Match 'Exchange governance: shared mailbox review completed'
        $script:collectorSource | Should -Match 'Exchange governance: forwarding policy review completed'
        $script:collectorSource | Should -Match 'Reviewing mailbox inbox rules for external forwarding'
        $script:collectorSource | Should -Match 'Mailbox \{0\}/\{1\}: \{2\} \| External rules \{3\} \| Lookup failures \{4\}'
        $script:collectorSource | Should -Not -Match 'Exchange governance: inbox rule review progress'
        $script:collectorSource | Should -Not -Match 'Exchange governance: slow inbox rule lookup'
        $script:collectorSource | Should -Match 'Export preparation: user and mailbox detail projection'
        $script:collectorSource | Should -Match 'Export preparation: inactive mailbox detail projection'
        $script:collectorSource | Should -Match 'Users: tenant license lookup unavailable, continuing without license and sign-in activity enrichment'
        $script:collectorSource | Should -Match 'Device management summaries: building compliance and supportability rollup'
        $script:collectorSource | Should -Match 'Tier B operational summaries: reviewing SharePoint tenant settings from Microsoft Graph'
        $script:collectorSource | Should -Match 'Tier B operational summaries: completed'
        $script:collectorSource | Should -Match 'External exposure summaries: reviewing sharing baseline and site override signals'
        $script:collectorSource | Should -Match 'External exposure summaries: completed'
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
        $script:collectorSource | Should -Match 'ApplicationSource'
        $script:collectorSource | Should -Match 'ApplicationSourceState'
        $script:collectorSource | Should -Match 'UnknownSourceApplications'
        $script:collectorSource | Should -Match 'OwnerCount'
        $script:collectorSource | Should -Match 'OwnerSignalState'
        $script:collectorSource | Should -Match 'RedirectUris'
        $script:collectorSource | Should -Match 'RedirectUriSignalState'
        $script:collectorSource | Should -Match 'InsecureRedirectUriCount'
        $script:collectorSource | Should -Match 'HasInsecureRedirectUris'
        $script:collectorSource | Should -Match 'AppCredentials'
        $script:collectorSource | Should -Match 'DelegatedLastSignIn'
        $script:collectorSource | Should -Match 'ApplicationLastSignIn'
        $script:collectorSource | Should -Match 'LastServicePrincipalSignInDateTime'
        $script:collectorSource | Should -Match 'ActivitySignalState'
        $script:collectorSource | Should -Match 'HasRecentActivity'
        $script:collectorSource | Should -Match 'ApplicationSourceCoverageState'
        $script:collectorSource | Should -Match 'OwnerSignalCoverageState'
        $script:collectorSource | Should -Match 'RedirectUriSignalCoverageState'
        $script:collectorSource | Should -Match 'ActivitySignalCoverageState'
        $script:collectorSource | Should -Match 'ApplicationsWithInsecureRedirectUris'
        $script:collectorSource | Should -Match 'FirstPartyAppsWithoutOwners'
        $script:collectorSource | Should -Match 'ThirdPartyAppsWithApplicationPerms'
        $script:collectorSource | Should -Match 'ApplicationsWithNoRecentActivity'
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

    It 'safely parses guest last sign-in timestamps and suppresses expected DNS / MSCommerce noise while preferring native MSCommerce imports' {
        $script:collectorSource | Should -Match "\$effectiveAuthenticationType -ne 'Interactive'"
        $script:collectorSource | Should -Match 'MSCommerce self-service purchase review is skipped for \$effectiveAuthenticationType assessment runs'
        $script:collectorSource | Should -Match '\$selfServiceErrorCountBeforeLookup = \$global:Error\.Count'
        $script:collectorSource | Should -Match 'while \(\$global:Error\.Count -gt \$selfServiceErrorCountBeforeLookup\)'
        $script:collectorSource | Should -Match '\$msCommerceModule = Get-Module -ListAvailable -Name MSCommerce \| Sort-Object Version -Descending \| Select-Object -First 1'
        $script:collectorSource | Should -Match '\$msCommerceImportTarget = if \(-not \[string\]::IsNullOrWhiteSpace\(\[string\]\$msCommerceModule\.Path\)\)'
        $script:collectorSource | Should -Match 'Import-Module -Name \$msCommerceImportTarget -ErrorAction Stop'
        $script:collectorSource | Should -Match 'Import-Module MSCommerce -UseWindowsPowerShell -ErrorAction Stop'
        $script:collectorSource | Should -Match 'Resolve-DnsName -Name \$domainName -Server 1\.1\.1\.1 -Type A -ErrorAction Ignore'
        $script:collectorSource | Should -Match 'Resolve-DnsName -Name \("\{0\}\._domainkey\.\{1\}" -f \$selector, \$domainName\) -Server 1\.1\.1\.1 -Type CNAME -ErrorAction Ignore'
    }

    It 'keeps MFA enforcement collection resilient when admin enrichment or flattened CA rows are imperfect' {
        $script:collectorSource | Should -Match 'Add-AssessmentNotePropertyIfPossible'
        $script:collectorSource | Should -Match 'Admin MFA review enrichment did not complete, but MFA enforcement summary data was retained'
        $script:collectorSource | Should -Match 'RequiresMfaEnforcement'
        $script:collectorSource | Should -Match 'UsesAuthenticationStrengthForMfa'
        $script:collectorSource | Should -Match 'guestCoveredPolicyNameLookup'
        $script:collectorSource | Should -Match 'guestExplicitScopePolicyNameLookup'
    }

    It 'avoids injecting unsupported $top paging into directory role expansion endpoints' {
        $script:collectorSource | Should -Match 'directoryRoles\?\`\$select=id,displayName,roleTemplateId'' -PageSize 0'
        $script:collectorSource | Should -Match 'members/microsoft\.graph\.user\?\`\$select=id'
        $script:collectorSource | Should -Match 'members/microsoft\.graph\.user\?\`\$select=id" -f \$directoryRoleId\) -PageSize 0'
        $script:collectorSource | Should -Match 'resolvedRoleTemplateIds'
        $script:collectorSource | Should -Match '\(\?i\)\(all\|none\|null\|\\\[\\\]\|\\\{\\\}\)'
    }

    It 'guards version-specific B2B and Teams activity calls with supported command names and cleanup' {
        $script:collectorSource | Should -Match 'Get-Command -Name ''Get-MgPolicyB2BManagementPolicy'' -ErrorAction Ignore'
        $script:collectorSource | Should -Match '\$b2bPolicyErrorCountBeforeLookup = \$global:Error\.Count'
        $script:collectorSource | Should -Match 'while \(\$global:Error\.Count -gt \$b2bPolicyErrorCountBeforeLookup\)'
        $script:collectorSource | Should -Match '-ServiceName ''TeamsUser'''
    }
}
