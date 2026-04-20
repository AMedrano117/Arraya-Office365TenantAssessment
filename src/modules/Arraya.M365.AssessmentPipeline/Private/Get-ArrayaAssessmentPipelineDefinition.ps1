function Get-ArrayaAssessmentPipelineDefinition {
    [CmdletBinding()]
    param()

    return @(
        [pscustomobject][ordered]@{
            Name            = 'Connection'
            Order           = 1
            DependsOn       = @()
            EmitsDataKeys   = @()
            Checkpointed    = $true
            Description     = 'Authenticates required workloads and records connection readiness.'
        }
        [pscustomobject][ordered]@{
            Name            = 'TenantOverview'
            Order           = 2
            DependsOn       = @('Connection')
            EmitsDataKeys   = @('TenantInfo', 'LicenseSKUs', 'AdConnectConfiguration')
            Checkpointed    = $true
            Description     = 'Collects tenant-wide summary and baseline identity metadata.'
        }
        [pscustomobject][ordered]@{
            Name            = 'Identity'
            Order           = 3
            DependsOn       = @('TenantOverview')
            EmitsDataKeys   = @('Users', 'Admins', 'EntraGroups', 'Domains', 'AuthenticationConfiguration', 'FederationConfiguration', 'ConditionalAccess', 'MfaRegistrationDetails')
            Checkpointed    = $true
            Description     = 'Collects Entra identity, authentication, and access policy data.'
        }
        [pscustomobject][ordered]@{
            Name            = 'Exchange'
            Order           = 4
            DependsOn       = @('Identity')
            EmitsDataKeys   = @('AllRecipients', 'MailboxFullDetails', 'ExchangeGroups', 'PublicFolders', 'HybridConfiguration', 'MailFlowConnectors', 'MailFlowRules', 'EmailActivity', 'ExchangeGovernanceSummaries')
            Checkpointed    = $true
            Description     = 'Collects Exchange inventory, hybrid, and mail-flow details.'
        }
        [pscustomobject][ordered]@{
            Name            = 'Collaboration'
            Order           = 5
            DependsOn       = @('Exchange')
            EmitsDataKeys   = @('UnifiedGroups', 'SharePoint', 'OneDrive', 'AllTeams', 'TeamsVoice')
            Checkpointed    = $true
            Description     = 'Collects Teams, SharePoint, OneDrive, and collaboration inventory.'
        }
        [pscustomobject][ordered]@{
            Name            = 'Endpoint'
            Order           = 6
            DependsOn       = @('Collaboration')
            EmitsDataKeys   = @('DeviceDetails', 'DeviceManagementSummary')
            Checkpointed    = $true
            Description     = 'Collects device and endpoint management details.'
        }
        [pscustomobject][ordered]@{
            Name            = 'Governance'
            Order           = 7
            DependsOn       = @('Endpoint')
            EmitsDataKeys   = @('SecuritySecureScore', 'PurviewCompliancePolicies', 'OperationalGovernanceSummaries', 'ExternalExposureFindings', 'OwnershipGovernance', 'AssessmentReportTables', 'ConfigurationSummaryTables')
            Checkpointed    = $true
            Description     = 'Builds security, governance, and cross-workload posture summaries.'
        }
        [pscustomobject][ordered]@{
            Name            = 'Export'
            Order           = 8
            DependsOn       = @('Governance')
            EmitsDataKeys   = @()
            Checkpointed    = $false
            Description     = 'Exports workbook, snapshot, and related artifacts from the latest checkpoint.'
        }
    )
}
