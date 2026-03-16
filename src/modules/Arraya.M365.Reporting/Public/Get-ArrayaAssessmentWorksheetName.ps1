function Get-ArrayaAssessmentWorksheetName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Anchor
    )

    switch ([string]$Anchor) {
        'licenses' { return 'LicenseSKUs' }
        'domains' { return 'Domains' }
        'domains-dns' { return 'Domains' }
        'identity-admins' { return 'Users' }
        'mailboxes' { return 'MailboxFullDetails' }
        'compliance-retention' { return 'MailboxFullDetails' }
        'inactive-mailboxes' { return 'InactiveMailboxDetails' }
        'sharepoint-onedrive' { return 'SharePoint / OneDrive' }
        'devices' { return 'DeviceDetails' }
        'ad-connect' { return 'AdConnectConfiguration' }
        'conditional-access-mfa' { return 'ConditionalAccessPolicies' }
        'exchange-hybrid' { return 'HybridConfiguration' }
        'cross-tenant-access' { return 'FederationConfiguration' }
        'secure-score' { return 'SecureScoreActions' }
        'ownership-governance' { return 'UnmanagedObjects' }
        'employee-experience-insights' { return 'EmployeeExpInsights' }
        'teams-collaboration' { return 'AllTeams' }
        default { return 'N/A' }
    }
}
