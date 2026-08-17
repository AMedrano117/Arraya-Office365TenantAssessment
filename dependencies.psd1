@{
    # Single source of truth for every PowerShell module this repository depends on.
    #
    # Version ranges are deliberate: MinimumVersion is the oldest version the assessment has
    # been validated against, and MaximumVersion pins the major line so an upstream breaking
    # release cannot silently change collection behaviour on a customer engagement.
    #
    # Scopes:
    #   Runtime    - required to execute a tenant assessment
    #   Onboarding - required to create the assessment app registration
    #   DevTools   - required to lint and test the repository
    #
    # Install with: ./tools/install-dependencies.ps1 -Scope All
    # Verify with:  ./tools/install-dependencies.ps1 -Scope All -Validate

    PowerShellVersion = '7.0'

    Modules = @(
        # --- Microsoft Graph PowerShell SDK (v2 line) -------------------------------------
        @{
            Name           = 'Microsoft.Graph.Authentication'
            MinimumVersion = '2.36.0'
            MaximumVersion = '2.99.99'
            Scope          = @('Runtime', 'Onboarding')
            Reason         = 'Connect-MgGraph, Invoke-MgGraphRequest, Get-MgContext.'
        },
        @{
            Name           = 'Microsoft.Graph.Applications'
            MinimumVersion = '2.36.0'
            MaximumVersion = '2.99.99'
            Scope          = @('Runtime', 'Onboarding')
            Reason         = 'App registration and service principal cmdlets used by onboarding and Entra app reporting.'
        },
        @{
            Name           = 'Microsoft.Graph.Identity.DirectoryManagement'
            MinimumVersion = '2.36.0'
            MaximumVersion = '2.99.99'
            Scope          = @('Runtime', 'Onboarding')
            Reason         = 'Get-MgOrganization, Get-MgDomain, Get-MgDirectoryRole, Get-MgSubscribedSku.'
        },
        @{
            Name           = 'Microsoft.Graph.Identity.Governance'
            MinimumVersion = '2.36.0'
            MaximumVersion = '2.99.99'
            Scope          = @('Runtime', 'Onboarding')
            Reason         = 'Directory role assignment cmdlets used to grant and audit assessment access.'
        },
        @{
            Name           = 'Microsoft.Graph.Identity.SignIns'
            MinimumVersion = '2.36.0'
            MaximumVersion = '2.99.99'
            Scope          = @('Runtime')
            Reason         = 'Conditional Access and cross-tenant access policy collection.'
        },
        @{
            Name           = 'Microsoft.Graph.Users'
            MinimumVersion = '2.36.0'
            MaximumVersion = '2.99.99'
            Scope          = @('Runtime')
            Reason         = 'Get-MgUser inventory collection.'
        },
        @{
            Name           = 'Microsoft.Graph.Reports'
            MinimumVersion = '2.36.0'
            MaximumVersion = '2.99.99'
            Scope          = @('Runtime')
            Reason         = 'Sign-in audit logs and MFA registration detail reports.'
        },
        @{
            Name           = 'Microsoft.Graph.Security'
            MinimumVersion = '2.36.0'
            MaximumVersion = '2.99.99'
            Scope          = @('Runtime')
            Reason         = 'Get-MgSecuritySecureScore.'
        },
        @{
            Name           = 'Microsoft.Graph.Sites'
            MinimumVersion = '2.36.0'
            MaximumVersion = '2.99.99'
            Scope          = @('Runtime')
            Reason         = 'SharePoint site discovery via Graph.'
        },
        @{
            Name           = 'Microsoft.Graph.Teams'
            MinimumVersion = '2.36.0'
            MaximumVersion = '2.99.99'
            Scope          = @('Runtime')
            Reason         = 'Teams discovery via Graph.'
        },
        @{
            Name           = 'Microsoft.Graph.DeviceManagement'
            MinimumVersion = '2.36.0'
            MaximumVersion = '2.99.99'
            Scope          = @('Runtime')
            Reason         = 'Intune managed device collection.'
        },

        # --- Workload-specific modules ----------------------------------------------------
        @{
            Name           = 'ExchangeOnlineManagement'
            MinimumVersion = '3.9.2'
            MaximumVersion = '3.99.99'
            Scope          = @('Runtime')
            Reason         = 'Exchange Online and Security/Compliance collection.'
        },
        @{
            Name           = 'MicrosoftTeams'
            MinimumVersion = '7.6.0'
            MaximumVersion = '7.99.99'
            Scope          = @('Runtime')
            Reason         = 'Teams policy and configuration collection.'
        },
        @{
            Name           = 'Microsoft.Online.SharePoint.PowerShell'
            MinimumVersion = '16.0.27011.12008'
            MaximumVersion = '16.99.99999.99999'
            Scope          = @('Runtime')
            Reason         = 'SharePoint Online tenant settings and site collection inventory.'
        },
        @{
            Name           = 'ImportExcel'
            MinimumVersion = '7.8.10'
            MaximumVersion = '7.99.99'
            Scope          = @('Runtime')
            Reason         = 'Workbook export.'
        },

        # --- Development and CI -----------------------------------------------------------
        @{
            Name           = 'Pester'
            MinimumVersion = '5.7.1'
            MaximumVersion = '5.99.99'
            Scope          = @('DevTools')
            Reason         = 'Test runner. Pester 5 configuration API is required by tools/run-pester.ps1.'
        },
        @{
            Name           = 'PSScriptAnalyzer'
            MinimumVersion = '1.24.0'
            MaximumVersion = '1.99.99'
            Scope          = @('DevTools')
            Reason         = 'Static analysis.'
        }
    )
}
