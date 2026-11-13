@{
    # Single source of truth for every PowerShell module this repository depends on.
    #
    # Version ranges are deliberate: MinimumVersion is the oldest version the assessment has
    # been validated against, and MaximumVersion caps the newest major line that has been
    # reviewed, so an unreviewed breaking release cannot silently change collection behaviour
    # on a customer engagement. Raise a cap only after reviewing that major version's release
    # notes against the cmdlets the collectors call.
    #
    # Last reviewed 2026-11-12 against the PowerShell Gallery: every module's latest release
    # is inside its range. MicrosoftTeams 8 (its only breaking change is to
    # New-CsAutoRecordingTemplate, which no collector calls) and Pester 6 were added.
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
            MaximumVersion = '8.99.99'
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
            MaximumVersion = '6.99.99'
            Scope          = @('DevTools')
            Reason         = 'Test runner. tools/run-pester.ps1 uses the configuration API introduced in Pester 5; the suite passes on 5.7.1 and 6.2.0.'
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
