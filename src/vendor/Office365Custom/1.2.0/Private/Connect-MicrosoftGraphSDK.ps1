    <#
    .SYNOPSIS
        Connects to various Office 365 services.
    .DESCRIPTION
        Based on the provided service names (or prompting the user), this function connects to Exchange Online, MSOnline, Microsoft Graph, SharePoint Online, Azure AD, and Teams.
    .PARAMETER ServiceName
        Specifies one or more service names to connect to.
    .PARAMETER WriteGraph
        Indicates graph connection with write permissions.
    .PARAMETER Force
        Forces re-authentication.
    .EXAMPLE
        Connect-Office365Services -ServiceName "ExchangeOnline","Graph"
    #>

    # Connect to Microsoft Graph using the Graph SDK
    function Connect-MicrosoftGraphSDK {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $false, HelpMessage = 'Choose if Read Only Graph Connection')]
            [switch]$WriteGraph,
            [Parameter(Mandatory = $false, HelpMessage = 'Provide the Graph authentication')]
            [ValidateSet('Delegate', 'Application', 'Certificate')]
            [string]$AuthenticationType = 'Delegate',
            [Parameter(Mandatory = $false)]
            [string]$TenantId,
            [Parameter(Mandatory = $false)]
            [string]$ClientId,
            [Parameter(Mandatory = $false)]
            [string]$CertificateThumbprint,
            [Parameter(Mandatory = $false, HelpMessage = 'Force Re-authentication')]
            [switch]$Force
        )
    
        # List of necessary Microsoft Graph modules
        $requiredModules = @(
            "Microsoft.Graph.Authentication",
            "Microsoft.Graph.Users",
            "Microsoft.Graph.Groups",
            "Microsoft.Graph.Sites",
            "Microsoft.Graph.Files",
            "Microsoft.Graph.DirectoryObjects",
            "Microsoft.Graph.Reports"
        )

        #Import Required Modules
        if ($PSVersionTable.PSVersion.Major -lt 7) {
            $requiredModules | foreach {
                Import-RequiredModule -ModuleName $_ -InstallMessage "Run 'Install-Module $_ -Scope CurrentUser -Repository PSGallery -Force' as an Administrator."
            }
        } else {
            Import-RequiredModule -ModuleName "Microsoft.Graph" -InstallMessage "Run 'Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force' as an Administrator."
        }
    
        # Connect to Microsoft Graph
        if ($Force) {
            Disconnect-MgGraph | out-null
        }
        try {
            # Check for existing Microsoft Graph connection
            $mgContext = Get-MgContext -ErrorAction Stop
            if ($mgContext) {
                $MGraphCompanyCheck = Get-MgOrganization -ErrorAction Stop
                Write-Verbose "Already Using Microsoft Graph PowerShell SDK - '$($mgContext.AuthType)' Authentication"
                if (Confirm-Tenant $MGraphCompanyCheck.DisplayName) {
                    return
                } else {
                    Disconnect-MgGraph | out-null
                }
            }
            switch ($AuthenticationType) {
                'application' { 
                    # Define tenant credentials
                    $targetTenantId = Read-Host "Tenant ID"
                    $targetClientSecretCredential = Get-Credential -Message "Enter Entra App client (app) ID as username and secret value of tenant as password"
    
                    # Authenticate to source tenant
                    Connect-MgGraph -TenantId $targetTenantId -ClientSecretCredential $targetClientSecretCredential }
                'certificate' {
                    if ([string]::IsNullOrWhiteSpace($TenantId)) {
                        $TenantId = Read-Host "Tenant ID"
                    }
                    if ([string]::IsNullOrWhiteSpace($ClientId)) {
                        $ClientId = Read-Host "Application (Client) ID"
                    }
                    if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
                        $CertificateThumbprint = Read-Host "Certificate Thumbprint"
                    }

                    if ([string]::IsNullOrWhiteSpace($TenantId) -or [string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
                        throw "Certificate authentication requires TenantId, ClientId, and CertificateThumbprint."
                    }

                    Write-Verbose "Connecting to Microsoft Graph using certificate authentication..."
                    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop
                }
                Default {
                    # Define Core Scopes for Microsoft Graph SDK
                    $CommonScopes = @(
                        "Directory.Read.All", "Group.Read.All", "GroupMember.Read.All",
                        "User.Read.All", "Sites.Read.All", "Files.Read.All",
                        "MailboxSettings.ReadWrite", "AuditLog.Read.All", "Policy.Read.All",
                        "Team.ReadBasic.All", "TeamSettings.Read.All", "TeamsTab.Read.All",
                        "LicenseAssignment.ReadWrite.All", "User.EnableDisableAccount.All", "User.Export.All",
                        "Device.Read.All", "SecurityEvents.Read.All", "SharePointTenantSettings.Read.All", "Organization.Read.All",
                        "OnPremDirectorySynchronization.Read.All"
                    )
    
                    $WriteScopes = @(
                        "Directory.ReadWrite.All", "Synchronization.ReadWrite.All", 
                        "Organization.ReadWrite.All", "IdentityRiskyUser.ReadWrite.All",
                        "IdentityUserFlow.ReadWrite.All", "User.ReadWrite.All", 
                        "UserAuthenticationMethod.ReadWrite.All", "User-LifeCycleInfo.ReadWrite.All",
                        "DeviceManagementManagedDevices.ReadWrite.All", "Domain.ReadWrite.All",
                        "Group.ReadWrite.All", "GroupMember.ReadWrite.All", "MailboxSettings.ReadWrite",
                        "Policy.Read.All", "Sites.ReadWrite.All"
                    )
    
                    $RequiredScopes = if ($WriteGraph) { $CommonScopes + $WriteScopes } else { $CommonScopes }
    
                    Write-Verbose "Connecting to Microsoft Graph using Delegate Permissions..."

                    # Import required modules
                    if ($PSVersionTable.PSVersion.Major -lt 7) {
                        Import-RequiredModule -ModuleName $requiredModules
                    } else {
                        Import-RequiredModule -ModuleName "Microsoft.Graph" -InstallMessage "Run 'Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force' as an Administrator."
                    }
                    Connect-MgGraph -Scopes $RequiredScopes -ErrorAction Stop
                }
            }        
        }
        catch {
            Write-Host "Error connecting to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
            Throw
        }
    
        # Output connection details
        $MGraphCompanyCheck = Get-MgOrganization -ErrorAction Stop
        Update-TitleBar $MGraphCompanyCheck.DisplayName
        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        Write-Host "Connected: '$($MGraphCompanyCheck.DisplayName)' tenant - Using '$($mgContext.AuthType)' Authentication" -ForegroundColor Green
    }
