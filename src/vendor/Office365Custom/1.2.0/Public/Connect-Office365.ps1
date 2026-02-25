function Connect-Office365 {
    <#
    .SYNOPSIS
        One-stop connection to Microsoft 365 services: Graph, ExchangeOnline, SharePointOnline, Teams.
    .DESCRIPTION
        Connects to Microsoft Graph (always first), discovers Tenant/SharePoint admin url, then connects to Exchange Online,
        SharePoint Online Admin, and Teams. Handles PS7 module expectations, optional re-authentication, and supports
        application or certificate-based authentication (Graph/ExchangeOnline).
    .PARAMETER WriteGraph
        Include write/elevated Graph scopes (delegate only).
    .PARAMETER Force
        Forces re-authentication (disconnects existing sessions where applicable).
    .PARAMETER TenantId
        Tenant ID (GUID). Required for Application/Certificate auth.
    .PARAMETER ClientSecretCredential
        Application client secret in PSCredential form (username = ClientId, password = client secret as SecureString).
        If not provided and required, will prompt.
    .PARAMETER CertificateThumbprint
        Thumbprint of the certificate (required for Certificate auth).
    .PARAMETER ClientId
        Application (Client) ID (required for Application/Certificate auth).
    .EXAMPLE
        Connect-Office365
    .EXAMPLE
        Connect-Office365 -WriteGraph
    .EXAMPLE
        Connect-Office365 -TenantId ... -ClientId ... -ClientSecretCredential (Get-Credential -UserName "AppId")
    .EXAMPLE
        Connect-Office365 -TenantId ... -ClientId ... -CertificateThumbprint 'xxxx'
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$WriteGraph,

        [Parameter()]
        [switch]$Force,

        [Parameter()]
        [string]$TenantId,

        [Parameter(Mandatory = $false, HelpMessage = "Enter a PSCredential where the username is your ClientId and the password is the app client secret as a SecureString. You can also run:`$ClientSecretCredential = Get-Credential -Message 'Please enter ClientId as username; password is the app client secret.'")]
        [ValidateNotNull()]
        [System.Management.Automation.Credential()]
        [System.Management.Automation.PSCredential]$ClientSecretCredential,

        [Parameter()]
        [string]$CertificateThumbprint,

        [Parameter()]
        [string]$ClientId
    )

    begin {
        $result = [ordered]@{
            Graph              = $false
            TenantName         = $null
            OnPremisesSyncEnabled = $false
            OnPremisesLastSyncDateTime = $null
            ExchangeOnline     = $false
            SharePointOnline   = $false
            Teams              = $false
            Tenant             = $null
            SharePointAdmin    = $null
            AuthenticationType = $null
            InitialDomain      = $null
        }

        $usingApplicationAuth = $false
        $AuthenticationType = if ($CertificateThumbprint) { 
            'Certificate' 
        } elseif ($ClientSecretCredential) { 
            'ClientSecret' 
        } else { 
            'Delegate' 
        }
        $result.AuthenticationType = $AuthenticationType

        if ($AuthenticationType -eq 'ClientSecret' -or $AuthenticationType -eq 'Certificate') {
            $usingApplicationAuth = $true

            if ($AuthenticationType -eq 'ClientSecret') {
                # Ensure PSCredential is present, prompt if absent/invalid
                if (
                    ($null -eq $ClientSecretCredential) -or
                    -not ($ClientSecretCredential -is [System.Management.Automation.PSCredential]) -or
                    ([string]::IsNullOrEmpty($ClientSecretCredential.GetNetworkCredential().UserName)) -or
                    ([string]::IsNullOrEmpty($ClientSecretCredential.GetNetworkCredential().Password))
                ) {
                    $ClientIdPrompt = Read-Host "Please enter ClientId (username for the App Registration)"
                    $ClientSecretCredential = Get-Credential -UserName $ClientIdPrompt -Message "Please enter the client secret as the password"
                }
            }
        }
        Write-Verbose "Authentication Type: $AuthenticationType"
    }

    process {
        # Disconnect previous sessions if -Force specified
        if ($Force) {
            Write-Host "Disconnecting existing sessions ..." -ForegroundColor Cyan -NoNewline
            try { 
                Write-Verbose "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan -NoNewline
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null 
            } catch {}
            try { 
                Write-Verbose "Disconnecting from Exchange Online..."
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue 
            } catch {}
            try { 
                Write-Verbose "Disconnecting from Microsoft Teams..."
                Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue 
            } catch {}
            Write-Host "✓ Disconnected" -ForegroundColor Green
        }

        # ===== PRE-LOAD EXCHANGEONLINE MANAGEMENT MODULE =====
        # Load ExchangeOnlineManagement before Graph to avoid MSAL assembly
        # version conflicts that can break Connect-ExchangeOnline in PS7.
        if (-not (Get-Module -ListAvailable -Name 'ExchangeOnlineManagement')) {
            Write-Host "✗ Exchange Online: The module 'ExchangeOnlineManagement' is not installed. Please install it before proceeding." -ForegroundColor Red
            return
        }
        else { 
            if (-not (Get-Module -Name 'ExchangeOnlineManagement')) {
                Write-Host "Importing ExchangeOnlineManagement module..." -ForegroundColor Cyan
                Import-Module 'ExchangeOnlineManagement' -ErrorAction Stop -WarningAction SilentlyContinue
                Write-Host "ExchangeOnlineManagement module imported successfully." -ForegroundColor Green
            }
        }

        # ===== CONNECT TO GRAPH FIRST =====
        $CommonScopes = @(
            "Organization.Read.All","User.Read.All","AuditLog.Read.All",
            "Group.Read.All","GroupMember.Read.All","RoleManagement.Read.Directory",
            "Domain.Read.All","Device.Read.All","Reports.Read.All",
            "Policy.Read.All","CrossTenantInformation.ReadBasic.All",
            "Security.Read.All","Application.Read.All",
            "Team.ReadBasic.All","Channel.ReadBasic.All"
        )
        $AdditionalScopes = @(
            "Directory.Read.All","Sites.Read.All","Files.Read.All",
            "MailboxSettings.ReadWrite","TeamSettings.Read.All","TeamsTab.Read.All",
            "LicenseAssignment.ReadWrite.All","User.EnableDisableAccount.All","User.Export.All",
            "SecurityEvents.Read.All","SharePointTenantSettings.Read.All",
            "Mail.Read","MailboxSettings.Read","DirectoryRecommendations.Read.All",
            "Policy.ReadWrite.CrossTenantAccess"
        )
        $WriteScopes = @(
            "Directory.ReadWrite.All","Synchronization.ReadWrite.All","Organization.ReadWrite.All",
            "IdentityRiskyUser.ReadWrite.All","IdentityUserFlow.ReadWrite.All","User.ReadWrite.All",
            "UserAuthenticationMethod.ReadWrite.All","User-LifeCycleInfo.ReadWrite.All",
            "DeviceManagementManagedDevices.ReadWrite.All","Domain.ReadWrite.All",
            "Group.ReadWrite.All","GroupMember.ReadWrite.All","Sites.ReadWrite.All"
        )
        $GraphScopes = if ($AuthenticationType -eq 'Delegate' -and $WriteGraph) { 
            $CommonScopes + $AdditionalScopes + $WriteScopes 
        } elseif ($AuthenticationType -eq 'Delegate') { 
            $CommonScopes 
        } else { 
            @() 
        }

        #region Microsoft Graph PowerShell Module Import
        try {
            Write-Verbose "Checking for Microsoft.Graph module and dependencies (handling PS 7 compatibility)..."

            if ($PSVersionTable.PSVersion.Major -ge 7) {
                #$graphModule = 'Microsoft.Graph'
                #if (Get-Module -ListAvailable -Name $graphModule) {
                #    if (-not (Get-Module -Name $graphModule)) {
                #        Write-Host "Importing $($graphModule) module..." -ForegroundColor Cyan
                #        Import-Module $graphModule -ErrorAction Stop -WarningAction SilentlyContinue
                #        Write-Host "$($graphModule) module imported successfully." -ForegroundColor Green
                #    } else {
                #        Write-Verbose "$graphModule module is already imported."
                #    }
                #} else {
                #     throw "$graphModule module is not installed."
                #}
            } else {
                $neededModules = @(
                    "Microsoft.Graph.Authentication","Microsoft.Graph.Users","Microsoft.Graph.Groups",
                    "Microsoft.Graph.Sites","Microsoft.Graph.Files","Microsoft.Graph.DirectoryObjects","Microsoft.Graph.Reports"
                )
                foreach ($mod in $neededModules) {
                    if (Get-Module -ListAvailable -Name $mod) {
                        if (-not (Get-Module -Name $mod)) {
                            Write-Verbose "Importing module $mod..."
                            Import-Module $mod -ErrorAction Stop
                        } else {
                            Write-Verbose "Module $mod is already imported."
                        }
                    } else {
                        throw "Module $mod is not installed."
                    }
                }
            }
        #endregion Microsoft Graph PowerShell Module Import

        #region Check if already connected to Microsoft Graph
            $existing = $null
            try { 
                Write-Verbose "Checking if already connected to Microsoft Graph..."
                $existing = Get-MgContext -ErrorAction Stop 
            } catch {}
            
            if ($existing -and -not $Force) {
                Write-Host "✓ Graph (already connected)" -ForegroundColor Green
                Write-Verbose "Existing Microsoft Graph session detected via '$($existing.AuthType)', skipping reconnect."
            } else {
                try {
                    switch ($AuthenticationType) {
                        'Certificate' {
                            Write-Verbose "Using certificate-based authentication for Graph."
                            if (-not $TenantId) { throw "Graph Certificate auth requires -TenantId." }
                            if (-not $ClientId) { throw "Graph Certificate auth requires -ClientId." }
                            Write-Host "Connecting to Graph (certificate)..." -ForegroundColor Cyan
                            Write-Verbose "Running Connect-MgGraph with -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint"
                            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop | Out-Null
                        }
                        'ClientSecret' {
                            Write-Verbose "Using application secret authentication for Graph."
                            if (-not $TenantId) { throw "Graph Application auth requires -TenantId." }
                            Write-Host "Connecting to Graph (application)..." -ForegroundColor Cyan
                            Write-Verbose "Running Connect-MgGraph with -TenantId $TenantId -ClientSecretCredential <PSCredential>"
                            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome -ErrorAction Stop | Out-Null
                        }
                        Default {
                            Write-Verbose "Using delegated authentication for Graph. Scopes: $($GraphScopes -join ', ')"
                            Write-Host "Connecting to Graph (delegate)..." -ForegroundColor Cyan
                            Write-Verbose "Running Connect-MgGraph with $($GraphScopes.Count) scopes"
                            Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop | Out-Null
                        }
                    }
                    Write-Host "✓ Graph connected" -ForegroundColor Green
                }
                catch {
                    Write-Host "✗ Graph: $($_.Exception.Message)" -ForegroundColor Red
                    $result.Graph = $false
                    return
                }
            }
            
            Write-Verbose "Retrieving organization info from Graph..."
            $org = Get-MgOrganization -ErrorAction Stop
            $result.OnPremisesSyncEnabled = $org.OnPremisesSyncEnabled
            $result.OnPremisesLastSyncDateTime = $org.OnPremisesLastSyncDateTime
            $result.Graph = $true
            $result.TenantName = $org.DisplayName
            Write-Verbose "Connected to tenant: '$($result.TenantName)'"
        }
        catch {
            Write-Host "✗ Graph: $($_.Exception.Message)" -ForegroundColor Red
            return
        }
        #endregion Check if already connected to Microsoft Graph

        #region Get Tenant Name from Graph if needed
        try {
            Write-Verbose "Extracting tenant name from organization domains..."
            $verifiedDomains = $org.VerifiedDomains
            $defaultDomain = $verifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -First 1
            if (-not $defaultDomain) { $defaultDomain = $verifiedDomains | Select-Object -First 1 }
            if (-not $defaultDomain) { 
                Write-Host "✗ Could not determine SharePoint tenant name: Unable to determine tenant default domain from Graph Organization object." -ForegroundColor Red
                return
            }
            
            $domainPart = $defaultDomain.Name.Split('.')[0]
            $result.InitialDomain = $defaultDomain.Name
            $TenantName = $domainPart
            Write-Verbose "Using tenant name '$TenantName' derived from domain '$($defaultDomain.Name)'"
        }
        catch {
            Write-Host "✗ Could not determine SharePoint tenant name: $($_.Exception.Message)" -ForegroundColor Red
            return
        }
        #endregion Get Tenant Name from Graph if needed

        # ===== CONNECT TO SHAREPOINT ONLINE (ADMIN) =====
        try {
            Write-Verbose "Checking for SharePoint module 'Microsoft.Online.SharePoint.PowerShell'..."
            
            if (-not (Get-Module -ListAvailable -Name 'Microsoft.Online.SharePoint.PowerShell')) {
                Write-Host "✗ SharePoint Online: The module 'Microsoft.Online.SharePoint.PowerShell' is not installed. Please install it before proceeding." -ForegroundColor Red
                return
            }

            if (-not (Get-Module -Name 'Microsoft.Online.SharePoint.PowerShell')) {
                Write-Host "Importing SharePoint module..." -ForegroundColor Cyan
                if ($PSVersionTable.PSVersion.Major -ge 7) {
                    Write-Verbose "Importing SharePoint module using Windows PowerShell context (for PS7+ compatibility)..."
                    Import-Module 'Microsoft.Online.SharePoint.PowerShell' -UseWindowsPowerShell -ErrorAction Stop -WarningAction SilentlyContinue
                } else {
                    Write-Verbose "Importing SharePoint module..."
                    Import-Module 'Microsoft.Online.SharePoint.PowerShell' -ErrorAction Stop -WarningAction SilentlyContinue
                }
                Write-Host "SharePoint module imported successfully." -ForegroundColor Green
            } else {
                Write-Verbose "SharePoint module already imported."
            }
            
            $spoAdminUrl = "https://$TenantName-admin.sharepoint.com"
            $existingAdminSite = $null
            try {
                Write-Verbose "Checking if already connected to SharePoint Online..."
                $tenant = Get-SPOTenant -ErrorAction Stop
                if ($tenant) {
                    $existingAdminSite = $spoAdminUrl
                }
            } catch {}
            
            if ($existingAdminSite -and -not $Force) {
                Write-Host "✓ SharePoint Online (already connected)" -ForegroundColor Green
                Write-Verbose "Existing SharePoint Online session detected at '$existingAdminSite', skipping reconnect."
                $result.SharePointOnline = $true
                $result.SharePointAdmin = $existingAdminSite
            }
            else {
                try {
                    switch ($authenticationType) {
                        'Certificate' {
                            Write-Verbose "Using certificate-based authentication for SharePoint Online ($spoAdminUrl)..."
                            if (-not $TenantId) { 
                                Write-Host "✗ SharePoint Online: SharePoint Online Certificate auth requires -TenantId." -ForegroundColor Red
                                return
                            }
                            if (-not $ClientId) { 
                                Write-Host "✗ SharePoint Online: SharePoint Online Certificate auth requires -ClientId." -ForegroundColor Red
                                return
                            }
                            Write-Host "Connecting to SharePoint Online (certificate)..." -ForegroundColor Cyan
                            Write-Verbose "Running Connect-SPOService with -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint"
                            Connect-SPOService -Url $spoAdminUrl -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop
                            $result.SharePointOnline = $true
                        }
                        'ClientSecret' {
                            Write-Warning "SharePoint Online does not support client secret authentication via Connect-SPOService. Will Rely on Microsoft Graph API instead."
                            #Write-Host "✗ SharePoint Online: SharePoint Online does not support client secret authentication via Connect-SPOService. Will Rely on Microsoft Graph API instead." -ForegroundColor Red
                            $result.SharePointOnline = $false

                            continue
                        }
                        Default {
                            Write-Verbose "Connecting to SharePoint Online ($spoAdminUrl) with Delegate authentication..."
                            Write-Host "Connecting to SharePoint Online (delegate)..." -ForegroundColor Cyan
                            Connect-SPOService -Url $spoAdminUrl -ErrorAction Stop
                            $result.SharePointOnline = $true
                        }
                    }

                    $result.SharePointAdmin = $spoAdminUrl
                    Write-Host "✓ SharePoint Online connected" -ForegroundColor Green
                } catch {
                    Write-Host "✗ SharePoint Online: $($_.Exception.Message)" -ForegroundColor Red
                    return
                }
            }
        }
        catch { 
            Write-Host "✗ SharePoint Online: $($_.Exception.Message)" -ForegroundColor Red
            return 
        }

        #region Connect to Exchange Online

        try {
            $existingEXOOrg = $null
            try {
                Write-Verbose "Checking if already connected to Exchange Online..."
                $existingEXOOrg = (Get-OrganizationConfig -ErrorAction Stop).Name
            } catch {}
            
            if ($existingEXOOrg -and -not $Force) {
                Write-Host "✓ Exchange Online (already connected)" -ForegroundColor Green
                Write-Verbose "Existing Exchange Online session detected for '$existingEXOOrg', skipping reconnect."
                $result.ExchangeOnline = $true
            } else {
                Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
                # Only support Certificate and delegate auth; warn and default to delegate if ClientSecretCredential is requested
                if ($authenticationType -eq 'Certificate') {
                    Write-Verbose "Using certificate-based authentication for Exchange Online."
                    if (-not $ClientId) { 
                        Write-Host "✗ Exchange Online: ExchangeOnline certificate auth requires -ClientId (AppId)." -ForegroundColor Red
                        return
                    }
                    if (-not $result.InitialDomain) { 
                        Write-Host "✗ Exchange Online: ExchangeOnline certificate auth requires Organization (initial domain)." -ForegroundColor Red
                        return
                    }
                    Write-Verbose "Running Connect-ExchangeOnline -AppId $ClientId -Organization $($result.InitialDomain) -CertificateThumbprint $CertificateThumbprint"
                    Connect-ExchangeOnline -AppId $ClientId -Organization $result.InitialDomain -CertificateThumbprint $CertificateThumbprint -ShowBanner:$false -ErrorAction Stop | Out-Null
                }
                elseif ($authenticationType -eq 'ClientSecret') {
                    Write-Warning "Exchange Online does not support ClientSecretCredential (App/Secret) authentication. Falling back to delegated authentication."
                    #Write-Verbose "Running Connect-ExchangeOnline with ShowBanner disabled (delegate authentication fallback)"
                    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
                }
                else {
                    Write-Verbose "Using delegated authentication for Exchange Online."
                    Write-Verbose "Running Connect-ExchangeOnline with ShowBanner disabled"
                    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
                }
                $result.ExchangeOnline = $true
                Write-Host "✓ Exchange Online connected" -ForegroundColor Green
            }
        }
        catch { 
            Write-Host "✗ Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
            return 
        }

        # ===== CONNECT TO TEAMS =====

        if (-not (Get-Module -ListAvailable -Name 'MicrosoftTeams')) {
            Write-Host "✗ Microsoft Teams: The module 'MicrosoftTeams' is not installed. Please install it before proceeding." -ForegroundColor Red
            return
        } else {
            if (-not (Get-Module -Name 'MicrosoftTeams')) {
                if ($PSVersionTable.PSVersion.Major -ge 7) {
                    Write-Verbose "Importing MicrosoftTeams module using Windows PowerShell context (for PS7+ compatibility)..."
                    Import-Module 'MicrosoftTeams' -UseWindowsPowerShell -ErrorAction Stop
                } else {
                    Write-Verbose "Importing MicrosoftTeams module..."
                    Import-Module 'MicrosoftTeams' -ErrorAction Stop
                }
            } else {
                Write-Verbose "ExchangeOnlineManagement module already imported."
            }
        }

        #region Import Microsoft Teams module
        if (-not (Get-Module -ListAvailable -Name 'MicrosoftTeams')) {
            # Already checked above, so do nothing (redundant check); keeping for structure
            return
        }
        else {
            if (-not (Get-Module -Name 'MicrosoftTeams')) {
                if ($PSVersionTable.PSVersion.Major -ge 7) {
                    Write-Verbose "Importing MicrosoftTeams module using Windows PowerShell context (for PS7+ compatibility)..."
                    Import-Module 'MicrosoftTeams' -UseWindowsPowerShell -ErrorAction Stop -WarningAction SilentlyContinue 
                } else {
                    Write-Verbose "Importing MicrosoftTeams module..."
                    Import-Module 'MicrosoftTeams' -SkipEditionCheck -ErrorAction Stop -WarningAction SilentlyContinue 
                }
                Write-Host "MicrosoftTeams module imported successfully." -ForegroundColor Green
            }
        }
        #endregion Import Microsoft Teams module

        #region Connect to Microsoft Teams
        try {
            $existingTeamsOrg = $null
            try {
                Write-Verbose "Checking if already connected to Microsoft Teams..."
                $existingTeamsOrg = (Get-CsTenant -ErrorAction Stop).DisplayName
            } catch {}
            
            if ($existingTeamsOrg -and -not $Force) {
                Write-Host "✓ Microsoft Teams (already connected)" -ForegroundColor Green
                Write-Verbose "Existing Microsoft Teams session detected for '$existingTeamsOrg', skipping reconnect."
                $result.Teams = $true
            } else {
                Write-Verbose "Connecting to Microsoft Teams (delegate authentication only)..."
                Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan
                if ($CertificateThumbprint -or $usingApplicationAuth) {
                    Write-Warning "Microsoft Teams connection currently supports delegate authentication only. Attempting connection..."
                }
                
                Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
                $result.Teams = $true
                Write-Host "✓ Microsoft Teams connected" -ForegroundColor Green
            }
        }
        catch { 
            Write-Host "✗ Microsoft Teams: $($_.Exception.Message)" -ForegroundColor Red
            return 
        }
        #endregion Connect to Microsoft Teams
    }

    end {
        if ($result.Graph -eq $false -or $result.SharePointOnline -eq $false -or $result.ExchangeOnline -eq $false -or $result.Teams -eq $false) {
            Throw "Failed to connect to tenant. One or more services failed to connect. Check errors above for connection problems encountered. Error: $($_.Exception.Message)"
        } else {
            Write-Verbose "Returning connection result object."
            Write-Host "`nConnection Summary:" -ForegroundColor Cyan
            Write-Host "  Tenant: $($result.TenantName)" -ForegroundColor White
            Write-Host "  Auth Type: $($result.AuthenticationType)" -ForegroundColor White
            Write-Host "  Initial Domain: $($result.InitialDomain)" -ForegroundColor White
            Write-Host "  SharePoint Admin: $($result.SharePointAdmin)" -ForegroundColor White
            Write-Host "  SharePoint Online: $($result.SharePointOnline)" -ForegroundColor White
            Write-Host "  Tenant DirSync Enabled: $($result.OnPremisesSyncEnabled)" -ForegroundColor White
            Write-Host "  Tenant DirSync Last Successful Sync: $($result.OnPremisesLastSyncDateTime)" -ForegroundColor White
        }
        #return [pscustomobject]$result
    }
}