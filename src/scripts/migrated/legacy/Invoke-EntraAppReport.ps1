#to do: Add filter for app with and without owners. Add a risk indicator for apps with no owners

<#PSScriptInfo
 
.VERSION 0.2.2
 
.GUID 175aa966-47dc-4a76-bbd1-b7ab11cd3079
 
.AUTHOR Daniel Bradley
 
.COMPANYNAME ourcloudnetwork.co.uk
 
.COPYRIGHT
 
.TAGS
    ourcloudnetwork
    Microsoft Entra
    Microsoft Graph
 
.LICENSEURI
    https://github.com/DanielBradley1/Invoke-EntraAppReport/blob/main/LICENSE
 
.PROJECTURI
    https://ourcloudnetwork.com/create-a-free-enterprise-app-permissions-report-in-microsoft-entra/
 
.EXTERNALMODULEDEPENDENCIES
    Microsoft.Graph.Authentication
 
.RELEASENOTES
    v0.1 - Initial release
    v0.2 - Added app registration owners
    v0.2.2 - Added risk indicators for insecure redirect URIs
#>

<#
.DESCRIPTION
 This script, created by Daniel Bradley at ourcloudnetwork.co.uk, generates a report of all applications in your Microsoft Entra tenant, including their permissions, credentials, and sign-in activity. The report includes a summary of the total number of applications, applications with delegated permissions, applications with application permissions, third-party applications, and inactive applications. The report also includes a list of applications with insecure sign-in methods, applications with application permissions, and applications with no sign-in activity. The report is generated in HTML format and can be saved to a file for further analysis.
 
.PARAMETER outpath
 Specified the output path of the report file.

.PARAMETER TenantId
 Optional tenant ID to connect to. When omitted, the script connects using the default tenant for the signed-in account.
 
.EXAMPLE
PS> Invoke-EntraAuthReport -outpath "C:\Reports\EntraAuthReport.html"
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$outpath,
    [Parameter()]
    [string]$TenantId,
    [Parameter()]
    [switch]$AllowUnsupportedLegacyExecution
)

$unsupportedLegacyScriptMessage = 'Invoke-EntraAppReport.ps1 is retained in this repo as an unsupported legacy reference script. Use Start-M365TenantAssessment.ps1 for supported production assessments. Re-run with -AllowUnsupportedLegacyExecution only when you intentionally need this standalone reference output for isolated testing.'
if (-not $AllowUnsupportedLegacyExecution) {
    throw $unsupportedLegacyScriptMessage
}

Write-Warning $unsupportedLegacyScriptMessage

# Display script start message
Write-Host "Starting Entra Application Permissions Report..." -ForegroundColor Cyan

# Check Microsoft Graph connection
Write-Host "Checking Microsoft Graph connection status..." -ForegroundColor Yellow

# Ensure required modules are loaded
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}
catch {
    Write-Error "The Microsoft.Graph.Authentication module is not installed or cannot be loaded. Please run 'Install-Module Microsoft.Graph.Authentication -Scope CurrentUser' and try again."
    exit
}

$state = Get-MgContext

# Define required permissions properly as an array of strings
$requiredPerms = @("AuditLog.Read.All", "Reports.Read.All", "Organization.Read.All", "Application.Read.All", "Directory.Read.All")

# Check if we're connected and have all required permissions
$hasAllPerms = $false
if ($state) {
    $missingPerms = @()
    foreach ($perm in $requiredPerms) {
        if ($state.Scopes -notcontains $perm) {
            $missingPerms += $perm
        }
    }
    if ($missingPerms.Count -eq 0) {
        $hasAllPerms = $true
        Write-output "Connected to Microsoft Graph with all required permissions"
    }
    else {
        Write-output "Missing required permissions: $($missingPerms -join ', ')"
        Write-output "Reconnecting with all required permissions..."
    }
}
else {
    Write-output "Not connected to Microsoft Graph. Connecting now..."
}

# Connect if we need to
if (-not $hasAllPerms -or ($TenantId -and $state -and $state.TenantId -ne $TenantId)) {
    try {
        if ($state) { Disconnect-MgGraph -ErrorAction SilentlyContinue }
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        $connectParams = @{ Scopes = $requiredPerms; ErrorAction = 'Stop'; NoWelcome = $true }
        if ($TenantId) { $connectParams['TenantId'] = $TenantId }
        Connect-MgGraph @connectParams
        Write-output "Successfully connected to Microsoft Graph"
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph: $_"
        exit
    }
}

# Initialize progress counter
$progressSteps = 8
$currentStep = 0

#Get organisation information
$currentStep++
Write-Progress -Activity "Gathering Entra Application Data" -Status "Getting organization information" -PercentComplete (($currentStep / $progressSteps) * 100)
$organisationInfo = (Invoke-MgGraphRequest -Uri "v1.0/organization" -OutputType PSObject | Select -Expand value)

# Get all service principals first to build our lookup dictionaries
$currentStep++
Write-Progress -Activity "Gathering Entra Application Data" -Status "Getting service principals and app role assignments" -PercentComplete (($currentStep / $progressSteps) * 100)
Write-Host "Retrieving service principals..." -ForegroundColor Yellow
$uri = "/v1.0/servicePrincipals?`$expand=appRoleAssignments&`$select=id,appId,displayName,appOwnerOrganizationId,replyUrls,appRoles&`$top=999"
$Result = Invoke-MgGraphRequest -Uri $uri -OutputType PSObject
$coreRoleAssignments = $Result.value
$NextLink = $Result."@odata.nextLink"
while ($NextLink -ne $null) {
    $Result = Invoke-MgGraphRequest -Method GET -Uri $NextLink -OutputType PSObject
    $coreRoleAssignments += $Result.value
    $NextLink = $Result."@odata.nextLink"
}

# Build dictionary of SP display names and AppRoles
$spDisplayNameDict = @{}
$appRoleDict = @{}
foreach ($sp in $coreRoleAssignments) {
    $spDisplayNameDict[$sp.id] = $sp.displayName
    if ($sp.appRoles) {
        foreach ($role in $sp.appRoles) {
            $appRoleDict[$role.id] = $role.value
        }
    }
}

# Get all OAuth2 delegated permissions for all apps
$currentStep++
Write-Progress -Activity "Gathering Entra Application Data" -Status "Getting OAuth2 delegated permissions" -PercentComplete (($currentStep / $progressSteps) * 100)
Write-Host "Retrieving OAuth2 delegated permissions..." -ForegroundColor Yellow
$uri = "v1.0/oauth2PermissionGrants?`$filter=ConsentType eq 'AllPrincipals'&`$top=999"
$Result = Invoke-MgGraphRequest -Uri $uri -OutputType PSObject
$coreDelegatedPermissions = $Result.value
$NextLink = $Result."@odata.nextLink"
while ($NextLink -ne $null) {
    $Result = Invoke-MgGraphRequest -Method GET -Uri $NextLink -OutputType PSObject
    $coreDelegatedPermissions += $Result.value
    $NextLink = $Result."@odata.nextLink"
}

# Process delegated permissions
$delegatedPermissions = @()
foreach ($grant in $coreDelegatedPermissions) {
    $spInfo = $coreRoleAssignments | Where-Object { $_.id -eq $grant.clientId }
    if (-not $spInfo) { continue }

    $resourceName = if ($grant.resourceId -and $spDisplayNameDict.ContainsKey($grant.resourceId)) { $spDisplayNameDict[$grant.resourceId] } else { "Unknown Resource ($($grant.resourceId))" }
    $scopes = $grant.Scope -split ' ' | Where-Object { $_ -ne '' }

    foreach ($scope in $scopes) {
        $delegatedPermissions += [PSCustomObject]@{
            ServicePrincipalId     = $spInfo.id
            DisplayName            = $spInfo.displayName
            appId                  = $spInfo.appId
            appOwnerOrganizationId = $spInfo.appOwnerOrganizationId
            ScopeType              = "Delegated"
            Scope                  = $scope
            ResourceId             = $grant.resourceId
            ResourceDisplayName    = $resourceName
        }
    }
}

# Process application role assignments
$currentStep++
Write-Progress -Activity "Gathering Entra Application Data" -Status "Processing application permissions" -PercentComplete (($currentStep / $progressSteps) * 100)
Write-Host "Processing application permissions..." -ForegroundColor Yellow
$appPermissions = @()
foreach ($sp in $coreRoleAssignments) {
    if ($sp.appRoleAssignments) {
        foreach ($assignment in $sp.appRoleAssignments) {
            if ($assignment.appRoleId -eq '00000000-0000-0000-0000-000000000000') { continue }

            $roleValue = if ($null -ne $assignment.appRoleId -and $appRoleDict.ContainsKey($assignment.appRoleId)) { $appRoleDict[$assignment.appRoleId] } else { $assignment.appRoleId }
            $resourceName = if ($null -ne $assignment.resourceId -and $spDisplayNameDict.ContainsKey($assignment.resourceId)) { $spDisplayNameDict[$assignment.resourceId] } else { $assignment.resourceDisplayName }
            if ([string]::IsNullOrEmpty($resourceName)) { $resourceName = "Unknown Resource ($($assignment.resourceId))" }

            $appPermissions += [PSCustomObject]@{
                ServicePrincipalId     = $sp.id
                DisplayName            = $sp.displayName
                appId                  = $sp.appId
                appOwnerOrganizationId = $sp.appOwnerOrganizationId
                ScopeType              = "Application"
                Scope                  = $roleValue
                ResourceId             = $assignment.resourceId
                ResourceDisplayName    = $resourceName
            }
        }
    }
}

# Group permissions by app and scope type
$groupedPermissions = ($delegatedPermissions + $appPermissions) | Group-Object ServicePrincipalId, ScopeType
$AllfilteredPermissions = @()

foreach ($group in $groupedPermissions) {
    $firstItem = $group.Group[0]

    $scopesByResource = @{}
    foreach ($item in $group.Group) {
        if (-not $scopesByResource.ContainsKey($item.ResourceDisplayName)) {
            $scopesByResource[$item.ResourceDisplayName] = @()
        }
        $scopesByResource[$item.ResourceDisplayName] += $item.Scope
    }

    $AllfilteredPermissions += [PSCustomObject]@{
        DisplayName            = $firstItem.DisplayName
        ServicePrincipalId     = $firstItem.ServicePrincipalId
        appId                  = $firstItem.appId
        appOwnerOrganizationId = $firstItem.appOwnerOrganizationId
        ScopeType              = $firstItem.ScopeType
        ScopesByResource       = $scopesByResource
    }
}
$AllfilteredPermissions = $AllfilteredPermissions | Sort-Object DisplayName

#Get ServicePrincipal Sign-in activity report
$currentStep++
Write-Progress -Activity "Gathering Entra Application Data" -Status "Getting service principal sign-in activity" -PercentComplete (($currentStep / $progressSteps) * 100)
Write-Host "Retrieving service principal sign-in activity..." -ForegroundColor Yellow
try {
    $response = Invoke-MgGraphRequest -uri "https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities" -OutputType PSObject -ErrorAction Stop | Select-Object -ExpandProperty Value
    $SignInActivityReport = $response
}
catch {
    Write-Warning "Could not retrieve sign-in activity report: $_"
    $SignInActivityReport = @()
}

#Function to get last sign-in activity
function GetLastActivityDate {
    param (
        [string]$appId
    )
    $lastsignin = $SignInActivityReport | Where-Object { $_.appId -eq $appId }
    $obj = [PSCustomObject]@{DelegatedLastSignIn = $lastsignin.delegatedClientSignInActivity.lastSignInDateTime; ApplicationLastSignIn = $lastsignin.applicationAuthenticationClientSignInActivity.lastSignInDateTime }
    return $obj
}

# Get Service Principal Audit Logs
Function Get-MgSpSignIns {
    param(
        $filter
    )

    process {
        try {
            $response = Invoke-MgGraphRequest -uri "https://graph.microsoft.com/beta/auditLogs/signIns?&source=sp&`$filter=$filter" -OutputType PSObject -ErrorAction Stop | Select-Object -ExpandProperty Value
            return $response
        }
        catch {
            Write-Warning "Could not retrieve sign-in logs: $_"
            return @()
        }
    }
}
$ServicePrincipalSignIns = Get-MgSpSignIns -filter "appOwnerTenantId ne '$($organisationInfo.id)'" | Select createdDateTime, appDisplayName, appId, clientCredentialType, appOwnerTenantId

#Add last sign-in activity to the permissions report
Write-Host "Adding sign-in activity data to applications..." -ForegroundColor Yellow
$AllfilteredPermissions | ForEach-Object {
    $app = $_
    $SignInInfo = GetLastActivityDate -appId $_.appId
    $LastSigninType = ($ServicePrincipalSignIns | Where-Object { $_.appId -eq $app.appId } | Sort-Object createdDateTime -Descending | Select-Object -First 1).clientCredentialType
    $_ | Add-Member -MemberType NoteProperty -Name "DelegatedLastSignIn" -Value $SignInInfo.DelegatedLastSignIn
    $_ | Add-Member -MemberType NoteProperty -Name "ApplicationLastSignIn" -Value $SignInInfo.ApplicationLastSignIn
    $_ | Add-Member -MemberType NoteProperty -Name "LastSignInType" -Value $LastSigninType
}

#Get password and key credentials for first-party applications
$currentStep++
Write-Progress -Activity "Gathering Entra Application Data" -Status "Getting application credentials" -PercentComplete (($currentStep / $progressSteps) * 100)
Write-Host "Retrieving application credentials..." -ForegroundColor Yellow
$credentials = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/applications/?`$select=displayName,id,appId,info,createdDateTime,keyCredentials,passwordCredentials,deletedDateTime,publicClient&`$count=true" -OutputType PSObject | Select -Expand Value
$credentials | ForEach-Object {
    $ownerdata = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/applications/$($_.id)/owners" -OutputType PSObject | Select -Expand Value
    $owner = $ownerdata.userPrincipalName
    $_ | Add-Member -MemberType NoteProperty -Name "Owners" -Value $owner
    $_ | Add-Member -MemberType NoteProperty -Name "RedirectUris" -Value $_.publicClient.RedirectUris
}


$AllfilteredPermissions | ForEach-Object {
    $app = $_
    $appcredentials = $credentials | Where-Object { $_.appId -eq $app.appId }
    $spdata = $spRoleAssignments | Where-Object { $_.appId -eq $app.appId }
    $credentialsToAdd = @()
    $KeyCert = $($appcredentials.keyCredentials).count
    $Password = $($appcredentials.passwordCredentials).count
    if ($($appcredentials.keyCredentials).count -gt 0) {
        $credentialsToAdd += "Client Certificate"
    }
    if ($($appcredentials.passwordCredentials).count -gt 0) {
        $credentialsToAdd += "Client Secret"
    }
    
    # Combine redirect URIs from both application and service principal
    $combinedRedirectUris = @()
    if ($null -ne $appcredentials.RedirectUris -and $appcredentials.RedirectUris.Count -gt 0) {
        $combinedRedirectUris += $appcredentials.RedirectUris
    }
    if ($null -ne $spdata.replyUrls -and $spdata.replyUrls.Count -gt 0) {
        $combinedRedirectUris += $spdata.replyUrls
    }
    # Remove duplicates and keep only unique URIs
    $uniqueRedirectUris = $combinedRedirectUris | Sort-Object -Unique
    
    $_ | Add-Member -MemberType NoteProperty -Name "App Credentials" -Value $credentialsToAdd
    $_ | Add-Member -MemberType NoteProperty -Name "Owners" -Value $appcredentials.Owners
    $_ | Add-Member -MemberType NoteProperty -Name "RedirectUris" -Value $uniqueRedirectUris
}

# Create additional stats
$currentStep++
Write-Progress -Activity "Gathering Entra Application Data" -Status "Creating security indicators" -PercentComplete (($currentStep / $progressSteps) * 100)
Write-Host "Generating security risk indicators..." -ForegroundColor Yellow

# Third-party apps with insecure sign-in methods
$WeakSignInMethods = @("clientSecret", "clientAssertion", "certificate")
$WeakAppCredentials = @("Client Secret", "Client Certificate")
$InsecureThirdPartyApps = $AllfilteredPermissions | Where { ($_.appOwnerOrganizationId -ne $organisationInfo.id) -and ($_.LastSignInType -in $WeakSignInMethods) } | Select DisplayName -Unique
# first-party apps with insecure sign-in methods
$InsecureFirstPartyApps = $AllfilteredPermissions | Where { ($_.appOwnerOrganizationId -eq $organisationInfo.id) -and ($_.'App Credentials' -in $WeakAppCredentials) } | Select DisplayName -Unique
# Third-party apps with application permissions
$ThirdPartyAppsWithAppPermissions = $AllfilteredPermissions | Where { ($_.appOwnerOrganizationId -ne $organisationInfo.id) -and ($_.ScopeType -eq "Application") } | Select DisplayName -Unique
# Applications with no sign-in activity
$NonActiveApps = $AllfilteredPermissions | Where { ($null -eq $_.DelegatedLastSignIn) -and ($null -eq $_.ApplicationLastSignIn) } | Select DisplayName -Unique
# First-party apps with no owners
$FirstPartyAppsWithNoOwners = $AllfilteredPermissions | Where { ($_.appOwnerOrganizationId -eq $organisationInfo.id) -and ([string]::IsNullOrEmpty($_.Owners) -or $_.Owners.Count -eq 0) } | Select DisplayName -Unique

# Applications with insecure redirect URIs
$AppsWithInsecureRedirectUris = $AllfilteredPermissions | Where-Object { 
    $hasInsecureUri = $false
    if ($null -ne $_.RedirectUris -and $_.RedirectUris.Count -gt 0) {
        foreach ($uri in $_.RedirectUris) {
            # Check for localhost patterns
            if ($uri -match "^https?://localhost" -or $uri -match "^https?://127\.0\.0\.1" -or $uri -match "^https?://\[::1\]") {
                $hasInsecureUri = $true
                break
            }
            # Check for wildcard domains (but not specific subdomains)
            if ($uri -match "^https?://\*\." -or $uri -match "://\*/" -or $uri -like "*//*") {
                $hasInsecureUri = $true
                break
            }
            # Check for Azure Web Apps (*.azurewebsites.net)
            if ($uri -match "^https?://.*\.azurewebsites\.net") {
                $hasInsecureUri = $true
                break
            }
            # Check for URL shorteners
            if ($uri -match "^https?://(bit\.ly|tinyurl\.com|t\.co|goo\.gl|ow\.ly|short\.link|tiny\.cc)") {
                $hasInsecureUri = $true
                break
            }
            # Check for insecure HTTP (non-HTTPS)
            if ($uri -match "^http://(?!localhost|127\.0\.0\.1|\[::1\])") {
                $hasInsecureUri = $true
                break
            }
        }
    }
    $hasInsecureUri
} | Select DisplayName -Unique

# Generate HTML Report for Entra Applications
function Export-EntraAppHTMLReport {
    param (
        [Parameter(Mandatory = $true)]
        [Array]$ReportData,
        
        [Parameter(Mandatory = $false)]
        [string]$ReportTitle = "Entra Application Permissions Report",
        
        [Parameter(Mandatory = $false)]
        [string]$OutputPath
    )
    
    # Define CSS for the report
    $css = @"
    <style>
        :root {
            --primary-color: #2563EB;
            --secondary-color: #1E40AF;
            --accent-color: #3B82F6;
            --background-color: #f8f9fa;
            --table-header-bg: #f2f2f2;
            --table-border: #ddd;
            --table-hover: #f1f1f1;
            --warning-color: #e74c3c;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            color: #333;
            margin: 0;
            padding: 0;
            background-color: var(--background-color);
        }
        .container {
            width: 95%;
            margin: 20px auto;
            background-color: white;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
            border-radius: 8px;
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
            position: relative;
            padding: 25px 20px;
            color: white;
            overflow: hidden;
        }
        .header::before {
            content: "";
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: radial-gradient(circle at 15% 50%, rgba(59, 130, 246, 0.3) 0%, transparent 50%),
                        radial-gradient(circle at 85% 30%, rgba(37, 99, 235, 0.2) 0%, transparent 50%);
        }
        .header h1 {
            margin: 0;
            font-size: 28px;
            font-weight: 600;
            position: relative;
            text-shadow: 0 1px 2px rgba(0, 0, 0, 0.1);
        }
        .header-subtitle {
            font-size: 14px;
            font-weight: 400;
            margin-top: 0px;
            margin-bottom: 10px;
            opacity: 0.9;
        }
        .header-content {
            display: flex;
            justify-content: space-between;
            position: relative;
        }
        .header-left-content {
            flex: 1;
        }
        .header-details {
            display: flex;
            justify-content: space-between;
            margin-top: 15px;
            font-size: 14px;
            position: relative;
            opacity: 0.9;
        }
        .author-info {
            margin-top: 12px;
            border-top: 1px solid rgba(255, 255, 255, 0.3);
            padding-top: 10px;
            display: flex;
            align-items: center;
            font-size: 13px;
        }
        .author-label {
            opacity: 0.8;
            margin-right: 6px;
        }
        .author-links {
            display: flex;
            align-items: center;
        }
        .author-link {
            color: white !important; /* Ensure text color is always white */
            text-decoration: none !important; /* Remove underline */
            display: inline-flex;
            align-items: center;
            border: 1px solid rgba(255, 255, 255, 0.5);
            padding: 4px 10px;
            border-radius: 4px;
            margin-right: 10px;
            transition: all 0.3s ease;
            background-color: rgba(255, 255, 255, 0.1);
            cursor: pointer;
            z-index: 10; /* Ensure links are above other elements */
            position: relative; /* Establish stacking context */
        }
        .author-link:hover {
            background-color: rgba(255, 255, 255, 0.3);
            border-color: rgba(255, 255, 255, 0.9);
            transform: translateY(-2px);
            box-shadow: 0 3px 5px rgba(0, 0, 0, 0.2);
        }
        .author-link:active {
            transform: translateY(0);
            box-shadow: 0 1px 2px rgba(0, 0, 0, 0.2);
        }
        .author-link svg {
            margin-right: 5px;
            flex-shrink: 0;
        }
        .report-info {
            text-align: right;
            font-size: 14px;
            display: flex;
            flex-direction: column;
            justify-content: flex-end;
            padding-bottom: 10px;
        }
        .report-date {
            font-weight: 500;
            margin-top: 5px;
        }
        .stats-container {
            display: flex;
            justify-content: space-between;
            margin: 20px;
            flex-wrap: wrap;
            gap: 15px;
        }
        .stat-box {
            background-color: white;
            border-radius: 10px;
            padding: 18px;
            box-shadow: 0 3px 10px rgba(0, 0, 0, 0.08);
            flex: 1;
            min-width: 200px;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
            border: none;
            position: relative;
            overflow: hidden;
        }
        .stat-box::after {
            content: '';
            position: absolute;
            bottom: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: linear-gradient(90deg,
                rgba(37, 99, 235, 0.8),
                rgba(59, 130, 246, 0.6),
                rgba(96, 165, 250, 0.4));
            border-bottom-left-radius: 10px;
            border-bottom-right-radius: 10px;
        }
        .stat-box:hover {
            transform: translateY(-3px);
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.1);
        }
        .stat-box h3 {
            margin-top: 0;
            color: #4b5563;
            font-size: 14px;
            font-weight: 500;
            letter-spacing: 0.3px;
            margin-bottom: 15px;
        }
        .stat-box p {
            margin-bottom: 0;
            font-size: 26px;
            font-weight: 600;
            color: #1f2937;
            background: linear-gradient(90deg, #2563EB, #4f46e5);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            text-fill-color: transparent;
        }
        .controls {
            padding: 15px 20px;
            background-color: #fff;
            border-bottom: 1px solid var(--table-border);
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            align-items: center;
        }
        .search-box {
            flex-grow: 1;
            position: relative;
            min-width: 250px;
        }
        .search-box input {
            width: 100%;
            padding: 8px 12px 8px 35px;
            border: 1px solid var(--table-border);
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
        }
        .search-box::before {
            content: "";
            position: absolute;
            left: 10px;
            top: 50%;
            transform: translateY(-50%);
            width: 14px;
            height: 14px;
            background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23666' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='11' cy='11' r='8'%3E%3C/circle%3E%3Cline x1='21' y1='21' x2='16.65' y2='16.65'%3E%3C/line%3E%3C/svg%3E");
            background-repeat: no-repeat;
            background-position: center;
        }
        select, button {
            padding: 8px 12px;
            border: 1px solid var(--table-border);
            border-radius: 4px;
            background-color: white;
            font-size: 14px;
        }
        button {
            cursor: pointer;
            background-color: var(--primary-color);
            color: white;
            border: none;
            transition: background-color 0.2s;
        }
        button:hover {
            background-color: var(--secondary-color);
        }
        .table-container {
            padding: 0 20px 20px;
            overflow-x: auto;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
        }
        th {
            background-color: var(--table-header-bg);
            color: #333;
            font-weight: 600;
            padding: 12px 15px;
            text-align: left;
            position: sticky;
            top: 0;
            cursor: pointer;
            user-select: none;
            border-bottom: 2px solid var(--table-border);
        }
        th:hover {
            background-color: #e6e6e6;
        }
        td {
            padding: 10px 15px;
            border-bottom: 1px solid var(--table-border);
            vertical-align: top;
        }
        tr:hover {
            background-color: var(--table-hover);
        }
        
        /* Expandable row styles */
        .expand-btn {
            background: none;
            border: none;
            cursor: pointer;
            padding: 4px 8px;
            border-radius: 4px;
            transition: all 0.2s ease;
            font-size: 14px;
            color: var(--primary-color);
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-width: 24px;
            height: 24px;
        }
        
        .expand-btn:hover {
            background-color: rgba(37, 99, 235, 0.1);
            transform: scale(1.1);
        }
        
        .expand-btn svg {
            transition: transform 0.2s ease;
            width: 16px;
            height: 16px;
        }
        
        .expand-btn.expanded svg {
            transform: rotate(90deg);
        }
        
        .expandable-row {
            cursor: pointer;
        }
        
        .expandable-row:hover {
            background-color: var(--table-hover);
        }
        
        .detail-row {
            display: none;
            background-color: #f8fafc;
            border-left: 3px solid var(--primary-color);
        }
        
        .detail-row.expanded {
            display: table-row;
        }
        
        .detail-row td {
            padding: 0;
            border-bottom: 1px solid var(--table-border);
        }
        
        .detail-content {
            padding: 20px;
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 15px;
            background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
            border-radius: 8px;
            margin: 10px;
        }
        
        .detail-item {
            background: white;
            padding: 16px 18px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
            border: 1px solid rgba(0, 0, 0, 0.05);
            transition: all 0.2s ease;
        }
        
        .detail-item:hover {
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
            transform: translateY(-1px);
        }
        
        .detail-label {
            font-weight: 600;
            color: #374151;
            font-size: 13px;
            margin-bottom: 6px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .detail-value {
            color: #1f2937;
            font-size: 14px;
            line-height: 1.4;
        }
        
        .detail-value .permission-list,
        .detail-value ul {
            margin: 0;
            padding-left: 16px;
        }
        
        .detail-value .no-activity {
            color: #6b7280;
            font-style: italic;
        }
        .permission-cell {
            max-width: 300px;
        }
        .permission-list {
            margin: 0;
            padding-left: 20px;
            word-break: break-word;
        }
        .date-cell {
            white-space: nowrap;
        }
        .no-activity {
            color: #999;
            font-style: italic;
        }
        .third-party {
            background-color: rgba(231, 76, 60, 0.1);
        }
        .third-party td:first-child::before {
            margin-right: 5px;
        }
        .tooltip {
            position: relative;
            display: inline-block;
            cursor: help;
        }
        .tooltip .tooltiptext {
            visibility: hidden;
            width: 200px;
            background-color: #333;
            color: #fff;
            text-align: center;
            border-radius: 6px;
            padding: 5px;
            position: absolute;
            z-index: 1;
            bottom: 125%;
            left: 50%;
            transform: translateX(-50%);
            opacity: 0;
            transition: opacity 0.3s;
        }
        .tooltip:hover .tooltiptext {
            visibility: visible;
            opacity: 1;
        }
        .footer {
            margin: 20px;
            text-align: center;
            font-size: 12px;
            color: #666;
            border-top: 1px solid var(--table-border);
            padding-top: 20px;
        }
        @media print {
            body {
                background-color: white;
            }
            .container {
                width: 100%;
                box-shadow: none;
            }
            .controls, button {
                display: none;
            }
            .tooltip .tooltiptext {
                display: none;
            }
        }
        @media (max-width: 768px) {
            .stats-container {
                flex-direction: column;
            }
            .stat-box {
                margin-bottom: 10px;
            }
            .header-details {
                flex-direction: column;
            }
        }
        .sort-icon::after {
            content: "";
            display: inline-block;
            margin-left: 5px;
            vertical-align: middle;
            border-left: 4px solid transparent;
            border-right: 4px solid transparent;
        }
        .sort-asc::after {
            border-bottom: 5px solid #333;
        }
        .sort-desc::after {
            border-top: 5px solid #333;
        }
         
        /* Add new styles for red flags section */
        .warning-icon {
            display: inline-block;
            width: 14px;
            height: 14px;
            background-color: #ffcc00;
            color: #000;
            border-radius: 50%;
            text-align: center;
            font-size: 11px;
            font-weight: bold;
            line-height: 14px;
            margin-right: 5px;
        }
            background-color: #FEECED;
            border-left: 4px solid #CD0000;
            margin: 20px;
            padding: 15px 20px;
            border-radius: 5px;
            box-shadow: 0 1px 5px rgba(0, 0, 0, 0.1);
        }
         
        .red-flag-title {
            color: #CD0000;
            font-weight: 600;
            margin-top: 0;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
        }
         
        .red-flag-title svg {
            margin-right: 10px;
            flex-shrink: 0;
        }
         
        .red-flag-item {
            margin-bottom: 15px;
            padding-bottom: 15px;
            border-bottom: 1px solid rgba(205, 0, 0, 0.2);
        }
         
        .red-flag-item:last-child {
            margin-bottom: 0;
            padding-bottom: 0;
            border-bottom: none;
        }
         
        .red-flag-heading {
            color: #CD0000;
            font-weight: 600;
            margin-top: 0;
            margin-bottom: 5px;
            font-size: 15px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .red-flag-heading-text {
            display: flex;
            align-items: center;
        }
        
        .minimize-btn {
            background: none;
            border: none;
            color: #777;
            cursor: pointer;
            font-size: 18px;
            padding: 0 5px;
            line-height: 1;
            transition: all 0.2s;
            margin-left: 10px;
        }
        
        .minimize-btn:hover {
            color: #CD0000;
            background: none;
        }
        
        .red-flag-content {
            transition: max-height 0.3s ease-out, opacity 0.3s ease-out;
            max-height: 500px;
            opacity: 1;
            overflow: hidden;
        }
        
        .red-flag-content.collapsed {
            max-height: 0;
            opacity: 0;
            margin-top: 0;
            margin-bottom: 0;
        }
         
        .red-flag-desc {
            margin-top: 0;
            margin-bottom: 10px;
        }
         
        .red-flag-count {
            display: inline-block;
            background-color: #CD0000;
            color: white;
            font-weight: bold;
            padding: 3px 8px;
            border-radius: 12px;
            font-size: 12px;
            margin-left: 8px;
        }
         
        .red-flag-apps {
            background-color: #FFDAD9;
            padding: 10px;
            border-radius: 4px;
            max-height: 120px;
            overflow-y: auto;
            font-size: 13px;
        }
         
        .red-flag-apps ul {
            margin: 0;
            padding-left: 20px;
        }
         
        .red-flag-apps li {
            margin-bottom: 3px;
        }
         
        .red-flag-link {
            color: #CD0000;
            cursor: pointer;
            text-decoration: underline;
            font-size: 13px;
        }
    </style>
"@

    # Format dates for display
    function Format-DateForDisplay {
        param($date)
        if ($null -eq $date -or $date -eq "") {
            return "<span class='no-activity'>No activity</span>"
        }
        else {
            try {
                $dt = [datetime]$date
                return $dt.ToString("yyyy-MM-dd HH:mm")
            }
            catch {
                return "<span class='no-activity'>Invalid date</span>"
            }
        }
    }

    # Calculate statistics for the report
    $totalApps = ($ReportData | Select-Object -Unique AppId).Count
    $delegatedApps = ($ReportData | Where-Object { $_.ScopeType -eq "Delegated" } | Select-Object -Unique AppId).Count
    $applicationApps = ($ReportData | Where-Object { $_.ScopeType -eq "Application" } | Select-Object -Unique AppId).Count
    $thirdPartyApps = ($ReportData | Where-Object { $_.appOwnerOrganizationId -ne $organisationInfo.id -and $null -ne $_.appOwnerOrganizationId } | Select-Object -Unique AppId).Count
    $inactiveApps = ($ReportData | Where-Object { $null -eq $_.DelegatedLastSignIn -and $null -eq $_.ApplicationLastSignIn } | Select-Object -Unique AppId).Count

    # Create rows for the report
    $tableRows = ""
    $rowIndex = 0
    foreach ($item in $ReportData) {
        $scopeList = ""
        if ($null -ne $item.ScopesByResource -and $item.ScopesByResource.Count -gt 0) {
            foreach ($resourceKey in $item.ScopesByResource.Keys | Sort-Object) {
                $scopeList += "<strong>$($resourceKey)</strong><ul class='permission-list'>"
                foreach ($scope in ($item.ScopesByResource[$resourceKey] | Sort-Object -Unique)) {
                    $scopeList += "<li>$scope</li>"
                }
                $scopeList += "</ul>"
            }
        }
        elseif ($null -ne $item.Scope -and $item.Scope.Count -gt 0) {
            $scopeList = "<ul class='permission-list'>"
            foreach ($scope in $item.Scope) {
                $scopeList += "<li>$scope</li>"
            }
            $scopeList += "</ul>"
        }
        else {
            $scopeList = "<span class='no-activity'>None</span>"
        }
        
        # Format App Credentials as bullet points similar to permissions
        $credentialsList = ""
        if ($null -ne $item.'App Credentials' -and $item.'App Credentials'.Count -gt 0) {
            $credentialsList = "<ul class='permission-list'>"
            foreach ($credential in $item.'App Credentials') {
                $credentialsList += "<li>$credential</li>"
            }
            $credentialsList += "</ul>"
        }
        else {
            $credentialsList = "<span class='no-activity'>None</span>"
        }
        
        # Format Owners as bullet points similar to permissions
        $ownersList = ""
        if ($null -ne $item.Owners -and $item.Owners.Count -gt 0) {
            $ownersList = "<ul class='permission-list'>"
            foreach ($owner in $item.Owners) {
                $ownersList += "<li>$owner</li>"
            }
            $ownersList += "</ul>"
        }
        else {
            $ownersList = "<span class='no-activity'>None</span>"
        }
        
        # Format Redirect URIs as bullet points similar to permissions
        $redirectUrisList = ""
        if ($null -ne $item.RedirectUris -and $item.RedirectUris.Count -gt 0) {
            $redirectUrisList = "<ul class='permission-list'>"
            foreach ($uri in $item.RedirectUris) {
                $redirectUrisList += "<li>$uri</li>"
            }
            $redirectUrisList += "</ul>"
        }
        else {
            $redirectUrisList = "<span class='no-activity'>None</span>"
        }
        
        $delegatedLastSignIn = Format-DateForDisplay $item.DelegatedLastSignIn
        $applicationLastSignIn = Format-DateForDisplay $item.ApplicationLastSignIn
        $lastSignInType = if ($item.LastSignInType) { $item.LastSignInType } else { "<span class='no-activity'>Unknown</span>" }
        
        # Determine app source and apply styling based on criteria
        $appSource = "First Party"
        $rowClass = ""
        $warningIcon = ""
        
        # Check if it's a third-party application
        $isThirdParty = (($null -eq $item.appOwnerOrganizationId) -or ($item.appOwnerOrganizationId -ne $organisationInfo.id))
        
        if ($isThirdParty) {
            $appSource = "Third Party"
            $rowClass = "class='third-party expandable-row'"
            
            # Only add warning icon for third-party apps with application permissions
            if ($item.ScopeType -eq "Application") {
                $warningIcon = "<span class='warning-icon' title='Third-Party App with Application Permissions'>!</span>"
            }
        }
        else {
            $rowClass = "class='expandable-row'"
        }
        
        # Main row with core information
        $tableRows += @"
        <tr $rowClass data-row-index="$rowIndex">
            <td>
                <button class="expand-btn" onclick="toggleRow($rowIndex)" title="Expand details">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <polyline points="9,18 15,12 9,6"></polyline>
                    </svg>
                </button>
                $warningIcon$($item.DisplayName)
            </td>
            <td>$($item.appId)</td>
            <td>$appSource</td>
            <td>$($item.ScopeType)</td>
            <td class="permission-cell">$scopeList</td>
        </tr>
        <tr class="detail-row" id="detail-$rowIndex">
            <td colspan="5">
                <div class="detail-content">
                    <div class="detail-item">
                        <div class="detail-label">Last Delegated Sign-in</div>
                        <div class="detail-value">$delegatedLastSignIn</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">Last Application Sign-in</div>
                        <div class="detail-value">$applicationLastSignIn</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">Last Sign-in Method</div>
                        <div class="detail-value">$lastSignInType</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">App Credentials</div>
                        <div class="detail-value">$credentialsList</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">Owners</div>
                        <div class="detail-value">$ownersList</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">Redirect URIs</div>
                        <div class="detail-value">$redirectUrisList</div>
                    </div>
                </div>
            </td>
        </tr>
"@
        $rowIndex++
    }

    # Create the HTML
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$ReportTitle</title>
    $css
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-content">
                <div class="header-left-content">
                    <h1>$ReportTitle</h1>
                    <div class="header-subtitle">Overview of application permissions in your tenant</div>
                    <div class="author-info">
                        <span class="author-label">Created by:</span>
                        <div class="author-links">
                            <a href="https://www.linkedin.com/in/danielbradley2/" class="author-link" onclick="window.open('https://www.linkedin.com/in/danielbradley2/', '_blank'); return false;">
                                <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="white">
                                    <path d="M19 0h-14c-2.761 0-5 2.239-5 5v14c0 2.761 2.239 5 5 5h14c2.762 0 5-2.239 5-5v-14c0-2.761-2.238-5-5-5zm-11 19h-3v-11h3v11zm-1.5-12.268c-.966 0-1.75-.79-1.75-1.764s.784-1.764 1.75-1.764 1.75.79 1.75 1.764-.783 1.764-1.75 1.764zm13.5 12.268h-3v-5.604c0-3.368-4-3.113-4 0v5.604h-3v-11h3v1.765c1.396-2.586 7-2.777 7 2.476v6.759z"/>
                                </svg>
                                Daniel Bradley
                            </a>
                            <a href="https://ourcloudnetwork.com" class="author-link" onclick="window.open('https://ourcloudnetwork.com', '_blank'); return false;">
                                <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="white">
                                    <path d="M21 13v10h-21v-19h12v2h-10v15h17v-8h2zm3-12h-10.988l4.035l4-6.977 7.07 2.828 2.828 6.977-7.07 4.125 4.172v-11z"/>
                                </svg>
                                ourcloudnetwork.com
                            </a>
                        </div>
                    </div>
                </div>
                <div class="report-info">
                    <div class="report-date">Generated: $(Get-Date -Format "MMMM d, yyyy")</div>
                    <div class="tenant">Tenant ID: $($organisationInfo.id)</div>
                </div>
            </div>
        </div>
 
        <div class="stats-container">
            <div class="stat-box">
                <h3>Total Applications</h3>
                <p>$totalApps</p>
            </div>
            <div class="stat-box">
                <h3>Apps with Delegated Permissions</h3>
                <p>$delegatedApps</p>
            </div>
            <div class="stat-box">
                <h3>Apps with Application Permissions</h3>
                <p>$applicationApps</p>
            </div>
            <div class="stat-box">
                <h3>Third-Party Apps</h3>
                <p>$thirdPartyApps</p>
            </div>
            <div class="stat-box">
                <h3>Apps with No Sign-In Activity</h3>
                <p>$inactiveApps</p>
            </div>
        </div>
         
"@

    # Add Red Flag section if needed
    if ($InsecureThirdPartyApps.Count -gt 0 -or $InsecureFirstPartyApps.Count -gt 0 -or $ThirdPartyAppsWithAppPermissions.Count -gt 0 -or $NonActiveApps.Count -gt 0 -or $FirstPartyAppsWithNoOwners.Count -gt 0 -or $AppsWithInsecureRedirectUris.Count -gt 0) {
        $html += @"
        <div class="red-flag-container">
            <h3 class="red-flag-title">
                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#CD0000" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path>
                    <line x1="12" y1="9" x2="12" y2="13"></line>
                    <line x1="12" y1="17" x2="12.01" y2="17"></line>
                </svg>
                Security Risk Indicators
            </h3>
"@
        
        if ($InsecureThirdPartyApps.Count -gt 0) {
            $html += @"
            <div class="red-flag-item">
                <h4 class="red-flag-heading">
                    <span class="red-flag-heading-text">Third-Party Applications Using Insecure Authentication <span class="red-flag-count">$($InsecureThirdPartyApps.Count)</span></span>
                    <button class="minimize-btn" onclick="toggleRedFlag(this)" title="Minimize">-</button>
                </h4>
                <div class="red-flag-content">
                    <p class="red-flag-desc">These third-party applications are using insecure authentication methods (client secrets, certificates or client assertion) which may pose security risks. Work with these vendors to assess the risk.</p>
                    <div style="margin-top: 8px;">
                        <span class="red-flag-link" onclick="filterByInsecureThirdPartyApps()">Filter to show these apps</span>
                    </div>
                </div>
            </div>
"@
        }
        
        if ($InsecureFirstPartyApps.Count -gt 0) {
            $html += @"
            <div class="red-flag-item">
                <h4 class="red-flag-heading">
                    <span class="red-flag-heading-text">First-Party Applications Using Insecure Authentication <span class="red-flag-count">$($InsecureFirstPartyApps.Count)</span></span>
                    <button class="minimize-btn" onclick="toggleRedFlag(this)" title="Minimize">-</button>
                </h4>
                <div class="red-flag-content">
                    <p class="red-flag-desc">These first-party applications are using insecure authentication methods (client secrets or certificates) which may pose security risks. Consider using a Managed Identity.</p>
                    <div style="margin-top: 8px;">
                        <span class="red-flag-link" onclick="filterByInsecureFirstPartyApps()">Filter to show these apps</span>
                    </div>
                </div>
            </div>
"@
        }
        
        if ($ThirdPartyAppsWithAppPermissions.Count -gt 0) {
            $html += @"
            <div class="red-flag-item">
                <h4 class="red-flag-heading">
                    <span class="red-flag-heading-text">Third-Party Applications with Application Permissions <span class="red-flag-count">$($ThirdPartyAppsWithAppPermissions.Count)</span></span>
                    <button class="minimize-btn" onclick="toggleRedFlag(this)" title="Minimize">-</button>
                </h4>
                <div class="red-flag-content">
                    <p class="red-flag-desc">These third-party applications have application-level permissions which grant access without user context. Review these regularly to ensure they follow the principle of least privilege.</p>
                    <div style="margin-top: 8px;">
                        <span class="red-flag-link" onclick="filterByThirdPartyWithAppPermissions()">Filter to show these apps</span>
                    </div>
                </div>
            </div>
"@
        }
        
        if ($NonActiveApps.Count -gt 0) {
            $html += @"
            <div class="red-flag-item">
                <h4 class="red-flag-heading">
                    <span class="red-flag-heading-text">Applications with No Recent Activity <span class="red-flag-count">$($NonActiveApps.Count)</span></span>
                    <button class="minimize-btn" onclick="toggleRedFlag(this)" title="Minimize">-</button>
                </h4>
                <div class="red-flag-content">
                    <p class="red-flag-desc">These applications show no sign-in activity. Consider reviewing and removing unused applications to reduce potential attack surface and improve security.</p>
                    <div style="margin-top: 8px;">
                        <span class="red-flag-link" onclick="filterByInactiveApps()">Filter to show these apps</span>
                    </div>
                </div>
            </div>
"@
        }
        
        if ($FirstPartyAppsWithNoOwners.Count -gt 0) {
            $html += @"
            <div class="red-flag-item">
                <h4 class="red-flag-heading">
                    <span class="red-flag-heading-text">First-Party Applications with No Owners <span class="red-flag-count">$($FirstPartyAppsWithNoOwners.Count)</span></span>
                    <button class="minimize-btn" onclick="toggleRedFlag(this)" title="Minimize">-</button>
                </h4>
                <div class="red-flag-content">
                    <p class="red-flag-desc">These first-party applications have no assigned owners. Applications without owners can become orphaned when team members leave and may pose governance risks.</p>
                    <div style="margin-top: 8px;">
                        <span class="red-flag-link" onclick="filterByFirstPartyAppsWithNoOwners()">Filter to show these apps</span>
                    </div>
                </div>
            </div>
"@
        }
        
        if ($AppsWithInsecureRedirectUris.Count -gt 0) {
            $html += @"
            <div class="red-flag-item">
                <h4 class="red-flag-heading">
                    <span class="red-flag-heading-text">Applications with Insecure Redirect URIs <span class="red-flag-count">$($AppsWithInsecureRedirectUris.Count)</span></span>
                    <button class="minimize-btn" onclick="toggleRedFlag(this)" title="Minimize">-</button>
                </h4>
                <div class="red-flag-content">
                    <p class="red-flag-desc">These applications have potentially insecure redirect URIs including localhost addresses, wildcard domains, Azure Web Apps, URL shorteners, or insecure HTTP redirects. These configurations may pose security risks and should be reviewed.</p>
                    <div style="margin-top: 8px;">
                        <span class="red-flag-link" onclick="filterByInsecureRedirectUris()">Filter to show these apps</span>
                    </div>
                </div>
            </div>
"@
        }
        
        $html += " </div>`n"
    }

    $html += @"
        <div class="controls">
            <div class="search-box">
                <input type="text" id="searchInput" placeholder="Search applications...">
            </div>
            <div class="search-box">
                <input type="text" id="permissionSearchInput" placeholder="Search permissions...">
            </div>
            <select id="sourceFilter">
                <option value="all">All Sources</option>
                <option value="First Party">First Party Only</option>
                <option value="Third Party">Third Party Only</option>
            </select>
            <select id="scopeTypeFilter">
                <option value="all">All Permission Types</option>
                <option value="Delegated">Delegated</option>
                <option value="Application">Application</option>
            </select>
            <select id="activityFilter">
                <option value="all">All Activity</option>
                <option value="active">Has Activity</option>
                <option value="inactive">No Activity</option>
            </select>
            <select id="credentialFilter">
                <option value="all">All Credentials</option>
                <option value="any">Any Credentials</option>
                <option value="none">No Credentials</option>
                <option value="cert">Client Certificate</option>
                <option value="secret">Client Secret</option>
                <option value="both">Certificate + Secret</option>
            </select>
            <select id="ownerFilter">
                <option value="all">All Owners</option>
                <option value="with">With Owners</option>
                <option value="without">No Owners</option>
            </select>
            <button onclick="exportCSV()">Export CSV</button>
            <button onclick="window.print()">Print Report</button>
        </div>
         
        <div class="table-container">
            <table id="appTable">
                <thead>
                    <tr>
                        <th class="sort-icon" onclick="sortTable(0)">Application Name</th>
                        <th class="sort-icon" onclick="sortTable(1)">App ID</th>
                        <th class="sort-icon" onclick="sortTable(2)">Source</th>
                        <th class="sort-icon" onclick="sortTable(3)">Permission Type</th>
                        <th>Permissions</th>
                    </tr>
                </thead>
                <tbody>
                    $tableRows
                </tbody>
            </table>
        </div>
         
        <div class="footer">
            <p>Report generated using Entra Application Permissions Report Tool | © $(Get-Date -Format "yyyy") ourcloudnetwork.com</p>
        </div>
    </div>
     
    <script>
        // Search & Filter Functionality
        document.getElementById('searchInput').addEventListener('keyup', filterTable);
        document.getElementById('permissionSearchInput').addEventListener('keyup', filterTable);
        document.getElementById('sourceFilter').addEventListener('change', filterTable);
        document.getElementById('scopeTypeFilter').addEventListener('change', filterTable);
        document.getElementById('activityFilter').addEventListener('change', filterTable);
        document.getElementById('credentialFilter').addEventListener('change', filterTable);
        document.getElementById('ownerFilter').addEventListener('change', filterTable);
        
        // Row expansion functionality
        function toggleRow(rowIndex) {
            const detailRow = document.getElementById('detail-' + rowIndex);
            const expandBtn = document.querySelector('tr[data-row-index="' + rowIndex + '"] .expand-btn');
            
            if (detailRow.classList.contains('expanded')) {
                detailRow.classList.remove('expanded');
                expandBtn.classList.remove('expanded');
            } else {
                detailRow.classList.add('expanded');
                expandBtn.classList.add('expanded');
            }
        }
         
        function filterTable() {
            const searchTerm = document.getElementById('searchInput').value.toLowerCase();
            const permissionSearchTerm = document.getElementById('permissionSearchInput').value.toLowerCase();
            const sourceFilter = document.getElementById('sourceFilter').value;
            const scopeFilter = document.getElementById('scopeTypeFilter').value;
            const activityFilter = document.getElementById('activityFilter').value;
            const credentialFilter = document.getElementById('credentialFilter').value;
            const ownerFilter = document.getElementById('ownerFilter').value;
            
            // Get all main rows (not detail rows)
            const rows = document.querySelectorAll('#appTable tbody tr.expandable-row');
             
            rows.forEach(row => {
                const rowIndex = row.getAttribute('data-row-index');
                const detailRow = document.getElementById('detail-' + rowIndex);
                
                const appName = row.cells[0].textContent.toLowerCase();
                const appId = row.cells[1].textContent.toLowerCase();
                const source = row.cells[2].textContent;
                const scopeType = row.cells[3].textContent;
                const permissions = row.cells[4].textContent.toLowerCase();
                
                // Get data from detail row for additional filtering with error handling
                let delegatedActivity = 'No activity';
                let applicationActivity = 'No activity';
                let signInMethod = 'Unknown';
                let credentials = 'none';
                let owners = 'none';
                let redirectUris = 'none';
                
                try {
                    const detailContent = detailRow.querySelector('.detail-content');
                    if (detailContent && detailContent.children.length >= 6) {
                        delegatedActivity = detailContent.children[0].querySelector('.detail-value').textContent;
                        applicationActivity = detailContent.children[1].querySelector('.detail-value').textContent;
                        signInMethod = detailContent.children[2].querySelector('.detail-value').textContent;
                        credentials = detailContent.children[3].querySelector('.detail-value').textContent.toLowerCase();
                        owners = detailContent.children[4].querySelector('.detail-value').textContent.toLowerCase();
                        redirectUris = detailContent.children[5].querySelector('.detail-value').textContent.toLowerCase();
                    }
                } catch (e) {
                    console.log('Error accessing detail content for row ' + rowIndex + ':', e);
                }
                 
                const matchesSearch = appName.includes(searchTerm) ||
                                     appId.includes(searchTerm) ||
                                     credentials.includes(searchTerm) ||
                                     signInMethod.toLowerCase().includes(searchTerm) ||
                                     redirectUris.includes(searchTerm);
                                      
                const matchesPermissionSearch = permissionSearchTerm === '' ||
                                              permissions.includes(permissionSearchTerm);
                                               
                const matchesSource = sourceFilter === 'all' || source === sourceFilter;
                const matchesScopeFilter = scopeFilter === 'all' || scopeType === scopeFilter;
                 
                let matchesActivity = true;
                if (activityFilter === 'active') {
                    matchesActivity = !(delegatedActivity.includes('No activity') && applicationActivity.includes('No activity'));
                } else if (activityFilter === 'inactive') {
                    matchesActivity = delegatedActivity.includes('No activity') && applicationActivity.includes('No activity');
                }
                 
                // Process credential filter
                let matchesCredential = true;
                const hasCert = credentials.includes('client certificate');
                const hasSecret = credentials.includes('client secret');
                const hasAnyCredential = hasCert || hasSecret;
                 
                switch (credentialFilter) {
                    case 'any':
                        matchesCredential = hasAnyCredential;
                        break;
                    case 'none':
                        matchesCredential = credentials.includes('none');
                        break;
                    case 'cert':
                        matchesCredential = hasCert;
                        break;
                    case 'secret':
                        matchesCredential = hasSecret;
                        break;
                    case 'both':
                        matchesCredential = hasCert && hasSecret;
                        break;
                    default: // 'all'
                        matchesCredential = true;
                }
                
                // Process owner filter
                let matchesOwner = true;
                const hasOwners = !owners.includes('none');
                
                switch (ownerFilter) {
                    case 'with':
                        matchesOwner = hasOwners;
                        break;
                    case 'without':
                        matchesOwner = !hasOwners;
                        break;
                    default: // 'all'
                        matchesOwner = true;
                }
                
                const shouldShow = (matchesSearch && matchesPermissionSearch &&
                                  matchesSource && matchesScopeFilter &&
                                  matchesActivity && matchesCredential &&
                                  matchesOwner);
                
                row.style.display = shouldShow ? '' : 'none';
                detailRow.style.display = shouldShow ? '' : 'none';
                
                // Collapse detail row if main row is hidden
                if (!shouldShow && detailRow.classList.contains('expanded')) {
                    detailRow.classList.remove('expanded');
                    const expandBtn = row.querySelector('.expand-btn');
                    if (expandBtn) {
                        expandBtn.classList.remove('expanded');
                    }
                }
            });
        }
         
        // Sorting Functionality - Updated for expandable rows
        let currentSortCol = -1;
        let currentSortDir = 'asc';
         
        function sortTable(columnIndex) {
            const table = document.getElementById('appTable');
            const headers = table.querySelectorAll('th');
             
            // Reset all headers
            headers.forEach(header => {
                header.classList.remove('sort-asc', 'sort-desc');
                if (header.classList.contains('sort-icon')) {
                    header.classList.add('sort-icon');
                }
            });
             
            // Set sort direction
            if (currentSortCol === columnIndex) {
                currentSortDir = currentSortDir === 'asc' ? 'desc' : 'asc';
            } else {
                currentSortDir = 'asc';
            }
             
            // Update header class
            headers[columnIndex].classList.add(currentSortDir === 'asc' ? 'sort-asc' : 'sort-desc');
             
            currentSortCol = columnIndex;
             
            // Get only main rows (not detail rows) for sorting
            const mainRows = Array.from(table.querySelectorAll('tbody tr.expandable-row'));
             
            const sortedRowPairs = mainRows.map(row => {
                const rowIndex = row.getAttribute('data-row-index');
                const detailRow = document.getElementById('detail-' + rowIndex);
                return { main: row, detail: detailRow };
            }).sort((a, b) => {
                let aVal = a.main.cells[columnIndex].textContent.trim();
                let bVal = b.main.cells[columnIndex].textContent.trim();
                
                // Clean up app name for sorting (remove warning icons)
                if (columnIndex === 0) {
                    aVal = aVal.replace('!', '').trim();
                    bVal = bVal.replace('!', '').trim();
                }
                 
                const comparison = aVal.localeCompare(bVal, undefined, {numeric: true, sensitivity: 'base'});
                return currentSortDir === 'asc' ? comparison : -comparison;
            });
             
            // Remove all existing rows
            const tbody = table.querySelector('tbody');
            while (tbody.firstChild) {
                tbody.removeChild(tbody.firstChild);
            }
             
            // Append sorted row pairs in order
            sortedRowPairs.forEach(pair => {
                tbody.appendChild(pair.main);
                tbody.appendChild(pair.detail);
            });
        }
         
        // CSV Export Functionality - Updated for expandable rows
        function exportCSV() {
            const table = document.getElementById('appTable');
            
            // Create CSV content with BOM for Excel UTF-8 support
            let csvContent = [];
            
            // Add header row
            const headers = ['Application Name', 'App ID', 'Source', 'Permission Type', 'Permissions', 
                           'Last Delegated Sign-in', 'Last Application Sign-in', 'Last Sign-in Method', 
                           'App Credentials', 'Owners', 'Redirect URIs'];
            csvContent.push(headers.map(h => '"' + h + '"').join(','));
            
            // Get all main rows (not detail rows)
            const mainRows = Array.from(table.querySelectorAll('tbody tr.expandable-row'));
             
            mainRows.forEach(row => {
                // Skip hidden rows
                if (row.style.display === 'none') return;
                
                const rowIndex = row.getAttribute('data-row-index');
                const detailRow = document.getElementById('detail-' + rowIndex);
                const detailContent = detailRow.querySelector('.detail-content');
                
                const rowData = [];
                
                // Get main row data
                for (let i = 0; i < 5; i++) {
                    let content = '';
                    const cell = row.cells[i];
                    
                    if (i === 0) {
                        // Application name - remove expand button and clean up
                        content = cell.textContent.replace('!', '').trim();
                        // Remove the expand button text/content
                        const expandBtn = cell.querySelector('.expand-btn');
                        if (expandBtn) {
                            content = content.replace(expandBtn.textContent || '', '').trim();
                        }
                    } else if (i === 4) {
                        // Permissions cell - handle list items
                        const items = Array.from(cell.querySelectorAll('li')).map(li => li.textContent.trim());
                        content = items.length > 0 ? items.join('; ') : 
                                 (cell.textContent.includes('None') ? 'None' : cell.textContent.trim());
                    } else {
                        content = cell.textContent.trim();
                    }
                    
                    // Escape quotes with double quotes (CSV standard)
                    content = content.replace(/"/g, '""');
                    rowData.push('"' + content + '"');
                }
                
                // Get detail row data
                const detailItems = detailContent.querySelectorAll('.detail-item .detail-value');
                detailItems.forEach(item => {
                    let content = '';
                    
                    // Handle list items in detail data
                    const items = Array.from(item.querySelectorAll('li')).map(li => li.textContent.trim());
                    if (items.length > 0) {
                        content = items.join('; ');
                    } else {
                        content = item.textContent.trim();
                    }
                    
                    // Escape quotes with double quotes (CSV standard)
                    content = content.replace(/"/g, '""');
                    rowData.push('"' + content + '"');
                });
                
                // Add row to CSV content
                csvContent.push(rowData.join(','));
            });
             
            // Create a BOM for UTF-8
            const BOM = '\uFEFF';
             
            // Join all rows with proper newlines and add BOM
            const csvString = BOM + csvContent.join('\r\n');
             
            // Create blob and download
            const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.setAttribute("href", url);
            link.setAttribute("download", "EntraAppReport_$(Get-Date -Format 'yyyy-MM-dd').csv");
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        }
         
        // Red Flag filtering functions - Updated for expandable rows
        function filterByInsecureThirdPartyApps() {
            // Reset UI filters
            document.getElementById('searchInput').value = '';
            document.getElementById('permissionSearchInput').value = '';
            document.getElementById('sourceFilter').value = 'Third Party';
            document.getElementById('scopeTypeFilter').value = 'all';
            document.getElementById('activityFilter').value = 'all';
            document.getElementById('credentialFilter').value = 'any';
            document.getElementById('ownerFilter').value = 'all';
             
            // Apply filtering
            const rows = document.querySelectorAll('#appTable tbody tr.expandable-row');
            rows.forEach(row => {
                const rowIndex = row.getAttribute('data-row-index');
                const detailRow = document.getElementById('detail-' + rowIndex);
                
                const source = row.cells[2].textContent;
                let signInMethod = 'Unknown';
                let credentials = 'none';
                
                try {
                    const detailContent = detailRow.querySelector('.detail-content');
                    if (detailContent && detailContent.children.length >= 6) {
                        signInMethod = detailContent.children[2].querySelector('.detail-value').textContent;
                        credentials = detailContent.children[3].querySelector('.detail-value').textContent;
                    }
                } catch (e) {
                    console.log('Error accessing detail content for red flag filter:', e);
                }
                 
                // Show only third party apps with client certificate or client secret authentication
                const isThirdParty = source === 'Third Party';
                const hasInsecureAuth = signInMethod.includes('clientSecret') ||
                                       signInMethod.includes('certificate') ||
                                       signInMethod.includes('clientAssertion') ||
                                       credentials.includes('Client Certificate') ||
                                       credentials.includes('Client Secret');
                 
                const shouldShow = isThirdParty && hasInsecureAuth;
                row.style.display = shouldShow ? '' : 'none';
                detailRow.style.display = shouldShow ? '' : 'none';
            });
        }
         
        function filterByInsecureFirstPartyApps() {
            // Reset UI filters
            document.getElementById('searchInput').value = '';
            document.getElementById('permissionSearchInput').value = '';
            document.getElementById('sourceFilter').value = 'First Party';
            document.getElementById('scopeTypeFilter').value = 'all';
            document.getElementById('activityFilter').value = 'all';
            document.getElementById('credentialFilter').value = 'any';
            document.getElementById('ownerFilter').value = 'all';
             
            // Apply filtering
            const rows = document.querySelectorAll('#appTable tbody tr.expandable-row');
            rows.forEach(row => {
                const rowIndex = row.getAttribute('data-row-index');
                const detailRow = document.getElementById('detail-' + rowIndex);
                
                const source = row.cells[2].textContent;
                let credentials = 'none';
                
                try {
                    const detailContent = detailRow.querySelector('.detail-content');
                    if (detailContent && detailContent.children.length >= 6) {
                        credentials = detailContent.children[3].querySelector('.detail-value').textContent;
                    }
                } catch (e) {
                    console.log('Error accessing detail content for red flag filter:', e);
                }
                 
                // Show only first party apps with client certificate or client secret credentials
                const isFirstParty = source === 'First Party';
                const hasInsecureAuth = credentials.includes('Client Certificate') ||
                                       credentials.includes('Client Secret');
                 
                const shouldShow = isFirstParty && hasInsecureAuth;
                row.style.display = shouldShow ? '' : 'none';
                detailRow.style.display = shouldShow ? '' : 'none';
            });
        }
         
        function filterByThirdPartyWithAppPermissions() {
            // Reset UI filters
            document.getElementById('searchInput').value = '';
            document.getElementById('permissionSearchInput').value = '';
            document.getElementById('sourceFilter').value = 'Third Party';
            document.getElementById('scopeTypeFilter').value = 'Application';
            document.getElementById('activityFilter').value = 'all';
            document.getElementById('credentialFilter').value = 'all';
            document.getElementById('ownerFilter').value = 'all';
             
            // The UI filters should handle this case automatically
            filterTable();
        }
         
        function filterByInactiveApps() {
            // Reset UI filters
            document.getElementById('searchInput').value = '';
            document.getElementById('permissionSearchInput').value = '';
            document.getElementById('sourceFilter').value = 'all';
            document.getElementById('scopeTypeFilter').value = 'all';
            document.getElementById('activityFilter').value = 'inactive';
            document.getElementById('credentialFilter').value = 'all';
            document.getElementById('ownerFilter').value = 'all';
             
            // The UI filters should handle this case automatically
            filterTable();
        }
        
        function filterByFirstPartyAppsWithNoOwners() {
            // Reset UI filters
            document.getElementById('searchInput').value = '';
            document.getElementById('permissionSearchInput').value = '';
            document.getElementById('sourceFilter').value = 'First Party';
            document.getElementById('scopeTypeFilter').value = 'all';
            document.getElementById('activityFilter').value = 'all';
            document.getElementById('credentialFilter').value = 'all';
            document.getElementById('ownerFilter').value = 'without';
            
            // The UI filters should handle this case automatically
            filterTable();
        }
        
        function filterByInsecureRedirectUris() {
            // Reset most filters but use search to find apps with insecure redirect URIs
            document.getElementById('searchInput').value = '';
            document.getElementById('permissionSearchInput').value = '';
            document.getElementById('sourceFilter').value = 'all';
            document.getElementById('scopeTypeFilter').value = 'all';
            document.getElementById('activityFilter').value = 'all';
            document.getElementById('credentialFilter').value = 'all';
            document.getElementById('ownerFilter').value = 'all';
            
            // Use custom filtering logic to show only apps with insecure redirect URIs
            const rows = document.querySelectorAll('#appTable tbody tr.expandable-row');
            
            rows.forEach(row => {
                const rowIndex = row.getAttribute('data-row-index');
                const detailRow = document.getElementById('detail-' + rowIndex);
                let hasInsecureUri = false;
                
                try {
                    const detailContent = detailRow.querySelector('.detail-content');
                    if (detailContent && detailContent.children.length >= 6) {
                        const redirectUrisElement = detailContent.children[5].querySelector('.detail-value');
                        const redirectUrisText = redirectUrisElement.textContent;
                        
                        // Split URIs and check each one with precise pattern matching
                        const uris = redirectUrisText.split(',').map(uri => uri.trim()).filter(uri => uri && uri !== 'None');
                        
                        for (const uri of uris) {
                            // Check for localhost patterns
                            if (/^https?:\/\/localhost/.test(uri) || /^https?:\/\/127\.0\.0\.1/.test(uri) || /^https?:\/\/\[::1\]/.test(uri)) {
                                hasInsecureUri = true;
                                break;
                            }
                            // Check for wildcard domains (but not specific subdomains)
                            if (/^https?:\/\/\*\./.test(uri) || /:\/\/\*\//.test(uri) || uri.includes('//*')) {
                                hasInsecureUri = true;
                                break;
                            }
                            // Check for Azure Web Apps (*.azurewebsites.net)
                            if (/^https?:\/\/.*\.azurewebsites\.net/.test(uri)) {
                                hasInsecureUri = true;
                                break;
                            }
                            // Check for URL shorteners
                            if (/^https?:\/\/(bit\.ly|tinyurl\.com|t\.co|goo\.gl|ow\.ly|short\.link|tiny\.cc)/.test(uri)) {
                                hasInsecureUri = true;
                                break;
                            }
                            // Check for insecure HTTP (non-HTTPS) excluding localhost
                            if (/^http:\/\/(?!localhost|127\.0\.0\.1|\[::1\])/.test(uri)) {
                                hasInsecureUri = true;
                                break;
                            }
                        }
                    }
                } catch (e) {
                    console.log('Error checking redirect URIs for row ' + rowIndex + ':', e);
                }
                
                row.style.display = hasInsecureUri ? '' : 'none';
                detailRow.style.display = hasInsecureUri ? '' : 'none';
                
                // Collapse detail row if main row is hidden
                if (!hasInsecureUri && detailRow.classList.contains('expanded')) {
                    detailRow.classList.remove('expanded');
                    const expandBtn = row.querySelector('.expand-btn');
                    if (expandBtn) {
                        expandBtn.classList.remove('expanded');
                    }
                }
            });
        }
         
        // Initialize sorting on application name column
        window.onload = function() {
            sortTable(0);
        };
        
        // Function to toggle red flag sections
        function toggleRedFlag(button) {
            const content = button.closest('.red-flag-item').querySelector('.red-flag-content');
            content.classList.toggle('collapsed');
            
            if (content.classList.contains('collapsed')) {
                button.textContent = '+';
                button.title = 'Expand';
            } else {
                button.textContent = '-';
                button.title = 'Minimize';
            }
        }
        
        // Make toggleRow available globally
        window.toggleRow = toggleRow;
    </script>
</body>
</html>
"@

    # Save the HTML file
    try {
        # If no output path is specified, use script root
        if ([string]::IsNullOrEmpty($OutputPath)) {
            $OutputPath = Join-Path $PSScriptRoot "EntraAppReport_$(Get-Date -Format 'yyyy-MM-dd_HH-mm').html"
        }
        
        # Ensure directory exists
        $directory = Split-Path -Path $OutputPath -Parent
        if (!(Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
        
        $html | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
        Write-Output "HTML Report saved to: $OutputPath"
        
        # Open the report in the default browser
        Start-Process $OutputPath
    }
    catch {
        Write-Error "Failed to save the HTML report: $_"
        return $null
    }
}

# Generate and open the HTML report after collecting the data
# Determine output path
Write-Progress -Activity "Generating Report" -Status "Creating HTML report" -PercentComplete 100
Write-Host "Generating HTML report..." -ForegroundColor Green
$reportOutputPath = $null
if ($outpath) {
    # Check if the provided path is a directory or a file
    if (Test-Path $outpath -PathType Container) {
        # It's a directory, append filename
        $reportOutputPath = Join-Path $outpath "EntraAppReport_$(Get-Date -Format 'yyyy-MM-dd_HH-mm').html"
    }
    else {
        # It's a file path or doesn't exist, use as is
        $reportOutputPath = $outpath
    }
}

Export-EntraAppHTMLReport -ReportData $AllfilteredPermissions -ReportTitle "Entra Application Permissions Report" -OutputPath $reportOutputPath

# Complete
Write-Progress -Activity "Generating Report" -Completed
Write-Host "Report generation complete!" -ForegroundColor Green
