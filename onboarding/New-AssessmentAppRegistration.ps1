<#
.SYNOPSIS
    Creates and configures the Entra app registration used by the Modern Workplace tenant assessment.

.DESCRIPTION
    Creates a single-tenant app registration, uploads a self-signed certificate credential,
    creates the service principal, grants admin consent for the Microsoft Graph application
    permissions required by the collector, adds the Office 365 Exchange Online
    Exchange.ManageAsApp permission, then configures one mutually exclusive authorization
    model. LeastPrivilege leaves Entra directory roles unassigned for workload-scoped read-only
    RBAC. GlobalReader assigns Global Reader and removes Exchange Administrator if present.

    The script is idempotent. Re-running it against an existing app of the same display name
    reuses that app, adds only the missing permissions, and skips consent that already exists.

.NOTES
    Required signed-in admin rights: Global Administrator, or Application Administrator plus
    Privileged Role Administrator. Granting admin consent and assigning a directory role both
    need privileged rights that Application Administrator alone does not provide.

    Delegated Graph scopes requested: Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All,
    RoleManagement.ReadWrite.Directory, Directory.Read.All.

    New-SelfSignedCertificate is Windows-only, so the certificate path requires Windows
    PowerShell 5.1 or PowerShell 7 on Windows.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$|^[A-Za-z0-9][A-Za-z0-9.-]*\.[A-Za-z]{2,}$')]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$DisplayName = 'Arraya M365 Tenant Assessment',

    [Parameter(Mandatory = $false)]
    [ValidateSet('Core', 'Standard', 'Extended')]
    [string]$PermissionSet = 'Standard',

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$CertificateSubject,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 5)]
    [int]$CertificateValidityYears = 2,

    [Parameter(Mandatory = $false)]
    [switch]$SkipCertificate,

    [Parameter(Mandatory = $false)]
    [switch]$SkipExchange,

    [Parameter(Mandatory = $false)]
    [ValidateSet('LeastPrivilege', 'GlobalReader')]
    [string]$AccessModel = 'LeastPrivilege',

    [Parameter(Mandatory = $false)]
    [switch]$SkipAdminConsent,

    [Parameter(Mandatory = $false)]
    [switch]$UseDeviceAuthentication,

    [Parameter(Mandatory = $false)]
    [switch]$DisableWebAccountManager
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$GraphAppId = '00000003-0000-0000-c000-000000000000'
$ExchangeOnlineAppId = '00000002-0000-0ff1-ce00-000000000000'
$LegacyExchangeRoleName = 'Exchange Administrator'
$GlobalReaderRoleName = 'Global Reader'

# Core permissions exercised by the collector preflight probes and the Graph scope planner.
$CorePermissions = @(
    'Application.Read.All'
    'AuditLog.Read.All'
    'Channel.ReadBasic.All'
    'CrossTenantInformation.ReadBasic.All'
    'Device.Read.All'
    'Domain.Read.All'
    'Group.Read.All'
    'GroupMember.Read.All'
    'OnPremDirectorySynchronization.Read.All'
    'Organization.Read.All'
    'Policy.Read.All'
    'Reports.Read.All'
    'ReportSettings.Read.All'
    'RoleManagement.Read.All'
    'SecurityEvents.Read.All'
    'SharePointTenantSettings.Read.All'
    'Sites.Read.All'
    'Team.ReadBasic.All'
    'User.Read.All'
)

# Teams member and guest count enrichment.
$StandardOnlyPermissions = @(
    'TeamMember.Read.All'
)

# Optional extended enrichment beyond the standard assessment path.
$ExtendedOnlyPermissions = @(
    'DeviceManagementApps.Read.All'
    'DeviceManagementConfiguration.Read.All'
    'DeviceManagementManagedDevices.Read.All'
    'DeviceManagementRBAC.Read.All'
    'DeviceManagementScripts.Read.All'
    'DeviceManagementServiceConfig.Read.All'
    'DirectoryRecommendations.Read.All'
    'IdentityRiskEvent.Read.All'
    'IdentityRiskyUser.Read.All'
    'LicenseAssignment.Read.All'
    'ServiceMessage.Read.All'
    'User.Export.All'
    'UserAuthenticationMethod.Read.All'
)

function Assert-M365GraphCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    if (-not (Get-Command -Name $Name -ErrorAction SilentlyContinue)) {
        throw "Microsoft Graph PowerShell cmdlet '$Name' is not available. Install or update the Microsoft.Graph module, then retry."
    }
}

function Get-AssessmentRequestedPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Core', 'Standard', 'Extended')]
        [string]$Set
    )

    $permissions = New-Object System.Collections.Generic.List[string]
    foreach ($permission in $CorePermissions) { $permissions.Add($permission) | Out-Null }

    if ($Set -in @('Standard', 'Extended')) {
        foreach ($permission in $StandardOnlyPermissions) { $permissions.Add($permission) | Out-Null }
    }

    if ($Set -eq 'Extended') {
        foreach ($permission in $ExtendedOnlyPermissions) { $permissions.Add($permission) | Out-Null }
    }

    return @($permissions | Select-Object -Unique)
}

function Resolve-AssessmentServicePrincipal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$AppId,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Description
    )

    $servicePrincipals = @(Get-MgServicePrincipal -Filter "appId eq '$AppId'" -All -ErrorAction Stop)
    if ($servicePrincipals.Count -eq 0) {
        throw "The $Description service principal (appId '$AppId') was not found in tenant '$TenantId'. It is normally provisioned automatically. Confirm the workload is licensed and enabled in this tenant."
    }

    return $servicePrincipals[0]
}

function Resolve-AssessmentAppRole {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $ResourceServicePrincipal,

        [Parameter(Mandatory = $true)]
        [string[]]$PermissionName
    )

    $resolved = New-Object System.Collections.Generic.List[pscustomobject]
    $unresolved = New-Object System.Collections.Generic.List[string]

    foreach ($name in $PermissionName) {
        $appRole = @(
            $ResourceServicePrincipal.AppRoles |
                Where-Object { $_.Value -eq $name -and $_.AllowedMemberTypes -contains 'Application' }
        ) | Select-Object -First 1

        if ($appRole) {
            $resolved.Add([pscustomobject]@{
                    Name  = $name
                    RoleId = $appRole.Id
                }) | Out-Null
        }
        else {
            $unresolved.Add($name) | Out-Null
        }
    }

    if ($unresolved.Count -gt 0) {
        Write-Warning ("These application permissions were not found on '{0}' and will be skipped: {1}" -f $ResourceServicePrincipal.DisplayName, ($unresolved -join ', '))
    }

    return @($resolved)
}

foreach ($commandName in @(
        'Connect-MgGraph',
        'Get-MgContext',
        'Get-MgApplication',
        'New-MgApplication',
        'Update-MgApplication',
        'Get-MgServicePrincipal',
        'New-MgServicePrincipal',
        'Get-MgServicePrincipalAppRoleAssignment',
        'New-MgServicePrincipalAppRoleAssignment',
        'Get-MgRoleManagementDirectoryRoleDefinition',
        'Get-MgRoleManagementDirectoryRoleAssignment',
        'New-MgRoleManagementDirectoryRoleAssignment',
        'Remove-MgRoleManagementDirectoryRoleAssignment'
    )) {
    Assert-M365GraphCommand -Name $commandName
}

if (-not $SkipCertificate -and -not (Get-Command -Name 'New-SelfSignedCertificate' -ErrorAction SilentlyContinue)) {
    throw 'New-SelfSignedCertificate is not available. Run this script on Windows, or pass -SkipCertificate and upload a certificate manually.'
}

if ([string]::IsNullOrWhiteSpace($CertificateSubject)) {
    $CertificateSubject = "CN=$DisplayName"
}

Write-Warning 'The signed-in admin must be able to grant admin consent and assign directory roles, normally Global Administrator or Application Administrator plus Privileged Role Administrator. If PIM is used, activate the roles before continuing.'

$connectParameters = @{
    TenantId    = $TenantId
    Scopes      = @(
        'Application.ReadWrite.All'
        'AppRoleAssignment.ReadWrite.All'
        'RoleManagement.ReadWrite.Directory'
        'Directory.Read.All'
    )
    NoWelcome   = $true
    ErrorAction = 'Stop'
}

if ($UseDeviceAuthentication) {
    $connectParameters['UseDeviceAuthentication'] = $true
}

# Web Account Manager is the default broker on Windows and requires a parent window handle.
# Hosts without one (background processes, some embedded terminals) fail with
# 'A window handle must be configured'. Disabling it falls back to the system browser.
if ($DisableWebAccountManager) {
    if (Get-Command -Name 'Set-MgGraphOption' -ErrorAction SilentlyContinue) {
        Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction Stop
        Write-Verbose 'Disabled Web Account Manager sign-in; the system browser will be used.'
    }
    else {
        Write-Warning 'Set-MgGraphOption is not available, so Web Account Manager could not be disabled. Update Microsoft.Graph.Authentication if interactive sign-in fails with a window handle error.'
    }
}

try {
    Connect-MgGraph @connectParameters | Out-Null
}
catch {
    throw "Could not connect to Microsoft Graph tenant '$TenantId'. Details: $($_.Exception.Message)"
}

$context = Get-MgContext
if (-not $context) {
    throw 'Microsoft Graph sign-in did not produce a context. Retry the connection.'
}

$resolvedTenantId = $context.TenantId
Write-Verbose ("Connected to tenant '{0}' as '{1}'." -f $resolvedTenantId, $context.Account)

$requestedPermissions = Get-AssessmentRequestedPermission -Set $PermissionSet
$graphServicePrincipal = Resolve-AssessmentServicePrincipal -AppId $GraphAppId -Description 'Microsoft Graph'
$graphRoles = Resolve-AssessmentAppRole -ResourceServicePrincipal $graphServicePrincipal -PermissionName $requestedPermissions

$resourceAccessPlan = New-Object System.Collections.Generic.List[pscustomobject]
$resourceAccessPlan.Add([pscustomobject]@{
        ServicePrincipal = $graphServicePrincipal
        ResourceAppId    = $GraphAppId
        Roles            = $graphRoles
    }) | Out-Null

$exchangeServicePrincipal = $null
if (-not $SkipExchange) {
    $exchangeServicePrincipal = Resolve-AssessmentServicePrincipal -AppId $ExchangeOnlineAppId -Description 'Office 365 Exchange Online'
    $exchangeRoles = Resolve-AssessmentAppRole -ResourceServicePrincipal $exchangeServicePrincipal -PermissionName @('Exchange.ManageAsApp')
    if ($exchangeRoles.Count -eq 0) {
        throw "The Exchange.ManageAsApp application permission was not found on the Office 365 Exchange Online service principal in tenant '$resolvedTenantId'. Re-run with -SkipExchange to provision a Graph-only app."
    }

    $resourceAccessPlan.Add([pscustomobject]@{
            ServicePrincipal = $exchangeServicePrincipal
            ResourceAppId    = $ExchangeOnlineAppId
            Roles            = $exchangeRoles
        }) | Out-Null
}

$requiredResourceAccess = @(
    foreach ($resource in $resourceAccessPlan) {
        @{
            resourceAppId  = $resource.ResourceAppId
            resourceAccess = @(
                foreach ($role in $resource.Roles) {
                    @{ id = $role.RoleId; type = 'Role' }
                }
            )
        }
    }
)

$escapedDisplayName = $DisplayName.Replace("'", "''")
$existingApplications = @(Get-MgApplication -Filter "displayName eq '$escapedDisplayName'" -All -ErrorAction Stop)
if ($existingApplications.Count -gt 1) {
    throw "More than one app registration named '$DisplayName' exists in tenant '$resolvedTenantId'. Remove the duplicates or pass a unique -DisplayName."
}

if ($existingApplications.Count -eq 1) {
    $application = $existingApplications[0]
    $applicationStatus = 'Existing'
    Write-Verbose ("Reusing existing app registration '{0}' (appId '{1}')." -f $application.DisplayName, $application.AppId)

    if ($PSCmdlet.ShouldProcess($application.DisplayName, 'Update required resource access')) {
        Update-MgApplication -ApplicationId $application.Id -RequiredResourceAccess $requiredResourceAccess -ErrorAction Stop
        $applicationStatus = 'Updated'
    }
}
else {
    if (-not $PSCmdlet.ShouldProcess($DisplayName, 'Create app registration')) {
        return
    }

    $application = New-MgApplication `
        -DisplayName $DisplayName `
        -SignInAudience 'AzureADMyOrg' `
        -RequiredResourceAccess $requiredResourceAccess `
        -Notes 'Read-only collector for the Modern Workplace tenant assessment.' `
        -ErrorAction Stop
    $applicationStatus = 'Created'
    Write-Verbose ("Created app registration '{0}' (appId '{1}')." -f $application.DisplayName, $application.AppId)
}

$certificateThumbprint = $null
$certificateNotAfter = $null
if (-not $SkipCertificate) {
    if ($PSCmdlet.ShouldProcess($CertificateSubject, 'Create and upload self-signed certificate')) {
        $certificate = New-SelfSignedCertificate `
            -Subject $CertificateSubject `
            -CertStoreLocation 'Cert:\CurrentUser\My' `
            -KeyExportPolicy Exportable `
            -KeySpec Signature `
            -KeyLength 2048 `
            -HashAlgorithm 'SHA256' `
            -NotAfter (Get-Date).AddYears($CertificateValidityYears) `
            -ErrorAction Stop

        $certificateThumbprint = $certificate.Thumbprint
        $certificateNotAfter = $certificate.NotAfter

        $existingKeyCredentials = @()
        if ($application.PSObject.Properties['KeyCredentials'] -and $application.KeyCredentials) {
            $existingKeyCredentials = @(
                foreach ($keyCredential in $application.KeyCredentials) {
                    @{
                        type        = $keyCredential.Type
                        usage       = $keyCredential.Usage
                        key         = $keyCredential.Key
                        displayName = $keyCredential.DisplayName
                    }
                }
            )
        }

        $newKeyCredential = @{
            type        = 'AsymmetricX509Cert'
            usage       = 'Verify'
            key         = $certificate.RawData
            displayName = $CertificateSubject
        }

        Update-MgApplication -ApplicationId $application.Id -KeyCredentials @($existingKeyCredentials + $newKeyCredential) -ErrorAction Stop
        Write-Verbose ("Uploaded certificate '{0}' with thumbprint '{1}'." -f $CertificateSubject, $certificateThumbprint)
    }
}

$servicePrincipals = @(Get-MgServicePrincipal -Filter "appId eq '$($application.AppId)'" -All -ErrorAction Stop)
if ($servicePrincipals.Count -gt 0) {
    $clientServicePrincipal = $servicePrincipals[0]
    $servicePrincipalStatus = 'Existing'
}
else {
    if (-not $PSCmdlet.ShouldProcess($application.DisplayName, 'Create service principal')) {
        return
    }

    $clientServicePrincipal = New-MgServicePrincipal -AppId $application.AppId -ErrorAction Stop
    $servicePrincipalStatus = 'Created'
}

# Directory replication is eventually consistent, so give the new service principal a moment
# to become visible to the appRoleAssignment and roleAssignment endpoints.
$propagationAttempt = 0
while ($propagationAttempt -lt 10) {
    $propagationAttempt++
    $probe = @(Get-MgServicePrincipal -Filter "appId eq '$($application.AppId)'" -All -ErrorAction SilentlyContinue)
    if ($probe.Count -gt 0) { break }
    Start-Sleep -Seconds 3
}

$consentResults = New-Object System.Collections.Generic.List[pscustomobject]
if ($SkipAdminConsent) {
    Write-Warning 'Admin consent was skipped. Grant consent in the Entra admin center before running the assessment, or the collector will fail its permission preflight.'
}
else {
    $existingAssignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $clientServicePrincipal.Id -All -ErrorAction Stop)

    foreach ($resource in $resourceAccessPlan) {
        foreach ($role in $resource.Roles) {
            $alreadyGranted = @(
                $existingAssignments |
                    Where-Object { $_.ResourceId -eq $resource.ServicePrincipal.Id -and $_.AppRoleId -eq $role.RoleId }
            ).Count -gt 0

            if ($alreadyGranted) {
                $consentResults.Add([pscustomobject]@{
                        Resource   = $resource.ServicePrincipal.DisplayName
                        Permission = $role.Name
                        Status     = 'AlreadyGranted'
                    }) | Out-Null
                continue
            }

            if (-not $PSCmdlet.ShouldProcess($role.Name, 'Grant admin consent')) {
                continue
            }

            try {
                New-MgServicePrincipalAppRoleAssignment `
                    -ServicePrincipalId $clientServicePrincipal.Id `
                    -PrincipalId $clientServicePrincipal.Id `
                    -ResourceId $resource.ServicePrincipal.Id `
                    -AppRoleId $role.RoleId `
                    -ErrorAction Stop | Out-Null

                $consentResults.Add([pscustomobject]@{
                        Resource   = $resource.ServicePrincipal.DisplayName
                        Permission = $role.Name
                        Status     = 'Granted'
                    }) | Out-Null
            }
            catch {
                Write-Warning ("Could not grant '{0}' on '{1}'. Details: {2}" -f $role.Name, $resource.ServicePrincipal.DisplayName, $_.Exception.Message)
                $consentResults.Add([pscustomobject]@{
                        Resource   = $resource.ServicePrincipal.DisplayName
                        Permission = $role.Name
                        Status     = 'Failed'
                    }) | Out-Null
            }
        }
    }
}

$directoryRoleName = $(if ($AccessModel -eq 'GlobalReader') { $GlobalReaderRoleName } else { $null })
$directoryRoleStatus = $(if ($AccessModel -eq 'GlobalReader') { 'Pending' } else { 'NotAssignedByDesign' })
$directoryRoleAssignmentId = $null
if ($AccessModel -eq 'GlobalReader') {
    $escapedRoleName = $GlobalReaderRoleName.Replace("'", "''")
    $roleDefinitions = @(Get-MgRoleManagementDirectoryRoleDefinition -Filter "displayName eq '$escapedRoleName'" -All -ErrorAction Stop)
    if ($roleDefinitions.Count -ne 1) {
        throw "Expected exactly one directory role definition named '$GlobalReaderRoleName' in tenant '$resolvedTenantId' but found $($roleDefinitions.Count)."
    }

    $roleDefinition = $roleDefinitions[0]
    $assignmentFilter = "principalId eq '$($clientServicePrincipal.Id)' and roleDefinitionId eq '$($roleDefinition.Id)'"
    $existingRoleAssignments = @(
        Get-MgRoleManagementDirectoryRoleAssignment -Filter $assignmentFilter -All -ErrorAction Stop |
            Where-Object { $_.DirectoryScopeId -eq '/' }
    )

    if ($existingRoleAssignments.Count -gt 0) {
        $directoryRoleStatus = 'AlreadyAssigned'
        $directoryRoleAssignmentId = $existingRoleAssignments[0].Id
    }
    elseif ($PSCmdlet.ShouldProcess($GlobalReaderRoleName, 'Assign read-only directory role to service principal')) {
        try {
            $roleAssignment = New-MgRoleManagementDirectoryRoleAssignment `
                -DirectoryScopeId '/' `
                -PrincipalId $clientServicePrincipal.Id `
                -RoleDefinitionId $roleDefinition.Id `
                -ErrorAction Stop

            $directoryRoleStatus = 'Assigned'
            $directoryRoleAssignmentId = $roleAssignment.Id
        }
        catch {
            $directoryRoleStatus = 'Failed'
            Write-Warning ("Could not assign '{0}' to the service principal. Details: {1}" -f $GlobalReaderRoleName, $_.Exception.Message)
        }
    }
}

# Access models are exclusive. Exchange Administrator was used by older versions of this
# script, so remove that legacy assignment when Global Reader is selected.
$removedRoleNames = New-Object System.Collections.Generic.List[string]
if ($AccessModel -eq 'GlobalReader') {
    $legacyRoleDefinitions = @(Get-MgRoleManagementDirectoryRoleDefinition -Filter "displayName eq '$LegacyExchangeRoleName'" -All -ErrorAction Stop)
    if ($legacyRoleDefinitions.Count -eq 1) {
        $legacyAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$($clientServicePrincipal.Id)' and roleDefinitionId eq '$($legacyRoleDefinitions[0].Id)'" -All -ErrorAction Stop | Where-Object DirectoryScopeId -eq '/')
        foreach ($legacyAssignment in $legacyAssignments) {
            if ($PSCmdlet.ShouldProcess($LegacyExchangeRoleName, 'Remove conflicting legacy directory role assignment')) {
                Remove-MgRoleManagementDirectoryRoleAssignment -UnifiedRoleAssignmentId $legacyAssignment.Id -ErrorAction Stop
                $removedRoleNames.Add($LegacyExchangeRoleName) | Out-Null
            }
        }
    }
}

return [pscustomobject]@{
    TenantId                 = $resolvedTenantId
    DisplayName              = $application.DisplayName
    AppId                    = $application.AppId
    ApplicationObjectId      = $application.Id
    ApplicationStatus        = $applicationStatus
    ServicePrincipalObjectId = $clientServicePrincipal.Id
    ServicePrincipalStatus   = $servicePrincipalStatus
    PermissionSet            = $PermissionSet
    AccessModel              = $AccessModel
    CertificateThumbprint    = $certificateThumbprint
    CertificateSubject       = $(if ($SkipCertificate) { $null } else { $CertificateSubject })
    CertificateNotAfter      = $certificateNotAfter
    DirectoryRoleName        = $directoryRoleName
    DirectoryRoleStatus      = $directoryRoleStatus
    DirectoryRoleAssignmentId = $directoryRoleAssignmentId
    RemovedDirectoryRoles    = @($removedRoleNames)
    WorkloadRbacRequired     = ($AccessModel -eq 'LeastPrivilege' -and -not $SkipExchange)
    ConsentResults           = @($consentResults)
}
