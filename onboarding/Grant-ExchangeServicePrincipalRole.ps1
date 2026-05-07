[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$CustomerTenantId,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$AppId,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$RoleName = 'Exchange Administrator'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

foreach ($commandName in @(
        'Connect-MgGraph',
        'Get-MgServicePrincipal',
        'Get-MgRoleManagementDirectoryRoleDefinition',
        'Get-MgRoleManagementDirectoryRoleAssignment',
        'New-MgRoleManagementDirectoryRoleAssignment'
    )) {
    Assert-M365GraphCommand -Name $commandName
}

Write-Warning 'The signed-in customer admin must have enough rights to assign directory roles, normally Privileged Role Administrator or Global Administrator. If PIM is used, activate the role before continuing.'

try {
    Connect-MgGraph -TenantId $CustomerTenantId -Scopes 'Application.Read.All', 'RoleManagement.ReadWrite.Directory' -ErrorAction Stop | Out-Null
}
catch {
    throw "Could not connect to Microsoft Graph tenant '$CustomerTenantId'. Sign in as a customer admin with RoleManagement.ReadWrite.Directory and sufficient directory-role privileges. Details: $($_.Exception.Message)"
}

$appIdFilter = "appId eq '$AppId'"
Write-Verbose ("Resolving customer service principal with filter: {0}" -f $appIdFilter)
$servicePrincipals = @(Get-MgServicePrincipal -Filter $appIdFilter -All -ErrorAction Stop)
if ($servicePrincipals.Count -eq 0) {
    throw "No service principal was found in customer tenant '$CustomerTenantId' for appId '$AppId'. Confirm tenant-wide admin consent completed for your multitenant app."
}

if ($servicePrincipals.Count -gt 1) {
    throw "More than one service principal was found for appId '$AppId' in tenant '$CustomerTenantId'. Resolve the duplicate service principals before assigning Exchange roles."
}

$servicePrincipal = $servicePrincipals[0]

$escapedRoleName = $RoleName.Replace("'", "''")
$roleFilter = "displayName eq '$escapedRoleName'"
Write-Verbose ("Resolving role definition with filter: {0}" -f $roleFilter)
$roleDefinitions = @(Get-MgRoleManagementDirectoryRoleDefinition -Filter $roleFilter -All -ErrorAction Stop)
if ($roleDefinitions.Count -eq 0) {
    throw "Directory role '$RoleName' was not found in tenant '$CustomerTenantId'. Use an Exchange-supported directory role such as 'Exchange Administrator'."
}

if ($roleDefinitions.Count -gt 1) {
    throw "More than one directory role definition matched '$RoleName' in tenant '$CustomerTenantId'. Use the exact Exchange-supported role display name."
}

$roleDefinition = $roleDefinitions[0]

$assignmentFilter = "principalId eq '$($servicePrincipal.Id)' and roleDefinitionId eq '$($roleDefinition.Id)'"
Write-Verbose ("Checking existing role assignments with filter: {0}" -f $assignmentFilter)
$existingAssignments = @(
    Get-MgRoleManagementDirectoryRoleAssignment -Filter $assignmentFilter -All -ErrorAction Stop |
        Where-Object { $_.DirectoryScopeId -eq '/' }
)

if ($existingAssignments.Count -gt 0) {
    $assignment = $existingAssignments[0]
    $assignmentStatus = 'AlreadyAssigned'
    Write-Verbose ("Service principal '{0}' already has role '{1}' at directory scope '/'." -f $servicePrincipal.Id, $RoleName)
}
else {
    Write-Verbose ("Assigning role '{0}' to service principal '{1}'." -f $RoleName, $servicePrincipal.Id)
    try {
        $assignment = New-MgRoleManagementDirectoryRoleAssignment `
            -DirectoryScopeId '/' `
            -PrincipalId $servicePrincipal.Id `
            -RoleDefinitionId $roleDefinition.Id `
            -ErrorAction Stop
        $assignmentStatus = 'Assigned'
    }
    catch {
        throw "Could not assign role '$RoleName' to service principal '$($servicePrincipal.Id)' in tenant '$CustomerTenantId'. Confirm the signed-in admin is Privileged Role Administrator or Global Administrator and has RoleManagement.ReadWrite.Directory. Details: $($_.Exception.Message)"
    }
}

return [pscustomobject]@{
    CustomerTenantId            = $CustomerTenantId
    AppId                       = $AppId
    ServicePrincipalObjectId    = $servicePrincipal.Id
    ServicePrincipalDisplayName = $servicePrincipal.DisplayName
    RoleName                    = $roleDefinition.DisplayName
    RoleDefinitionId            = $roleDefinition.Id
    RoleAssignmentId            = $assignment.Id
    RoleAssignmentStatus        = $assignmentStatus
}
