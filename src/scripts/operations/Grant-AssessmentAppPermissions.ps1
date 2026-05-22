<#
.SYNOPSIS
    Enforces the exact set of Microsoft Graph application permissions required by the
    M365 Tenant Assessment tool — adds missing permissions and removes any extras.

.DESCRIPTION
    Connects to Microsoft Graph, compares the app's current Graph permissions against the
    required list, then:
      - Grants any permissions that are missing
      - Revokes any Graph permissions that are present but not in the required list
      - Updates RequiredResourceAccess on the app registration to match

    Only Microsoft Graph permissions are touched. Permissions granted to other APIs on the
    same app registration are left alone.

    Requires: Microsoft.Graph PowerShell module (Install-Module Microsoft.Graph)
    Requires: Global Administrator or Application Administrator with
              Application.ReadWrite.All and AppRoleAssignment.ReadWrite.All

.PARAMETER TenantId
    Entra tenant ID (GUID).

.PARAMETER ClientId
    App registration client ID (GUID) for the assessment app.

.PARAMETER WhatIf
    Show what would be added/removed without making any changes.

.EXAMPLE
    .\Grant-AssessmentAppPermissions.ps1 -TenantId "25f736c7-..." -ClientId "b80ca099-..."

.EXAMPLE
    .\Grant-AssessmentAppPermissions.ps1 -TenantId "25f736c7-..." -ClientId "b80ca099-..." -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$ClientId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Exact required permission set — nothing more, nothing less
# ---------------------------------------------------------------------------

$RequiredPermissions = @(
    # Core
    "User.Read.All"
    "AuditLog.Read.All"
    "Organization.Read.All"
    "Domain.Read.All"
    "Group.Read.All"
    "GroupMember.Read.All"                        # group member counts + license group membership
    "Device.Read.All"
    "DeviceManagementManagedDevices.Read.All"     # Intune device compliance, management agent, model
    "RoleManagement.Read.Directory"
    "OnPremDirectorySynchronization.Read.All"     # detailed sync features (writeback, version, etc.)
    "Policy.Read.All"
    "Application.Read.All"
    "Reports.Read.All"
    "ReportSettings.Read.All"                     # detect report anonymization (GUIDs vs names in usage reports)
    "Team.ReadBasic.All"
    "TeamMember.Read.All"
    "Channel.ReadBasic.All"
    "SharePointTenantSettings.Read.All"
    "SecurityEvents.Read.All"
    # Optional — tool degrades gracefully without these, but they unlock additional data
    "IdentityRiskyUser.Read.All"
    "MailboxSettings.Read"
    "Place.Read.All"
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Write-Line {
    param([string]$Symbol, [string]$Color, [string]$Message)
    Write-Host "  " -NoNewline
    Write-Host $Symbol -ForegroundColor $Color -NoNewline
    Write-Host "  $Message"
}

# ---------------------------------------------------------------------------
# Connect
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "  M365 Assessment - Enforce App Permissions" -ForegroundColor Cyan
Write-Host "  ==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Tenant : $TenantId"
Write-Host "  App    : $ClientId"
Write-Host ""

Connect-MgGraph -TenantId $TenantId `
    -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All" `
    -NoWelcome

# ---------------------------------------------------------------------------
# Resolve principals
# ---------------------------------------------------------------------------

$appSp = Get-MgServicePrincipal -Filter "appId eq '$ClientId'" -ErrorAction SilentlyContinue
if (-not $appSp) {
    Write-Host "  Creating service principal for app..." -ForegroundColor DarkGray
    $appSp = New-MgServicePrincipal -AppId $ClientId
}

$graphAppId = "00000003-0000-0000-c000-000000000000"
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -Property Id,AppRoles

# Build name -> AppRole lookup
$byName = @{}
foreach ($r in $graphSp.AppRoles) { $byName[$r.Value] = $r }

# Resolve every required name to its AppRole — fail fast if a name is wrong
$requiredRoles = foreach ($name in $RequiredPermissions) {
    if (-not $byName.ContainsKey($name)) {
        Write-Error "Permission '$name' not found in Microsoft Graph app roles. Check the name."
        return
    }
    $byName[$name]
}
$requiredRoleIds = $requiredRoles | Select-Object -ExpandProperty Id | ForEach-Object { $_.ToString() }

# ---------------------------------------------------------------------------
# Compare current vs required
# ---------------------------------------------------------------------------

# Only look at assignments for Microsoft Graph (ignore other APIs)
$currentAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appSp.Id -All |
    Where-Object { $_.ResourceId -eq $graphSp.Id }

$currentRoleIds = $currentAssignments | Select-Object -ExpandProperty AppRoleId |
    ForEach-Object { $_.ToString() }

$toAdd    = $requiredRoles    | Where-Object { $_.Id.ToString() -notin $currentRoleIds }
$toRemove = $currentAssignments | Where-Object { $_.AppRoleId.ToString() -notin $requiredRoleIds }

# Build a name lookup for display (roleId -> name)
$idToName = @{}
foreach ($r in $graphSp.AppRoles) { $idToName[$r.Id.ToString()] = $r.Value }

# ---------------------------------------------------------------------------
# Print plan
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "  Current Graph permissions on this app:" -ForegroundColor DarkGray
if ($currentAssignments) {
    foreach ($a in ($currentAssignments | Sort-Object { $idToName[$_.AppRoleId.ToString()] })) {
        $name = $idToName[$a.AppRoleId.ToString()] ?? $a.AppRoleId.ToString()
        $flag = if ($a.AppRoleId.ToString() -in $requiredRoleIds) { "v" } else { "x" }
        $col  = if ($flag -eq "v") { "Green" } else { "Red" }
        Write-Line $flag $col $name
    }
} else {
    Write-Host "    (none)" -ForegroundColor DarkGray
}

Write-Host ""
if ($toAdd) {
    Write-Host "  To add ($($toAdd.Count)):" -ForegroundColor DarkGray
    foreach ($r in ($toAdd | Sort-Object Value)) { Write-Line "+" "Cyan" $r.Value }
} else {
    Write-Host "  Nothing to add." -ForegroundColor DarkGray
}

Write-Host ""
if ($toRemove) {
    Write-Host "  To remove ($($toRemove.Count)):" -ForegroundColor DarkGray
    foreach ($a in $toRemove) {
        $name = $idToName[$a.AppRoleId.ToString()] ?? $a.AppRoleId.ToString()
        Write-Line "-" "Yellow" $name
    }
} else {
    Write-Host "  Nothing to remove." -ForegroundColor DarkGray
}

if (-not $toAdd -and -not $toRemove) {
    Write-Host ""
    Write-Host "  App permissions are already correct. No changes needed." -ForegroundColor Green
    Disconnect-MgGraph | Out-Null
    return
}

Write-Host ""
if ($WhatIfPreference) {
    Write-Host "  WhatIf: no changes made." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    return
}

# ---------------------------------------------------------------------------
# Apply changes
# ---------------------------------------------------------------------------

Write-Host "  Applying changes..." -ForegroundColor DarkGray
Write-Host ""

$addedCount   = 0
$removedCount = 0

foreach ($role in $toAdd) {
    if ($PSCmdlet.ShouldProcess($role.Value, "Grant application permission")) {
        try {
            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $appSp.Id `
                -PrincipalId        $appSp.Id `
                -ResourceId         $graphSp.Id `
                -AppRoleId          $role.Id `
                | Out-Null
            Write-Line "+" "Cyan" "$($role.Value)  granted"
            $addedCount++
        } catch {
            Write-Line "x" "Red" "$($role.Value)  FAILED: $_"
        }
    }
}

foreach ($assignment in $toRemove) {
    $name = $idToName[$assignment.AppRoleId.ToString()] ?? $assignment.AppRoleId.ToString()
    if ($PSCmdlet.ShouldProcess($name, "Revoke application permission")) {
        try {
            Remove-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId    $appSp.Id `
                -AppRoleAssignmentId   $assignment.Id
            Write-Line "-" "Yellow" "$name  revoked"
            $removedCount++
        } catch {
            Write-Line "x" "Red" "$name  FAILED to remove: $_"
        }
    }
}

# ---------------------------------------------------------------------------
# Sync RequiredResourceAccess on the app registration to match
# ---------------------------------------------------------------------------

$appReg = Get-MgApplication -Filter "appId eq '$ClientId'"

# Build the exact ResourceAccess list for Microsoft Graph
$exactGraphAccess = $requiredRoleIds | ForEach-Object {
    @{ Id = [guid]$_; Type = "Role" }
}

# Preserve ResourceAccess entries for all other APIs untouched
$otherApiAccess = $appReg.RequiredResourceAccess |
    Where-Object { $_.ResourceAppId -ne $graphAppId }

$newRequiredAccess = @(
    @{
        ResourceAppId    = $graphAppId
        ResourceAccess   = $exactGraphAccess
    }
) + @($otherApiAccess)

Update-MgApplication -ApplicationId $appReg.Id -RequiredResourceAccess $newRequiredAccess
Write-Host ""
Write-Host "  App registration manifest updated." -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "  Done. $addedCount added, $removedCount removed." -ForegroundColor Cyan
Write-Host ""

Disconnect-MgGraph | Out-Null
