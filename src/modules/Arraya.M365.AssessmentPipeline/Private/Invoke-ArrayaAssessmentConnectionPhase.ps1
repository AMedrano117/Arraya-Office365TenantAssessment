function Resolve-ArrayaAssessmentPipelineRequestedAuthMode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$AuthMode,
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,
        [Parameter(Mandatory = $false)]
        [string]$ClientSecret
    )

    if (-not [string]::IsNullOrWhiteSpace($AuthMode)) { return $AuthMode }
    if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) { return 'Certificate' }
    if (-not [string]::IsNullOrWhiteSpace($ClientSecret)) { return 'ClientSecret' }
    return 'Interactive'
}

function Get-ArrayaAssessmentPipelineGraphDelegatedScopes {
    [CmdletBinding()]
    param()

    return @(
        'Organization.Read.All',
        'User.Read.All',
        'AuditLog.Read.All',
        'Group.Read.All',
        'GroupMember.Read.All',
        'RoleManagement.Read.Directory',
        'Domain.Read.All',
        'Device.Read.All',
        'Reports.Read.All',
        'ReportSettings.Read.All',
        'Policy.Read.All',
        'CrossTenantInformation.ReadBasic.All',
        'SecurityEvents.Read.All',
        'Application.Read.All',
        'Sites.Read.All',
        'SharePointTenantSettings.Read.All',
        'Team.ReadBasic.All',
        'Channel.ReadBasic.All',
        'OnPremDirectorySynchronization.Read.All'
    )
}

function Resolve-ArrayaAssessmentPipelineAuthWorkloadPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
        [string]$AuthMode,
        [Parameter(Mandatory = $false)]
        [bool]$NeedsGovernanceCompliancePolicies = $false,
        [Parameter(Mandatory = $false)]
        [bool]$NeedsTeamsInventory = $false,
        [Parameter(Mandatory = $false)]
        [bool]$NeedsTeamsVoice = $false,
        [Parameter(Mandatory = $false)]
        [bool]$NeedsSharePointData = $true
    )

    $requiredWorkloads = New-Object System.Collections.Generic.List[string]
    $requiredWorkloads.Add('Graph') | Out-Null
    $requiredWorkloads.Add('ExchangeOnline') | Out-Null
    if ($NeedsGovernanceCompliancePolicies) {
        $requiredWorkloads.Add('PurviewCompliance') | Out-Null
    }

    $sharePointMode = if ($AuthMode -eq 'Interactive' -and $NeedsSharePointData) { 'OptionalModuleConnect' } else { 'GraphFallback' }
    $teamsMode = if ($AuthMode -eq 'Interactive' -and ($NeedsTeamsInventory -or $NeedsTeamsVoice)) { 'OptionalPowerShellConnect' } else { 'GraphOnly' }

    return [pscustomobject][ordered]@{
        AuthenticationType = $AuthMode
        RequiredWorkloads  = @($requiredWorkloads)
        FallbackWorkloads  = @(
            if ($sharePointMode -eq 'GraphFallback') { 'SharePointOnline' }
            if ($teamsMode -eq 'GraphOnly' -and ($NeedsTeamsInventory -or $NeedsTeamsVoice)) { 'Teams' }
        )
        Workloads = [ordered]@{
            Graph = [ordered]@{ Required = $true; Connect = $true; Mode = 'Required' }
            ExchangeOnline = [ordered]@{ Required = $true; Connect = $true; Mode = 'Required' }
            PurviewCompliance = [ordered]@{
                Required = [bool]$NeedsGovernanceCompliancePolicies
                Connect  = [bool]$NeedsGovernanceCompliancePolicies
                Mode     = if ($NeedsGovernanceCompliancePolicies) { 'Required' } else { 'SkippedByProfile' }
            }
            SharePointOnline = [ordered]@{
                Required = $false
                Connect  = [bool]$NeedsSharePointData
                Mode     = $sharePointMode
            }
            Teams = [ordered]@{
                Required = $false
                Connect  = [bool]($NeedsTeamsInventory -or $NeedsTeamsVoice)
                Mode     = $teamsMode
            }
        }
    }
}

function Test-ArrayaAssessmentPipelineExchangeCmdletsAvailable {
    [CmdletBinding()]
    param()

    return (
        [bool](Get-Command -Name 'Get-EXOMailbox' -ErrorAction SilentlyContinue) -and
        [bool](Get-Command -Name 'Get-UnifiedGroup' -ErrorAction SilentlyContinue)
    )
}

function Test-ArrayaAssessmentPipelinePurviewCmdletsAvailable {
    [CmdletBinding()]
    param()

    return (
        [bool](Get-Command -Name 'Get-RetentionCompliancePolicy' -ErrorAction SilentlyContinue) -and
        [bool](Get-Command -Name 'Get-DlpCompliancePolicy' -ErrorAction SilentlyContinue)
    )
}

function Test-ArrayaAssessmentPipelinePurviewSessionReady {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$ProbeCommands
    )

    if (-not (Test-ArrayaAssessmentPipelinePurviewCmdletsAvailable)) {
        return $false
    }

    if (-not $ProbeCommands) {
        return $true
    }

    foreach ($commandName in @('Get-RetentionCompliancePolicy', 'Get-DlpCompliancePolicy')) {
        try {
            switch ($commandName) {
                'Get-RetentionCompliancePolicy' { @(Get-RetentionCompliancePolicy -ErrorAction Stop | Select-Object -First 1) | Out-Null }
                'Get-DlpCompliancePolicy' { @(Get-DlpCompliancePolicy -ErrorAction Stop | Select-Object -First 1) | Out-Null }
            }
        }
        catch {
            return $false
        }
    }

    return $true
}

function Get-ArrayaAssessmentPipelineSharePointAdminUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$InitialDomain
    )

    if ([string]::IsNullOrWhiteSpace($InitialDomain)) { return $null }
    $tenantName = $InitialDomain
    if ($tenantName -match '^(?<prefix>[^.]+)\.onmicrosoft\.com$') { $tenantName = $Matches['prefix'] }
    if ([string]::IsNullOrWhiteSpace($tenantName)) { return $null }
    return ('https://{0}-admin.sharepoint.com' -f $tenantName)
}

function Get-ArrayaAssessmentPipelineGraphResource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [Parameter(Mandatory = $false)]
        [int]$PageSize = 999,
        [Parameter(Mandatory = $false)]
        [string]$Activity = 'Fetching data from Microsoft Graph',
        [Parameter(Mandatory = $false)]
        [switch]$SuppressProgress,
        [Parameter(Mandatory = $false)]
        [switch]$SuppressAccessDeniedWarning
    )

    return Office365Custom\Get-GraphData `
        -Uri $Uri `
        -PageSize $PageSize `
        -Activity $Activity `
        -MaxRetries 5 `
        -SuppressProgress:$SuppressProgress `
        -SuppressAccessDeniedWarning:$SuppressAccessDeniedWarning
}

function Export-ArrayaAssessmentPipelineGraphReportCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [Parameter(Mandatory = $false)]
        [string]$Activity = 'Downloading Microsoft Graph report CSV'
    )

    $tempCsvPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("arraya-pipeline-graph-report-" + [guid]::NewGuid().ToString('N') + '.csv')
    try {
        if (Get-Command -Name 'Invoke-MgGraphRequest' -ErrorAction SilentlyContinue) {
            Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputFilePath $tempCsvPath -ProgressAction SilentlyContinue -ErrorAction Stop | Out-Null
        }
        else {
            $headers = if ($global:GraphHeaders) { $global:GraphHeaders } else { @{} }
            if ($headers.Count -eq 0) {
                throw "No Graph authentication context is available for '$Activity'."
            }
            $response = Invoke-WebRequest -Uri $Uri -Headers $headers -Method GET -MaximumRedirection 5 -ErrorAction Stop
            Set-Content -Path $tempCsvPath -Value $response.Content -Encoding UTF8 -Force
        }

        return @(Import-Csv -Path $tempCsvPath -ErrorAction Stop)
    }
    finally {
        if (Test-Path -Path $tempCsvPath) {
            Remove-Item -Path $tempCsvPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-ArrayaAssessmentPipelineGraphAdminReportSettings {
    [CmdletBinding()]
    param()

    try {
        $response = Get-ArrayaAssessmentPipelineGraphResource -Uri 'https://graph.microsoft.com/v1.0/admin/reportSettings' -Activity 'Admin report settings' -SuppressProgress
        $displayConcealedNames = $null
        if ($response -and $response.PSObject.Properties['displayConcealedNames']) {
            try { $displayConcealedNames = [bool]$response.displayConcealedNames } catch {}
        }

        return [pscustomobject]@{
            Available = ($null -ne $response)
            DisplayConcealedNames = $displayConcealedNames
            ErrorMessage = $null
        }
    }
    catch {
        return [pscustomobject]@{
            Available = $false
            DisplayConcealedNames = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Get-ArrayaAssessmentPipelineTenantOrganization {
    [CmdletBinding()]
    param()

    $organization = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
    $initialDomain = $null
    if ($organization -and $organization.VerifiedDomains) {
        $initialDomainRow = @($organization.VerifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -First 1)
        if ($initialDomainRow.Count -gt 0) {
            $initialDomain = [string](Get-ArrayaObjectValue -Object $initialDomainRow[0] -Names @('Name', 'Id'))
        }
    }

    if ($organization -and -not $organization.PSObject.Properties['InitialDomain']) {
        $organization | Add-Member -MemberType NoteProperty -Name InitialDomain -Value $initialDomain -Force
    }

    return $organization
}

function ConvertFrom-ArrayaAssessmentPipelineJwtPayload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Token
    )

    if ([string]::IsNullOrWhiteSpace($Token)) { return $null }
    $tokenParts = $Token.Split('.')
    if ($tokenParts.Count -lt 2) { return $null }
    $payloadSegment = $tokenParts[1].Replace('-', '+').Replace('_', '/')
    switch ($payloadSegment.Length % 4) {
        2 { $payloadSegment += '==' }
        3 { $payloadSegment += '=' }
    }
    try {
        $payloadBytes = [System.Convert]::FromBase64String($payloadSegment)
        $payloadJson = [System.Text.Encoding]::UTF8.GetString($payloadBytes)
        return ($payloadJson | ConvertFrom-Json -Depth 20)
    }
    catch {
        return $null
    }
}

function Get-ArrayaAssessmentPipelineGrantedGraphPermissions {
    [CmdletBinding()]
    param()

    $permissionSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $addPermission = {
        param([string]$PermissionName)
        if (-not [string]::IsNullOrWhiteSpace($PermissionName)) {
            $null = $permissionSet.Add($PermissionName.Trim())
        }
    }

    $mgContext = Get-MgContext -ErrorAction SilentlyContinue
    if ($mgContext -and $mgContext.Scopes) {
        foreach ($scope in @($mgContext.Scopes)) { & $addPermission ([string]$scope) }
    }

    $graphTokenVariable = Get-Variable -Name GraphToken -Scope Global -ErrorAction SilentlyContinue
    $graphToken = if ($graphTokenVariable) { [string]$graphTokenVariable.Value } else { $null }
    if ([string]::IsNullOrWhiteSpace([string]$graphToken)) {
        try { $graphToken = Get-MgAccessToken -ErrorAction Stop } catch {}
    }

    $claims = ConvertFrom-ArrayaAssessmentPipelineJwtPayload -Token $graphToken
    if ($claims) {
        if ($claims.PSObject.Properties['scp']) {
            foreach ($scope in @([string]$claims.scp -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
                & $addPermission ([string]$scope)
            }
        }
        if ($claims.PSObject.Properties['roles']) {
            foreach ($role in @($claims.roles)) {
                & $addPermission ([string]$role)
            }
        }
    }

    return [pscustomobject]@{
        PermissionSet = $permissionSet
        Context       = $mgContext
    }
}

function Test-ArrayaAssessmentPipelinePermissionPreflight {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,
        [Parameter(Mandatory = $true)]
        [pscustomobject]$ConnectionResult,
        [Parameter(Mandatory = $true)]
        [pscustomobject]$WorkloadPlan,
        [Parameter(Mandatory = $false)]
        [ValidateSet('All', 'Graph', 'ExchangeOnline', 'Purview')]
        [string]$Workload = 'All'
    )

    $grantedPermissions = (Get-ArrayaAssessmentPipelineGrantedGraphPermissions).PermissionSet
    $failures = New-Object System.Collections.Generic.List[object]
    $warnings = New-Object System.Collections.Generic.List[object]

    $addResult = {
        param(
            [bool]$Blocking,
            [string]$Area,
            [string]$Requirement,
            [string]$NeededFor,
            [string]$Details
        )

        $resultRecord = [pscustomobject]@{
            Area        = $Area
            Requirement = $Requirement
            NeededFor   = $NeededFor
            Details     = $Details
        }

        if ($Blocking) {
            $failures.Add($resultRecord) | Out-Null
        }
        else {
            $warnings.Add($resultRecord) | Out-Null
        }
    }

    $graphChecks = @(
        @{ Requirement = 'Organization.Read.All'; NeededFor = 'tenant organization metadata'; Uri = 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName'; PageSize = 0; Blocking = $true }
        @{ Requirement = 'User.Read.All'; NeededFor = 'user inventory and hygiene analysis'; Uri = 'https://graph.microsoft.com/v1.0/users?$top=1&$select=id,userType,accountEnabled'; Blocking = $true }
        @{ Requirement = 'Group.Read.All'; NeededFor = 'group inventory'; Uri = 'https://graph.microsoft.com/v1.0/groups?$top=1&$select=id,displayName'; Blocking = $true }
        @{ Requirement = 'RoleManagement.Read.Directory'; NeededFor = 'admin role assignments'; Uri = 'https://graph.microsoft.com/v1.0/directoryRoles?$select=id,displayName,roleTemplateId'; PageSize = 0; Blocking = $true }
        @{ Requirement = 'Domain.Read.All'; NeededFor = 'accepted domain inventory'; Uri = 'https://graph.microsoft.com/v1.0/domains?$top=1&$select=id,isVerified'; Blocking = $true }
        @{ Requirement = 'Device.Read.All'; NeededFor = 'device inventory'; Uri = 'https://graph.microsoft.com/v1.0/devices?$top=1&$select=id,displayName'; Blocking = $true }
        @{ Requirement = 'Policy.Read.All'; NeededFor = 'Conditional Access and auth policy review'; Uri = 'https://graph.microsoft.com/v1.0/policies/authorizationPolicy?$select=id,allowInvitesFrom'; PageSize = 0; Blocking = $true }
        @{ Requirement = 'Application.Read.All'; NeededFor = 'enterprise application posture review'; Uri = 'https://graph.microsoft.com/v1.0/servicePrincipals?$top=1&$select=id,displayName'; Blocking = $true }
        @{ Requirement = 'Sites.Read.All'; NeededFor = 'SharePoint and OneDrive site inventory'; Uri = 'https://graph.microsoft.com/v1.0/sites/root?$select=id,webUrl'; PageSize = 0; Blocking = $true }
        @{ Requirement = 'SharePointTenantSettings.Read.All'; NeededFor = 'tenant-level SharePoint settings'; Uri = 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings'; PageSize = 0; Blocking = $true }
        @{ Requirement = 'SecurityEvents.Read.All'; NeededFor = 'Microsoft Secure Score collection'; Uri = 'https://graph.microsoft.com/v1.0/security/secureScores?$top=1'; Blocking = $true }
        @{ Requirement = 'Reports.Read.All'; NeededFor = 'usage and activity reports'; ReportUri = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D7')"; Blocking = $true }
        @{ Requirement = 'ReportSettings.Read.All'; NeededFor = 'report settings validation'; ReportSettings = $true; Blocking = $true }
        @{ Requirement = 'OnPremDirectorySynchronization.Read.All'; NeededFor = 'directory synchronization and password lifecycle review'; Uri = 'https://graph.microsoft.com/v1.0/directory/onPremisesSynchronization'; PageSize = 0; Blocking = $false }
    )
    if ([bool]$WorkloadPlan.Workloads.Teams.Connect) {
        $graphChecks += @{ Requirement = 'Team.ReadBasic.All'; NeededFor = 'Teams inventory'; Uri = 'https://graph.microsoft.com/v1.0/teams?$top=1&$select=id,displayName'; Blocking = $false }
        $graphChecks += @{ Requirement = 'Channel.ReadBasic.All'; NeededFor = 'Teams channel inventory'; Uri = 'https://graph.microsoft.com/v1.0/teams?$top=1&$select=id'; Blocking = $false }
    }

    $exchangeChecks = @(
        @{ Requirement = 'Exchange mailbox read access'; NeededFor = 'Exchange mailbox inventory'; Probe = { @(Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Select-Object -First 1) | Out-Null } }
    )
    if ([bool]$WorkloadPlan.Workloads.Teams.Connect -or [bool]$WorkloadPlan.Workloads.SharePointOnline.Connect) {
        $exchangeChecks += @{ Requirement = 'Exchange unified group read access'; NeededFor = 'group-backed messaging review'; Probe = { @(Get-UnifiedGroup -ResultSize 1 -ErrorAction Stop | Select-Object -First 1) | Out-Null } }
    }

    $selectedGraphChecks = @()
    $selectedExchangeChecks = @()
    $runPurviewCheck = $false
    switch ($Workload) {
        'Graph' { $selectedGraphChecks = $graphChecks }
        'ExchangeOnline' { $selectedExchangeChecks = $exchangeChecks }
        'Purview' { $runPurviewCheck = [bool]$WorkloadPlan.Workloads.PurviewCompliance.Required }
        default {
            $selectedGraphChecks = $graphChecks
            $selectedExchangeChecks = $exchangeChecks
            $runPurviewCheck = [bool]$WorkloadPlan.Workloads.PurviewCompliance.Required
        }
    }

    $progressId = 91
    $progressIndex = 0
    $progressTotal = @($selectedGraphChecks).Count + @($selectedExchangeChecks).Count + $(if ($runPurviewCheck) { 1 } else { 0 })

    try {
        foreach ($check in @($selectedGraphChecks)) {
            $progressIndex++
            Write-ProgressHelper -Total ([Math]::Max($progressTotal, 1)) -Id $progressId -Index $progressIndex -Activity ("Permission preflight: Graph: {0}" -f $check.Requirement)
            try {
                if ($check.ContainsKey('ReportUri')) {
                    Export-ArrayaAssessmentPipelineGraphReportCsv -Uri $check.ReportUri -Activity ('Permission preflight: ' + $check.Requirement) | Out-Null
                }
                elseif ($check.ContainsKey('ReportSettings')) {
                    $settings = Get-ArrayaAssessmentPipelineGraphAdminReportSettings
                    if (-not $settings.Available) {
                        throw $settings.ErrorMessage
                    }
                }
                elseif ($check.Requirement -eq 'Channel.ReadBasic.All') {
                    $teamsResponse = Get-ArrayaAssessmentPipelineGraphResource -Uri 'https://graph.microsoft.com/v1.0/teams?$top=1&$select=id' -Activity 'Permission preflight: Team lookup' -SuppressProgress -SuppressAccessDeniedWarning
                    $firstTeam = @($teamsResponse | Select-Object -First 1)
                    if ($firstTeam.Count -gt 0 -and $firstTeam[0].PSObject.Properties['id']) {
                        Get-ArrayaAssessmentPipelineGraphResource -Uri ('https://graph.microsoft.com/v1.0/teams/{0}/allChannels?$top=1&$select=displayName,membershipType' -f $firstTeam[0].id) -Activity 'Permission preflight: Channel.ReadBasic.All' -SuppressProgress -SuppressAccessDeniedWarning | Out-Null
                    }
                }
                else {
                    $pageSize = if ($check.ContainsKey('PageSize')) { [int]$check.PageSize } else { 999 }
                    Get-ArrayaAssessmentPipelineGraphResource -Uri $check.Uri -PageSize $pageSize -Activity ('Permission preflight: ' + $check.Requirement) -SuppressProgress -SuppressAccessDeniedWarning | Out-Null
                }
            }
            catch {
                & $addResult $check.Blocking 'Graph' $check.Requirement $check.NeededFor $_.Exception.Message
            }
        }

        foreach ($check in @($selectedExchangeChecks)) {
            $progressIndex++
            Write-ProgressHelper -Total ([Math]::Max($progressTotal, 1)) -Id $progressId -Index $progressIndex -Activity ("Permission preflight: Exchange Online: {0}" -f $check.Requirement)
            try {
                & $check.Probe
            }
            catch {
                & $addResult $true 'Exchange Online' $check.Requirement $check.NeededFor 'Ensure Exchange Online access is available for the selected auth mode.'
            }
        }

        if ($runPurviewCheck) {
            $progressIndex++
            Write-ProgressHelper -Total ([Math]::Max($progressTotal, 1)) -Id $progressId -Index $progressIndex -Activity 'Permission preflight: Purview: retention and DLP policy access'
            if (-not (Test-ArrayaAssessmentPipelinePurviewSessionReady -ProbeCommands)) {
                & $addResult $true 'Purview' 'Retention and DLP policy access' 'retention and DLP policy collection' 'Connect-IPPSSession did not expose the required compliance cmdlets in this session. Assign the service principal the Exchange Administrator role or run Connect-IPPSSession successfully in the current session and rerun the script.'
            }
        }
    }
    finally {
        Write-ProgressHelper -Total ([Math]::Max($progressTotal, 1)) -Id $progressId -Activity 'Permission preflight' -Completed
    }

    $summaryLabel = switch ($Workload) {
        'Graph' { 'Graph preflight summary' }
        'ExchangeOnline' { 'Exchange preflight summary' }
        'Purview' { 'Purview preflight summary' }
        default { 'Permission preflight summary' }
    }
    $successCount = [Math]::Max($progressTotal - ($failures.Count + $warnings.Count), 0)
    $summaryText = ("{0}: {1} successful, {2} remaining ({3} blocking, {4} non-blocking)." -f $summaryLabel, $successCount, ($failures.Count + $warnings.Count), $failures.Count, $warnings.Count)
    Write-Host $summaryText -ForegroundColor Cyan

    if ($failures.Count -gt 0) {
        Write-Host ($summaryLabel -replace 'summary', 'failed.') -ForegroundColor Red
        foreach ($failure in $failures) {
            Write-Host (" - [{0}] {1}" -f $failure.Area, $failure.Requirement) -ForegroundColor Yellow
            Write-Host ("   Needed for: {0}" -f $failure.NeededFor) -ForegroundColor DarkYellow
            Write-Host ("   Current issue: {0}" -f $failure.Details) -ForegroundColor DarkYellow
        }
        throw (($failures | ForEach-Object { "[{0}] {1}: {2}" -f $_.Area, $_.Requirement, $_.Details }) -join [Environment]::NewLine)
    }

    if ($warnings.Count -gt 0) {
        Write-Host ($summaryLabel -replace 'summary', 'warnings:') -ForegroundColor Yellow
        foreach ($warning in $warnings) {
            Write-Host (" - [{0}] {1}" -f $warning.Area, $warning.Requirement) -ForegroundColor Yellow
            Write-Host ("   Needed for: {0}" -f $warning.NeededFor) -ForegroundColor DarkYellow
            Write-Host ("   Current issue: {0}" -f $warning.Details) -ForegroundColor DarkYellow
        }
        Write-Host 'Hybrid sync and password lifecycle fields may be marked as not validated in current auth mode.' -ForegroundColor DarkYellow
    }

    return [pscustomobject]@{
        SummaryText     = $summaryText
        SuccessfulCount = $successCount
        FailureCount    = $failures.Count
        WarningCount    = $warnings.Count
    }
}

function Test-IsArrayaAssessmentPipelineExchangeInteractiveAuthFailure {
    [CmdletBinding()]
    param([AllowNull()][string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) { return $false }
    return (
        $Message -like '*RuntimeBroker*' -or
        $Message -like '*Object reference not set to an instance of an object*' -or
        $Message -like '*Error Acquiring Token*' -or
        $Message -like '*A window handle must be configured*' -or
        $Message -like '*broker*' -or
        $Message -like '*MSAL*'
    )
}

function Invoke-ArrayaAssessmentPipelineExchangeDelegatedConnect {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$BaseParameters,
        [Parameter(Mandatory = $false)]
        [string]$GraphAccount
    )

    $connectCommand = Get-Command -Name 'Connect-ExchangeOnline' -ErrorAction Stop
    $attempts = New-Object System.Collections.Generic.List[hashtable]
    $attempts.Add(@{ Label = 'interactive authentication'; Params = @{} }) | Out-Null
    if ($connectCommand.Parameters.ContainsKey('DisableWAM')) {
        $attempts.Add(@{ Label = 'interactive authentication with -DisableWAM'; Params = @{ DisableWAM = $true } }) | Out-Null
    }
    if ($connectCommand.Parameters.ContainsKey('Device')) {
        $deviceParams = @{ Device = $true }
        if (-not [string]::IsNullOrWhiteSpace($GraphAccount) -and $connectCommand.Parameters.ContainsKey('UserPrincipalName')) {
            $deviceParams.UserPrincipalName = $GraphAccount
        }
        $attempts.Add(@{ Label = 'device code authentication'; Params = $deviceParams }) | Out-Null
    }

    $lastErrorMessage = $null
    foreach ($attempt in $attempts) {
        $attemptParams = @{}
        foreach ($key in $BaseParameters.Keys) { $attemptParams[$key] = $BaseParameters[$key] }
        foreach ($key in $attempt.Params.Keys) { $attemptParams[$key] = $attempt.Params[$key] }
        try {
            Connect-ExchangeOnline @attemptParams | Out-Null
            return
        }
        catch {
            $lastErrorMessage = $_.Exception.Message
            if (-not (Test-IsArrayaAssessmentPipelineExchangeInteractiveAuthFailure -Message $lastErrorMessage)) {
                throw
            }
        }
    }

    throw "All delegated Exchange Online authentication attempts failed. Last error: $lastErrorMessage"
}

function Connect-ArrayaAssessmentPipelineGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
        [string]$AuthenticationType,
        [Parameter(Mandatory = $false)]
        [string]$TenantId,
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,
        [Parameter(Mandatory = $false)]
        [string]$ClientSecret
    )

    $existing = Get-MgContext -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host 'Graph already connected for this session.' -ForegroundColor Green
    }
    else {
        Write-Host 'Connecting assessment Graph session...' -ForegroundColor Cyan
        switch ($AuthenticationType) {
            'Certificate' {
                Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop | Out-Null
            }
            'ClientSecret' {
                $secureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
                $clientSecretCredential = [System.Management.Automation.PSCredential]::new($ClientId, $secureClientSecret)
                Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $clientSecretCredential -NoWelcome -ErrorAction Stop | Out-Null
            }
            default {
                $graphConnectParams = @{
                    Scopes      = (Get-ArrayaAssessmentPipelineGraphDelegatedScopes)
                    NoWelcome   = $true
                    ErrorAction = 'Stop'
                }
                if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $graphConnectParams.TenantId = $TenantId }
                try {
                    Connect-MgGraph @graphConnectParams | Out-Null
                }
                catch {
                    if (
                        $_.Exception.Message -like '*InteractiveBrowserCredential*' -and
                        (Get-Command -Name 'Connect-MgGraph' -ErrorAction SilentlyContinue).Parameters.ContainsKey('UseDeviceCode')
                    ) {
                        Write-Warning 'Interactive browser authentication failed. Retrying Microsoft Graph sign-in with device code.'
                        $graphConnectParams.UseDeviceCode = $true
                        Connect-MgGraph @graphConnectParams | Out-Null
                    }
                    else {
                        throw
                    }
                }
            }
        }
        Write-Host 'Graph connected for assessment collection.' -ForegroundColor Green
    }

    $tenantOrganization = Get-ArrayaAssessmentPipelineTenantOrganization

    try {
        $graphAccessToken = Get-MgAccessToken -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace([string]$graphAccessToken)) {
            $global:GraphToken = [string]$graphAccessToken
            $global:GraphHeaders = @{
                'Content-Type'     = 'application/json'
                'Authorization'    = "Bearer $graphAccessToken"
                'ConsistencyLevel' = 'eventual'
            }
        }
    }
    catch {}

    return [pscustomobject][ordered]@{
        Graph         = $true
        TenantName    = $tenantOrganization.DisplayName
        InitialDomain = $tenantOrganization.InitialDomain
    }
}

function Connect-ArrayaAssessmentPipelineExchange {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
        [string]$AuthenticationType,
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,
        [Parameter(Mandatory = $false)]
        [string]$InitialDomain
    )

    if (-not (Get-Module -ListAvailable -Name 'ExchangeOnlineManagement')) {
        throw "Exchange Online: The module 'ExchangeOnlineManagement' is not installed. Please install it before proceeding."
    }
    if (-not (Get-Module -Name 'ExchangeOnlineManagement')) {
        Import-Module 'ExchangeOnlineManagement' -ErrorAction Stop -WarningAction SilentlyContinue
    }

    $existingExchangeConnection = @()
    if (Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue) {
        try { $existingExchangeConnection = @(Get-ConnectionInformation -ErrorAction Stop) } catch {}
    }
    if ($existingExchangeConnection.Count -gt 0 -and (Test-ArrayaAssessmentPipelineExchangeCmdletsAvailable)) {
        Write-Host 'Exchange Online already connected for this session.' -ForegroundColor Green
        return [pscustomobject]@{ ExchangeOnline = $true }
    }

    Write-Host 'Connecting assessment Exchange Online session...' -ForegroundColor Cyan
    $exchangeConnectParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    if ($AuthenticationType -eq 'Certificate') {
        if ([string]::IsNullOrWhiteSpace($ClientId)) { throw 'Exchange Online certificate auth requires -ClientId.' }
        if ([string]::IsNullOrWhiteSpace($InitialDomain)) { throw 'Exchange Online certificate auth requires the tenant initial domain.' }
        $exchangeConnectParams.AppId = $ClientId
        $exchangeConnectParams.Organization = $InitialDomain
        $exchangeConnectParams.CertificateThumbprint = $CertificateThumbprint
        Connect-ExchangeOnline @exchangeConnectParams | Out-Null
    }
    else {
        $graphAccount = $null
        try { $graphAccount = (Get-MgContext -ErrorAction Stop).Account } catch {}
        if ($AuthenticationType -eq 'ClientSecret') {
            Write-Warning 'Exchange Online does not support client-secret-only app auth in this workflow. Falling back to delegated Exchange sign-in.'
        }
        Invoke-ArrayaAssessmentPipelineExchangeDelegatedConnect -BaseParameters $exchangeConnectParams -GraphAccount $graphAccount
    }

    Write-Host 'Exchange Online connected for assessment collection.' -ForegroundColor Green
    return [pscustomobject]@{ ExchangeOnline = $true }
}

function Connect-ArrayaAssessmentPipelineSharePoint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
        [string]$AuthenticationType,
        [Parameter(Mandatory = $false)]
        [string]$InitialDomain
    )

    if ($AuthenticationType -in @('Certificate', 'ClientSecret')) {
        return [pscustomobject]@{ SharePointOnline = $false; Status = 'GraphFallback'; Message = 'SharePoint module login is skipped by design for app-based auth. Graph fallback remains active.' }
    }
    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Online.SharePoint.PowerShell')) {
        return [pscustomobject]@{ SharePointOnline = $false; Status = 'SkippedByDesign'; Message = 'Microsoft.Online.SharePoint.PowerShell is not installed. Graph fallback remains active.' }
    }
    if (-not (Get-Module -Name 'Microsoft.Online.SharePoint.PowerShell')) {
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            Import-Module 'Microsoft.Online.SharePoint.PowerShell' -UseWindowsPowerShell -ErrorAction Stop
        }
        else {
            Import-Module 'Microsoft.Online.SharePoint.PowerShell' -ErrorAction Stop
        }
    }
    $spoAdminUrl = Get-ArrayaAssessmentPipelineSharePointAdminUrl -InitialDomain $InitialDomain
    if ([string]::IsNullOrWhiteSpace($spoAdminUrl)) {
        return [pscustomobject]@{ SharePointOnline = $false; Status = 'GraphFallback'; Message = 'Unable to resolve the SharePoint admin URL. Graph fallback remains active.' }
    }
    try {
        Write-Host 'Connecting assessment SharePoint admin session...' -ForegroundColor Cyan
        Connect-SPOService -Url $spoAdminUrl -ErrorAction Stop
        Write-Host 'SharePoint admin session connected.' -ForegroundColor Green
        return [pscustomobject]@{ SharePointOnline = $true; Status = 'Connected'; Message = 'Connected to SharePoint Online admin PowerShell.' }
    }
    catch {
        return [pscustomobject]@{ SharePointOnline = $false; Status = 'GraphFallback'; Message = $_.Exception.Message }
    }
}

function Connect-ArrayaAssessmentPipelineTeams {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
        [string]$AuthenticationType
    )

    if ($AuthenticationType -ne 'Interactive') {
        return [pscustomobject]@{ Teams = $false; Status = 'GraphOnly'; Message = 'Teams PowerShell login is skipped by design for app-based auth.' }
    }
    if (-not (Get-Module -ListAvailable -Name 'MicrosoftTeams')) {
        return [pscustomobject]@{ Teams = $false; Status = 'GraphOnly'; Message = 'MicrosoftTeams module is not installed. Graph-only Teams collection remains active.' }
    }
    if (-not (Get-Module -Name 'MicrosoftTeams')) {
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            Import-Module 'MicrosoftTeams' -UseWindowsPowerShell -ErrorAction Stop
        }
        else {
            Import-Module 'MicrosoftTeams' -ErrorAction Stop
        }
    }
    try {
        Write-Host 'Connecting assessment Teams PowerShell session...' -ForegroundColor Cyan
        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
        Write-Host 'Teams PowerShell connected.' -ForegroundColor Green
        return [pscustomobject]@{ Teams = $true; Status = 'Connected'; Message = 'Connected to Microsoft Teams PowerShell.' }
    }
    catch {
        return [pscustomobject]@{ Teams = $false; Status = 'GraphOnly'; Message = $_.Exception.Message }
    }
}

function Connect-ArrayaAssessmentPipelinePurview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
        [string]$AuthenticationType,
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,
        [Parameter(Mandatory = $false)]
        [string]$InitialDomain
    )

    if (Test-ArrayaAssessmentPipelinePurviewSessionReady) {
        Write-Host 'Purview compliance session ready for assessment collection.' -ForegroundColor Green
        return [pscustomobject]@{ PurviewCmdletsAvailable = $true }
    }
    if (-not (Get-Module -ListAvailable -Name 'ExchangeOnlineManagement')) {
        throw "Connect-IPPSSession is unavailable. Install or import ExchangeOnlineManagement before running the assessment."
    }
    if (-not (Get-Module -Name 'ExchangeOnlineManagement')) {
        Import-Module 'ExchangeOnlineManagement' -ErrorAction Stop -WarningAction SilentlyContinue
    }
    if (-not (Get-Command -Name 'Connect-IPPSSession' -ErrorAction SilentlyContinue)) {
        throw "Connect-IPPSSession is unavailable in the current session. Install or import ExchangeOnlineManagement before running the assessment."
    }
    if ($AuthenticationType -eq 'ClientSecret') {
        throw 'Client secret authentication is not supported for Purview compliance PowerShell in this workflow. Assign the service principal the Exchange Administrator role or run Connect-IPPSSession successfully in the current session and rerun the script.'
    }

    Write-Host 'Connecting assessment Purview compliance session...' -ForegroundColor Cyan
    $connectParams = @{ ErrorAction = 'Stop' }
    $connectCommand = Get-Command -Name 'Connect-IPPSSession' -ErrorAction Stop
    if ($connectCommand.Parameters.ContainsKey('CommandName')) {
        $connectParams.CommandName = @('Get-RetentionCompliancePolicy', 'Get-DlpCompliancePolicy')
    }
    if ($connectCommand.Parameters.ContainsKey('ShowBanner')) {
        $connectParams.ShowBanner = $false
    }

    if ($AuthenticationType -eq 'Certificate') {
        if ([string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($CertificateThumbprint) -or [string]::IsNullOrWhiteSpace($InitialDomain)) {
            throw 'Purview certificate auth requires -ClientId, -CertificateThumbprint, and the tenant initial domain.'
        }
        $connectParams.AppId = $ClientId
        $connectParams.CertificateThumbprint = $CertificateThumbprint
        $connectParams.Organization = $InitialDomain
        Connect-IPPSSession @connectParams | Out-Null
    }
    else {
        try {
            Connect-IPPSSession @connectParams | Out-Null
        }
        catch {
            if ($connectCommand.Parameters.ContainsKey('Device')) {
                Write-Warning 'Interactive Purview authentication failed. Retrying with device code.'
                $connectParams.Device = $true
                Connect-IPPSSession @connectParams | Out-Null
            }
            else {
                throw
            }
        }
    }

    if (-not (Test-ArrayaAssessmentPipelinePurviewCmdletsAvailable)) {
        throw 'Connect-IPPSSession completed, but the required Purview compliance cmdlets were not available. Assign the service principal the Exchange Administrator role or run Connect-IPPSSession successfully in the current session and rerun the script.'
    }

    Write-Host 'Purview compliance session ready for assessment collection.' -ForegroundColor Green
    return [pscustomobject]@{ PurviewCmdletsAvailable = $true }
}

function Test-ArrayaAssessmentPipelineExistingSessions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$WorkloadPlan
    )

    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        throw 'SkipAuth was requested, but no existing Microsoft Graph session was found. Connect first, then rerun with -SkipAuth.'
    }
    if (-not (Test-ArrayaAssessmentPipelineExchangeCmdletsAvailable)) {
        throw 'SkipAuth was requested, but Exchange Online cmdlets are not available in the current session. Connect to Exchange Online first, then rerun with -SkipAuth.'
    }
    if ($WorkloadPlan.Workloads.PurviewCompliance.Required -and -not (Test-ArrayaAssessmentPipelinePurviewSessionReady -ProbeCommands)) {
        throw 'SkipAuth was requested, but Purview compliance session is not usable in the current session. Connect to Purview compliance PowerShell first, confirm Get-RetentionCompliancePolicy and Get-DlpCompliancePolicy work, then rerun with -SkipAuth.'
    }

    $connectedWorkloads = @('Graph', 'ExchangeOnline')
    if ($WorkloadPlan.Workloads.PurviewCompliance.Required) { $connectedWorkloads += 'PurviewCompliance' }
    $skippedWorkloads = @()
    if (-not $WorkloadPlan.Workloads.PurviewCompliance.Required) { $skippedWorkloads += 'PurviewCompliance' }
    $fallbackWorkloads = @()
    if ($WorkloadPlan.Workloads.SharePointOnline.Mode -eq 'GraphFallback') { $fallbackWorkloads += 'SharePointOnline' } else { $skippedWorkloads += 'SharePointOnline' }
    if ($WorkloadPlan.Workloads.Teams.Mode -eq 'GraphOnly') { $fallbackWorkloads += 'Teams' } else { $skippedWorkloads += 'Teams' }

    return [pscustomobject][ordered]@{
        AuthenticationType       = 'ExistingSession'
        RequiredWorkloads        = @($WorkloadPlan.RequiredWorkloads)
        ConnectedWorkloads       = @($connectedWorkloads)
        SkippedWorkloads         = @($skippedWorkloads)
        FallbackWorkloads        = @($fallbackWorkloads)
        InitialDomain            = $null
        Graph                    = $true
        GraphContextAvailable    = $true
        ExchangeOnline           = $true
        ExchangeCmdletsAvailable = $true
        PurviewCmdletsAvailable  = [bool](Test-ArrayaAssessmentPipelinePurviewCmdletsAvailable)
        SharePointOnline         = $false
        Teams                    = $false
    }
}

function Write-ArrayaAssessmentPipelineConnectionSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$ConnectionResult,
        [Parameter(Mandatory = $false)]
        [switch]$PermissionPreflightSkipped
    )

    $connectedWorkloads = @($ConnectionResult.ConnectedWorkloads | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $fallbackWorkloads = @($ConnectionResult.FallbackWorkloads | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $skippedWorkloads = @($ConnectionResult.SkippedWorkloads | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })

    Write-Host ''
    Write-Host 'Connection / preflight ready.' -ForegroundColor Green
    Write-Host ("  Connected workloads: {0}" -f $(if ($connectedWorkloads.Count -gt 0) { $connectedWorkloads -join ', ' } else { 'None' })) -ForegroundColor DarkGray
    if ($fallbackWorkloads.Count -gt 0) { Write-Host ("  Fallback workloads: {0}" -f ($fallbackWorkloads -join ', ')) -ForegroundColor DarkGray }
    if ($skippedWorkloads.Count -gt 0) { Write-Host ("  Skipped workloads: {0}" -f ($skippedWorkloads -join ', ')) -ForegroundColor Yellow }
    $permissionSummary = if ($PermissionPreflightSkipped) { 'Skipped by request' } else { 'Passed for required workloads' }
    Write-Host ("  Permission checks: {0}" -f $permissionSummary) -ForegroundColor $(if ($PermissionPreflightSkipped) { 'Yellow' } else { 'DarkGray' })
}

function Invoke-ArrayaAssessmentConnectionPhase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context
    )

    Write-Host ''
    Write-Host '[Connection] Connection / Preflight' -ForegroundColor White
    Write-Host ('-' * 72) -ForegroundColor DarkGray

    $resolvedAuthMode = Resolve-ArrayaAssessmentPipelineRequestedAuthMode -AuthMode $Context.AuthMode -CertificateThumbprint $Context.CertificateThumbprint -ClientSecret $Context.ClientSecret
    $workloadPlan = Resolve-ArrayaAssessmentPipelineAuthWorkloadPlan `
        -AuthMode $resolvedAuthMode `
        -NeedsGovernanceCompliancePolicies $true `
        -NeedsTeamsInventory $true `
        -NeedsTeamsVoice $true `
        -NeedsSharePointData $true

    $authResult = [ordered]@{
        AuthenticationType       = $workloadPlan.AuthenticationType
        RequiredWorkloads        = @($workloadPlan.RequiredWorkloads)
        ConnectedWorkloads       = @()
        SkippedWorkloads         = @()
        FallbackWorkloads        = @($workloadPlan.FallbackWorkloads)
        InitialDomain            = $null
        Graph                    = $false
        GraphContextAvailable    = $false
        ExchangeOnline           = $false
        ExchangeCmdletsAvailable = $false
        PurviewCmdletsAvailable  = $false
        SharePointOnline         = $false
        Teams                    = $false
    }

    Write-Host ("Assessment login mode: {0}" -f $workloadPlan.AuthenticationType) -ForegroundColor Cyan
    Write-Host ("Required auth workloads: {0}" -f ($workloadPlan.RequiredWorkloads -join ', ')) -ForegroundColor DarkGray
    if ([bool]$Context.SkipPermissionPreflight) {
        Write-Host 'Workload preflight skipped by request. Collection will continue and may fail later where access is missing.' -ForegroundColor Yellow
    }

    if ([bool]$Context.SkipAuth) {
        Write-Host 'Reusing existing workload sessions for this run.' -ForegroundColor Cyan
        $existingSessionResult = Test-ArrayaAssessmentPipelineExistingSessions -WorkloadPlan $workloadPlan
        foreach ($property in @($existingSessionResult.PSObject.Properties)) {
            $authResult[$property.Name] = $property.Value
        }
        if (-not [bool]$Context.SkipPermissionPreflight) {
            Test-ArrayaAssessmentPipelinePermissionPreflight -Context $Context -ConnectionResult ([pscustomobject]$authResult) -WorkloadPlan $workloadPlan -Workload Graph | Out-Null
            Test-ArrayaAssessmentPipelinePermissionPreflight -Context $Context -ConnectionResult ([pscustomobject]$authResult) -WorkloadPlan $workloadPlan -Workload ExchangeOnline | Out-Null
            if ($workloadPlan.Workloads.PurviewCompliance.Required) {
                Test-ArrayaAssessmentPipelinePermissionPreflight -Context $Context -ConnectionResult ([pscustomobject]$authResult) -WorkloadPlan $workloadPlan -Workload Purview | Out-Null
            }
        }
    }
    else {
        $graphResult = Connect-ArrayaAssessmentPipelineGraph -AuthenticationType $workloadPlan.AuthenticationType -TenantId $Context.TenantId -ClientId $Context.ClientId -CertificateThumbprint $Context.CertificateThumbprint -ClientSecret $Context.ClientSecret
        $authResult.Graph = [bool]$graphResult.Graph
        $authResult.GraphContextAvailable = [bool](Get-MgContext -ErrorAction SilentlyContinue)
        $authResult.InitialDomain = $graphResult.InitialDomain
        $authResult.ConnectedWorkloads += 'Graph'
        if (-not [bool]$Context.SkipPermissionPreflight) {
            Test-ArrayaAssessmentPipelinePermissionPreflight -Context $Context -ConnectionResult ([pscustomobject]$authResult) -WorkloadPlan $workloadPlan -Workload Graph | Out-Null
        }

        $exchangeResult = Connect-ArrayaAssessmentPipelineExchange -AuthenticationType $workloadPlan.AuthenticationType -ClientId $Context.ClientId -CertificateThumbprint $Context.CertificateThumbprint -InitialDomain $authResult.InitialDomain
        $authResult.ExchangeOnline = [bool]$exchangeResult.ExchangeOnline
        $authResult.ExchangeCmdletsAvailable = [bool](Test-ArrayaAssessmentPipelineExchangeCmdletsAvailable)
        $authResult.ConnectedWorkloads += 'ExchangeOnline'
        if (-not [bool]$Context.SkipPermissionPreflight) {
            Test-ArrayaAssessmentPipelinePermissionPreflight -Context $Context -ConnectionResult ([pscustomobject]$authResult) -WorkloadPlan $workloadPlan -Workload ExchangeOnline | Out-Null
        }

        if ($workloadPlan.Workloads.PurviewCompliance.Required) {
            $purviewResult = Connect-ArrayaAssessmentPipelinePurview -AuthenticationType $workloadPlan.AuthenticationType -ClientId $Context.ClientId -CertificateThumbprint $Context.CertificateThumbprint -InitialDomain $authResult.InitialDomain
            $authResult.PurviewCmdletsAvailable = [bool]$purviewResult.PurviewCmdletsAvailable
            $authResult.ConnectedWorkloads += 'PurviewCompliance'
            if (-not [bool]$Context.SkipPermissionPreflight) {
                Test-ArrayaAssessmentPipelinePermissionPreflight -Context $Context -ConnectionResult ([pscustomobject]$authResult) -WorkloadPlan $workloadPlan -Workload Purview | Out-Null
            }
        }
        else {
            $authResult.SkippedWorkloads += 'PurviewCompliance'
        }
    }

    if ($workloadPlan.Workloads.SharePointOnline.Connect) {
        $sharePointResult = Connect-ArrayaAssessmentPipelineSharePoint -AuthenticationType $workloadPlan.AuthenticationType -InitialDomain $authResult.InitialDomain
        $authResult.SharePointOnline = [bool]$sharePointResult.SharePointOnline
        switch ($sharePointResult.Status) {
            'Connected' { $authResult.ConnectedWorkloads += 'SharePointOnline' }
            'GraphFallback' {
                if ($authResult.FallbackWorkloads -notcontains 'SharePointOnline') { $authResult.FallbackWorkloads += 'SharePointOnline' }
                Write-Host ("SharePoint auth: Graph fallback active. {0}" -f $sharePointResult.Message) -ForegroundColor DarkGray
            }
            default {
                $authResult.SkippedWorkloads += 'SharePointOnline'
                Write-Host ("SharePoint auth skipped. {0}" -f $sharePointResult.Message) -ForegroundColor DarkGray
            }
        }
    }

    if ($workloadPlan.Workloads.Teams.Connect) {
        $teamsResult = Connect-ArrayaAssessmentPipelineTeams -AuthenticationType $workloadPlan.AuthenticationType
        $authResult.Teams = [bool]$teamsResult.Teams
        switch ($teamsResult.Status) {
            'Connected' { $authResult.ConnectedWorkloads += 'Teams' }
            'GraphOnly' {
                if ($authResult.FallbackWorkloads -notcontains 'Teams') { $authResult.FallbackWorkloads += 'Teams' }
                Write-Host ("Teams auth: Graph-only path active. {0}" -f $teamsResult.Message) -ForegroundColor DarkGray
            }
            default {
                $authResult.SkippedWorkloads += 'Teams'
                Write-Host ("Teams auth skipped. {0}" -f $teamsResult.Message) -ForegroundColor DarkGray
            }
        }
    }

    $connectionResult = [pscustomobject]$authResult
    $tenantOrganization = $null
    try {
        $tenantOrganization = Get-ArrayaAssessmentPipelineTenantOrganization
    }
    catch {}
    if (
        $tenantOrganization -and
        [string]::IsNullOrWhiteSpace([string]$connectionResult.InitialDomain) -and
        -not [string]::IsNullOrWhiteSpace([string]$tenantOrganization.InitialDomain)
    ) {
        $connectionResult.InitialDomain = [string]$tenantOrganization.InitialDomain
    }
    Write-ArrayaAssessmentPipelineConnectionSummary -ConnectionResult $connectionResult -PermissionPreflightSkipped:([bool]$Context.SkipPermissionPreflight)

    $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Connection'
    $snapshot = New-ArrayaTenantSnapshot -Metadata ([ordered]@{
        GeneratedAt        = (Get-Date).ToString('o')
        OutputProfile      = $Context.OutputProfile
        OutputProfileLabel = $Context.OutputProfileLabel
        ReportingMode      = $Context.ReportingMode
        Authentication     = [ordered]@{
            BootstrapMode      = [string]$connectionResult.AuthenticationType
            RequiredWorkloads  = @($connectionResult.RequiredWorkloads)
            ConnectedWorkloads = @($connectionResult.ConnectedWorkloads)
            SkippedWorkloads   = @($connectionResult.SkippedWorkloads)
            FallbackWorkloads  = @($connectionResult.FallbackWorkloads)
        }
        Tenant = [ordered]@{
            DisplayName       = if ($tenantOrganization) { [string]$tenantOrganization.DisplayName } else { $null }
            DefaultDomainName = $connectionResult.InitialDomain
        }
    })
    Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $checkpointPath

    return [pscustomobject]@{
        Phase              = 'Connection'
        CheckpointPath     = $checkpointPath
        ExportFileLocation = $Context.ExportPath
        ConnectionResult   = $connectionResult
    }
}
