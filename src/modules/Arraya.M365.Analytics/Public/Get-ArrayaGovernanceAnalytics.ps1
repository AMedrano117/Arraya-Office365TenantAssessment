# Conditional Access, MFA posture, privileged access, and collaboration cleanup analysis.
# Shared by the HTML assessment report and the tenant improvement plan so both
# deliverables describe the same tenant from one implementation.

function New-ArrayaConditionalAccessOptimizationRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$ConditionalAccessPolicies = @(),
        [Parameter(Mandatory = $false)]$ConditionalAccessSummary
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $policyRows = @($ConditionalAccessPolicies | Where-Object { $null -ne $_ })
    function Test-PolicyEnabledLocal { param($Policy) ([string](Get-ArrayaObjectValue -Object $Policy -Names @('State', 'state'))).ToLowerInvariant() -eq 'enabled' }
    function Test-PolicyReportOnlyLocal { param($Policy) ([string](Get-ArrayaObjectValue -Object $Policy -Names @('State', 'state'))).ToLowerInvariant() -match 'report' }
    function Test-PolicyRequiresMfaLocal {
        param($Policy)
        $explicit = Convert-ArrayaToBoolean (Get-ArrayaObjectValue -Object $Policy -Names @('RequiresMfaEnforcement', 'RequiresMfa', 'UsesAuthenticationStrengthForMfa'))
        if ($explicit -eq $true) { return $true }
        return (($Policy | ConvertTo-Json -Depth 8 -Compress) -match '(?i)mfa|multiFactorAuthentication|authenticationStrength')
    }
    function Test-PolicySignalLocal {
        param($Policy, [string[]]$BooleanNames, [string]$Regex)
        foreach ($name in @($BooleanNames)) {
            $value = Convert-ArrayaToBoolean (Get-ArrayaObjectValue -Object $Policy -Names @($name))
            if ($value -eq $true) { return $true }
        }
        $nameText = [string](Get-ArrayaObjectValue -Object $Policy -Names @('DisplayName', 'Name'))
        $jsonText = ($Policy | ConvertTo-Json -Depth 8 -Compress)
        return (($nameText -match $Regex) -or ($jsonText -match $Regex))
    }
    function Add-OptimizationRowLocal {
        param([string]$Signal, [string]$Status, [string]$Severity, [object[]]$Policies, [string]$CurrentState, [string]$Recommendation, [string]$Evidence)
        $policyNames = @($Policies | ForEach-Object { [string](Get-ArrayaObjectValue -Object $_ -Names @('DisplayName', 'Name')) } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 8)
        $rows.Add([pscustomobject]@{
            Signal          = $Signal
            Status          = $Status
            Severity        = $Severity
            PolicyCount     = @($Policies).Count
            PolicyNames     = ($policyNames -join '; ')
            CurrentState    = $CurrentState
            Recommendation  = $Recommendation
            Evidence        = $Evidence
            SourceWorksheet = 'ConditionalAccessPolicies'
        }) | Out-Null
    }

    $enabledPolicies = @($policyRows | Where-Object { Test-PolicyEnabledLocal -Policy $_ })
    $reportOnlyPolicies = @($policyRows | Where-Object { Test-PolicyReportOnlyLocal -Policy $_ })
    $disabledPolicies = @($policyRows | Where-Object { ([string](Get-ArrayaObjectValue -Object $_ -Names @('State', 'state'))).ToLowerInvariant() -eq 'disabled' })
    $adminMfaPolicies = @($enabledPolicies | Where-Object { (Test-PolicyRequiresMfaLocal -Policy $_) -and (Test-PolicySignalLocal -Policy $_ -BooleanNames @('TargetsPrivilegedRoles') -Regex '(?i)admin|privileged|global administrator|directoryRole') })
    $legacyPolicies = @($enabledPolicies | Where-Object { Test-PolicySignalLocal -Policy $_ -BooleanNames @('BlocksLegacyAuth') -Regex '(?i)legacy|exchangeActiveSync|other clients|clientAppTypes' })
    $guestPolicies = @($enabledPolicies | Where-Object { Test-PolicySignalLocal -Policy $_ -BooleanNames @('TargetsGuestsOrExternalUsers') -Regex '(?i)guest|external|b2b|includeGuestsOrExternalUsers' })
    $riskPolicies = @($enabledPolicies | Where-Object { Test-PolicySignalLocal -Policy $_ -BooleanNames @('UsesRiskSignals') -Regex '(?i)risk|signInRisk|userRisk' })
    $devicePolicies = @($enabledPolicies | Where-Object { Test-PolicySignalLocal -Policy $_ -BooleanNames @('RequiresCompliantDevice') -Regex '(?i)compliantDevice|device compliance' })
    $exclusionPolicies = @($policyRows | Where-Object { (Convert-ArrayaToBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasExclusions'))) -eq $true -or (($_ | ConvertTo-Json -Depth 8 -Compress) -match '(?i)"exclude') })

    Add-OptimizationRowLocal -Signal 'Admin MFA baseline coverage' -Status $(if ($adminMfaPolicies.Count -gt 0) { 'Observed' } else { 'Review' }) -Severity $(if ($adminMfaPolicies.Count -gt 0) { 'Info' } else { 'High' }) -Policies $adminMfaPolicies -CurrentState "$($adminMfaPolicies.Count) enabled admin/privileged MFA policy signal(s)" -Recommendation 'Validate privileged roles are protected by enabled Conditional Access requiring strong MFA or authentication strength, with only documented emergency-access exclusions.' -Evidence 'TargetsPrivilegedRoles and MFA grant/authentication-strength signals.'
    Add-OptimizationRowLocal -Signal 'Legacy authentication blocking' -Status $(if ($legacyPolicies.Count -gt 0) { 'Observed' } else { 'Review' }) -Severity $(if ($legacyPolicies.Count -gt 0) { 'Info' } else { 'High' }) -Policies $legacyPolicies -CurrentState "$($legacyPolicies.Count) enabled legacy-authentication protection signal(s)" -Recommendation 'Validate legacy client authentication is blocked through an enabled policy and exceptions are explicit.' -Evidence 'BlocksLegacyAuth, client app, grant-control, and policy-name signals.'
    Add-OptimizationRowLocal -Signal 'Guest and external access coverage' -Status $(if ($guestPolicies.Count -gt 0) { 'Observed' } else { 'Review' }) -Severity $(if ($guestPolicies.Count -gt 0) { 'Info' } else { 'Medium' }) -Policies $guestPolicies -CurrentState "$($guestPolicies.Count) enabled guest/external policy signal(s)" -Recommendation 'Confirm guest/external users have a deliberate access baseline, including MFA or trusted inbound MFA where appropriate.' -Evidence 'Guest/external target signals.'
    Add-OptimizationRowLocal -Signal 'Risk-based access coverage' -Status $(if ($riskPolicies.Count -gt 0) { 'Observed' } else { 'Review' }) -Severity $(if ($riskPolicies.Count -gt 0) { 'Info' } else { 'Medium' }) -Policies $riskPolicies -CurrentState "$($riskPolicies.Count) enabled risk-based policy signal(s)" -Recommendation 'Review user-risk and sign-in-risk policy coverage for high-value users and high-risk access paths.' -Evidence 'User risk and sign-in risk condition signals.'
    Add-OptimizationRowLocal -Signal 'Compliant-device access scenarios' -Status $(if ($devicePolicies.Count -gt 0) { 'Observed' } else { 'Review' }) -Severity 'Low' -Policies $devicePolicies -CurrentState "$($devicePolicies.Count) enabled compliant-device requirement signal(s)" -Recommendation 'Decide which managed-access scenarios should require compliant devices and validate those controls in report-only or pilot before enforcement.' -Evidence 'Grant-control compliant-device signals.'
    Add-OptimizationRowLocal -Signal 'Report-only policy backlog' -Status $(if ($reportOnlyPolicies.Count -gt 0) { 'Review' } else { 'Observed' }) -Severity $(if ($reportOnlyPolicies.Count -gt 0) { 'Medium' } else { 'Info' }) -Policies $reportOnlyPolicies -CurrentState "$($reportOnlyPolicies.Count) report-only policy/policies" -Recommendation 'Review report-only impact, document expected behavior, and move validated policies into enabled enforcement.' -Evidence 'Policy state.'
    Add-OptimizationRowLocal -Signal 'Disabled policy backlog' -Status $(if ($disabledPolicies.Count -gt 0) { 'Review' } else { 'Observed' }) -Severity 'Low' -Policies $disabledPolicies -CurrentState "$($disabledPolicies.Count) disabled policy/policies" -Recommendation 'Retire obsolete disabled policies or document why they remain as staged/preserved policy definitions.' -Evidence 'Policy state.'
    Add-OptimizationRowLocal -Signal 'Exclusion-heavy policy review' -Status $(if ($exclusionPolicies.Count -ge 3) { 'Review' } else { 'Observed' }) -Severity $(if ($exclusionPolicies.Count -ge 3) { 'Medium' } else { 'Info' }) -Policies $exclusionPolicies -CurrentState "$($exclusionPolicies.Count) policy/policies with exclusion signals" -Recommendation 'Minimize exclusions and reserve them for break-glass or validated operational exceptions.' -Evidence 'Exclude users/groups/roles/app/client signals.'

    $signatureGroups = @(
        $policyRows |
            Group-Object {
                $state = ([string](Get-ArrayaObjectValue -Object $_ -Names @('State', 'state'))).ToLowerInvariant()
                $includeUsers = [string](Get-ArrayaObjectValue -Object $_ -Names @('IncludedUsers', 'IncludeUsers', 'Users'))
                $includeGroups = [string](Get-ArrayaObjectValue -Object $_ -Names @('IncludedGroups', 'IncludeGroups'))
                $includeRoles = [string](Get-ArrayaObjectValue -Object $_ -Names @('IncludedRoles', 'IncludeRoles'))
                $applications = [string](Get-ArrayaObjectValue -Object $_ -Names @('Applications', 'IncludedApplications', 'CloudApps'))
                $clientApps = [string](Get-ArrayaObjectValue -Object $_ -Names @('ClientAppTypes', 'ClientApps'))
                $risk = [string](Get-ArrayaObjectValue -Object $_ -Names @('SignInRiskLevels', 'UserRiskLevels'))
                $platforms = [string](Get-ArrayaObjectValue -Object $_ -Names @('IncludePlatforms', 'Platforms'))
                $grant = [string](Get-ArrayaObjectValue -Object $_ -Names @('GrantControls'))
                $session = [string](Get-ArrayaObjectValue -Object $_ -Names @('SessionControls'))
                "$state|$includeUsers|$includeGroups|$includeRoles|$applications|$clientApps|$risk|$platforms|$grant|$session".ToLowerInvariant()
            } |
            Where-Object { $_.Count -gt 1 -and -not [string]::IsNullOrWhiteSpace([string]$_.Name) }
    )
    foreach ($signatureGroup in @($signatureGroups)) {
        Add-OptimizationRowLocal -Signal 'Possible overlapping Conditional Access policies' -Status 'Review' -Severity 'Low' -Policies @($signatureGroup.Group) -CurrentState "$($signatureGroup.Count) policy/policies share a similar target, condition, grant, and session-control signature" -Recommendation 'Compare these policies for duplicate scope, conflicting controls, and opportunities to consolidate or clarify ownership.' -Evidence 'Deterministic policy-signature grouping.'
    }

    return @($rows.ToArray())
}

function New-ArrayaMfaMethodPostureSummaryRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$MfaEnrollmentSummary,
        [Parameter(Mandatory = $false)][object[]]$MfaRegistrationDetails = @()
    )

    $summaryRecord = if ($MfaEnrollmentSummary) { Get-ArrayaObjectValue -Object $MfaEnrollmentSummary -Names @('Summary', 'Configuration') } else { $null }
    if (-not $summaryRecord -and $MfaEnrollmentSummary) { $summaryRecord = $MfaEnrollmentSummary }
    $totalUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $summaryRecord -Names @('TotalUsers', 'UserCount', 'TotalUserCount'))
    $registeredUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $summaryRecord -Names @('RegisteredUsers', 'RegisteredUserCount', 'MfaRegisteredUsers'))
    $notRegisteredUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $summaryRecord -Names @('NotRegisteredUsers', 'UsersNotRegisteredForMfa', 'NotRegisteredUserCount'))
    if ($null -eq $notRegisteredUsers -and $null -ne $totalUsers -and $null -ne $registeredUsers) { $notRegisteredUsers = [Math]::Max(($totalUsers - $registeredUsers), 0) }
    $weakMethodOnlyUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $summaryRecord -Names @('UsersWithWeakMethodsOnly'))
    $weakDefaultMethodUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $summaryRecord -Names @('UsersWithWeakDefaultMethod'))
    $phishingResistantUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $summaryRecord -Names @('UsersWithPhishingResistantMethods', 'PhishingResistantUsers'))
    $weakBreakdown = [string](Get-ArrayaObjectValue -Object $summaryRecord -Names @('WeakMethodBreakdown'))
    $registeredBreakdown = [string](Get-ArrayaObjectValue -Object $summaryRecord -Names @('RegisteredMethodBreakdown'))
    if (($null -eq $totalUsers -or $null -eq $registeredUsers) -and @($MfaRegistrationDetails).Count -gt 0) {
        $totalUsers = @($MfaRegistrationDetails).Count
        $registeredUsers = @($MfaRegistrationDetails | Where-Object { (Convert-ArrayaToBoolean (Get-ArrayaObjectValue -Object $_ -Names @('IsMfaRegistered', 'IsRegistered', 'MfaRegistered'))) -eq $true }).Count
        $notRegisteredUsers = [Math]::Max(($totalUsers - $registeredUsers), 0)
    }

    return @(
        [pscustomobject]@{ Signal = 'MFA registration coverage'; Status = if ($null -ne $notRegisteredUsers -and $notRegisteredUsers -gt 0) { 'Review' } else { 'Observed' }; UserCount = $registeredUsers; CurrentState = ("{0} registered / {1} not registered" -f $(if ($null -eq $registeredUsers) { 'Unknown' } else { [int]$registeredUsers }), $(if ($null -eq $notRegisteredUsers) { 'Unknown' } else { [int]$notRegisteredUsers })); Recommendation = 'Keep enrollment separate from enforcement and drive remaining users through registration readiness.'; SourceWorksheet = 'MfaEnrollmentSummary' },
        [pscustomobject]@{ Signal = 'Weak MFA method reliance'; Status = if ((($null -ne $weakMethodOnlyUsers -and $weakMethodOnlyUsers -gt 0) -or ($null -ne $weakDefaultMethodUsers -and $weakDefaultMethodUsers -gt 0))) { 'Review' } else { 'Observed' }; UserCount = $(if ($null -ne $weakMethodOnlyUsers) { [int]$weakMethodOnlyUsers } else { $null }); CurrentState = ("{0} weak-only user(s); {1} weak default user(s); {2}" -f $(if ($null -eq $weakMethodOnlyUsers) { 'Unknown' } else { [int]$weakMethodOnlyUsers }), $(if ($null -eq $weakDefaultMethodUsers) { 'Unknown' } else { [int]$weakDefaultMethodUsers }), $(if ([string]::IsNullOrWhiteSpace($weakBreakdown)) { 'Weak-method breakdown not surfaced' } else { $weakBreakdown })); Recommendation = 'Reduce SMS, voice, and email OTP reliance; encourage Microsoft Authenticator, passwordless, FIDO2/passkeys, or other phishing-resistant methods for privileged access.'; SourceWorksheet = 'MfaEnrollmentSummary' },
        [pscustomobject]@{ Signal = 'Phishing-resistant method adoption'; Status = if ($null -ne $phishingResistantUsers -and $phishingResistantUsers -gt 0) { 'Observed' } else { 'Review' }; UserCount = $phishingResistantUsers; CurrentState = $(if ($null -eq $phishingResistantUsers) { 'Phishing-resistant method count not surfaced' } else { "$([int]$phishingResistantUsers) user(s)" }); Recommendation = 'Prioritize phishing-resistant authentication for privileged and high-risk populations where licensing and device readiness support it.'; SourceWorksheet = 'MfaEnrollmentSummary' },
        [pscustomobject]@{ Signal = 'Registered method mix'; Status = 'Observed'; UserCount = $registeredUsers; CurrentState = $(if ([string]::IsNullOrWhiteSpace($registeredBreakdown)) { 'Registered method mix not surfaced' } else { $registeredBreakdown }); Recommendation = 'Use the method mix to target communications, registration campaigns, and weak-method retirement waves.'; SourceWorksheet = 'MfaEnrollmentSummary' }
    )
}

function New-ArrayaPrivilegedAccessRemediationSummaryRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Admins = @(),
        [Parameter(Mandatory = $false)]$AdminMfaSummary,
        [Parameter(Mandatory = $false)][object[]]$AdminMfaRegistrationGaps = @(),
        [Parameter(Mandatory = $false)][object[]]$AdminMfaEnforcementGaps = @(),
        [Parameter(Mandatory = $false)][int]$MaxGlobalAdmins = 5
    )

    $enabledAdmins = @($Admins | Where-Object { (Convert-ArrayaToBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))) -ne $false })
    $globalAdmins = @($enabledAdmins | Where-Object { [string](Get-ArrayaObjectValue -Object $_ -Names @('Role', 'DirectoryRole', 'AdminRole', 'RoleName')) -match 'Global Administrator' })
    $staleAdmins = @($enabledAdmins | Where-Object {
        $lastSignIn = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastSignInDateTime', 'LastSuccessfulSignInDateTime'))
        $lastSignIn -and $lastSignIn -lt (Get-Date).AddDays(-90)
    })
    $cloudOnlyGlobalAdmins = @($globalAdmins | Where-Object { (Convert-ArrayaToBoolean (Get-ArrayaObjectValue -Object $_ -Names @('OnPremisesSyncEnabled', 'DirSyncEnabled', 'IsSynced'))) -ne $true })
    $summaryRecord = if ($AdminMfaSummary) { Get-ArrayaObjectValue -Object $AdminMfaSummary -Names @('Summary', 'Configuration') } else { $null }
    if (-not $summaryRecord -and $AdminMfaSummary) { $summaryRecord = $AdminMfaSummary }
    $notRegistered = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $summaryRecord -Names @('AdminUsersNotRegisteredForMfa'))
    if ($null -eq $notRegistered) { $notRegistered = @($AdminMfaRegistrationGaps).Count }
    $notCovered = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $summaryRecord -Names @('AdminUsersNotCoveredByMfaEnforcement'))
    if ($null -eq $notCovered) { $notCovered = @($AdminMfaEnforcementGaps).Count }

    return @(
        [pscustomobject]@{ Signal = 'Standing Global Administrator assignments'; Status = if ($globalAdmins.Count -gt $MaxGlobalAdmins) { 'Review' } else { 'Observed' }; Count = $globalAdmins.Count; CurrentState = "$($globalAdmins.Count) enabled Global Administrator assignment(s)"; RecommendedAction = 'Reduce standing Global Administrators, remove stale assignments, use lower-privilege roles such as Global Reader where appropriate, and prefer PIM eligible activation with approval where supported.'; SourceWorksheet = 'Admins' },
        [pscustomobject]@{ Signal = 'Stale privileged admin accounts'; Status = if ($staleAdmins.Count -gt 0) { 'Review' } else { 'Observed' }; Count = $staleAdmins.Count; CurrentState = "$($staleAdmins.Count) enabled admin account(s) appear stale"; RecommendedAction = 'Validate stale admin accounts and remove, disable, or downgrade privileges that are no longer required.'; SourceWorksheet = 'Admins' },
        [pscustomobject]@{ Signal = 'Admin MFA registration gaps'; Status = if ($notRegistered -gt 0) { 'Review' } else { 'Observed' }; Count = $notRegistered; CurrentState = "$([int]$notRegistered) enabled admin account(s) not registered for MFA"; RecommendedAction = 'Drive admin registration completion separately from Conditional Access enforcement validation.'; SourceWorksheet = 'AdminMfaRegistrationGaps' },
        [pscustomobject]@{ Signal = 'Admin MFA enforcement gaps'; Status = if ($notCovered -gt 0) { 'Review' } else { 'Observed' }; Count = $notCovered; CurrentState = "$([int]$notCovered) enabled admin account(s) not covered by active MFA enforcement"; RecommendedAction = 'Validate enabled Conditional Access enforcement for privileged roles and minimize exclusions to approved emergency-access accounts.'; SourceWorksheet = 'AdminMfaEnforcementGaps' },
        [pscustomobject]@{ Signal = 'Cloud-only emergency access coverage'; Status = if ($cloudOnlyGlobalAdmins.Count -lt 2) { 'Review' } else { 'Observed' }; Count = $cloudOnlyGlobalAdmins.Count; CurrentState = "$($cloudOnlyGlobalAdmins.Count) enabled cloud-only Global Administrator account(s) surfaced"; RecommendedAction = 'Maintain at least two cloud-only emergency-access accounts and test their Conditional Access exclusion and monitoring process.'; SourceWorksheet = 'Admins' }
    )
}

function New-ArrayaTeamsGroupsCleanupCandidateRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Teams = @(),
        [Parameter(Mandatory = $false)][object[]]$Groups = @()
    )

    $rows = New-Object System.Collections.Generic.List[object]
    function Add-CleanupCandidateLocal {
        param($ObjectType, $Name, $OwnerCount, $MemberCount, $GuestCount, $LastActivityDate, $StorageGB, $IsManagingLicenses, $RiskSignal, $RecommendedAction, $SourceWorksheet)
        $rows.Add([pscustomobject]@{
            ObjectType         = $ObjectType
            Name               = $Name
            OwnerCount         = $OwnerCount
            MemberCount        = $MemberCount
            GuestCount         = $GuestCount
            LastActivityDate   = $LastActivityDate
            StorageGB          = $StorageGB
            IsManagingLicenses = $IsManagingLicenses
            RiskSignal         = $RiskSignal
            RecommendedAction  = $RecommendedAction
            SourceWorksheet    = $SourceWorksheet
        }) | Out-Null
    }
    function Resolve-MemberCountLocal {
        param($Record)

        $explicitCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $Record -Names @('MemberCount', 'MembersCount', 'GroupMemberCount', 'MemberTotalCount'))
        if ($null -ne $explicitCount) {
            return $explicitCount
        }

        $membersValue = Get-ArrayaObjectValue -Object $Record -Names @('Members', 'GroupMembers')
        if ($null -eq $membersValue -or $membersValue -is [string]) {
            return $null
        }

        return @(Convert-ArrayaObjectToArray $membersValue).Count
    }

    foreach ($team in @($Teams)) {
        $name = [string](Get-ArrayaObjectValue -Object $team -Names @('DisplayName', 'Name'))
        $ownerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $team -Names @('OwnerCount', 'Owners'))
        $memberCount = Resolve-MemberCountLocal -Record $team
        $guestCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $team -Names @('GuestCount', 'Guests'))
        $privateChannelCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $team -Names @('PrivateChannelCount'))
        $sharedChannelCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $team -Names @('SharedChannelCount'))
        $lastActivity = Get-ArrayaObjectValue -Object $team -Names @('LastActivityDate', 'LastActivityDateTime')
        $storage = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $team -Names @('SiteSize-GB', 'StorageUsedGB'))
        $riskSignals = @()
        if ($null -ne $ownerCount -and $ownerCount -le 0) { $riskSignals += 'Ownerless Team' }
        if ($null -ne $memberCount -and $memberCount -le 0) { $riskSignals += 'No members' }
        if ($null -ne $guestCount -and $null -ne $memberCount -and $memberCount -gt 0 -and (($guestCount / $memberCount) -ge 0.4)) { $riskSignals += 'Guest-heavy Team' }
        if (($null -ne $privateChannelCount -and $privateChannelCount -ge 5) -or ($null -ne $sharedChannelCount -and $sharedChannelCount -ge 1)) { $riskSignals += 'Private/shared channel review' }
        try { if ($lastActivity -and ([datetime]$lastActivity) -lt (Get-Date).AddDays(-90)) { $riskSignals += 'Dormant Team' } } catch {}
        if ($riskSignals.Count -gt 0) {
            Add-CleanupCandidateLocal -ObjectType 'Team' -Name $name -OwnerCount $ownerCount -MemberCount $memberCount -GuestCount $guestCount -LastActivityDate $lastActivity -StorageGB $storage -IsManagingLicenses $false -RiskSignal ($riskSignals -join '; ') -RecommendedAction 'Confirm owner accountability, guest access, channel lifecycle, and whether this Team should remain active, be archived, or be retired.' -SourceWorksheet 'AllTeams'
        }
    }

    foreach ($group in @($Groups)) {
        $name = [string](Get-ArrayaObjectValue -Object $group -Names @('DisplayName', 'Name'))
        $ownerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $group -Names @('OwnerCount', 'Owners'))
        $memberCount = Resolve-MemberCountLocal -Record $group
        $isManagingLicenses = Test-ArrayaGroupManagesLicenses -Group $group
        $riskSignals = @()
        if ($null -ne $ownerCount -and $ownerCount -le 0) { $riskSignals += 'Ownerless group' }
        if ($null -ne $memberCount -and $memberCount -le 0) { $riskSignals += 'No members' }
        if ($isManagingLicenses -eq $true -and ($null -eq $ownerCount -or $ownerCount -le 0)) { $riskSignals += 'License-managing group without owner' }
        if ($riskSignals.Count -gt 0) {
            Add-CleanupCandidateLocal -ObjectType 'Microsoft 365 / Entra group' -Name $name -OwnerCount $ownerCount -MemberCount $memberCount -GuestCount $null -LastActivityDate (Get-ArrayaObjectValue -Object $group -Names @('LastActivityDate', 'LastActivityDateTime')) -StorageGB $null -IsManagingLicenses $isManagingLicenses -RiskSignal ($riskSignals -join '; ') -RecommendedAction 'Confirm owner accountability, membership purpose, and whether the group should remain active, be assigned lifecycle ownership, or be retired.' -SourceWorksheet 'EntraIDGroups'
        }
    }

    return @($rows.ToArray())
}

function Get-ArrayaConditionalAccessPolicyState {
    <#
    .SYNOPSIS
        Normalizes a Conditional Access policy state to one of four known values.

    .DESCRIPTION
        Graph reports policy state as enabled / enabledForReportingButNotEnforced / disabled,
        but snapshot round-trips and older collectors also produce reportOnly and report-only.
        A state that is missing or unrecognised returns unknown rather than disabled, so a
        policy whose state was never collected is not counted as switched off.

        This replaced three copies that normalized state identically: one in the HTML report
        helper, one in the legacy collector, and one nested inside the employee-experience
        analysis in Arraya.M365.Reporting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $Policy
    )

    $rawState = [string](Get-ArrayaObjectValue -Object $Policy -Names @('State', 'state', 'PolicyState', 'policyState'))
    if ([string]::IsNullOrWhiteSpace($rawState)) {
        return 'unknown'
    }

    switch ($rawState.Trim().ToLowerInvariant()) {
        'enabled' { return 'enabled' }
        'enabledforreportingbutnotenforced' { return 'enabledForReportingButNotEnforced' }
        'reportonly' { return 'enabledForReportingButNotEnforced' }
        'report-only' { return 'enabledForReportingButNotEnforced' }
        'disabled' { return 'disabled' }
        default { return 'unknown' }
    }
}
