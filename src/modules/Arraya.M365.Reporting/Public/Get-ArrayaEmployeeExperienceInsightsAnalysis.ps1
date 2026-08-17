function Get-ArrayaEmployeeExperienceInsightsAnalysis {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$EmailActivitySummary,
        [Parameter(Mandatory = $false)]
        [object]$EmployeeExperienceInsightsSummary,
        [Parameter(Mandatory = $false)]
        [array]$TopSenders = @(),
        [Parameter(Mandatory = $false)]
        [array]$TopReceivers = @(),
        [Parameter(Mandatory = $false)]
        [array]$TeamsTopUsers = @(),
        [Parameter(Mandatory = $false)]
        [array]$GroupsTopGroups = @(),
        [Parameter(Mandatory = $false)]
        [object]$AdminReportSettings,
        [Parameter(Mandatory = $false)]
        [object]$AuthConfig,
        [Parameter(Mandatory = $false)]
        [array]$ConditionalAccessPolicies = @(),
        [Parameter(Mandatory = $false)]
        [int]$TotalUsers = 0
    )

    function Get-SummaryValue {
        param(
            [AllowNull()]
            $Record,
            [Parameter(Mandatory)]
            [string[]]$Key,
            [AllowNull()]
            $Default = $null
        )

        if ($null -eq $Record) { return $Default }
        foreach ($candidateKey in @($Key)) {
            if ([string]::IsNullOrWhiteSpace($candidateKey)) { continue }
            if ($Record -is [System.Collections.IDictionary] -and ([System.Collections.IDictionary]$Record).Contains($candidateKey)) { return $Record[$candidateKey] }
            if ($Record -is [System.Collections.Specialized.OrderedDictionary] -and $Record.Contains($candidateKey)) { return $Record[$candidateKey] }
            if ($Record.PSObject -and $Record.PSObject.Properties[$candidateKey]) { return $Record.PSObject.Properties[$candidateKey].Value }
        }
        return $Default
    }

    function Convert-ToInt {
        param(
            [AllowNull()]
            $Value,
            [int]$Default = 0
        )

        if ($null -eq $Value) {
            return $Default
        }

        try {
            if ($Value -is [string]) {
                $normalized = ($Value -replace ',', '').Trim()
                if ([string]::IsNullOrWhiteSpace($normalized)) {
                    return $Default
                }
                return [int]$normalized
            }
            return [int]$Value
        }
        catch {
            return $Default
        }
    }

    $findings = New-Object 'System.Collections.Generic.List[object]'

    function Add-Finding {
        param(
            [Parameter(Mandatory)]
            [ValidateSet('Risk', 'Warning', 'Info')]
            [string]$Type,
            [Parameter(Mandatory)]
            [string]$Category,
            [Parameter(Mandatory)]
            [string]$Message,
            [int]$Priority = 3
        )

        [void]$findings.Add(@{
            Type     = $Type
            Category = $Category
            Message  = $Message
            Anchor   = 'employee-experience-insights'
            Priority = $Priority
        })
    }

    function Get-ConditionalAccessPolicyState {
        param(
            [AllowNull()]
            $Policy
        )

        $state = [string](Get-SummaryValue -Record $Policy -Key @('State', 'state', 'PolicyState', 'policyState') -Default '')
        switch ($state.Trim().ToLowerInvariant()) {
            'enabled'                           { return 'enabled' }
            'enabledforreportingbutnotenforced' { return 'enabledForReportingButNotEnforced' }
            'reportonly'                        { return 'enabledForReportingButNotEnforced' }
            'report-only'                       { return 'enabledForReportingButNotEnforced' }
            'disabled'                          { return 'disabled' }
            default                             { return 'unknown' }
        }
    }

    function Convert-ToBoolean {
        param(
            [AllowNull()]
            $Value
        )

        if ($null -eq $Value) { return $null }
        if ($Value -is [bool]) { return [bool]$Value }

        switch (([string]$Value).Trim().ToLowerInvariant()) {
            { $_ -in @('true', 'yes', 'enabled', 'on', '1') } { return $true }
            { $_ -in @('false', 'no', 'disabled', 'off', '0') } { return $false }
            default { return $null }
        }
    }

    function Convert-ToStringArray {
        param(
            [AllowNull()]
            $Value
        )

        if ($null -eq $Value) { return @() }
        if ($Value -is [string]) {
            return @(
                $Value -split '[,;|]' |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
        }
        if ($Value -is [System.Collections.IEnumerable]) {
            return @(
                $Value |
                    ForEach-Object { [string]$_ } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
        }
        return @([string]$Value)
    }

    function Test-ConditionalAccessPolicyRequiresMfa {
        param(
            [AllowNull()]
            $Policy
        )

        if ($null -eq $Policy) { return $false }

        $explicitRequirement = Convert-ToBoolean -Value (Get-SummaryValue -Record $Policy -Key @('RequiresMfaEnforcement', 'requiresMfaEnforcement', 'RequiresMfa', 'requiresMfa') -Default $null)
        if ($explicitRequirement -eq $true) { return $true }

        $grantControls = Get-SummaryValue -Record $Policy -Key @('GrantControls', 'grantControls') -Default $null
        $builtInControls = Get-SummaryValue -Record $Policy -Key @('GrantControls_BuiltInControls', 'grantControls_BuiltInControls', 'BuiltInControls', 'builtInControls') -Default $null
        if ($null -eq $builtInControls -and $null -ne $grantControls) {
            $builtInControls = Get-SummaryValue -Record $grantControls -Key @('BuiltInControls', 'builtInControls') -Default $null
        }
        if (@(Convert-ToStringArray -Value $builtInControls | Where-Object { $_ -match '^(?i:mfa|multiFactorAuthentication)$' }).Count -gt 0) {
            return $true
        }

        $usesAuthenticationStrength = Convert-ToBoolean -Value (Get-SummaryValue -Record $Policy -Key @('UsesAuthenticationStrengthForMfa', 'usesAuthenticationStrengthForMfa') -Default $null)
        if ($usesAuthenticationStrength -eq $true) { return $true }

        $authenticationStrength = Get-SummaryValue -Record $Policy -Key @('GrantControls_AuthenticationStrength', 'grantControls_AuthenticationStrength', 'AuthenticationStrength', 'authenticationStrength') -Default $null
        if ($null -eq $authenticationStrength -and $null -ne $grantControls) {
            $authenticationStrength = Get-SummaryValue -Record $grantControls -Key @('AuthenticationStrength', 'authenticationStrength') -Default $null
        }
        if ($null -eq $authenticationStrength) { return $false }
        if ($authenticationStrength -is [System.Collections.IDictionary]) { return ([System.Collections.IDictionary]$authenticationStrength).Count -gt 0 }
        if ($authenticationStrength.PSObject -and @($authenticationStrength.PSObject.Properties).Count -gt 0 -and -not ($authenticationStrength -is [string])) { return $true }

        $authenticationStrengthText = ([string]$authenticationStrength).Trim()
        return (-not [string]::IsNullOrWhiteSpace($authenticationStrengthText) -and $authenticationStrengthText -notin @('{}', 'null'))
    }

    $summaryRecord = $EmailActivitySummary
    if ($summaryRecord -is [array]) {
        $summaryRecord = @($summaryRecord | Select-Object -First 1)
        $summaryRecord = if ($summaryRecord.Count -gt 0) { $summaryRecord[0] } else { $null }
    }

    $reportRows = Convert-ToInt -Value (Get-SummaryValue -Record $summaryRecord -Key 'ReportRows' -Default 0)
    $activeUsers = Convert-ToInt -Value (Get-SummaryValue -Record $summaryRecord -Key 'ActiveUsers' -Default 0)
    $period = [string](Get-SummaryValue -Record $summaryRecord -Key 'PeriodDuration' -Default 'Unknown')
    if ([string]::IsNullOrWhiteSpace($period)) {
        $period = 'Unknown'
    }

    $teamsReportRows = Convert-ToInt -Value (Get-SummaryValue -Record $EmployeeExperienceInsightsSummary -Key 'TeamsReportRows' -Default 0)
    $teamsActiveUsers = Convert-ToInt -Value (Get-SummaryValue -Record $EmployeeExperienceInsightsSummary -Key 'TeamsActiveUsers' -Default 0)
    $groupsReportRows = Convert-ToInt -Value (Get-SummaryValue -Record $EmployeeExperienceInsightsSummary -Key 'GroupsReportRows' -Default 0)
    $activeGroups = Convert-ToInt -Value (Get-SummaryValue -Record $EmployeeExperienceInsightsSummary -Key 'ActiveGroups' -Default 0)

    if ($teamsReportRows -eq 0 -and @($TeamsTopUsers).Count -gt 0) {
        $teamsReportRows = @($TeamsTopUsers).Count
    }
    if ($groupsReportRows -eq 0 -and @($GroupsTopGroups).Count -gt 0) {
        $groupsReportRows = @($GroupsTopGroups).Count
    }

    if (
        $reportRows -eq 0 -and
        @($TopSenders).Count -eq 0 -and
        @($TopReceivers).Count -eq 0 -and
        $teamsReportRows -eq 0 -and
        $groupsReportRows -eq 0
    ) {
        Add-Finding -Type Warning -Category 'Security Telemetry Coverage' -Priority 2 -Message 'No collaboration activity detail rows were collected from Microsoft 365 usage reports. This limits behavior-baseline and anomaly-review usefulness.'
    }
    else {
        Add-Finding -Type Info -Category 'Security Telemetry Coverage' -Priority 3 -Message "Telemetry collected for baseline analysis (period=$period; emailRows=$reportRows; teamsRows=$teamsReportRows; groupRows=$groupsReportRows)."
    }

    if ($TotalUsers -gt 0 -and $activeUsers -ge 0) {
        $coveragePct = [math]::Round((($activeUsers / [double]$TotalUsers) * 100), 1)
        if ($coveragePct -lt 25) {
            Add-Finding -Type Warning -Category 'Security Telemetry Coverage' -Priority 2 -Message "Only $activeUsers of $TotalUsers users show activity in the selected period ($coveragePct%). Validate report scope, period, and data-retention visibility."
        }
    }

    $displayConcealedNames = $null
    if ($AdminReportSettings) {
        $rawDisplayConcealed = Get-SummaryValue -Record $AdminReportSettings -Key 'DisplayConcealedNames' -Default $null
        if ($null -ne $rawDisplayConcealed -and -not [string]::IsNullOrWhiteSpace([string]$rawDisplayConcealed)) {
            try { $displayConcealedNames = [bool]$rawDisplayConcealed } catch { $displayConcealedNames = $null }
        }
    }

    if ($displayConcealedNames -eq $true) {
        Add-Finding -Type Warning -Category 'Report Identity Visibility' -Priority 2 -Message 'Microsoft 365 usage reports are configured to conceal user names. This reduces actor-level investigation fidelity for security and incident response.'
    }
    elseif ($displayConcealedNames -eq $false) {
        Add-Finding -Type Info -Category 'Report Identity Visibility' -Priority 3 -Message 'Microsoft 365 usage reports are configured to show identifiable names for authorized security and operations analysis.'
    }

    $allPolicies = @($ConditionalAccessPolicies | Where-Object { $null -ne $_ })
    $enabledPolicies = @(
        $allPolicies | Where-Object {
            (Get-ConditionalAccessPolicyState -Policy $_) -eq 'enabled'
        }
    )
    $reportOnlyPolicies = @(
        $allPolicies | Where-Object {
            (Get-ConditionalAccessPolicyState -Policy $_) -eq 'enabledForReportingButNotEnforced'
        }
    )

    if ($allPolicies.Count -eq 0) {
        Add-Finding -Type Warning -Category 'Conditional Access Enforcement' -Priority 2 -Message 'No Conditional Access policies were collected. Validate Graph permissions and ensure baseline zero-trust controls are configured.'
    }
    elseif ($enabledPolicies.Count -eq 0) {
        Add-Finding -Type Warning -Category 'Conditional Access Enforcement' -Priority 2 -Message "Conditional Access policies were collected ($($allPolicies.Count)) but none are in an enabled state."
    }
    else {
        Add-Finding -Type Info -Category 'Conditional Access Enforcement' -Priority 3 -Message "Conditional Access enabled policies detected: $($enabledPolicies.Count) of $($allPolicies.Count) collected."
    }

    if ($reportOnlyPolicies.Count -gt 0) {
        Add-Finding -Type Warning -Category 'Conditional Access Staging' -Priority 2 -Message "$($reportOnlyPolicies.Count) Conditional Access policy/policies are in report-only state and do not currently enforce access controls."
    }

    $enabledMfaPolicies = @(
        $enabledPolicies | Where-Object { Test-ConditionalAccessPolicyRequiresMfa -Policy $_ }
    )
    $reportOnlyMfaPolicies = @(
        $reportOnlyPolicies | Where-Object { Test-ConditionalAccessPolicyRequiresMfa -Policy $_ }
    )

    if ($enabledMfaPolicies.Count -gt 0) {
        Add-Finding -Type Info -Category 'MFA Coverage' -Priority 3 -Message "Active MFA enforcement is signaled by $($enabledMfaPolicies.Count) enabled Conditional Access policy/policies with an MFA or authentication-strength grant."
    }
    elseif ($reportOnlyMfaPolicies.Count -gt 0) {
        Add-Finding -Type Warning -Category 'MFA Coverage' -Priority 1 -Message "$($reportOnlyMfaPolicies.Count) MFA-grant Conditional Access policy/policies are report-only. Authentication methods may be available, but enabled MFA Conditional Access enforcement evidence was not detected."
    }
    else {
        Add-Finding -Type Warning -Category 'MFA Coverage' -Priority 2 -Message 'No enabled Conditional Access policy with an MFA or authentication-strength grant was detected. Authentication-method availability does not establish enforcement; validate Security Defaults or other enforcement evidence separately.'
    }

    $passwordlessMethods = @()
    if ($AuthConfig) {
        $rawPasswordlessMethods = Get-SummaryValue -Record $AuthConfig -Key 'PasswordlessMethods' -Default @()
        if ($rawPasswordlessMethods -is [string]) {
            if (-not [string]::IsNullOrWhiteSpace($rawPasswordlessMethods)) {
                $passwordlessMethods = @($rawPasswordlessMethods)
            }
        }
        else {
            $passwordlessMethods = @($rawPasswordlessMethods)
        }
    }
    $passwordlessMethodCount = @($passwordlessMethods | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count
    if ($passwordlessMethodCount -eq 0) {
        Add-Finding -Type Warning -Category 'Passwordless Readiness' -Priority 2 -Message 'No passwordless method signals were detected in the authentication summary. Consider phishing-resistant authentication adoption planning.'
    }
    else {
        Add-Finding -Type Info -Category 'Passwordless Readiness' -Priority 3 -Message "Passwordless methods detected in authentication summary: $passwordlessMethodCount."
    }

    $topSenderRows = @($TopSenders)
    if ($topSenderRows.Count -gt 1) {
        $orderedSenders = @(
            $topSenderRows |
                Sort-Object {
                    Convert-ToInt -Value (Get-SummaryValue -Record $_ -Key 'SendCount' -Default 0)
                } -Descending
        )
        $topSender = $orderedSenders | Select-Object -First 1
        $topSenderCount = Convert-ToInt -Value (Get-SummaryValue -Record $topSender -Key 'SendCount' -Default 0)
        $topTenTotal = Convert-ToInt -Value (($orderedSenders | Select-Object -First 10 | Measure-Object -Property SendCount -Sum).Sum)
        if ($topTenTotal -gt 0) {
            $topSenderPct = [math]::Round((($topSenderCount / [double]$topTenTotal) * 100), 1)
            if ($topSenderPct -ge 45) {
                Add-Finding -Type Warning -Category 'Threat Surface Baseline' -Priority 2 -Message "A single account contributes $topSenderPct% of top sender volume in the sampled telemetry. Validate whether this is expected service-account behavior."
            }
        }
    }

    if ($groupsReportRows -gt 0 -and $activeGroups -eq 0) {
        Add-Finding -Type Info -Category 'Threat Surface Baseline' -Priority 3 -Message 'Office 365 Groups activity report was collected but no active groups were detected in the selected period.'
    }

    return @{
        Findings = @($findings.ToArray())
    }
}
