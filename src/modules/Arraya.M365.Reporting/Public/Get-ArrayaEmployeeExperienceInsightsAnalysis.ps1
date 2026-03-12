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
        [int]$TotalUsers = 0
    )

    function Get-SummaryValue {
        param(
            [AllowNull()]
            $Record,
            [Parameter(Mandatory)]
            [string]$Key,
            [AllowNull()]
            $Default = $null
        )

        if ($null -eq $Record) { return $Default }
        if ($Record -is [System.Collections.IDictionary] -and $Record.Contains($Key)) { return $Record[$Key] }
        if ($Record -is [System.Collections.Specialized.OrderedDictionary] -and $Record.Contains($Key)) { return $Record[$Key] }
        if ($Record.PSObject -and $Record.PSObject.Properties[$Key]) { return $Record.$Key }
        return $Default
    }

    $findings = @()
    $summaryRecord = $EmailActivitySummary
    if ($summaryRecord -is [array]) {
        $summaryRecord = @($summaryRecord | Select-Object -First 1)
        $summaryRecord = if ($summaryRecord.Count -gt 0) { $summaryRecord[0] } else { $null }
    }

    $reportRows = 0
    $activeUsers = 0
    $period = 'Unknown'
    $teamsReportRows = 0
    $teamsActiveUsers = 0
    $groupsReportRows = 0
    $activeGroups = 0
    if ($summaryRecord) {
        try { $reportRows = [int](Get-SummaryValue -Record $summaryRecord -Key 'ReportRows' -Default 0) } catch { $reportRows = 0 }
        try { $activeUsers = [int](Get-SummaryValue -Record $summaryRecord -Key 'ActiveUsers' -Default 0) } catch { $activeUsers = 0 }
        $rawPeriod = [string](Get-SummaryValue -Record $summaryRecord -Key 'PeriodDuration' -Default '')
        if (-not [string]::IsNullOrWhiteSpace($rawPeriod)) { $period = $rawPeriod }
    }
    if ($EmployeeExperienceInsightsSummary) {
        try { $teamsReportRows = [int](Get-SummaryValue -Record $EmployeeExperienceInsightsSummary -Key 'TeamsReportRows' -Default 0) } catch { $teamsReportRows = 0 }
        try { $teamsActiveUsers = [int](Get-SummaryValue -Record $EmployeeExperienceInsightsSummary -Key 'TeamsActiveUsers' -Default 0) } catch { $teamsActiveUsers = 0 }
        try { $groupsReportRows = [int](Get-SummaryValue -Record $EmployeeExperienceInsightsSummary -Key 'GroupsReportRows' -Default 0) } catch { $groupsReportRows = 0 }
        try { $activeGroups = [int](Get-SummaryValue -Record $EmployeeExperienceInsightsSummary -Key 'ActiveGroups' -Default 0) } catch { $activeGroups = 0 }
    }
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
        $findings += @{
            Type     = 'Warning'
            Category = 'Usage Telemetry Coverage'
            Message  = 'No collaboration activity detail rows were collected from Microsoft 365 usage reports.'
            Anchor   = 'employee-experience-insights'
            Priority = 2
        }
    }
    else {
        $findings += @{
            Type     = 'Info'
            Category = 'Collaboration Activity Baseline'
            Message  = "Collaboration activity telemetry collected (period=$period; emailRows=$reportRows; teamsRows=$teamsReportRows; groupRows=$groupsReportRows)."
            Anchor   = 'employee-experience-insights'
            Priority = 3
        }
    }

    if ($TotalUsers -gt 0 -and $activeUsers -ge 0) {
        $coveragePct = [math]::Round((($activeUsers / [double]$TotalUsers) * 100), 1)
        if ($coveragePct -lt 25) {
            $findings += @{
                Type     = 'Warning'
                Category = 'Usage Telemetry Coverage'
                Message  = "Only $activeUsers of $TotalUsers users show activity in the selected period ($coveragePct%). Validate scope, period, and telemetry completeness."
                Anchor   = 'employee-experience-insights'
                Priority = 2
            }
        }
    }

    if ($TotalUsers -gt 0 -and $teamsActiveUsers -gt 0) {
        $teamsCoveragePct = [math]::Round((($teamsActiveUsers / [double]$TotalUsers) * 100), 1)
        if ($teamsCoveragePct -lt 20) {
            $findings += @{
                Type     = 'Info'
                Category = 'Usage Telemetry Coverage'
                Message  = "Teams activity coverage is $teamsCoveragePct% of users ($teamsActiveUsers of $TotalUsers) for period $period."
                Anchor   = 'employee-experience-insights'
                Priority = 3
            }
        }
    }

    if ($groupsReportRows -gt 0) {
        if ($activeGroups -eq 0) {
            $findings += @{
                Type     = 'Info'
                Category = 'Collaboration Activity Baseline'
                Message  = 'Office 365 Groups activity report was collected but no active groups were detected in the selected period.'
                Anchor   = 'employee-experience-insights'
                Priority = 3
            }
        }
        elseif ($activeGroups -lt 10) {
            $findings += @{
                Type     = 'Info'
                Category = 'Collaboration Activity Baseline'
                Message  = "Only $activeGroups group(s) show recent activity in the selected period."
                Anchor   = 'employee-experience-insights'
                Priority = 3
            }
        }
    }

    $displayConcealedNames = $null
    if ($AdminReportSettings) {
        if ($AdminReportSettings -is [System.Collections.IDictionary] -and $AdminReportSettings.Contains('DisplayConcealedNames')) {
            $rawDisplayConcealed = $AdminReportSettings['DisplayConcealedNames']
            if ($null -ne $rawDisplayConcealed -and -not [string]::IsNullOrWhiteSpace([string]$rawDisplayConcealed)) {
                try { $displayConcealedNames = [bool]$rawDisplayConcealed } catch { $displayConcealedNames = $null }
            }
        }
        elseif ($AdminReportSettings.PSObject -and $AdminReportSettings.PSObject.Properties['DisplayConcealedNames']) {
            $rawDisplayConcealed = $AdminReportSettings.DisplayConcealedNames
            if ($null -ne $rawDisplayConcealed -and -not [string]::IsNullOrWhiteSpace([string]$rawDisplayConcealed)) {
                try { $displayConcealedNames = [bool]$rawDisplayConcealed } catch { $displayConcealedNames = $null }
            }
        }
    }

    if ($displayConcealedNames -eq $true) {
        $findings += @{
            Type     = 'Warning'
            Category = 'Report Identity Visibility'
            Message  = 'Microsoft 365 usage reports are configured to conceal names, which reduces actor-level visibility in activity insights.'
            Anchor   = 'employee-experience-insights'
            Priority = 2
        }
    }
    elseif ($displayConcealedNames -eq $false) {
        $findings += @{
            Type     = 'Info'
            Category = 'Report Identity Visibility'
            Message  = 'Microsoft 365 usage reports are configured to show identifiable names for authorized analysis.'
            Anchor   = 'employee-experience-insights'
            Priority = 3
        }
    }
    elseif (
        $displayConcealedNames -eq $null -and
        $AdminReportSettings -and
        $AdminReportSettings.PSObject -and
        $AdminReportSettings.PSObject.Properties['Available'] -and
        [bool]$AdminReportSettings.Available -eq $false
    ) {
        $reason = if ($AdminReportSettings.PSObject.Properties['ErrorMessage'] -and -not [string]::IsNullOrWhiteSpace([string]$AdminReportSettings.ErrorMessage)) {
            [string]$AdminReportSettings.ErrorMessage
        } else {
            'Access to admin report settings is unavailable.'
        }
        $findings += @{
            Type     = 'Info'
            Category = 'Report Identity Visibility'
            Message  = "Could not determine report identity-visibility setting. $reason"
            Anchor   = 'employee-experience-insights'
            Priority = 3
        }
    }

    return @{
        Findings = @($findings)
    }
}
