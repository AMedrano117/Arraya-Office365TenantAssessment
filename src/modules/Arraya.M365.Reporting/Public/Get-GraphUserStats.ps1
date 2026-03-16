function Get-GraphUserStats {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateSet('D90', 'D180', 'D60', 'D30')]
        [string]$PeriodDuration = 'D90',
        [Parameter(Mandatory = $false)]
        $Context
    )

    function Get-ReportRows {
        param(
            [Parameter(Mandatory = $true)]
            [string]$ServiceName,
            [Parameter(Mandatory = $true)]
            [string]$DataName
        )

        try {
            Write-Host "Fetching $DataName from the Graph..." -NoNewline -ForegroundColor Cyan
            $data = @(Office365Custom\Get-GraphAPIActivityReport -ServiceName $ServiceName -PeriodDuration $PeriodDuration)
            Write-Host 'Completed' -ForegroundColor Green
            return $data
        }
        catch {
            Write-Log -Type ERROR -Message "[Get-GraphUserStats] Unable to fetch $DataName with Graph API. Exception: $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
            return @()
        }
    }

    function Get-SignInRows {
        try {
            Write-Host 'Fetching User Sign-In Data from the Graph...' -NoNewline -ForegroundColor Cyan
            $data = @(Office365Custom\Get-GraphData -Uri "https://graph.microsoft.com/v1.0/users?`$select=displayName,userPrincipalName,mail,id,createdDateTime,signInActivity,userType" -PageSize 999 -Activity 'User Sign-In Data')
            Write-Host 'Completed' -ForegroundColor Green
            return $data
        }
        catch {
            Write-Log -Type ERROR -Message "[Get-GraphUserStats] Unable to fetch User Sign-In Data with Graph API. Exception: $($_.Exception.Message)" -ExportFileLocation $exportDetails -CaptureError -ErrorRecordVar $_
            return @()
        }
    }

    function Get-OrCreateUserRecord {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Upn
        )

        if (-not $userWorkloadMap.ContainsKey($Upn)) {
            $userWorkloadMap[$Upn] = [ordered]@{
                UPN                          = $Upn
                DisplayName                  = ''
                Status                       = 'Account unused'
                LastSignIn                   = 'No sign in data found'
                DaysSinceSignIn              = 'N/A'
                EXOLastActive                = 'No Activity'
                EXODaysSinceActive           = 'N/A'
                EXOQuotaUsed                 = 0
                EXOItems                     = 0
                EXOSendCount                 = 0
                EXOReadCount                 = 0
                EXOReceiveCount              = 0
                TeamsLastActive              = 'No Activity'
                TeamsDaysSinceActive         = 'N/A'
                TeamsUsage                   = 'No Activity'
                TeamsChannelChat             = 0
                TeamsPrivateChat             = 0
                TeamsMeetings                = 0
                TeamsCalls                   = 0
                SPOLastActive                = 'No Activity'
                SPODaysSinceActive           = 'N/A'
                SPOViewedEditedFiles         = 0
                SPOSyncedFiles               = 0
                SPOSharedExtFiles            = 0
                SPOSharedIntFiles            = 0
                SPOVisitedPages              = 0
                OneDriveLastActive           = 'No Activity'
                OneDriveDaysSinceActive      = 'N/A'
                OneDriveFiles                = 0
                OneDriveStorage              = 0
                OneDriveQuota                = 1024
                YammerLastActive             = 'No Activity'
                YammerDaysSinceActive        = 'N/A'
                YammerPosts                  = 0
                YammerReads                  = 0
                YammerLikes                  = 0
                License                      = ''
                OneDriveSite                 = ''
                IsDeleted                    = $null
                EXOReportDate                = ''
                TeamsReportDate              = ''
                'AllServices-AverageDaysSinceUse' = 365
                _Days                        = @{
                    Exo    = 365
                    Teams  = 365
                    SPO    = 365
                    OD     = 365
                    Yammer = 365
                }
            }
        }

        return $userWorkloadMap[$Upn]
    }

    function Convert-ToActivityDays {
        param(
            [AllowNull()]
            $RawDate
        )

        if ([string]::IsNullOrWhiteSpace([string]$RawDate)) {
            return @{
                Display = 'No Activity'
                Days    = 'N/A'
            }
        }

        $activityDate = Get-Date $RawDate -Format 'dd-MMM-yyyy'
        return @{
            Display = $activityDate
            Days    = (New-TimeSpan($activityDate)).Days
        }
    }

    $Context = Resolve-ArrayaReportingCollectorContext -Context $Context
    $tenantStatsHash = $Context.TenantStats
    $exportDetails = $Context.ExportFileLocation
    $initialStart = Get-ArrayaReportingCollectorStartTime -Context $Context
    $tenantStatsHash['AllGraphUserStats'] = @{}
    $userWorkloadMap = @{}
    $signInLookup = @{}
    $Context.Runtime['GraphUserStatsWorkloadMap'] = $userWorkloadMap

    $startTime1 = Get-Date
    Write-Log -Type INFO -Message '[Get-GraphUserStats] Gathering all User Combined Summary Details from Graph' -ExportFileLocation $exportDetails
    Write-Host ''

    $teamsRows = Get-ReportRows -ServiceName 'TeamsUser' -DataName 'Teams User Report'
    $oneDriveRows = Get-ReportRows -ServiceName 'OneDriveUsage' -DataName 'OneDrive Usage Report'
    $emailRows = Get-ReportRows -ServiceName 'EmailActivity' -DataName 'Exchange Activity Report'
    $mailboxRows = Get-ReportRows -ServiceName 'MailboxUsage' -DataName 'Mailbox Usage Report'
    $spoRows = Get-ReportRows -ServiceName 'SharePointUser' -DataName 'SharePoint Activity Report'
    $yammerRows = Get-ReportRows -ServiceName 'YammerActivity' -DataName 'Yammer Activity Report'
    $signInRows = Get-SignInRows

    Write-Log -Type INFO -Message '[Get-GraphUserStats] Processing User Sign-In Data fetched from Graph' -ExportFileLocation $exportDetails
    foreach ($user in @($signInRows | Where-Object { $_.UserType -eq 'Member' } | Sort-Object UserPrincipalName -Unique)) {
        if ([string]::IsNullOrWhiteSpace([string]$user.UserPrincipalName)) {
            continue
        }

        $signInLookup[[string]$user.UserPrincipalName] = if ($user.SignInActivity.LastSignInDateTime) {
            Get-Date($user.SignInActivity.LastSignInDateTime)
        }
        else {
            $null
        }
    }

    Write-Host 'Processing activity data fetched from the Graph...'
    Write-Log -Type INFO -Message '[Get-GraphUserStats] Processing activity data fetched from the Graph' -ExportFileLocation $exportDetails

    foreach ($row in $teamsRows) {
        $upn = [string](Get-ArrayaObjectValue -Object $row -Names @('User Principal Name'))
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }
        $record = Get-OrCreateUserRecord -Upn $upn
        $activity = Convert-ToActivityDays -RawDate (Get-ArrayaObjectValue -Object $row -Names @('Last Activity Date'))
        $record.TeamsLastActive = $activity.Display
        $record.TeamsDaysSinceActive = $activity.Days
        if ($activity.Days -ne 'N/A') { $record._Days.Teams = [int]$activity.Days }
        $record.TeamsReportDate = if (Get-ArrayaObjectValue -Object $row -Names @('Report Refresh Date')) { Get-Date((Get-ArrayaObjectValue -Object $row -Names @('Report Refresh Date'))) -Format 'dd-MMM-yyyy' } else { '' }
        $record.License = [string](Get-ArrayaObjectValue -Object $row -Names @('Assigned Products'))
        $record.TeamsChannelChat = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Team Chat Message Count')))
        $record.TeamsPrivateChat = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Private Chat Message Count')))
        $record.TeamsCalls = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Call Count')))
        $record.TeamsMeetings = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Meeting Count')))
    }

    foreach ($row in $emailRows) {
        $upn = [string](Get-ArrayaObjectValue -Object $row -Names @('User Principal Name'))
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }
        $record = Get-OrCreateUserRecord -Upn $upn
        $record.DisplayName = [string](Get-ArrayaObjectValue -Object $row -Names @('Display Name'))
        $activity = Convert-ToActivityDays -RawDate (Get-ArrayaObjectValue -Object $row -Names @('Last Activity Date'))
        $record.EXOLastActive = $activity.Display
        $record.EXODaysSinceActive = $activity.Days
        if ($activity.Days -ne 'N/A') { $record._Days.Exo = [int]$activity.Days }
        $record.EXOReportDate = if (Get-ArrayaObjectValue -Object $row -Names @('Report Refresh Date')) { Get-Date((Get-ArrayaObjectValue -Object $row -Names @('Report Refresh Date'))) -Format 'dd-MMM-yyyy' } else { '' }
        $record.EXOSendCount = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Send Count')))
        $record.EXOReadCount = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Read Count')))
        $record.EXOReceiveCount = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Receive Count')))
        $record.IsDeleted = Get-ArrayaObjectValue -Object $row -Names @('Is Deleted')
    }

    foreach ($row in $mailboxRows) {
        $upn = [string](Get-ArrayaObjectValue -Object $row -Names @('User Principal Name'))
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }
        $record = Get-OrCreateUserRecord -Upn $upn
        if ([string]::IsNullOrWhiteSpace([string]$record.DisplayName)) {
            $record.DisplayName = [string](Get-ArrayaObjectValue -Object $row -Names @('Display Name'))
        }
        $activity = Convert-ToActivityDays -RawDate (Get-ArrayaObjectValue -Object $row -Names @('Last Activity Date'))
        $record.EXOLastActive = $activity.Display
        $record.EXODaysSinceActive = $activity.Days
        if ($activity.Days -ne 'N/A') { $record._Days.Exo = [int]$activity.Days }
        $record.EXOQuotaUsed = [math]::Round(([double](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Storage Used (Byte)', 'Storage Used (Bytes)')))) / 1GB, 2)
        $record.EXOItems = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Item Count')))
    }

    foreach ($row in $spoRows) {
        $upn = [string](Get-ArrayaObjectValue -Object $row -Names @('User Principal Name'))
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }
        $record = Get-OrCreateUserRecord -Upn $upn
        $activity = Convert-ToActivityDays -RawDate (Get-ArrayaObjectValue -Object $row -Names @('Last Activity Date'))
        $record.SPOLastActive = $activity.Display
        $record.SPODaysSinceActive = $activity.Days
        if ($activity.Days -ne 'N/A') { $record._Days.SPO = [int]$activity.Days }
        $record.SPOViewedEditedFiles = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Viewed or Edited File Count')))
        $record.SPOSyncedFiles = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Synced File Count')))
        $record.SPOSharedExtFiles = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Shared Externally File Count')))
        $record.SPOSharedIntFiles = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Shared Internally File Count')))
        $record.SPOVisitedPages = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Visited Page Count')))
    }

    foreach ($row in $oneDriveRows) {
        $upn = [string](Get-ArrayaObjectValue -Object $row -Names @('Owner Principal Name'))
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }
        $record = Get-OrCreateUserRecord -Upn $upn
        if ([string]::IsNullOrWhiteSpace([string]$record.DisplayName)) {
            $record.DisplayName = [string](Get-ArrayaObjectValue -Object $row -Names @('Owner Display Name'))
        }
        $activity = Convert-ToActivityDays -RawDate (Get-ArrayaObjectValue -Object $row -Names @('Last Activity Date'))
        $record.OneDriveLastActive = $activity.Display
        $record.OneDriveDaysSinceActive = $activity.Days
        if ($activity.Days -ne 'N/A') { $record._Days.OD = [int]$activity.Days }
        $record.OneDriveSite = [string](Get-ArrayaObjectValue -Object $row -Names @('Site URL'))
        $record.OneDriveFiles = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('File Count')))
        $record.OneDriveStorage = [math]::Round(([double](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Storage Used (Byte)', 'Storage Used (Bytes)')))) / 1GB, 4)
        $record.OneDriveQuota = [math]::Round(([double](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Storage Allocated (Byte)', 'Storage Allocated (Bytes)')))) / 1GB, 2)
    }

    foreach ($row in $yammerRows) {
        $upn = [string](Get-ArrayaObjectValue -Object $row -Names @('User Principal Name'))
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }
        $record = Get-OrCreateUserRecord -Upn $upn
        if ([string]::IsNullOrWhiteSpace([string]$record.DisplayName)) {
            $record.DisplayName = [string](Get-ArrayaObjectValue -Object $row -Names @('Display Name'))
        }
        $activity = Convert-ToActivityDays -RawDate (Get-ArrayaObjectValue -Object $row -Names @('Last Activity Date'))
        $record.YammerLastActive = $activity.Display
        $record.YammerDaysSinceActive = $activity.Days
        if ($activity.Days -ne 'N/A') { $record._Days.Yammer = [int]$activity.Days }
        $record.YammerPosts = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Posted Count')))
        $record.YammerReads = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Read Count')))
        $record.YammerLikes = [int](Convert-ArrayaToNumber -Value (Get-ArrayaObjectValue -Object $row -Names @('Liked Count')))
    }

    $graphUserExtractionProgressId = 84
    $userNumber = 0
    $totalUsers = $userWorkloadMap.Count
    foreach ($upn in @($userWorkloadMap.Keys | Sort-Object)) {
        $userNumber++
        Write-ProgressHelper -Id $graphUserExtractionProgressId -Activity 'Extracting User Data' -Operation "Processing user: $upn" -Index $userNumber -Total ([Math]::Max($totalUsers, 1))
        Write-Log -Type INFO -Message "[Get-GraphUserStats] Processing User Data for $upn" -ExportFileLocation $exportDetails

        $record = $userWorkloadMap[$upn]
        $lastAccountSignIn = if ($signInLookup.ContainsKey($upn)) { $signInLookup[$upn] } else { $null }
        if ($null -eq $lastAccountSignIn) {
            $record.LastSignIn = 'No sign in data found'
            $record.DaysSinceSignIn = 'N/A'
        }
        else {
            $record.LastSignIn = Get-Date($lastAccountSignIn) -Format g
            $record.DaysSinceSignIn = (New-TimeSpan($lastAccountSignIn)).Days
        }

        if ($record.TeamsDaysSinceActive -ne 'N/A' -and [int]$record.TeamsDaysSinceActive -le 30 -and $record.TeamsPrivateChat -gt 50) {
            $record.TeamsUsage = 'Active'
        }
        elseif ($record.TeamsDaysSinceActive -ne 'N/A' -and [int]$record.TeamsDaysSinceActive -ge 31 -and [int]$record.TeamsDaysSinceActive -le 60 -and $record.TeamsPrivateChat -ge 100) {
            $record.TeamsUsage = 'Low'
        }
        elseif ($record.TeamsDaysSinceActive -ne 'N/A' -and [int]$record.TeamsDaysSinceActive -ge 31 -and [int]$record.TeamsDaysSinceActive -le 90 -and $record.TeamsPrivateChat -le 100) {
            $record.TeamsUsage = 'Minimal'
        }

        $averageDaysSinceUse = [math]::Round((($record._Days.Exo + $record._Days.Teams + $record._Days.SPO + $record._Days.OD) / 4), 2)
        $record['AllServices-AverageDaysSinceUse'] = $averageDaysSinceUse

        switch ($averageDaysSinceUse) {
            { $PSItem -le 8 } { $record.Status = 'Heavy usage' }
            { $PSItem -ge 9 -and $PSItem -le 50 } { $record.Status = 'Moderate usage' }
            { $PSItem -ge 51 -and $PSItem -le 120 } { $record.Status = 'Poor usage' }
            { $PSItem -ge 121 -and $PSItem -le 300 } { $record.Status = 'Review account' }
            default { $record.Status = 'Account unused' }
        }

        if (
            $record._Days.Exo -le 14 -or
            $record._Days.Teams -le 14 -or
            $record._Days.SPO -le 14 -or
            $record._Days.OD -le 14 -or
            $record._Days.Yammer -le 14
        ) {
            $record.Status = 'Account in use'
        }

        $output = [ordered]@{}
        foreach ($key in $record.Keys) {
            if ($key -eq '_Days') { continue }
            $output[$key] = $record[$key]
        }
        $tenantStatsHash['AllGraphUserStats'][$upn] = [PSCustomObject]$output
    }

    Write-ProgressHelper -Id $graphUserExtractionProgressId -Total ([Math]::Max($totalUsers, 1)) -Activity 'Extracting User Data' -Completed

    $endTime = Get-Date
    $graphTime = $endTime - $startTime1
    $accountsPerMinute = if ($graphTime.TotalSeconds -gt 0) {
        [math]::Round(($tenantStatsHash['AllGraphUserStats'].Values.Count / ($graphTime.TotalSeconds / 60)), 2)
    }
    else {
        0
    }

    Write-Verbose ''
    Write-Verbose 'Statistics for Graph Report Script'
    Write-Verbose '----------------------------------'
    Write-Verbose "Total time for script:                   $($graphTime.Minutes):$($graphTime.Seconds)"
    Write-Verbose "Total accounts processed:                $($tenantStatsHash['AllGraphUserStats'].Values.Count)"
    Write-Verbose "Accounts processed per minute:           $accountsPerMinute"
    Write-Verbose ''

    $Context.Runtime['GraphUserStatsSummary'] = [PSCustomObject]@{
        PeriodDuration      = $PeriodDuration
        AccountsProcessed   = [int]$tenantStatsHash['AllGraphUserStats'].Values.Count
        TeamsReportRows     = [int]@($teamsRows).Count
        EmailReportRows     = [int]@($emailRows).Count
        MailboxReportRows   = [int]@($mailboxRows).Count
        SharePointReportRows = [int]@($spoRows).Count
        OneDriveReportRows  = [int]@($oneDriveRows).Count
        YammerReportRows    = [int]@($yammerRows).Count
        SignInRows          = [int]@($signInRows).Count
        StartedAt           = $initialStart
        CompletedAt         = $endTime
    }

    return $tenantStatsHash['AllGraphUserStats']
}
