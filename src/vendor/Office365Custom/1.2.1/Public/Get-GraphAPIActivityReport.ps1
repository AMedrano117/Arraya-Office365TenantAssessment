<#
.SYNOPSIS
    Connect to Microsoft Graph API and get the Activity Report for various services
.Example
    Get-GraphAPIActivityReport -ServiceName SharePointSites -PeriodDuration D90
    Get-GraphAPIActivityReport -ServiceName OneDriveUsage -PeriodDuration D90 
    Get-GraphAPIActivityReport -ServiceName MailboxUsage -PeriodDuration D90
#>

function Get-GraphAPIActivityReport {
    param (
        [Parameter()]
        [ValidateSet('D7', 'D30', 'D90', 'D180')]
        [String]
        $PeriodDuration = "D90",
        [Parameter(Mandatory=$True, HelpMessage='Provide the Kind of Activity Report')]
        [ValidateSet('SharePointSites', 'MailboxUsage', 'SharePointUser', 'OneDriveUsage', 'OneDriveActivity', 'EmailActivity', 'YammerActivity', 'BrowserUsage', 'SkypeForBusinessActivity',
        'Office365ActiveUser', 'Office365ServicesUserCounts', 'Office365ActivationCounts', 'Office365ActivationUser', 'Office365ActiveUserCounts', 
        'Office365GroupsActivity','Office365GroupsActivityCounts', 'Office365GroupsActivityFileCounts', 'Office365GroupsActivityGroupCounts', 'Office365GroupsActivityStorage', 'Office365GroupsActivityUser',
        "TeamsTeamActivityDetail", 'TeamsUser', "TeamsTeamActivityDistributionCounts", "TeamsTeamCounts"

        )]
        [string]$ServiceName,
        [Parameter(Mandatory=$false, HelpMessage='Use Beta API version')]
        [switch]$UseBeta
    )

    # Determine the API version
    $apiVersion = if ($UseBeta) { "beta" } else { "v1.0" }

    # Construct the Activity Report URL based on the ServiceName and PeriodDuration
    switch ($ServiceName) {
        SharePointSites { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getSharePointSiteUsageDetail(period='$($PeriodDuration)')" }
        SharePointUser { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getSharePointActivityUserDetail(period='$($PeriodDuration)')" }
        TeamsUser { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getTeamsUserActivityUserDetail(period='$($PeriodDuration)')" }
        TeamsTeamActivityDetail { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getTeamsTeamActivityDetail(period='$($PeriodDuration)')" }
        TeamsTeamActivityDistributionCounts { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getTeamsTeamActivityDistributionCounts(period='$($PeriodDuration)')" }
        TeamsTeamCounts { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getTeamsTeamCounts(period='$($PeriodDuration)')" }
        OneDriveUsage { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getOneDriveUsageAccountDetail(period='$($PeriodDuration)')" }
        OneDriveActivity { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getOneDriveActivityUserDetail(period='$($PeriodDuration)')" }
        MailboxUsage { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getMailboxUsageDetail(period='$($PeriodDuration)')" }
        EmailActivity { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getEmailActivityUserDetail(period='$($PeriodDuration)')" }
        YammerActivity { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getYammerActivityUserDetail(period='$($PeriodDuration)')" }
        SkypeForBusinessActivity { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getSkypeForBusinessActivityUserDetail(period='$($PeriodDuration)')" }
        BrowserUsage { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getBrowserUsageDetail(period='$($PeriodDuration)')" }
        Office365ActiveUser { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getOffice365ActiveUserDetail(period='$($PeriodDuration)')" }
        Office365ServicesUserCounts { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getOffice365ServicesUserCounts(period='$($PeriodDuration)')" }
        Office365ActivationCounts { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getOffice365ActivationCounts(period='$($PeriodDuration)')" }
        Office365ActivationUser { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getOffice365ActivationUserDetail(period='$($PeriodDuration)')" }
        Office365ActiveUserCounts { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getOffice365ActiveUserCounts(period='$($PeriodDuration)')" }
        Office365GroupsActivity { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getOffice365GroupsActivityDetail(period='$($PeriodDuration)')" }
        Office365GroupsActivityCounts { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getOffice365GroupsActivityCounts(period='$($PeriodDuration)')" }
        Office365GroupsActivityFileCounts { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getOffice365GroupsActivityFileCounts(period='$($PeriodDuration)')" }
        Office365GroupsActivityGroupCounts { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getOffice365GroupsActivityGroupCounts(period='$($PeriodDuration)')" }
        Office365GroupsActivityStorage { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getOffice365GroupsActivityStorage(period='$($PeriodDuration)')" }
        Office365GroupsActivityUser { $ActivityReport = "https://graph.microsoft.com/$apiVersion/reports/getOffice365GroupsActivityUserDetail(period='$($PeriodDuration)')" }
        default { Write-Host "Invalid Service Name!" -ForegroundColor Red; break }
    }

    # Call the Graph API and get the data
    # $result = Get-GraphData -Uri $ActivityReport
    # Convert the result from CSV format
    # return $result | ConvertFrom-Csv
    $tempFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ([System.IO.Path]::GetRandomFileName() + ".csv"))
    $headers = $null
    if ($global:GraphHeaders) {
        $headers = $global:GraphHeaders
    }
    elseif ($global:GraphToken) {
        $headers = @{
            'Content-Type'     = 'application/json'
            'Authorization'    = "Bearer $global:GraphToken"
            'ConsistencyLevel' = 'eventual'
        }
    }

    $canUseSdk = $false
    if (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($mgContext) {
            $canUseSdk = $true
        }
    }

    try {
        Write-Verbose "Saving report to temporary file: $tempFile"
        if ($canUseSdk) {
            try {
                Invoke-MgGraphRequest -Method GET -Uri $ActivityReport -ProgressAction SilentlyContinue -ErrorAction Stop -OutputFilePath $tempFile | Out-Null
            }
            catch {
                if (-not $headers) {
                    throw
                }
                $canUseSdk = $false
            }
        }

        if (-not $canUseSdk) {
            if (-not $headers) {
                throw 'No Microsoft Graph authentication context is available for activity report retrieval.'
            }

            $savedProgressPreference = $ProgressPreference
            try {
                $ProgressPreference = 'SilentlyContinue'
                $response = Invoke-WebRequest -Uri $ActivityReport -Headers $headers -Method GET -MaximumRedirection 5 -ErrorAction Stop
            }
            finally {
                $ProgressPreference = $savedProgressPreference
            }

            if ($null -eq $response -or [string]::IsNullOrWhiteSpace($response.Content)) {
                throw "No activity report content was returned for service '$ServiceName'."
            }

            Set-Content -Path $tempFile -Value $response.Content -Encoding UTF8 -Force
        }

        if (-not (Test-Path -Path $tempFile)) {
            throw "Activity report download did not create a temporary CSV file for service '$ServiceName'."
        }

        return @(Import-Csv -Path $tempFile -ErrorAction Stop)
    }
    catch {
        if ($_.Exception.Message -like '*Authentication needed. Please call Connect-MgGraph*') {
            Write-Error 'Error retrieving data from Microsoft Graph. Please call Connect-MgGraph first.'
        }
        else {
            Write-Error $_.Exception.Message
        }
    }
    finally {
        if (Test-Path -Path $tempFile) {
            Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
}
