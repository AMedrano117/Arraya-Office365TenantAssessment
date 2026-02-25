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
    try {
        # Generate a temporary file name with a .csv extension
        $tempFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ([System.IO.Path]::GetRandomFileName() + ".csv"))
        Write-Verbose "Saving report to temporary file: $tempFile"

        # Invoke the Graph API request and output the result to the temporary file
        $resulttmp = Invoke-MgGraphRequest -Method GET -Uri $ActivityReport -ErrorAction Stop -OutputFilePath $tempFile

        # Import the CSV data from the temporary file
        $result = Import-Csv -Path $tempFile

        # Clean up: remove the temporary file
        Remove-Item -Path $tempFile -Force

        return $result

    } catch {
        if ($_.Exception.Message -like '*Authentication needed. Please call Connect-MgGraph*') {
            Write-Error "Error retrieving data from Microsoft Graph. Please call Connect-MgGraph first."
            #throw $_.Exception.Message 
        } else {
            #Write-Warning "Error retrieving data from Microsoft Graph."
            Write-Error $_.Exception.Message
        }
    }
}
