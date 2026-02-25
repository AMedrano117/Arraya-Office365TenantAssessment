# Wrapper script that uses the local Office365Custom module implementation.

$officeModuleLoaderPath = Join-Path -Path $PSScriptRoot -ChildPath 'Import-Office365CustomLocal.ps1'
if (-not (Test-Path -Path $officeModuleLoaderPath)) {
    throw "Required loader script not found: $officeModuleLoaderPath"
}
. $officeModuleLoaderPath
Import-Office365CustomLocal -RepoRoot $PSScriptRoot -RequiredCommands @(
    'Office365Custom\Get-GraphAPIActivityReport'
) | Out-Null

function Get-GraphAPIActivityReport {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateSet('D7', 'D30', 'D90', 'D180')]
        [string]$PeriodDuration = 'D90',
        [Parameter(Mandatory = $true, HelpMessage = 'Provide the kind of activity report')]
        [ValidateSet(
            'SharePointSites',
            'MailboxUsage',
            'SharePointUser',
            'TeamsUser',
            'OneDriveUsage',
            'OneDriveActivity',
            'EmailActivity',
            'YammerActivity',
            'BrowserUsage',
            'SkypeForBusinessActivity',
            'Office365ActiveUser',
            'Office365ServicesUserCounts',
            'Office365ActivationCounts',
            'Office365ActivationUser',
            'Office365ActiveUserCounts',
            'Office365GroupsActivity',
            'Office365GroupsActivityCounts',
            'Office365GroupsActivityFileCounts',
            'Office365GroupsActivityGroupCounts',
            'Office365GroupsActivityStorage',
            'Office365GroupsActivityUser'
        )]
        [string]$ServiceName,
        [Parameter(Mandatory = $false, HelpMessage = 'Use beta API version')]
        [switch]$UseBeta
    )

    Office365Custom\Get-GraphAPIActivityReport -PeriodDuration $PeriodDuration -ServiceName $ServiceName -UseBeta:$UseBeta
}

# Example:
# . .\Get-GraphAPIActivityReport.ps1
# $usageData = Get-GraphAPIActivityReport -ServiceName SharePointSites -PeriodDuration D90
# $usageData | Export-Csv "C:\Temp\siteusage.csv" -NoTypeInformation
