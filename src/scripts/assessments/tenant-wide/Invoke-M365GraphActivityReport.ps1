[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
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
    [Parameter(Mandatory = $false)]
    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$PeriodDuration = 'D90',
    [Parameter(Mandatory = $false)]
    [switch]$UseBeta
)

$resolveRepoRootHelperPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\shared\Resolve-ArrayaRepoRoot.ps1'))
if (-not (Test-Path -Path $resolveRepoRootHelperPath)) {
    throw "Repo-root helper script not found: $resolveRepoRootHelperPath"
}
. $resolveRepoRootHelperPath

$repoRoot = Resolve-ArrayaRepoRoot -StartPath $PSScriptRoot
$runnerManifestPath = Join-Path $repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
if (-not (Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction SilentlyContinue)) {
    Import-Module -Name $runnerManifestPath -ErrorAction Stop
}

Invoke-GraphActivityAssessment -ServiceName $ServiceName -PeriodDuration $PeriodDuration -UseBeta:$UseBeta
