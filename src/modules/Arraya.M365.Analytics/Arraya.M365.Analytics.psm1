# Analysis helpers shared by the HTML assessment report and the tenant improvement plan.
#
# These functions previously existed as two parallel copies, one prefixed Assessment* in
# the HTML helper and one prefixed Improvement* in the plan generator. The copies had begun
# to diverge - different accessor helpers, different parameter names - which meant two
# deliverables generated from a single run could disagree about the same tenant. One
# implementation lives here so that cannot happen again.

$commonModuleManifestPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\Arraya.M365.Common\Arraya.M365.Common.psd1'))
if (Test-Path -Path $commonModuleManifestPath) {
    Import-Module -Name $commonModuleManifestPath -ErrorAction Stop -WarningAction SilentlyContinue -DisableNameChecking
}

$private = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter *.ps1 -ErrorAction SilentlyContinue)
$public = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public') -Filter *.ps1 -ErrorAction SilentlyContinue)
foreach ($file in @($private) + @($public)) {
    . $file.FullName
}

Export-ModuleMember -Function @(
    'Add-ArrayaGroupLicensingEffectiveCounts',
    'Get-ArrayaConditionalAccessPolicyState',
    'Get-ArrayaDateText',
    'Get-ArrayaGroupLicenseAssignedUserCountLookup',
    'Get-ArrayaLatestDateFromRows',
    'Get-ArrayaLicenseSkuLookup',
    'Get-ArrayaLicenseSuiteFamily',
    'Get-ArrayaNormalizedPrincipalKey',
    'Get-ArrayaUserLicenseActivityState',
    'Get-ArrayaUserLicenseAssignmentRows',
    'New-ArrayaActivityLookup',
    'New-ArrayaConditionalAccessOptimizationRows',
    'New-ArrayaGroupLicensingSummaryRows',
    'New-ArrayaLicenseOptimizationCandidateRows',
    'New-ArrayaMfaMethodPostureSummaryRows',
    'New-ArrayaPrivilegedAccessRemediationSummaryRows',
    'New-ArrayaTeamsGroupsCleanupCandidateRows',
    'Resolve-ArrayaLicenseSkuName',
    'Test-ArrayaGroupManagesLicenses'
)
