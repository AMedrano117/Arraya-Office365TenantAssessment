[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AssessmentJsonPath,
    [Parameter(Mandatory = $false)]
    [string]$ExportPath,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
    [string[]]$OutputProfile = @('SolutionsEngineer'),
    [Parameter(Mandatory = $false)]
    [switch]$SkipHtmlReport,
    [Parameter(Mandatory = $false)]
    [switch]$SkipPdfReport,
    [Parameter(Mandatory = $false)]
    [switch]$SkipJsonReport,
    [Parameter(Mandatory = $false)]
    [switch]$SkipImprove,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeLegacyArtifacts,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeLegacyAssessmentArtifacts,
    [Parameter(Mandatory = $false)]
    [string]$ImproveOutputFolder,
    [Parameter(Mandatory = $false)]
    [Alias('LiveRefresh')]
    [switch]$UseGraphFallback
)

$resolveRepoRootHelperPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\shared\Resolve-ArrayaRepoRoot.ps1'))
if (-not (Test-Path -Path $resolveRepoRootHelperPath)) {
    throw "Repo-root helper script not found: $resolveRepoRootHelperPath"
}
. $resolveRepoRootHelperPath

$repoRoot = Resolve-ArrayaRepoRoot -StartPath $PSScriptRoot
$runnerManifestPath = Join-Path $repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
if (-not (Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction SilentlyContinue)) {
    Import-Module -Name $runnerManifestPath -WarningAction SilentlyContinue -ErrorAction Stop
}

$invokeParams = @{ AssessmentJsonPath = $AssessmentJsonPath }
if ($PSBoundParameters.ContainsKey('ExportPath')) { $invokeParams.ExportPath = $ExportPath }
if ($PSBoundParameters.ContainsKey('OutputProfile')) { $invokeParams.OutputProfile = $OutputProfile }
if ($PSBoundParameters.ContainsKey('SkipHtmlReport')) { $invokeParams.SkipHtmlReport = $SkipHtmlReport }
if ($PSBoundParameters.ContainsKey('SkipPdfReport')) { $invokeParams.SkipPdfReport = $SkipPdfReport }
if ($PSBoundParameters.ContainsKey('SkipJsonReport')) { $invokeParams.SkipJsonReport = $SkipJsonReport }
if ($PSBoundParameters.ContainsKey('SkipImprove')) { $invokeParams.SkipImprove = $SkipImprove }
if ($PSBoundParameters.ContainsKey('IncludeLegacyArtifacts')) { $invokeParams.IncludeLegacyArtifacts = $IncludeLegacyArtifacts }
if ($PSBoundParameters.ContainsKey('IncludeLegacyAssessmentArtifacts')) { $invokeParams.IncludeLegacyAssessmentArtifacts = $IncludeLegacyAssessmentArtifacts }
if ($PSBoundParameters.ContainsKey('ImproveOutputFolder')) { $invokeParams.ImproveOutputFolder = $ImproveOutputFolder }
if ($PSBoundParameters.ContainsKey('UseGraphFallback')) { $invokeParams.UseGraphFallback = $UseGraphFallback }

Invoke-M365TenantAssessmentExport @invokeParams
