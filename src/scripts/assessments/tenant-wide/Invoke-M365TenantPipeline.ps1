[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Full', 'CollectOnly', 'ExportOnly', 'PreflightOnly')]
    [string]$Mode = 'Full',
    [Parameter(Mandatory = $false)]
    [ValidateSet('Connection', 'TenantOverview', 'Identity', 'Exchange', 'Collaboration', 'Endpoint', 'Governance', 'Export')]
    [string]$ThroughPhase,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Connection', 'TenantOverview', 'Identity', 'Exchange', 'Collaboration', 'Endpoint', 'Governance', 'Export')]
    [string]$Phase,
    [Parameter(Mandatory = $false)]
    [string]$CheckpointRoot,
    [Parameter(Mandatory = $false)]
    [string]$ResumeFromCheckpointRoot,
    [Parameter(Mandatory = $false)]
    [string]$ExportPath,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
    [string[]]$OutputProfile = @('SolutionsEngineer'),
    [Parameter(Mandatory = $false)]
    [string]$AssessmentJsonPath,
    [Parameter(Mandatory = $false)]
    [switch]$SkipHtmlReport,
    [Parameter(Mandatory = $false)]
    [switch]$SkipPdfReport,
    [Parameter(Mandatory = $false)]
    [switch]$SkipJsonReport,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeLegacyAssessmentArtifacts,
    [Parameter(Mandatory = $false)]
    [switch]$StoreTenantStatsGlobal,
    [Parameter(Mandatory = $false)]
    [string]$TenantStatsVariableName = 'ArrayaTenantStats',
    [Parameter(Mandatory = $false)]
    [switch]$SkipAuth,
    [Parameter(Mandatory = $false)]
    [switch]$SkipPermissionPreflight,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
    [string]$AuthMode,
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret
)

$resolveRepoRootHelperPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\shared\Resolve-ArrayaRepoRoot.ps1'))
if (-not (Test-Path -Path $resolveRepoRootHelperPath)) {
    throw "Repo-root helper script not found: $resolveRepoRootHelperPath"
}
. $resolveRepoRootHelperPath

$repoRoot = Resolve-ArrayaRepoRoot -StartPath $PSScriptRoot
$runnerManifestPath = Join-Path $repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
if (-not (Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction SilentlyContinue)) {
    Import-Module -Name $runnerManifestPath -WarningAction SilentlyContinue -ErrorAction Stop
}

$invokeParams = @{}
foreach ($key in $PSBoundParameters.Keys) {
    $invokeParams[$key] = $PSBoundParameters[$key]
}

Invoke-M365TenantPipeline @invokeParams
