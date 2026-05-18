[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ExportPath,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
    [string[]]$OutputProfile = @('SolutionsEngineer'),
    [Parameter(Mandatory = $false)]
    [switch]$RunImprove,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeLegacyArtifacts,
    [Parameter(Mandatory = $false)]
    [string]$ImproveOutputFolder,
    [Parameter(Mandatory = $false)]
    [Alias('LiveRefresh')]
    [switch]$UseGraphFallback,
    [Parameter(Mandatory = $false)]
    [switch]$StoreTenantStatsGlobal,
    [Parameter(Mandatory = $false)]
    [string]$TenantStatsVariableName = 'ArrayaTenantStats',
    [Parameter(Mandatory = $false)]
    [switch]$SkipAuth,
    [Parameter(Mandatory = $false)]
    [switch]$SkipPermissionPreflight,
    [Parameter(Mandatory = $false)]
    [switch]$UseExistingConnections,
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
    [string]$ClientSecret,
    [Parameter(Mandatory = $false)]
    [pscredential]$ClientSecretCredential,
    [Parameter(Mandatory = $false)]
    [securestring]$ClientSecretSecure
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
if ($PSBoundParameters.ContainsKey('ExportPath')) { $invokeParams.ExportPath = $ExportPath }
if ($PSBoundParameters.ContainsKey('OutputProfile')) { $invokeParams.OutputProfile = $OutputProfile }
if ($PSBoundParameters.ContainsKey('RunImprove')) { $invokeParams.RunImprove = $RunImprove }
if ($PSBoundParameters.ContainsKey('IncludeLegacyArtifacts')) { $invokeParams.IncludeLegacyArtifacts = $IncludeLegacyArtifacts }
if ($PSBoundParameters.ContainsKey('ImproveOutputFolder')) { $invokeParams.ImproveOutputFolder = $ImproveOutputFolder }
if ($PSBoundParameters.ContainsKey('UseGraphFallback')) { $invokeParams.UseGraphFallback = $UseGraphFallback }
if ($PSBoundParameters.ContainsKey('StoreTenantStatsGlobal')) { $invokeParams.StoreTenantStatsGlobal = $StoreTenantStatsGlobal }
if ($PSBoundParameters.ContainsKey('TenantStatsVariableName')) { $invokeParams.TenantStatsVariableName = $TenantStatsVariableName }
if ($PSBoundParameters.ContainsKey('SkipAuth')) { $invokeParams.SkipAuth = $SkipAuth }
if ($PSBoundParameters.ContainsKey('SkipPermissionPreflight')) { $invokeParams.SkipPermissionPreflight = $SkipPermissionPreflight }
if ($PSBoundParameters.ContainsKey('UseExistingConnections')) { $invokeParams.UseExistingConnections = $UseExistingConnections }
if ($PSBoundParameters.ContainsKey('AuthMode')) { $invokeParams.AuthMode = $AuthMode }
if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }
if ($PSBoundParameters.ContainsKey('ClientSecretCredential')) { $invokeParams.ClientSecretCredential = $ClientSecretCredential }
if ($PSBoundParameters.ContainsKey('ClientSecretSecure')) { $invokeParams.ClientSecretSecure = $ClientSecretSecure }

Invoke-M365TenantDataCollection @invokeParams
