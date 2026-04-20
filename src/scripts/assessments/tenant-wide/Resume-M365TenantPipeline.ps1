[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CheckpointRoot,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Connection', 'TenantOverview', 'Identity', 'Exchange', 'Collaboration', 'Endpoint', 'Governance', 'Export')]
    [string]$ThroughPhase,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Connection', 'TenantOverview', 'Identity', 'Exchange', 'Collaboration', 'Endpoint', 'Governance', 'Export')]
    [string]$Phase
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

$invokeParams = @{
    CheckpointRoot = $CheckpointRoot
}
if ($PSBoundParameters.ContainsKey('ThroughPhase')) { $invokeParams.ThroughPhase = $ThroughPhase }
if ($PSBoundParameters.ContainsKey('Phase')) { $invokeParams.Phase = $Phase }

Resume-M365TenantPipeline @invokeParams
