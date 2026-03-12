[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ExportPath,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
    [string[]]$OutputProfile = @('SolutionsEngineer'),
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret
)

function Resolve-ArrayaRepoRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$StartPath
    )

    $candidate = (Resolve-Path -Path $StartPath).Path
    while ($true) {
        $runnerManifestPath = Join-Path $candidate 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
        if (Test-Path -Path $runnerManifestPath) {
            return $candidate
        }

        $parent = Split-Path -Path $candidate -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $candidate) {
            break
        }
        $candidate = $parent
    }

    throw "Could not resolve repository root from: $StartPath"
}

$repoRoot = Resolve-ArrayaRepoRoot -StartPath $PSScriptRoot
$runnerManifestPath = Join-Path $repoRoot 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
if (-not (Get-Module -Name 'Arraya.M365.AssessmentRunner' -ErrorAction SilentlyContinue)) {
    Import-Module -Name $runnerManifestPath -ErrorAction Stop
}

$invokeParams = @{}
if ($PSBoundParameters.ContainsKey('ExportPath')) { $invokeParams.ExportPath = $ExportPath }
if ($PSBoundParameters.ContainsKey('OutputProfile')) { $invokeParams.OutputProfile = $OutputProfile }
if ($PSBoundParameters.ContainsKey('TenantId')) { $invokeParams.TenantId = $TenantId }
if ($PSBoundParameters.ContainsKey('CertificateThumbprint')) { $invokeParams.CertificateThumbprint = $CertificateThumbprint }
if ($PSBoundParameters.ContainsKey('ClientId')) { $invokeParams.ClientId = $ClientId }
if ($PSBoundParameters.ContainsKey('ClientSecret')) { $invokeParams.ClientSecret = $ClientSecret }

Invoke-M365TenantDataCollection @invokeParams
