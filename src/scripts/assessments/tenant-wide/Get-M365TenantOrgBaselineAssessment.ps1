[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'AppCertificate')]
    [string]$AuthMode = 'Interactive',
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeRaw,
    [Parameter(Mandatory = $false)]
    [ValidateSet('Presales', 'SolutionsEngineer', 'ExecutiveLevel', 'TenantToTenantMigration', 'Geek', 'Machine')]
    [string]$OutputProfile = 'SolutionsEngineer'
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

if ($AuthMode -ne 'Interactive' -or $TenantId -or $ClientId -or $CertificateThumbprint) {
    Write-Warning 'AuthMode/TenantId/ClientId/CertificateThumbprint are ignored by this compatibility shim. Use the full tenant script for custom auth flows.'
}
if ($IncludeRaw) {
    Write-Warning 'IncludeRaw is ignored by this compatibility shim.'
}

$invokeParams = @{ ExportPath = $OutputPath }
if ($PSBoundParameters.ContainsKey('OutputProfile')) {
    $invokeParams.OutputProfile = $OutputProfile
}

Invoke-M365TenantAssessment @invokeParams
