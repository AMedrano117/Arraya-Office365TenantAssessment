function Resolve-ArrayaRepoRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$StartPath,
        [Parameter(Mandatory = $false)]
        [string]$SentinelRelativePath = 'src\modules\Arraya.M365.AssessmentRunner\Arraya.M365.AssessmentRunner.psd1'
    )

    $candidate = (Resolve-Path -Path $StartPath).Path
    while ($true) {
        $sentinelPath = Join-Path $candidate $SentinelRelativePath
        if (Test-Path -Path $sentinelPath) {
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
