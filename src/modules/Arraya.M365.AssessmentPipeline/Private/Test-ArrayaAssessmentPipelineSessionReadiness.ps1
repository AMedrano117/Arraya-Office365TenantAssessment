function Test-ArrayaAssessmentPipelineSessionReadiness {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,
        [Parameter(Mandatory = $false)]
        [string]$PhaseName
    )

    $graphReady = $false
    try {
        $graphReady = ($null -ne (Get-MgContext -ErrorAction SilentlyContinue))
    }
    catch {}

    $exchangeReady = [bool](Get-Command -Name 'Get-EXOMailbox' -ErrorAction SilentlyContinue) -and
        [bool](Get-Command -Name 'Get-UnifiedGroup' -ErrorAction SilentlyContinue)

    if (-not ($graphReady -and $exchangeReady)) {
        return $false
    }

    if ([string]$PhaseName -eq 'Governance') {
        $purviewReady = [bool](Get-Command -Name 'Get-RetentionCompliancePolicy' -ErrorAction SilentlyContinue) -or
            [bool](Get-Command -Name 'Get-DlpCompliancePolicy' -ErrorAction SilentlyContinue)
        return $purviewReady
    }

    return $true
}
