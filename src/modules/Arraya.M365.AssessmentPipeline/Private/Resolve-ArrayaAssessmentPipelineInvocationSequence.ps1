function Resolve-ArrayaAssessmentPipelineInvocationSequence {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,
        [Parameter(Mandatory = $false)]
        [string]$ThroughPhase,
        [Parameter(Mandatory = $false)]
        [string]$Phase
    )

    $definitions = @(Get-ArrayaAssessmentPipelineDefinition | Sort-Object -Property Order)
    if (-not [string]::IsNullOrWhiteSpace($Phase)) {
        return @($definitions | Where-Object { $_.Name -eq $Phase } | Select-Object -ExpandProperty Name)
    }

    if (-not [string]::IsNullOrWhiteSpace($ThroughPhase)) {
        $target = $definitions | Where-Object { $_.Name -eq $ThroughPhase } | Select-Object -First 1
        if (-not $target) {
            throw "Unknown pipeline phase: $ThroughPhase"
        }

        return @($definitions | Where-Object { [int]$_.Order -le [int]$target.Order } | Select-Object -ExpandProperty Name)
    }

    switch ([string]$Context.Mode) {
        'PreflightOnly' { return @('Connection') }
        'CollectOnly'   { return @($definitions | Where-Object { $_.Name -ne 'Export' } | Select-Object -ExpandProperty Name) }
        'ExportOnly'    { return @('Export') }
        default         { return @($definitions | Select-Object -ExpandProperty Name) }
    }
}
