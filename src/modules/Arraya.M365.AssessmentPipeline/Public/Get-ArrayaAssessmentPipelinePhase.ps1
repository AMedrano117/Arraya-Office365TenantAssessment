function Get-ArrayaAssessmentPipelinePhase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Name
    )

    Import-AssessmentPipelineDependencies
    $definitions = @(Get-ArrayaAssessmentPipelineDefinition | Sort-Object -Property Order)
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $definitions
    }

    return @($definitions | Where-Object { $_.Name -eq $Name })
}
