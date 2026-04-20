function Invoke-AssessmentPipelineScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,
        [Parameter(Mandatory = $false)]
        [hashtable]$Parameters
    )

    & {
        Set-StrictMode -Off

        if ($Parameters) {
            & $ScriptPath @Parameters
        }
        else {
            & $ScriptPath
        }
    }
}
