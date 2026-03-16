function Invoke-QuietCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $savedProgressPreference = $ProgressPreference
    try {
        $ProgressPreference = 'SilentlyContinue'
        & $ScriptBlock
    }
    finally {
        $ProgressPreference = $savedProgressPreference
    }
}
