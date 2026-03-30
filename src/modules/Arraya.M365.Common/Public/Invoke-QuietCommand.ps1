function Invoke-QuietCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $savedProgressPreference = $ProgressPreference
    $savedProgressDefaults = @{}
    $progressDefaultKeys = @(
        'Get-Mg*:ProgressAction',
        'Find-Mg*:ProgressAction',
        'Invoke-MgGraphRequest:ProgressAction',
        'Update-Mg*:ProgressAction',
        'New-Mg*:ProgressAction',
        'Remove-Mg*:ProgressAction'
    )
    try {
        $ProgressPreference = 'SilentlyContinue'
        foreach ($key in $progressDefaultKeys) {
            if ($global:PSDefaultParameterValues.ContainsKey($key)) {
                $savedProgressDefaults[$key] = $global:PSDefaultParameterValues[$key]
            }
            $global:PSDefaultParameterValues[$key] = 'SilentlyContinue'
        }
        & $ScriptBlock
    }
    finally {
        foreach ($key in $progressDefaultKeys) {
            if ($savedProgressDefaults.ContainsKey($key)) {
                $global:PSDefaultParameterValues[$key] = $savedProgressDefaults[$key]
            }
            else {
                $null = $global:PSDefaultParameterValues.Remove($key)
            }
        }
        $ProgressPreference = $savedProgressPreference
        Write-Progress -Id 0 -Activity 'Graph Request' -Completed -ErrorAction SilentlyContinue
    }
}
