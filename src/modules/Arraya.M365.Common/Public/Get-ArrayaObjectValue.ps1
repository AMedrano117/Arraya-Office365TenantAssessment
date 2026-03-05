function Get-ArrayaObjectValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Object,
        [Parameter(Mandatory = $true)]
        [string[]]$Names
    )

    if ($null -eq $Object) {
        return $null
    }

    foreach ($name in $Names) {
        if ($Object -is [System.Collections.IDictionary] -and $Object.Contains($name)) {
            return $Object[$name]
        }

        if ($Object.PSObject -and ($Object.PSObject.Properties.Name -contains $name)) {
            return $Object.$name
        }
    }

    return $null
}
