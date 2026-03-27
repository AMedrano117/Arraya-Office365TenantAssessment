function Convert-ArrayaObjectToArray {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $InputObject
    )

    if ($null -eq $InputObject) {
        return @()
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $values = @($InputObject.Values)
        $containsComplexValues = $false
        foreach ($value in $values) {
            if ($null -eq $value) {
                continue
            }

            if ($value -is [System.Collections.IDictionary]) {
                $containsComplexValues = $true
                break
            }

            if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
                $containsComplexValues = $true
                break
            }

            if (
                $value.PSObject -and
                $value.PSObject.Properties.Count -gt 0 -and
                -not ($value -is [datetime]) -and
                -not ($value -is [System.ValueType])
            ) {
                $containsComplexValues = $true
                break
            }
        }

        if ($containsComplexValues) {
            return $values
        }

        return @($InputObject)
    }

    if ($InputObject -is [string]) {
        return @($InputObject)
    }

    if ($InputObject -is [System.Collections.IEnumerable]) {
        return @($InputObject)
    }

    return @($InputObject)
}
