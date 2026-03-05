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
        return @($InputObject.Values)
    }

    if ($InputObject -is [string]) {
        return @($InputObject)
    }

    if ($InputObject -is [System.Collections.IEnumerable]) {
        return @($InputObject)
    }

    return @($InputObject)
}
