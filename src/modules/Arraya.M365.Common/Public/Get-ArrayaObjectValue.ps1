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

    $normalizedLookup = $null

    foreach ($name in $Names) {
        if ($Object -is [System.Collections.IDictionary] -and ([System.Collections.IDictionary]$Object).Contains($name)) {
            return $Object[$name]
        }

        if ($Object.PSObject -and ($Object.PSObject.Properties.Name -contains $name)) {
            return $Object.$name
        }

        $normalizedName = ([string]$name -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($normalizedName)) {
            continue
        }

        if ($null -eq $normalizedLookup) {
            $normalizedLookup = @{}

            if ($Object -is [System.Collections.IDictionary]) {
                foreach ($key in $Object.Keys) {
                    $currentName = ([string]$key -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
                    if (-not [string]::IsNullOrWhiteSpace($currentName) -and -not $normalizedLookup.ContainsKey($currentName)) {
                        $normalizedLookup[$currentName] = $Object[$key]
                    }
                }
            }

            if ($Object.PSObject) {
                foreach ($property in $Object.PSObject.Properties) {
                    $currentName = ([string]$property.Name -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
                    if (-not [string]::IsNullOrWhiteSpace($currentName) -and -not $normalizedLookup.ContainsKey($currentName)) {
                        $normalizedLookup[$currentName] = $property.Value
                    }
                }
            }
        }

        if ($normalizedLookup.ContainsKey($normalizedName)) {
            return $normalizedLookup[$normalizedName]
        }
    }

    return $null
}
