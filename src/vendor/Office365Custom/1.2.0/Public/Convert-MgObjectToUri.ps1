<#
.SYNOPSIS
    Converts Microsoft Graph object parameters into a valid URI.
.DESCRIPTION
    Constructs a URI string for Microsoft Graph API requests based on the provided ObjectType, optional objectID,
    filter, and selected properties. Handles query string encoding as needed.
.PARAMETER ObjectType
    Specifies the object type (e.g. "users" or "groups").
.PARAMETER objectID
    Optional identifier (ID or UPN) for the object.
.PARAMETER Filter
    Optional filter string for narrowing the query.
.PARAMETER Properties
    Optional array of properties to include in the select query.
.PARAMETER Version
    Specifies the API version ("beta" or "v1.0"). Default is "v1.0".
.EXAMPLE
    PS C:\> Convert-MgObjectToUri -ObjectType "users" -Filter "startswith(displayName,'John')" -Properties "id","displayName"
#>
function Convert-MgObjectToUri {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("users","groups","teams")]
        [string]$ObjectType,

        [Parameter(Mandatory = $false)]
        [string]$ObjectID,

        [Parameter(Mandatory = $false)]
        [string]$Filter,

        [Parameter(Mandatory = $false)]
        [string[]]$Properties,

        [Parameter(Mandatory = $false)]
        [ValidateSet("beta","v1.0")]
        [string]$Version = "v1.0"
    )

    # Base URI
    $uri = "https://graph.microsoft.com/$Version/$ObjectType"
    if ($ObjectID) {
        $uri += "/$ObjectID"
    }

    # If a filter is present, add "?`$filter=..."
    if ($Filter) {
        $encodedFilter = [System.Uri]::EscapeDataString($Filter)
        # Triple backticks produce a single literal backtick in the final string
        $uri += "?```$filter=$encodedFilter"
    }

    # If properties are present, add "&`$select=..." or "?`$select=..."
    if ($Properties) {
        $selectProps = $Properties -join ','
        if ($Filter) {
            $uri += "&```$select=$selectProps"
        }
        else {
            $uri += "?```$select=$selectProps"
        }
    }

    return $uri
}
