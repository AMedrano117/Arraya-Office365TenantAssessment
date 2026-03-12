function Get-ArrayaGraphAdminReportSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers,
        [Parameter(Mandatory = $false)]
        [switch]$PreferRest
    )

    $uri = 'https://graph.microsoft.com/v1.0/admin/reportSettings'
    try {
        $response = Get-ArrayaGraphResource -Uri $uri -Activity 'Admin report settings' -PreferRest:$PreferRest -Headers $Headers
    }
    catch {
        return [PSCustomObject]@{
            Available                 = $false
            DisplayConcealedNames     = $null
            IdentifiableNamesInReports = $null
            Source                    = 'Graph /admin/reportSettings'
            ErrorMessage              = $_.Exception.Message
            RetrievedAt               = (Get-Date).ToString('o')
        }
    }

    if ($null -eq $response) {
        return [PSCustomObject]@{
            Available                 = $false
            DisplayConcealedNames     = $null
            IdentifiableNamesInReports = $null
            Source                    = 'Graph /admin/reportSettings'
            ErrorMessage              = 'No response was returned from Graph.'
            RetrievedAt               = (Get-Date).ToString('o')
        }
    }

    $displayConcealedNames = $null
    if ($response -is [System.Collections.IDictionary]) {
        if ($response.Contains('displayConcealedNames')) {
            try { $displayConcealedNames = [bool]$response['displayConcealedNames'] } catch { $displayConcealedNames = $null }
        }
    }
    elseif ($response.PSObject -and $response.PSObject.Properties['displayConcealedNames']) {
        try { $displayConcealedNames = [bool]$response.displayConcealedNames } catch { $displayConcealedNames = $null }
    }

    return [PSCustomObject]@{
        Available                 = $true
        DisplayConcealedNames     = $displayConcealedNames
        IdentifiableNamesInReports = if ($displayConcealedNames -is [bool]) { -not $displayConcealedNames } else { $null }
        Source                    = 'Graph /admin/reportSettings'
        ErrorMessage              = $null
        RetrievedAt               = (Get-Date).ToString('o')
    }
}

