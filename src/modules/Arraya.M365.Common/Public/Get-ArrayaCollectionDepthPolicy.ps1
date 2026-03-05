function Get-ArrayaCollectionDepthPolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('Minimum', 'Combined', 'All', 'Geek')]
        [string]$ReportingMode = 'Minimum'
    )

    $mode = $ReportingMode.ToLowerInvariant()
    $isMinimum = $mode -eq 'minimum'
    $isCombined = $mode -eq 'combined'
    $isAll = $mode -eq 'all'
    $isGeek = $mode -eq 'geek'

    [PSCustomObject]@{
        ReportingMode                   = $ReportingMode
        CollectEntraGroupDeepDetails    = (-not $isMinimum)
        CollectEntraGroupMemberCounts   = (-not $isMinimum)
        CollectEntraGroupOwnerCounts    = (-not $isMinimum)
        CollectEntraGroupLicenseChecks  = (-not $isMinimum)
        CollectSsoApplicationDetails    = (-not $isMinimum)
        CollectAuthenticationMethodRows = $true
        CollectSecureScoreMappings      = $true
        CollectFullMailboxNormalization = $true
        CollectFullSharePointDetail     = ($isAll -or $isGeek)
        CollectExtendedGraphEnrichment  = ($isAll -or $isGeek)
        IsMinimum                       = $isMinimum
        IsCombined                      = $isCombined
        IsAll                           = $isAll
        IsGeek                          = $isGeek
    }
}
