function Resolve-ArrayaRequestedOutputProfiles {
    <#
    .SYNOPSIS
        Returns the output profiles a run was asked for, in a fixed canonical order.

    .DESCRIPTION
        A merged selection reaches the collector as a label such as
        'Merged(Presales+TenantToTenantMigration)', optionally prefixed by the action
        ('Collection-', 'Export-'). This function recovers the individual profiles from the
        label, falls back to -OutputProfile, removes duplicates and unknown names, and orders
        the result canonically so the same request always yields the same list.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Returns the set of requested profiles; a singular noun would misdescribe the result.')]
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$OutputProfile,
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$OutputProfileLabel
    )

    $canonicalOrder = @('Presales', 'TenantToTenantMigration', 'SolutionsEngineer', 'ExecutiveLevel', 'Geek', 'Machine')
    $requested = New-Object System.Collections.Generic.List[string]

    $mergedMatch = [regex]::Match([string]$OutputProfileLabel, 'Merged\((?<Profiles>[^)]*)\)')
    if ($mergedMatch.Success) {
        foreach ($name in ($mergedMatch.Groups['Profiles'].Value -split '[+,]')) {
            $trimmed = $name.Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmed)) { $requested.Add($trimmed) | Out-Null }
        }
    }
    if ($requested.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($OutputProfile)) {
        foreach ($name in ($OutputProfile -split '[+,]')) {
            $trimmed = $name.Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmed)) { $requested.Add($trimmed) | Out-Null }
        }
    }

    return [string[]]@(
        $canonicalOrder | Where-Object {
            $candidate = $_
            @($requested | Where-Object { [string]::Equals($_, $candidate, [System.StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
        }
    )
}
