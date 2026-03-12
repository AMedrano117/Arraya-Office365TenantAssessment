function Update-ArrayaTenantSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Snapshot,
        [Parameter(Mandatory = $true)]
        [ValidateSet('Metadata', 'CollectionPlan', 'Data', 'Derived', 'Diagnostics')]
        [string]$Section,
        [Parameter(Mandatory = $true)]
        [hashtable]$Values,
        [Parameter(Mandatory = $false)]
        [switch]$Replace
    )

    if (-not $Snapshot.Contains($Section) -or -not ($Snapshot[$Section] -is [System.Collections.IDictionary])) {
        $Snapshot[$Section] = [ordered]@{}
    }

    if ($Replace) {
        $newSection = [ordered]@{}
        foreach ($key in $Values.Keys) {
            $newSection[[string]$key] = $Values[$key]
        }
        $Snapshot[$Section] = $newSection
        return $Snapshot
    }

    foreach ($key in $Values.Keys) {
        $Snapshot[$Section][[string]$key] = $Values[$key]
    }

    return $Snapshot
}
