function New-ArrayaTenantSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [hashtable]$Metadata,
        [Parameter(Mandatory = $false)]
        [hashtable]$CollectionPlan,
        [Parameter(Mandatory = $false)]
        [hashtable]$Data,
        [Parameter(Mandatory = $false)]
        [hashtable]$Derived,
        [Parameter(Mandatory = $false)]
        [hashtable]$Diagnostics
    )

    $snapshot = [ordered]@{
        SchemaVersion  = 2
        Metadata       = [ordered]@{
            GeneratedAt = (Get-Date).ToString('o')
        }
        CollectionPlan = [ordered]@{}
        Data           = [ordered]@{
            Exchange      = [ordered]@{}
            Identity      = [ordered]@{}
            Collaboration = [ordered]@{}
            Security      = [ordered]@{}
            Governance    = [ordered]@{}
            Tenant        = [ordered]@{}
            Other         = [ordered]@{}
        }
        Derived        = [ordered]@{}
        Diagnostics    = [ordered]@{
            WarningCount   = 0
            ErrorCount     = 0
            WarningSummary = @()
            ErrorSummary   = @()
            SourceCoverage = [ordered]@{}
            CollectorStats = [ordered]@{}
        }
    }

    if ($Metadata) {
        foreach ($key in $Metadata.Keys) {
            $snapshot.Metadata[[string]$key] = $Metadata[$key]
        }
    }
    if ($CollectionPlan) {
        foreach ($key in $CollectionPlan.Keys) {
            $snapshot.CollectionPlan[[string]$key] = $CollectionPlan[$key]
        }
    }
    if ($Data) {
        foreach ($key in $Data.Keys) {
            $snapshot.Data[[string]$key] = $Data[$key]
        }
    }
    if ($Derived) {
        foreach ($key in $Derived.Keys) {
            $snapshot.Derived[[string]$key] = $Derived[$key]
        }
    }
    if ($Diagnostics) {
        foreach ($key in $Diagnostics.Keys) {
            $snapshot.Diagnostics[[string]$key] = $Diagnostics[$key]
        }
    }

    return $snapshot
}
