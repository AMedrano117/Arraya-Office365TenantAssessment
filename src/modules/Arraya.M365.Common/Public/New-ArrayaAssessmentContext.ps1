function New-ArrayaAssessmentContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ExportFileLocation,
        [Parameter(Mandatory = $false)]
        [string]$ReportingMode = 'Minimum',
        [Parameter(Mandatory = $false)]
        [hashtable]$TenantStats,
        [Parameter(Mandatory = $false)]
        [hashtable]$Policies,
        [Parameter(Mandatory = $false)]
        [hashtable]$Runtime,
        [Parameter(Mandatory = $false)]
        [hashtable]$Metadata
    )

    if (-not $TenantStats) {
        $TenantStats = [ordered]@{}
    }
    if (-not $Policies) {
        $Policies = [ordered]@{}
    }
    if (-not $Runtime) {
        $Runtime = [ordered]@{}
    }
    if (-not $Metadata) {
        $Metadata = [ordered]@{}
    }

    if (-not $Policies.Contains('ReportingMode')) {
        $Policies['ReportingMode'] = $ReportingMode
    }
    if (-not $Metadata.Contains('StartedAt')) {
        $Metadata['StartedAt'] = Get-Date
    }

    return [PSCustomObject]@{
        TenantStats        = $TenantStats
        ExportFileLocation = $ExportFileLocation
        Policies           = $Policies
        Runtime            = $Runtime
        Metadata           = $Metadata
    }
}
