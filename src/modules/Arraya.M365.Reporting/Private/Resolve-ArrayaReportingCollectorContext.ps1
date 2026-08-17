function Resolve-ArrayaReportingCollectorContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Context
    )

    if ($null -eq $Context) {
        $Context = New-ArrayaAssessmentContext -ExportFileLocation $null -TenantStats ([ordered]@{}) -Metadata ([ordered]@{ StartedAt = (Get-Date) })
    }

    if (-not ($Context.TenantStats -is [System.Collections.IDictionary])) {
        $Context | Add-Member -NotePropertyName TenantStats -NotePropertyValue ([ordered]@{}) -Force
    }
    if (-not ($Context.Policies -is [System.Collections.IDictionary])) {
        $Context | Add-Member -NotePropertyName Policies -NotePropertyValue ([ordered]@{}) -Force
    }
    if (-not ($Context.Runtime -is [System.Collections.IDictionary])) {
        $Context | Add-Member -NotePropertyName Runtime -NotePropertyValue ([ordered]@{}) -Force
    }
    if (-not ($Context.Metadata -is [System.Collections.IDictionary])) {
        $Context | Add-Member -NotePropertyName Metadata -NotePropertyValue ([ordered]@{}) -Force
    }
    if (-not $Context.PSObject.Properties['ExportFileLocation']) {
        $Context | Add-Member -NotePropertyName ExportFileLocation -NotePropertyValue $null -Force
    }
    if (-not ([System.Collections.IDictionary]$Context.Metadata).Contains('StartedAt')) {
        $Context.Metadata['StartedAt'] = Get-Date
    }

    return $Context
}

function Get-ArrayaReportingCollectorStartTime {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Context
    )

    if (
        $Context.Metadata -is [System.Collections.IDictionary] -and
        ([System.Collections.IDictionary]$Context.Metadata).Contains('StartedAt')
    ) {
        return [datetime]$Context.Metadata['StartedAt']
    }

    return (Get-Date)
}
