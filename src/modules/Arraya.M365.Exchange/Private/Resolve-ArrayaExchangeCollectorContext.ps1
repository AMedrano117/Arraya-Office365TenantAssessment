function Get-ArrayaExchangeCollectionDepthPolicy {
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

    return [PSCustomObject]@{
        ReportingMode                   = $ReportingMode
        CollectUnifiedGroupMailboxStats = (-not $isMinimum)
        IsMinimum                       = $isMinimum
        IsCombined                      = $isCombined
        IsAll                           = $isAll
        IsGeek                          = $isGeek
    }
}

function Resolve-ArrayaExchangeCollectorContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Context,
        [Parameter(Mandatory = $false)]
        [string]$DetailLevel = 'minimum'
    )

    $resolvedMode = (Get-Culture).TextInfo.ToTitleCase($DetailLevel.ToLowerInvariant())
    if ($null -eq $Context) {
        $Context = New-ArrayaAssessmentContext `
            -ExportFileLocation $null `
            -ReportingMode $resolvedMode `
            -TenantStats ([ordered]@{}) `
            -Policies ([ordered]@{}) `
            -Runtime ([ordered]@{}) `
            -Metadata ([ordered]@{ StartedAt = (Get-Date) })
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

    if (-not $Context.Policies.Contains('ReportingMode')) {
        $Context.Policies['ReportingMode'] = $resolvedMode
    }
    if (-not $Context.Policies.Contains('CollectionDepth')) {
        $Context.Policies['CollectionDepth'] = Get-ArrayaExchangeCollectionDepthPolicy -ReportingMode ([string]$Context.Policies['ReportingMode'])
    }
    if (-not $Context.Metadata.Contains('StartedAt')) {
        $Context.Metadata['StartedAt'] = Get-Date
    }

    return $Context
}

function Get-ArrayaExchangeCollectorStartTime {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Context
    )

    if ($Context.Metadata -is [System.Collections.IDictionary] -and $Context.Metadata.Contains('StartedAt')) {
        return [datetime]$Context.Metadata['StartedAt']
    }

    return (Get-Date)
}
