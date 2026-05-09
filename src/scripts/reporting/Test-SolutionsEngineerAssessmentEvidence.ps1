[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SnapshotPath,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$MatrixPath = (Join-Path -Path $PSScriptRoot -ChildPath '..\..\config\baseline\solutions-engineer-assessment-objectives.json'),

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$AsJson,

    [Parameter(Mandatory = $false)]
    [switch]$FailOnMissingEvidence,

    [Parameter(Mandatory = $false)]
    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Import-SeJsonFile {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$Label was not found: $Path"
    }

    try {
        return (Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json -Depth 100)
    }
    catch {
        throw "Unable to parse $Label as JSON: $($_.Exception.Message)"
    }
}

function Test-SeSimpleValue {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value
    )

    return (
        $null -eq $Value -or
        $Value -is [string] -or
        $Value -is [char] -or
        $Value -is [bool] -or
        $Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [int] -or
        $Value -is [int64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal] -or
        $Value -is [datetime]
    )
}

function Test-SeCollectionLikeValue {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $false
    }

    if (Test-SeSimpleValue -Value $Value) {
        return $false
    }

    if ($Value -is [System.Collections.IDictionary]) {
        return $true
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        return $true
    }

    $properties = @($Value.PSObject.Properties | Where-Object { $_.MemberType -in @('NoteProperty', 'Property', 'AliasProperty') })
    return ($properties.Count -gt 0)
}

function Get-SeDatasetRecordCount {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return 0
    }

    if ($Value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Value)) {
            return 0
        }

        return 1
    }

    if ($Value -is [System.Collections.IDictionary]) {
        return $Value.Count
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        return @($Value).Count
    }

    $properties = @($Value.PSObject.Properties | Where-Object { $_.MemberType -in @('NoteProperty', 'Property', 'AliasProperty') })
    if ($properties.Count -eq 0) {
        return 0
    }

    $collectionLikeProperties = @(
        $properties |
            Where-Object {
                $propertyValue = $_.Value
                Test-SeCollectionLikeValue -Value $propertyValue
            }
    )

    if ($properties.Count -gt 1 -and $collectionLikeProperties.Count -eq $properties.Count) {
        return $properties.Count
    }

    return 1
}

function Add-SeDatasetMatches {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [int]$Depth,

        [Parameter(Mandatory = $false)]
        [ValidateNotNull()]
        [System.Collections.Generic.List[object]]$Matches
    )

    if ($Depth -gt 60 -or $null -eq $Value) {
        return
    }

    if (Test-SeSimpleValue -Value $Value) {
        return
    }

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in @($Value.Keys)) {
            $keyText = [string]$key
            $childValue = $Value[$key]
            $childPath = "$Path.$keyText"

            if ($keyText.Equals($Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                $Matches.Add([pscustomobject]@{
                        Path        = $childPath
                        Value       = $childValue
                        RecordCount = Get-SeDatasetRecordCount -Value $childValue
                    }) | Out-Null
            }

            Add-SeDatasetMatches -Value $childValue -Name $Name -Path $childPath -Depth ($Depth + 1) -Matches $Matches
        }

        return
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $index = 0
        foreach ($item in @($Value)) {
            Add-SeDatasetMatches -Value $item -Name $Name -Path ("{0}[{1}]" -f $Path, $index) -Depth ($Depth + 1) -Matches $Matches
            $index++
        }

        return
    }

    $properties = @($Value.PSObject.Properties | Where-Object { $_.MemberType -in @('NoteProperty', 'Property', 'AliasProperty') })
    foreach ($property in $properties) {
        $propertyPath = "$Path.$($property.Name)"
        if ($property.Name.Equals($Name, [System.StringComparison]::OrdinalIgnoreCase)) {
            $Matches.Add([pscustomobject]@{
                    Path        = $propertyPath
                    Value       = $property.Value
                    RecordCount = Get-SeDatasetRecordCount -Value $property.Value
                }) | Out-Null
        }

        Add-SeDatasetMatches -Value $property.Value -Name $Name -Path $propertyPath -Depth ($Depth + 1) -Matches $Matches
    }
}

function Get-SeDatasetEvidence {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Snapshot,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DatasetName
    )

    $matches = [System.Collections.Generic.List[object]]::new()
    Add-SeDatasetMatches -Value $Snapshot -Name $DatasetName -Path '$' -Depth 0 -Matches $matches

    $bestMatch = @($matches | Sort-Object -Property RecordCount -Descending | Select-Object -First 1)
    if ($bestMatch.Count -eq 0) {
        return [pscustomobject]@{
            Name        = $DatasetName
            Present     = $false
            Populated   = $false
            RecordCount = 0
            Path        = $null
        }
    }

    return [pscustomobject]@{
        Name        = $DatasetName
        Present     = $true
        Populated   = ([int]$bestMatch[0].RecordCount -gt 0)
        RecordCount = [int]$bestMatch[0].RecordCount
        Path        = [string]$bestMatch[0].Path
    }
}

function Add-SeDatasetEvidenceRecord {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [object]$Index
    )

    $recordCount = Get-SeDatasetRecordCount -Value $Value
    if ($Index.ContainsKey($Name) -and [int]$Index[$Name].RecordCount -ge $recordCount) {
        return
    }

    $Index[$Name] = [pscustomobject]@{
        Name        = $Name
        Present     = $true
        Populated   = ($recordCount -gt 0)
        RecordCount = $recordCount
        Path        = $Path
    }
}

function Add-SeDatasetIndexMatches {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [int]$Depth,

        [Parameter(Mandatory = $true)]
        [object]$DatasetNames,

        [Parameter(Mandatory = $true)]
        [object]$Index
    )

    if ($Depth -gt 60 -or $null -eq $Value) {
        return
    }

    if (Test-SeSimpleValue -Value $Value) {
        return
    }

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in @($Value.Keys)) {
            $keyText = [string]$key
            $childValue = $Value[$key]
            $childPath = "$Path.$keyText"

            if ($DatasetNames.Contains($keyText)) {
                Add-SeDatasetEvidenceRecord -Name $keyText -Path $childPath -Value $childValue -Index $Index
                continue
            }

            Add-SeDatasetIndexMatches -Value $childValue -Path $childPath -Depth ($Depth + 1) -DatasetNames $DatasetNames -Index $Index
        }

        return
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $itemIndex = 0
        foreach ($item in @($Value)) {
            Add-SeDatasetIndexMatches -Value $item -Path ("{0}[{1}]" -f $Path, $itemIndex) -Depth ($Depth + 1) -DatasetNames $DatasetNames -Index $Index
            $itemIndex++
        }

        return
    }

    $properties = @($Value.PSObject.Properties | Where-Object { $_.MemberType -in @('NoteProperty', 'Property', 'AliasProperty') })
    foreach ($property in $properties) {
        $propertyPath = "$Path.$($property.Name)"
        if ($DatasetNames.Contains($property.Name)) {
            Add-SeDatasetEvidenceRecord -Name $property.Name -Path $propertyPath -Value $property.Value -Index $Index
            continue
        }

        Add-SeDatasetIndexMatches -Value $property.Value -Path $propertyPath -Depth ($Depth + 1) -DatasetNames $DatasetNames -Index $Index
    }
}

function New-SeDatasetEvidenceIndex {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Snapshot,

        [Parameter(Mandatory = $true)]
        [object]$DatasetNames
    )

    $index = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
    Add-SeDatasetIndexMatches -Value $Snapshot -Path '$' -Depth 0 -DatasetNames $DatasetNames -Index $index
    return $index
}

function Get-SeDatasetEvidenceFromIndex {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Index,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DatasetName
    )

    if (-not $Index.ContainsKey($DatasetName)) {
        return [pscustomobject]@{
            Name        = $DatasetName
            Present     = $false
            Populated   = $false
            RecordCount = 0
            Path        = $null
        }
    }

    $record = $Index[$DatasetName]
    return [pscustomobject]@{
        Name        = $DatasetName
        Present     = $true
        Populated   = [bool]$record.Populated
        RecordCount = [int]$record.RecordCount
        Path        = [string]$record.Path
    }
}

function Get-SeObjectiveStatus {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$DatasetEvidence,

        [Parameter(Mandatory = $false)]
        [string[]]$ExpectedEmptyDatasets = @()
    )

    $expectedEmptyLookup = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($datasetName in @($ExpectedEmptyDatasets)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$datasetName)) {
            $null = $expectedEmptyLookup.Add([string]$datasetName)
        }
    }

    $presentCount = @($DatasetEvidence | Where-Object { $_.Present }).Count
    $missingCount = @($DatasetEvidence | Where-Object { -not $_.Present }).Count
    $usableEvidenceCount = @(
        $DatasetEvidence |
            Where-Object {
                $_.Populated -or ($_.Present -and $expectedEmptyLookup.Contains([string]$_.Name))
            }
    ).Count

    if ($presentCount -eq 0) {
        return 'Missing'
    }

    if ($missingCount -gt 0) {
        return 'Partial'
    }

    if ($usableEvidenceCount -eq 0) {
        return 'PresentEmpty'
    }

    return 'Covered'
}

function Get-SeObjectiveConfidence {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Status,

        [Parameter(Mandatory = $true)]
        [object[]]$DatasetEvidence,

        [Parameter(Mandatory = $false)]
        [string[]]$ExpectedEmptyDatasets = @()
    )

    $expectedEmptyLookup = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($datasetName in @($ExpectedEmptyDatasets)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$datasetName)) {
            $null = $expectedEmptyLookup.Add([string]$datasetName)
        }
    }

    switch ($Status) {
        'Covered' {
            $emptyCount = @(
                $DatasetEvidence |
                    Where-Object {
                        $_.Present -and
                        -not $_.Populated -and
                        -not $expectedEmptyLookup.Contains([string]$_.Name)
                    }
            ).Count
            if ($emptyCount -gt 0) {
                return 'Review'
            }

            return 'Strong'
        }
        'PresentEmpty' {
            $unexpectedEmptyCount = @(
                $DatasetEvidence |
                    Where-Object {
                        $_.Present -and
                        -not $_.Populated -and
                        -not $expectedEmptyLookup.Contains([string]$_.Name)
                    }
            ).Count
            if ($unexpectedEmptyCount -eq 0) {
                return 'Strong'
            }

            return 'Review'
        }
        'Partial' { return 'Partial' }
        default { return 'Missing' }
    }
}

function New-SeObjectiveCoverage {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Objective,

        [Parameter(Mandatory = $true)]
        [object]$DatasetEvidenceIndex
    )

    $datasetEvidence = @(
        foreach ($dataset in @($Objective.ExpectedDatasets)) {
            Get-SeDatasetEvidenceFromIndex -Index $DatasetEvidenceIndex -DatasetName ([string]$dataset)
        }
    )

    $expectedEmptyDatasets = @()
    if ($Objective.PSObject.Properties['ExpectedEmptyDatasets']) {
        $expectedEmptyDatasets = @($Objective.ExpectedEmptyDatasets)
    }
    $status = Get-SeObjectiveStatus -DatasetEvidence $datasetEvidence -ExpectedEmptyDatasets $expectedEmptyDatasets
    $confidence = Get-SeObjectiveConfidence -Status $status -DatasetEvidence $datasetEvidence -ExpectedEmptyDatasets $expectedEmptyDatasets
    $missingDatasets = @($datasetEvidence | Where-Object { -not $_.Present } | ForEach-Object { $_.Name })
    $emptyDatasets = @($datasetEvidence | Where-Object { $_.Present -and -not $_.Populated } | ForEach-Object { $_.Name })
    $expectedEmptyLookup = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($datasetName in @($expectedEmptyDatasets)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$datasetName)) {
            $null = $expectedEmptyLookup.Add([string]$datasetName)
        }
    }
    $expectedEmptyDatasetsFound = @($datasetEvidence | Where-Object { $_.Present -and -not $_.Populated -and $expectedEmptyLookup.Contains([string]$_.Name) } | ForEach-Object { $_.Name })
    $unexpectedEmptyDatasets = @($datasetEvidence | Where-Object { $_.Present -and -not $_.Populated -and -not $expectedEmptyLookup.Contains([string]$_.Name) } | ForEach-Object { $_.Name })
    $presentDatasets = @($datasetEvidence | Where-Object { $_.Present } | ForEach-Object { $_.Name })

    return [pscustomobject][ordered]@{
        ObjectiveId          = [string]$Objective.ObjectiveId
        Area                 = [string]$Objective.Area
        Goal                 = [string]$Objective.Goal
        Status               = $status
        Confidence           = $confidence
        ExpectedDatasetCount = @($Objective.ExpectedDatasets).Count
        PresentDatasetCount  = @($datasetEvidence | Where-Object { $_.Present }).Count
        PopulatedDatasetCount = @($datasetEvidence | Where-Object { $_.Populated }).Count
        MissingDatasets      = $missingDatasets
        EmptyDatasets        = $emptyDatasets
        ExpectedEmptyDatasets = $expectedEmptyDatasetsFound
        UnexpectedEmptyDatasets = $unexpectedEmptyDatasets
        PresentDatasets      = $presentDatasets
        DatasetEvidence      = $datasetEvidence
        RelatedFindings      = @($Objective.RelatedFindings)
        ChecksFor            = @($Objective.ChecksFor)
        ConfidenceNotes      = [string]$Objective.ConfidenceNotes
    }
}

function New-SeCoverageReport {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Matrix,

        [Parameter(Mandatory = $true)]
        [object]$Snapshot,

        [Parameter(Mandatory = $true)]
        [string]$ResolvedSnapshotPath,

        [Parameter(Mandatory = $true)]
        [string]$ResolvedMatrixPath
    )

    $allDatasetNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($objective in @($Matrix.Objectives)) {
        foreach ($dataset in @($objective.ExpectedDatasets)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$dataset)) {
                $null = $allDatasetNames.Add([string]$dataset)
            }
        }
    }

    $datasetEvidenceIndex = New-SeDatasetEvidenceIndex -Snapshot $Snapshot -DatasetNames $allDatasetNames

    $objectiveCoverage = @(
        foreach ($objective in @($Matrix.Objectives)) {
            New-SeObjectiveCoverage -Objective $objective -DatasetEvidenceIndex $datasetEvidenceIndex
        }
    )

    $missingDatasets = @(
        $objectiveCoverage |
            ForEach-Object { $_.MissingDatasets } |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            Sort-Object -Unique
    )
    $emptyDatasets = @(
        $objectiveCoverage |
            ForEach-Object { $_.EmptyDatasets } |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            Sort-Object -Unique
    )
    $unexpectedEmptyDatasets = @(
        $objectiveCoverage |
            ForEach-Object { $_.UnexpectedEmptyDatasets } |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            Sort-Object -Unique
    )

    $statusSummary = [ordered]@{}
    foreach ($statusName in @('Covered', 'Review', 'Partial', 'PresentEmpty', 'Missing')) {
        $statusSummary[$statusName] = 0
    }

    foreach ($objectiveResult in $objectiveCoverage) {
        if ($objectiveResult.Status -eq 'Covered' -and $objectiveResult.Confidence -eq 'Review') {
            $statusSummary['Review']++
            continue
        }

        if (-not $statusSummary.Contains($objectiveResult.Status)) {
            $statusSummary[$objectiveResult.Status] = 0
        }

        $statusSummary[$objectiveResult.Status]++
    }

    return [pscustomobject][ordered]@{
        SchemaVersion                  = 1
        GeneratedAtUtc                 = (Get-Date).ToUniversalTime().ToString('o')
        OutputProfile                  = [string]$Matrix.OutputProfile
        ReportingMode                  = [string]$Matrix.ReportingMode
        SnapshotPath                   = $ResolvedSnapshotPath
        MatrixPath                     = $ResolvedMatrixPath
        ObjectiveCount                 = $objectiveCoverage.Count
        CoveredObjectiveCount          = @($objectiveCoverage | Where-Object { $_.Status -eq 'Covered' -and $_.Confidence -ne 'Review' }).Count
        ReviewObjectiveCount           = @($objectiveCoverage | Where-Object { $_.Status -eq 'Covered' -and $_.Confidence -eq 'Review' }).Count
        PartialObjectiveCount          = @($objectiveCoverage | Where-Object { $_.Status -eq 'Partial' }).Count
        PresentEmptyObjectiveCount     = @($objectiveCoverage | Where-Object { $_.Status -eq 'PresentEmpty' }).Count
        MissingObjectiveCount          = @($objectiveCoverage | Where-Object { $_.Status -eq 'Missing' }).Count
        MissingOrPartialObjectiveCount = @($objectiveCoverage | Where-Object { $_.Status -in @('Missing', 'Partial') }).Count
        MissingDatasetCount            = $missingDatasets.Count
        EmptyDatasetCount              = $emptyDatasets.Count
        UnexpectedEmptyDatasetCount    = $unexpectedEmptyDatasets.Count
        MissingDatasets                = $missingDatasets
        EmptyDatasets                  = $emptyDatasets
        UnexpectedEmptyDatasets        = $unexpectedEmptyDatasets
        StatusSummary                  = [pscustomobject]$statusSummary
        Objectives                     = $objectiveCoverage
        CoverageGaps                   = @($Matrix.CoverageGaps)
    }
}

$resolvedSnapshotPath = (Resolve-Path -LiteralPath $SnapshotPath).Path
$resolvedMatrixPath = (Resolve-Path -LiteralPath $MatrixPath).Path
$matrix = Import-SeJsonFile -Path $resolvedMatrixPath -Label 'Solutions Engineer objective matrix'
$snapshot = Import-SeJsonFile -Path $resolvedSnapshotPath -Label 'assessment snapshot'

if ([string]$matrix.OutputProfile -ne 'SolutionsEngineer') {
    throw "The matrix at $resolvedMatrixPath is for '$($matrix.OutputProfile)', not SolutionsEngineer."
}

$coverageReport = New-SeCoverageReport -Matrix $matrix -Snapshot $snapshot -ResolvedSnapshotPath $resolvedSnapshotPath -ResolvedMatrixPath $resolvedMatrixPath

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $resolvedOutputPath = if ([System.IO.Path]::IsPathRooted($OutputPath)) {
        [System.IO.Path]::GetFullPath($OutputPath)
    }
    else {
        [System.IO.Path]::GetFullPath((Join-Path -Path (Get-Location).Path -ChildPath $OutputPath))
    }
    $resolvedOutputDirectory = Split-Path -Path $resolvedOutputPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($resolvedOutputDirectory) -and -not (Test-Path -LiteralPath $resolvedOutputDirectory)) {
        New-Item -ItemType Directory -Path $resolvedOutputDirectory -Force | Out-Null
    }

    $coverageReport | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $resolvedOutputPath -Encoding UTF8
}

if ($FailOnMissingEvidence -and $coverageReport.MissingOrPartialObjectiveCount -gt 0) {
    throw ("Solutions Engineer evidence coverage found {0} missing or partial objective(s). Missing datasets: {1}" -f $coverageReport.MissingOrPartialObjectiveCount, (($coverageReport.MissingDatasets | Select-Object -First 20) -join ', '))
}

if ($PassThru) {
    return $coverageReport
}

if ($AsJson) {
    $coverageReport | ConvertTo-Json -Depth 100
    return
}

$coverageReport.Objectives |
    Select-Object ObjectiveId, Area, Status, Confidence, ExpectedDatasetCount, PresentDatasetCount, PopulatedDatasetCount |
    Format-Table -AutoSize

Write-Output ("Summary: Covered={0}; Review={1}; Partial={2}; PresentEmpty={3}; Missing={4}; MissingDatasets={5}; EmptyDatasets={6}" -f $coverageReport.CoveredObjectiveCount, $coverageReport.ReviewObjectiveCount, $coverageReport.PartialObjectiveCount, $coverageReport.PresentEmptyObjectiveCount, $coverageReport.MissingObjectiveCount, $coverageReport.MissingDatasetCount, $coverageReport.EmptyDatasetCount)
