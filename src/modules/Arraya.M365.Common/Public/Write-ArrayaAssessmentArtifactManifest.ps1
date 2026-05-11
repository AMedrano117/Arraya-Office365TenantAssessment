function New-ArrayaAssessmentOperatorSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object[]]$Artifacts = @(),

        [Parameter(Mandatory = $false)]
        [string]$OutputProfileLabel,

        [Parameter(Mandatory = $false)]
        [string]$ReportingMode,

        [Parameter(Mandatory = $false)]
        [bool]$CollectionOnly = $false,

        [Parameter(Mandatory = $false)]
        [bool]$ExportOnly = $false
    )

    function Select-ArtifactRecordByType {
        param(
            [Parameter(Mandatory = $false)]
            [object[]]$Rows = @(),
            [Parameter(Mandatory = $true)]
            [string[]]$Types
        )

        $selected = New-Object System.Collections.Generic.List[object]
        foreach ($typeName in @($Types)) {
            foreach ($row in @($Rows)) {
                if ($null -eq $row) { continue }
                $rowType = [string]$row.Type
                if (-not [string]::Equals($rowType, $typeName, [System.StringComparison]::OrdinalIgnoreCase)) {
                    continue
                }

                $selected.Add([pscustomobject][ordered]@{
                        Type   = [string]$row.Type
                        Path   = [string]$row.Path
                        Exists = [bool]$row.Exists
                    }) | Out-Null
            }
        }

        return $selected.ToArray()
    }

    function Import-EvidenceCoverageSummary {
        param(
            [Parameter(Mandatory = $false)]
            [object[]]$Rows = @()
        )

        $coverageArtifact = @(
            $Rows |
                Where-Object {
                    $null -ne $_ -and
                    [string]::Equals([string]$_.Type, 'Solutions Engineer Evidence Coverage', [System.StringComparison]::OrdinalIgnoreCase) -and
                    -not [string]::IsNullOrWhiteSpace([string]$_.Path) -and
                    (Test-Path -Path ([string]$_.Path) -PathType Leaf)
                } |
                Select-Object -First 1
        )

        if ($coverageArtifact.Count -eq 0) {
            return $null
        }

        try {
            $coverage = Get-Content -Path ([string]$coverageArtifact[0].Path) -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 100
        }
        catch {
            return [pscustomobject][ordered]@{
                Present = $true
                Path    = [string]$coverageArtifact[0].Path
                Status  = 'Unreadable'
                Error   = $_.Exception.Message
            }
        }

        $status = if ([int]$coverage.MissingObjectiveCount -gt 0 -or [int]$coverage.MissingDatasetCount -gt 0) {
            'MissingEvidence'
        }
        elseif ([int]$coverage.PartialObjectiveCount -gt 0) {
            'PartialEvidence'
        }
        elseif ([int]$coverage.ReviewObjectiveCount -gt 0 -or [int]$coverage.UnexpectedEmptyDatasetCount -gt 0) {
            'Review'
        }
        else {
            'Covered'
        }

        return [pscustomobject][ordered]@{
            Present                     = $true
            Path                        = [string]$coverageArtifact[0].Path
            Status                      = $status
            ObjectiveCount              = [int]$coverage.ObjectiveCount
            CoveredObjectiveCount       = [int]$coverage.CoveredObjectiveCount
            ReviewObjectiveCount        = [int]$coverage.ReviewObjectiveCount
            PartialObjectiveCount       = [int]$coverage.PartialObjectiveCount
            MissingObjectiveCount       = [int]$coverage.MissingObjectiveCount
            MissingDatasetCount         = [int]$coverage.MissingDatasetCount
            UnexpectedEmptyDatasetCount = [int]$coverage.UnexpectedEmptyDatasetCount
            EmptyDatasetCount           = [int]$coverage.EmptyDatasetCount
        }
    }

    $artifactRows = @($Artifacts | Where-Object { $null -ne $_ })
    $primaryTypes = @(
        'Workbook',
        'Customer Assessment Report',
        'Roadmap Remediation Plan',
        'Engineer Action Pack'
    )
    $supportTypes = @(
        'Assessment Snapshot JSON',
        'Solutions Engineer Evidence Coverage',
        'Customer Assessment Report Markdown',
        'Improvement Plan JSON',
        'Remediation Snippets',
        'Best Practices HTML',
        'Full HTML',
        'PDF',
        'Questionnaire'
    )

    $primaryDeliverables = @(Select-ArtifactRecordByType -Rows $artifactRows -Types $primaryTypes)
    $supportArtifacts = @(Select-ArtifactRecordByType -Rows $artifactRows -Types $supportTypes)
    $evidenceCoverage = Import-EvidenceCoverageSummary -Rows $artifactRows

    $recommendedStart = @(
        $primaryDeliverables |
            Where-Object { $_.Exists } |
            ForEach-Object { [string]$_.Type }
    )

    return [pscustomobject][ordered]@{
        SchemaVersion       = 1
        OutputProfile       = $OutputProfileLabel
        ReportingMode       = $ReportingMode
        CollectionOnly      = $CollectionOnly
        ExportOnly          = $ExportOnly
        RecommendedStart    = $recommendedStart
        PrimaryDeliverables = $primaryDeliverables
        SupportArtifacts    = $supportArtifacts
        EvidenceCoverage    = $evidenceCoverage
        OperatorNote        = 'Start with PrimaryDeliverables. Use SupportArtifacts for evidence, replay, troubleshooting, and engineering handoff.'
    }
}

function Write-ArrayaAssessmentArtifactManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseExportPath,
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Artifacts,
        [Parameter(Mandatory = $true)]
        [string]$OutputProfileLabel,
        [Parameter(Mandatory = $true)]
        [string]$ReportingMode,
        [Parameter(Mandatory = $true)]
        [bool]$CollectionOnly,
        [Parameter(Mandatory = $true)]
        [bool]$ExportOnly
    )

    $resolvedBasePath = [System.IO.Path]::GetFullPath($BaseExportPath)
    $baseDirectory = Split-Path -Path $resolvedBasePath -Parent
    if ([string]::IsNullOrWhiteSpace($baseDirectory)) {
        $baseDirectory = (Get-Location).Path
    }
    $supportDirectory = Join-Path -Path $baseDirectory -ChildPath 'Support'
    if (-not (Test-Path -Path $supportDirectory)) {
        $null = New-Item -ItemType Directory -Path $supportDirectory -Force
    }

    $baseLeaf = [System.IO.Path]::GetFileNameWithoutExtension($resolvedBasePath)
    foreach ($suffix in @('-Assess', ' - Tenant Details', '-Tenant Details')) {
        if ($baseLeaf.EndsWith($suffix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $baseLeaf = $baseLeaf.Substring(0, $baseLeaf.Length - $suffix.Length)
            break
        }
    }
    $manifestFileName = if (
        [string]::IsNullOrWhiteSpace($baseLeaf) -or
        [string]::Equals($baseLeaf, 'Assess', [System.StringComparison]::OrdinalIgnoreCase) -or
        [string]::Equals($baseLeaf, 'Tenant Details', [System.StringComparison]::OrdinalIgnoreCase)
    ) {
        'Run.manifest.json'
    }
    else {
        '{0}-Run.manifest.json' -f $baseLeaf
    }
    $manifestPath = Join-Path -Path $supportDirectory -ChildPath $manifestFileName

    $artifactRecords = @()
    foreach ($entry in $Artifacts.GetEnumerator()) {
        if ([string]::IsNullOrWhiteSpace([string]$entry.Value)) {
            continue
        }

        $artifactPath = [string]$entry.Value
        $artifactType = [string]$entry.Key
        $exists = Test-Path -Path $artifactPath
        $sizeBytes = $null
        if ($exists) {
            try {
                $sizeBytes = (Get-Item -Path $artifactPath -ErrorAction Stop).Length
            }
            catch {}
        }

        $artifactRecords += [ordered]@{
            Type      = $artifactType
            Path      = $artifactPath
            Exists    = [bool]$exists
            SizeBytes = $sizeBytes
        }
    }

    $payload = [ordered]@{
        SchemaVersion  = 2
        GeneratedAt    = (Get-Date).ToString('o')
        OutputProfile  = $OutputProfileLabel
        ReportingMode  = $ReportingMode
        CollectionOnly = $CollectionOnly
        ExportOnly     = $ExportOnly
        BaseExportPath = $resolvedBasePath
        OperatorSummary = (New-ArrayaAssessmentOperatorSummary -Artifacts $artifactRecords -OutputProfileLabel $OutputProfileLabel -ReportingMode $ReportingMode -CollectionOnly $CollectionOnly -ExportOnly $ExportOnly)
        Artifacts      = $artifactRecords
    }

    $json = $payload | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($manifestPath, $json, [System.Text.UTF8Encoding]::new($false))
    return $manifestPath
}
