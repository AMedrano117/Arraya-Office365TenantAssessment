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
    $manifestPath = $resolvedBasePath -replace '\.xlsx$', '.manifest.json'
    if ($manifestPath -eq $resolvedBasePath) {
        $manifestPath = "$resolvedBasePath.manifest.json"
    }

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
        Artifacts      = $artifactRecords
    }

    $json = $payload | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($manifestPath, $json, [System.Text.UTF8Encoding]::new($false))
    return $manifestPath
}

