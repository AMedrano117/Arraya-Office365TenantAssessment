function Resolve-ArrayaAssessmentPipelineCheckpointRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$CheckpointRoot,
        [Parameter(Mandatory = $false)]
        [string]$ExportPath
    )

    if (-not [string]::IsNullOrWhiteSpace($CheckpointRoot)) {
        return [System.IO.Path]::GetFullPath($CheckpointRoot)
    }

    $resolvedExportPath = if ([string]::IsNullOrWhiteSpace($ExportPath)) {
        Get-ArrayaAssessmentOutputRoot -FallbackPath $script:RepoRoot
    }
    else {
        $ExportPath
    }

    $fullExportPath = [System.IO.Path]::GetFullPath($resolvedExportPath)
    $runRoot = if ([System.IO.Path]::HasExtension($fullExportPath)) {
        Split-Path -Path $fullExportPath -Parent
    }
    else {
        $fullExportPath
    }

    return (Join-Path -Path $runRoot -ChildPath 'Support\Pipeline')
}
