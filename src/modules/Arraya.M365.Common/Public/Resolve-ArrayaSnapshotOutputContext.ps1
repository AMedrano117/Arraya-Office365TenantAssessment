function Resolve-ArrayaSnapshotOutputContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PrimaryInputPath,
        [Parameter(Mandatory = $false)]
        [string]$OutputFolder,
        [Parameter(Mandatory = $false)]
        [string]$OutputPrefix
    )

    $resolvedPrimaryInputPath = if (Test-Path -Path $PrimaryInputPath) {
        (Resolve-Path -Path $PrimaryInputPath).Path
    }
    else {
        [System.IO.Path]::GetFullPath($PrimaryInputPath)
    }

    if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
        $defaultOutputFolder = Split-Path -Path $resolvedPrimaryInputPath -Parent
        if ([string]::Equals((Split-Path -Path $defaultOutputFolder -Leaf), 'Support', [System.StringComparison]::OrdinalIgnoreCase)) {
            $parentFolder = Split-Path -Path $defaultOutputFolder -Parent
            if (-not [string]::IsNullOrWhiteSpace($parentFolder)) {
                $defaultOutputFolder = $parentFolder
            }
        }
        $OutputFolder = $defaultOutputFolder
    }
    if (-not (Test-Path -Path $OutputFolder)) {
        $null = New-Item -ItemType Directory -Path $OutputFolder -Force
    }
    $resolvedOutputFolder = (Resolve-Path -Path $OutputFolder).Path

    if ([string]::IsNullOrWhiteSpace($OutputPrefix)) {
        $OutputPrefix = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPrimaryInputPath)
    }

    return [PSCustomObject]@{
        PrimaryInputPath = $resolvedPrimaryInputPath
        OutputFolder     = $resolvedOutputFolder
        OutputPrefix     = $OutputPrefix
    }
}
