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

    function Get-NormalizedArrayaOutputPrefix {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Candidate
        )

        $normalized = $Candidate
        foreach ($suffix in @('-AssessmentSnapshot', '-Snapshot')) {
            if ($normalized.EndsWith($suffix, [System.StringComparison]::Ordinal)) {
                $normalized = $normalized.Substring(0, $normalized.Length - $suffix.Length)
                break
            }
        }

        $normalized = $normalized.Replace(' Tenant Discovery Report-', '-')

        $replacements = [ordered]@{
            '-TenantToTenantMigration_' = '-T2T_'
            '-SolutionsEngineer_'       = '-SE_'
            '-ExecutiveLevel_'          = '-Exec_'
            '-TenantToTenantMigration'  = '-T2T'
            '-SolutionsEngineer'        = '-SE'
            '-ExecutiveLevel'           = '-Exec'
        }

        foreach ($entry in $replacements.GetEnumerator()) {
            $normalized = $normalized.Replace([string]$entry.Key, [string]$entry.Value)
        }

        return $normalized
    }

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
    $OutputPrefix = Get-NormalizedArrayaOutputPrefix -Candidate $OutputPrefix

    return [PSCustomObject]@{
        PrimaryInputPath = $resolvedPrimaryInputPath
        OutputFolder     = $resolvedOutputFolder
        OutputPrefix     = $OutputPrefix
    }
}
