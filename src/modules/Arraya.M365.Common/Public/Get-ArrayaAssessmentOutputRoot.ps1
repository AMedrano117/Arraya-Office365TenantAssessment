function Get-ArrayaAssessmentOutputRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$FallbackPath
    )

    $localAppData = [Environment]::GetFolderPath('LocalApplicationData')
    if (-not [string]::IsNullOrWhiteSpace($localAppData)) {
        $defaultRoot = Join-Path -Path $localAppData -ChildPath 'Arraya\M365TenantAssessment\Outputs'
        return [System.IO.Path]::GetFullPath($defaultRoot)
    }

    if ([string]::IsNullOrWhiteSpace($FallbackPath)) {
        return (Get-Location).Path
    }

    if ([System.IO.Path]::IsPathRooted($FallbackPath)) {
        return [System.IO.Path]::GetFullPath($FallbackPath)
    }

    return [System.IO.Path]::GetFullPath((Join-Path -Path (Get-Location).Path -ChildPath $FallbackPath))
}
