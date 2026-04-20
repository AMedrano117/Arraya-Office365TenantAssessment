Set-StrictMode -Version Latest

$public = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public') -Filter *.ps1 -ErrorAction SilentlyContinue)
$privatePath = Join-Path $PSScriptRoot 'Private'
$private = if (Test-Path -Path $privatePath -PathType Container) {
    @(Get-ChildItem -Path $privatePath -Filter *.ps1 -ErrorAction SilentlyContinue)
}
else {
    @()
}

foreach ($file in @($private) + @($public)) {
    . $file.FullName
}

Export-ModuleMember -Function @(
    'Get-ArrayaAssessmentPipelinePhase',
    'New-ArrayaAssessmentPipelineContext',
    'Invoke-ArrayaAssessmentPipeline',
    'Resume-ArrayaAssessmentPipeline'
)
