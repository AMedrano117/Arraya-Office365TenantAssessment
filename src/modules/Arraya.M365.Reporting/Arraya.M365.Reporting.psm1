$commonModuleManifestPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\Arraya.M365.Common\Arraya.M365.Common.psd1'))
if (Test-Path -Path $commonModuleManifestPath) {
    Import-Module -Name $commonModuleManifestPath -ErrorAction Stop -WarningAction SilentlyContinue -DisableNameChecking
}

# Conditional Access state normalization is shared with the HTML report and the improvement
# plan, so it lives in Arraya.M365.Analytics rather than being reimplemented here.
$analyticsModuleManifestPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\Arraya.M365.Analytics\Arraya.M365.Analytics.psd1'))
if (Test-Path -Path $analyticsModuleManifestPath) {
    Import-Module -Name $analyticsModuleManifestPath -ErrorAction Stop -WarningAction SilentlyContinue -DisableNameChecking
}

$private = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter *.ps1 -ErrorAction SilentlyContinue)
$public = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public') -Filter *.ps1 -ErrorAction SilentlyContinue)
foreach ($file in @($private) + @($public)) {
    . $file.FullName
}

Export-ModuleMember -Function @(
    'Get-ArrayaAssessmentRecommendationText',
    'Get-ArrayaAssessmentWorksheetName',
    'Get-ArrayaEmployeeExperienceInsightsAnalysis',
    'Get-GraphUserStats'
)
