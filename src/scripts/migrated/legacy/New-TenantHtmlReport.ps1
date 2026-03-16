# Compatibility wrapper: canonical HTML/report helper implementations now live in:
#   src/scripts/assessments/HTML Scripts/Invoke-HTMLHelperFunctions.ps1

$htmlHelperFunctionsPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\assessments\HTML Scripts\Invoke-HTMLHelperFunctions.ps1'))
if (-not (Test-Path -Path $htmlHelperFunctionsPath)) {
    throw "Required HTML helper script not found: $htmlHelperFunctionsPath"
}

. $htmlHelperFunctionsPath
