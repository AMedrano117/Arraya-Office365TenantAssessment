$commonModuleManifestPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\Arraya.M365.Common\Arraya.M365.Common.psd1'))
if (Test-Path -Path $commonModuleManifestPath) {
    Import-Module -Name $commonModuleManifestPath -ErrorAction Stop -WarningAction SilentlyContinue
}

$public = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public') -Filter *.ps1 -ErrorAction SilentlyContinue)
$private = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter *.ps1 -ErrorAction SilentlyContinue)
foreach ($file in @($private) + @($public)) { . $file.FullName }
Export-ModuleMember -Function @(
    'Get-EntraIDGroups'
)
