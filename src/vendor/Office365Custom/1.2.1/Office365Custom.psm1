$moduleRoot = Split-Path -Parent $PSCommandPath

# Import private functions
Get-ChildItem -Path (Join-Path $moduleRoot 'Private') -Filter '*.ps1' -ErrorAction SilentlyContinue |
    Sort-Object -Property Name | ForEach-Object {
        . $_.FullName
    }

# Import public functions
$publicScripts = Get-ChildItem -Path (Join-Path $moduleRoot 'Public') -Filter '*.ps1' -ErrorAction SilentlyContinue |
    Sort-Object -Property Name
foreach ($script in $publicScripts) {
    . $script.FullName
}

Export-ModuleMember -Function ($publicScripts | Select-Object -ExpandProperty BaseName)
