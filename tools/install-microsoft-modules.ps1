[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDevTools,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

$modules = @(
    'Microsoft.Graph.Authentication',
    'ExchangeOnlineManagement',
    'MicrosoftTeams',
    'Microsoft.Online.SharePoint.PowerShell',
    'ImportExcel'
)

if ($IncludeDevTools) {
    $modules += @(
        'Pester',
        'PSScriptAnalyzer'
    )
}

$gallery = Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue
if ($gallery -and $gallery.InstallationPolicy -ne 'Trusted') {
    Write-Host 'Setting PSGallery to Trusted for this install session...' -ForegroundColor Cyan
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
}

foreach ($moduleName in $modules) {
    $installedModule = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
    if ($installedModule -and -not $Force) {
        Write-Host ("Already installed: {0} {1}" -f $installedModule.Name, $installedModule.Version) -ForegroundColor Green
        continue
    }

    Write-Host ("Installing: {0}" -f $moduleName) -ForegroundColor Cyan
    $installParams = @{
        Name         = $moduleName
        Scope        = 'CurrentUser'
        Repository   = 'PSGallery'
        AllowClobber = $true
        ErrorAction  = 'Stop'
    }

    if ($Force) {
        $installParams.Force = $true
    }

    Install-Module @installParams
    $newestInstalled = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
    Write-Host ("Installed: {0} {1}" -f $newestInstalled.Name, $newestInstalled.Version) -ForegroundColor Green
}

Write-Host 'Microsoft module installation complete.' -ForegroundColor Green
