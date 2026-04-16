$commonModuleManifestPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\Arraya.M365.Common\Arraya.M365.Common.psd1'))
if (Test-Path -Path $commonModuleManifestPath) {
    Import-Module -Name $commonModuleManifestPath -ErrorAction Stop -WarningAction SilentlyContinue -DisableNameChecking
}

function Write-ArrayaExchangeCollectorBanner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$ExportFileLocation
    )

    if (-not [string]::IsNullOrWhiteSpace($ExportFileLocation) -and (Get-Command -Name Write-Log -ErrorAction SilentlyContinue)) {
        Write-Log -Type INFO -Message $Message -ExportFileLocation $ExportFileLocation
    }
}

function Write-ArrayaExchangeCollectorSubstep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$ForegroundColor = 'DarkCyan'
    )

    Write-Host ("    > {0}" -f $Message) -ForegroundColor $ForegroundColor
}

function Write-ArrayaExchangeCollectorCompletionBanner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$ExportFileLocation
    )

    if (-not [string]::IsNullOrWhiteSpace($ExportFileLocation) -and (Get-Command -Name Write-Log -ErrorAction SilentlyContinue)) {
        Write-Log -Type INFO -Message $Message -ExportFileLocation $ExportFileLocation
    }
}

$private = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter *.ps1 -ErrorAction SilentlyContinue)
$public = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public') -Filter *.ps1 -ErrorAction SilentlyContinue)
foreach ($file in @($private) + @($public)) {
    . $file.FullName
}

Export-ModuleMember -Function @(
    'Get-AllExchangeMailboxDetails',
    'Get-AllPublicFolderDetails',
    'Get-AllRecipientDetails',
    'Get-ExchangeGroupDetails',
    'Get-ExchangeHybridConfiguration',
    'Get-MailFlowRulesandConnectors',
    'Get-SMTPRelayConfiguration',
    'Get-ThirdPartySpamFilteringConfig'
)
