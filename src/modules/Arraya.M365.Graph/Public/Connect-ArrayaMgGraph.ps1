function Connect-ArrayaMgGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Scopes,
        [ValidateSet('Interactive','Certificate')][string]$AuthMode='Interactive',
        [string]$TenantId,[string]$ClientId,[string]$CertificateThumbprint
    )
    if ($AuthMode -eq 'Interactive') {
        Connect-MgGraph -Scopes $Scopes -NoWelcome | Out-Null
    } else {
        if ([string]::IsNullOrWhiteSpace($TenantId) -or [string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
            throw 'Certificate authentication requires TenantId, ClientId, and CertificateThumbprint.'
        }
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome | Out-Null
    }
    Get-MgContext
}
