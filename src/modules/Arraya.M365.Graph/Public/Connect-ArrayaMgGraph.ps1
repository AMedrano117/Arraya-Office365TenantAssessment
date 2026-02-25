function Connect-ArrayaMgGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Scopes,
        [ValidateSet('Interactive','AppCertificate')][string]$AuthMode='Interactive',
        [string]$TenantId,[string]$ClientId,[string]$CertificateThumbprint
    )
    if ($AuthMode -eq 'Interactive') {
        Connect-MgGraph -Scopes $Scopes -NoWelcome | Out-Null
    } else {
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome | Out-Null
    }
    Get-MgContext
}
