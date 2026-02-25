function New-ArrayaFinding {
    [CmdletBinding()]
    param(
        [string]$TenantId,
        [Parameter(Mandatory)][string]$Workload,
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$Control,
        [string]$CurrentState,[string]$RecommendedState,
        [ValidateSet('Critical','High','Medium','Low','Info')][string]$Severity='Info',
        [ValidateSet('Compliant','NonCompliant','Informational','NotAssessed','NotLicensed','PermissionDenied','Error')][string]$Status='Informational',
        [string]$Risk,[string]$Evidence,[string]$Remediation
    )
    [pscustomobject]@{
        TenantId=$TenantId; Workload=$Workload; Category=$Category; Control=$Control; CurrentState=$CurrentState; RecommendedState=$RecommendedState;
        Severity=$Severity; Status=$Status; Risk=$Risk; Evidence=$Evidence; Remediation=$Remediation; AssessedAt=(Get-Date).ToString('o')
    }
}
