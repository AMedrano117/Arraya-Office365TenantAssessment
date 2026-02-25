[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$OutputPath,
    [ValidateSet('Interactive','AppCertificate')][string]$AuthMode='Interactive',
    [string]$TenantId,[string]$ClientId,[string]$CertificateThumbprint,[switch]$IncludeRaw
)
$ErrorActionPreference='Stop'
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$runFolder = Join-Path $OutputPath "run-$timestamp"
$exports = Join-Path $runFolder 'exports'
$raw = Join-Path $runFolder 'raw'
New-Item -ItemType Directory -Path $exports,$raw -Force | Out-Null
if ($AuthMode -eq 'Interactive') { Connect-MgGraph -Scopes 'Organization.Read.All' -NoWelcome | Out-Null }
else { Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome | Out-Null }
$ctx = Get-MgContext
$org = Get-MgOrganization -ErrorAction Stop
$findings = @([pscustomobject]@{TenantId=$ctx.TenantId;Workload='Tenant';Category='Metadata';Control='Organization Metadata Retrieval';CurrentState='Success';RecommendedState='';Severity='Info';Status='Informational';Risk='None';Evidence='Get-MgOrganization returned organization object.';Remediation='';AssessedAt=(Get-Date).ToString('o')})
$findings | Export-Csv (Join-Path $exports 'findings.csv') -NoTypeInformation -Encoding UTF8
[pscustomobject]@{RunId=[guid]::NewGuid().Guid;ScriptName=$MyInvocation.MyCommand.Name;ScriptVersion='0.1.0';StartTime=(Get-Date).ToString('o');EndTime=(Get-Date).ToString('o');DurationSeconds=0;TenantId=$ctx.TenantId;TenantDisplayName=$org.DisplayName;VerifiedDomains=@($org.VerifiedDomains.Name);AuthMode=$AuthMode;GraphScopes=@($ctx.Scopes);PowerShellVersion=$PSVersionTable.PSVersion.ToString();Host=$env:COMPUTERNAME;ExecutedBy="$env:USERDOMAIN\$env:USERNAME";FindingCount=$findings.Count;ErrorCount=0} | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $exports 'execution-metadata.json') -Encoding UTF8
@() | ConvertTo-Json | Set-Content (Join-Path $exports 'errors.json') -Encoding UTF8
if ($IncludeRaw) { $org | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $raw 'raw-data.json') -Encoding UTF8 }
Disconnect-MgGraph | Out-Null
Write-Output ([pscustomobject]@{OutputFolder=$runFolder;TenantId=$ctx.TenantId;Findings=$findings.Count})
