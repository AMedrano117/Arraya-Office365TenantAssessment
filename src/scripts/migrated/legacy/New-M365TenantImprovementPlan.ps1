[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AssessmentJsonPath,
    [Parameter(Mandatory = $false)]
    [string]$OutputFolder,
    [Parameter(Mandatory = $false)]
    [string]$OutputPrefix,
    [Parameter(Mandatory = $false)]
    [ValidateRange(30, 1095)]
    [int]$StaleDeviceDays = 180,
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 50)]
    [int]$MaxGlobalAdmins = 5,
    [Parameter(Mandatory = $false)]
    [switch]$UseGraphFallback,
    [Parameter(Mandatory = $false)]
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
$graphFallbackEnabled = $UseGraphFallback.IsPresent
$graphFallbackUnavailableMessageShown = $false

function Import-ArrayaCommonModuleForPlanningScripts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$GraphFallbackRequired
    )

    $commonModuleManifestPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'))
    if (-not (Test-Path -Path $commonModuleManifestPath)) {
        if ($GraphFallbackRequired) {
            Write-Warning "Graph fallback requested but common module manifest was not found: $commonModuleManifestPath"
        }
        return $false
    }

    $resolvedCommonManifestPath = (Resolve-Path -Path $commonModuleManifestPath).Path
    $loadedCommonModule = Get-Module -Name 'Arraya.M365.Common' -ErrorAction SilentlyContinue | Select-Object -First 1
    $requiredCommonCommands = @(
        'Convert-ArrayaObjectToArray',
        'Get-ArrayaObjectValue',
        'Convert-ArrayaToNumber',
        'Convert-ArrayaToDate',
        'Get-ArrayaGraphResource'
    )
    $missingCommonCommands = @(
        $requiredCommonCommands | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) }
    )
    if (
        -not $loadedCommonModule -or
        $loadedCommonModule.Path -ne $resolvedCommonManifestPath -or
        $missingCommonCommands.Count -gt 0
    ) {
        Import-Module -Name $resolvedCommonManifestPath -Force -ErrorAction Stop
    }

    return $true
}

if (-not (Import-ArrayaCommonModuleForPlanningScripts -GraphFallbackRequired:$graphFallbackEnabled)) {
    $graphFallbackEnabled = $false
}

function New-Finding {
    param(
        [string]$RuleId,
        [string]$Category,
        [string]$Severity,
        [string]$Finding,
        [string]$Recommendation,
        [string]$Value,
        [string]$Target
    )

    [PSCustomObject]@{
        RuleId         = $RuleId
        Category       = $Category
        Severity       = $Severity
        Finding        = $Finding
        Recommendation = $Recommendation
        Value          = $Value
        Target         = $Target
    }
}

function Get-ImprovementPlanDataset {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $DataRoot,
        [Parameter(Mandatory = $true)]
        [string[]]$Names,
        [Parameter(Mandatory = $false)]
        [string]$GraphUri,
        [Parameter(Mandatory = $false)]
        [string]$Activity = 'Graph fallback dataset fetch',
        [Parameter(Mandatory = $false)]
        [switch]$UseGraphFallback
    )

    $rows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $DataRoot -Names $Names)
    if ($rows.Count -gt 0 -or -not $UseGraphFallback -or [string]::IsNullOrWhiteSpace($GraphUri)) {
        return $rows
    }

    $graphHelper = Get-Command -Name 'Get-ArrayaGraphResource' -ErrorAction SilentlyContinue
    if (-not $graphHelper) {
        if (-not $script:graphFallbackUnavailableMessageShown) {
            Write-Warning 'Graph fallback requested but Get-ArrayaGraphResource is unavailable in the current session.'
            $script:graphFallbackUnavailableMessageShown = $true
        }
        return @()
    }

    try {
        $graphResult = Get-ArrayaGraphResource -Uri $GraphUri -Activity $Activity -PageSize 999
        return Convert-ArrayaObjectToArray $graphResult
    }
    catch {
        Write-Warning "Graph fallback failed for '$Activity': $($_.Exception.Message)"
        return @()
    }
}

if (-not (Test-Path -Path $AssessmentJsonPath)) {
    throw "Assessment JSON not found: $AssessmentJsonPath"
}

$jsonRoot = Get-Content -Raw -Path $AssessmentJsonPath | ConvertFrom-Json
$tenantData = if ($jsonRoot.PSObject.Properties.Name -contains 'Data') { $jsonRoot.Data } else { $jsonRoot }

if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $OutputFolder = Split-Path -Path $AssessmentJsonPath -Parent
}
if (-not (Test-Path -Path $OutputFolder)) {
    $null = New-Item -ItemType Directory -Path $OutputFolder -Force
}

if ([string]::IsNullOrWhiteSpace($OutputPrefix)) {
    $OutputPrefix = [IO.Path]::GetFileNameWithoutExtension($AssessmentJsonPath)
}

$findings = New-Object System.Collections.Generic.List[object]

# Rule: Secure Score
$secureScoreRows = Get-ImprovementPlanDataset -DataRoot $tenantData -Names @('SecuritySecureScore', 'SecureScore') -GraphUri '/v1.0/security/secureScores?$top=10' -Activity 'Secure Score fallback' -UseGraphFallback:$graphFallbackEnabled
if ($secureScoreRows.Count -gt 0) {
    $latestSecureScore = $secureScoreRows |
        Sort-Object { Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('CreatedDateTime', 'createdDateTime')) } -Descending |
        Select-Object -First 1

    $currentScore = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $latestSecureScore -Names @('CurrentScore', 'currentScore'))
    $maxScore = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $latestSecureScore -Names @('MaxScore', 'maxScore'))

    if ($null -ne $currentScore -and $null -ne $maxScore -and $maxScore -gt 0) {
        $pct = [math]::Round(($currentScore / $maxScore) * 100, 2)
        if ($pct -lt 40) {
            $findings.Add((New-Finding -RuleId 'SEC-001' -Category 'Security Posture' -Severity 'Critical' -Finding 'Secure Score is below 40%.' -Recommendation 'Prioritize Microsoft Secure Score improvement actions with highest impact first.' -Value "$pct%" -Target '>= 60%')) | Out-Null
        } elseif ($pct -lt 60) {
            $findings.Add((New-Finding -RuleId 'SEC-001' -Category 'Security Posture' -Severity 'High' -Finding 'Secure Score is below 60%.' -Recommendation 'Implement top score improvement actions in identity, email, and endpoint controls.' -Value "$pct%" -Target '>= 60%')) | Out-Null
        } elseif ($pct -lt 75) {
            $findings.Add((New-Finding -RuleId 'SEC-001' -Category 'Security Posture' -Severity 'Medium' -Finding 'Secure Score can be improved.' -Recommendation 'Review unresolved improvement actions and automate policy deployment where possible.' -Value "$pct%" -Target '>= 75%')) | Out-Null
        } else {
            $findings.Add((New-Finding -RuleId 'SEC-001' -Category 'Security Posture' -Severity 'Info' -Finding 'Secure Score is in a healthy range.' -Recommendation 'Sustain score improvements and track trend monthly.' -Value "$pct%" -Target 'Maintain > 75%')) | Out-Null
        }
    }
}

# Rule: Conditional Access baseline
$caPolicies = Get-ImprovementPlanDataset -DataRoot $tenantData -Names @('ConditionalAccessPolicies', 'ConditionalAccess') -GraphUri '/v1.0/identity/conditionalAccess/policies' -Activity 'Conditional Access policy fallback' -UseGraphFallback:$graphFallbackEnabled
$enabledCaPolicies = @(
    $caPolicies | Where-Object {
        $state = (Get-ArrayaObjectValue -Object $_ -Names @('State', 'state'))
        $null -ne $state -and $state.ToString().ToLower().Contains('enabled')
    }
)
if ($caPolicies.Count -eq 0) {
    $findings.Add((New-Finding -RuleId 'CA-001' -Category 'Identity Access' -Severity 'Critical' -Finding 'No Conditional Access policies were found.' -Recommendation 'Create baseline Conditional Access policies for MFA, legacy auth blocking, and admin protection.' -Value '0 policies' -Target '>= 3 enabled baseline policies')) | Out-Null
} elseif ($enabledCaPolicies.Count -eq 0) {
    $findings.Add((New-Finding -RuleId 'CA-001' -Category 'Identity Access' -Severity 'High' -Finding 'Conditional Access policies exist but none appear enabled.' -Recommendation 'Enable validated baseline policies in stages using report-only then enforce mode.' -Value "$($caPolicies.Count) policies, 0 enabled" -Target 'At least baseline policies enabled')) | Out-Null
}

# Rule: MFA registration
$mfaSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('MfaRegistrationSummary', 'MFARegistrationSummary', 'MfaRegistration', 'MFARegistration')
if ($mfaSummary) {
    $registeredUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaSummary -Names @('RegisteredUsers', 'RegisteredUserCount', 'MfaRegisteredUsers'))
    $totalUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaSummary -Names @('TotalUsers', 'UserCount', 'TotalUserCount'))
    $registrationPct = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaSummary -Names @('RegistrationPercent', 'RegisteredPercent', 'MfaRegistrationPercent'))

    if ($null -eq $registrationPct -and $null -ne $registeredUsers -and $null -ne $totalUsers -and $totalUsers -gt 0) {
        $registrationPct = [math]::Round(($registeredUsers / $totalUsers) * 100, 2)
    }

    if ($null -ne $registrationPct) {
        if ($registrationPct -lt 70) {
            $findings.Add((New-Finding -RuleId 'MFA-001' -Category 'Identity Access' -Severity 'High' -Finding 'MFA registration appears low.' -Recommendation 'Run an MFA registration campaign and enforce MFA with Conditional Access.' -Value "$registrationPct%" -Target '>= 90%')) | Out-Null
        } elseif ($registrationPct -lt 90) {
            $findings.Add((New-Finding -RuleId 'MFA-001' -Category 'Identity Access' -Severity 'Medium' -Finding 'MFA registration is moderate.' -Recommendation 'Complete MFA enrollment for remaining user segments.' -Value "$registrationPct%" -Target '>= 90%')) | Out-Null
        }
    }
}

# Rule: Global admin count
$adminRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('AllOffice365Admins', 'Office365Admins', 'Admins'))
$globalAdmins = @(
    $adminRows | Where-Object {
        $role = (Get-ArrayaObjectValue -Object $_ -Names @('Role', 'RoleName', 'DirectoryRole', 'AdminRole'))
        $null -ne $role -and $role.ToString() -match 'Global Administrator|Company Administrator'
    }
)
$globalAdminCount = if ($globalAdmins.Count -gt 0) { $globalAdmins.Count } else { $adminRows.Count }
if ($globalAdminCount -gt $MaxGlobalAdmins) {
    $findings.Add((New-Finding -RuleId 'ADMIN-001' -Category 'Privileged Access' -Severity 'High' -Finding 'Global administrator count exceeds recommended threshold.' -Recommendation 'Reduce standing Global Admin accounts and move to least privilege with PIM/JIT.' -Value "$globalAdminCount accounts" -Target "<= $MaxGlobalAdmins accounts")) | Out-Null
}

# Rule: Domain verification
$domainRows = Get-ImprovementPlanDataset -DataRoot $tenantData -Names @('Domains') -GraphUri '/v1.0/domains' -Activity 'Domain fallback' -UseGraphFallback:$graphFallbackEnabled
$unverifiedDomains = @(
    $domainRows | Where-Object {
        $isVerified = Get-ArrayaObjectValue -Object $_ -Names @('IsVerified', 'Verified', 'isVerified')
        if ($isVerified -is [bool]) { return (-not $isVerified) }
        if ($null -eq $isVerified) { return $false }
        return ($isVerified.ToString().ToLower() -notin @('true', 'verified'))
    }
)
if ($unverifiedDomains.Count -gt 0) {
    $findings.Add((New-Finding -RuleId 'DOMAIN-001' -Category 'Domain Hygiene' -Severity 'Medium' -Finding 'Unverified domains were detected.' -Recommendation 'Verify all required domains and remove stale/unneeded domain entries.' -Value "$($unverifiedDomains.Count) unverified" -Target '0 unverified')) | Out-Null
}

# Rule: License pressure
$licenseRows = Get-ImprovementPlanDataset -DataRoot $tenantData -Names @('LicenseSKUs', 'Licenses') -GraphUri '/v1.0/subscribedSkus' -Activity 'Subscribed SKU fallback' -UseGraphFallback:$graphFallbackEnabled
$highUtilSkus = New-Object System.Collections.Generic.List[string]
foreach ($sku in $licenseRows) {
    $skuName = (Get-ArrayaObjectValue -Object $sku -Names @('SkuPartNumber', 'DisplayName', 'ProductName', 'SkuId'))
    if ([string]::IsNullOrWhiteSpace($skuName)) { $skuName = 'UnknownSKU' }

    $consumed = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $sku -Names @('ConsumedUnits', 'Consumed', 'Assigned'))
    $active = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $sku -Names @('ActiveUnits', 'EnabledUnits', 'TotalUnits'))

    if ($null -eq $active) {
        $prepaid = Get-ArrayaObjectValue -Object $sku -Names @('PrepaidUnits')
        if ($prepaid) {
            $active = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $prepaid -Names @('Enabled', 'enabled'))
        }
    }

    if ($null -ne $consumed -and $null -ne $active -and $active -gt 0) {
        $util = [math]::Round(($consumed / $active) * 100, 2)
        if ($util -ge 95) {
            $highUtilSkus.Add("$skuName ($util%)") | Out-Null
        }
    }
}
if ($highUtilSkus.Count -gt 0) {
    $findings.Add((New-Finding -RuleId 'LIC-001' -Category 'Licensing' -Severity 'Medium' -Finding 'One or more license SKUs are near/full capacity.' -Recommendation 'Reclaim inactive licenses and procure additional seats for constrained SKUs.' -Value ($highUtilSkus -join '; ') -Target '< 95% utilization per SKU')) | Out-Null
}

# Rule: Stale devices
$deviceRows = Get-ImprovementPlanDataset -DataRoot $tenantData -Names @('DeviceDetails', 'Devices') -GraphUri '/v1.0/devices?$select=id,displayName,approximateLastSignInDateTime' -Activity 'Device fallback' -UseGraphFallback:$graphFallbackEnabled
if ($deviceRows.Count -gt 0) {
    $staleCutoff = (Get-Date).AddDays(-1 * $StaleDeviceDays)
    $staleDevices = @(
        $deviceRows | Where-Object {
            $lastSignIn = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('ApproximateLastSignInDateTime', 'LastSignInDateTime', 'LastSignInDate', 'LastLogonDateTime'))
            $null -ne $lastSignIn -and $lastSignIn -lt $staleCutoff
        }
    )

    $stalePct = [math]::Round((($staleDevices.Count / $deviceRows.Count) * 100), 2)
    if ($stalePct -ge 30) {
        $findings.Add((New-Finding -RuleId 'DEV-001' -Category 'Device Hygiene' -Severity 'High' -Finding 'A high percentage of devices appear stale.' -Recommendation 'Review and disable/remove stale devices and enforce lifecycle policies.' -Value "$stalePct% stale ($($staleDevices.Count)/$($deviceRows.Count))" -Target '< 15% stale')) | Out-Null
    } elseif ($stalePct -ge 15) {
        $findings.Add((New-Finding -RuleId 'DEV-001' -Category 'Device Hygiene' -Severity 'Medium' -Finding 'Device stale rate is above target.' -Recommendation 'Investigate stale inventory and tune automated cleanup process.' -Value "$stalePct% stale ($($staleDevices.Count)/$($deviceRows.Count))" -Target '< 15% stale')) | Out-Null
    }
}

if ($findings.Count -eq 0) {
    $findings.Add((New-Finding -RuleId 'GEN-000' -Category 'General' -Severity 'Info' -Finding 'No actionable findings were generated from detected fields.' -Recommendation 'Validate input JSON completeness and rerun assessment with full mode.' -Value 'N/A' -Target 'N/A')) | Out-Null
}

$severityWeight = @{
    Critical = 5
    High     = 4
    Medium   = 3
    Low      = 2
    Info     = 1
}

$sortedFindings = $findings |
    Sort-Object @{ Expression = { $severityWeight[$_.Severity] }; Descending = $true }, Category, RuleId

$jsonOutPath = Join-Path -Path $OutputFolder -ChildPath "$OutputPrefix-ImprovementPlan.json"
$csvOutPath = Join-Path -Path $OutputFolder -ChildPath "$OutputPrefix-ImprovementPlan.csv"
$mdOutPath = Join-Path -Path $OutputFolder -ChildPath "$OutputPrefix-ImprovementPlan.md"
$snippetOutPath = Join-Path -Path $OutputFolder -ChildPath "$OutputPrefix-RemediationSnippets.ps1"

$payload = [PSCustomObject]@{
    GeneratedAt = (Get-Date).ToString('o')
    SourceFile  = (Resolve-Path -Path $AssessmentJsonPath).Path
    Thresholds  = [PSCustomObject]@{
        StaleDeviceDays = $StaleDeviceDays
        MaxGlobalAdmins = $MaxGlobalAdmins
    }
    Findings    = $sortedFindings
}

$payload | ConvertTo-Json -Depth 8 | Set-Content -Path $jsonOutPath -Encoding UTF8
$sortedFindings | Export-Csv -Path $csvOutPath -NoTypeInformation -Encoding UTF8

$severityCounts = $sortedFindings | Group-Object Severity | Sort-Object Name
$summaryLines = @(
    "# Microsoft 365 Tenant Improvement Plan",
    "",
    "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
    "Source: $AssessmentJsonPath",
    "",
    "## Severity Summary"
)
foreach ($group in $severityCounts) {
    $summaryLines += "- $($group.Name): $($group.Count)"
}

$summaryLines += @(
    "",
    "## Findings",
    "",
    "| Severity | Rule | Category | Finding | Value | Target |",
    "|---|---|---|---|---|---|"
)
foreach ($finding in $sortedFindings) {
    $summaryLines += "| $($finding.Severity) | $($finding.RuleId) | $($finding.Category) | $($finding.Finding) | $($finding.Value) | $($finding.Target) |"
}
Set-Content -Path $mdOutPath -Value ($summaryLines -join [Environment]::NewLine) -Encoding UTF8

$snippetLibrary = @{
    'SEC-001' = @"
# Secure Score review
Connect-MgGraph -Scopes 'SecurityEvents.Read.All','SecurityEvents.ReadWrite.All'
Get-MgSecuritySecureScore -Top 1 | Format-List *
"@
    'CA-001' = @"
# Conditional Access baseline validation
Connect-MgGraph -Scopes 'Policy.Read.All','Policy.ReadWrite.ConditionalAccess'
Get-MgIdentityConditionalAccessPolicy | Select-Object DisplayName, State
"@
    'MFA-001' = @"
# MFA registration review
Connect-MgGraph -Scopes 'Reports.Read.All'
Get-MgReportAuthenticationMethodUserRegistrationDetail -All |
    Select-Object UserPrincipalName, IsMfaRegistered
"@
    'ADMIN-001' = @"
# Global Administrator review
Connect-MgGraph -Scopes 'RoleManagement.Read.Directory','Directory.Read.All'
$gaRole = Get-MgDirectoryRole -Filter \"displayName eq 'Global Administrator'\"
Get-MgDirectoryRoleMember -DirectoryRoleId $gaRole.Id -All
"@
    'DOMAIN-001' = @"
# Domain verification review
Connect-MgGraph -Scopes 'Domain.Read.All'
Get-MgDomain | Select-Object Id, IsVerified, IsDefault
"@
    'LIC-001' = @"
# License utilization review
Connect-MgGraph -Scopes 'Organization.Read.All'
Get-MgSubscribedSku | Select-Object SkuPartNumber, ConsumedUnits, @{N='Enabled';E={$_.PrepaidUnits.Enabled}}
"@
    'DEV-001' = @"
# Stale device review
Connect-MgGraph -Scopes 'Device.Read.All'
$cutoff = (Get-Date).AddDays(-180)
Get-MgDevice -All -Property Id,DisplayName,ApproximateLastSignInDateTime |
    Where-Object { $_.ApproximateLastSignInDateTime -and $_.ApproximateLastSignInDateTime -lt $cutoff } |
    Select-Object DisplayName, ApproximateLastSignInDateTime
"@
}

$presentRuleIds = $sortedFindings.RuleId | Sort-Object -Unique
$snippetBlocks = @(
    "<#",
    "Remediation snippets generated by New-M365TenantImprovementPlan.ps1.",
    "Review and test each command in a non-production scope before enforcing changes.",
    "#>",
    ""
)
foreach ($ruleId in $presentRuleIds) {
    if ($snippetLibrary.ContainsKey($ruleId)) {
        $snippetBlocks += "##############################"
        $snippetBlocks += "# Rule: $ruleId"
        $snippetBlocks += "##############################"
        $snippetBlocks += $snippetLibrary[$ruleId].Trim()
        $snippetBlocks += ""
    }
}
Set-Content -Path $snippetOutPath -Value ($snippetBlocks -join [Environment]::NewLine) -Encoding UTF8

$outputSummary = [PSCustomObject]@{
    FindingsCount      = $sortedFindings.Count
    JsonPath           = $jsonOutPath
    CsvPath            = $csvOutPath
    MarkdownPath       = $mdOutPath
    RemediationPs1Path = $snippetOutPath
}

Write-Host "Improvement plan generated."
Write-Host "  JSON: $jsonOutPath"
Write-Host "  CSV: $csvOutPath"
Write-Host "  MD : $mdOutPath"
Write-Host "  PS1: $snippetOutPath"

if ($PassThru) {
    $outputSummary
}
