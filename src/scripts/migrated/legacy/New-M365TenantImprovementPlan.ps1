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
    [Alias('LiveRefresh')]
    [switch]$UseGraphFallback,
    [Parameter(Mandatory = $false)]
    [switch]$IncludeLegacyArtifacts,
    [Parameter(Mandatory = $false)]
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
$graphFallbackEnabled = $UseGraphFallback.IsPresent
$graphFallbackUnavailableMessageShown = $false

function Write-ImproveConsoleWarning {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Message)

    if ($Message -match '(?i)conditional access|inbox rule') {
        Write-Verbose $Message
        return
    }

    Write-Warning $Message
}

function Import-ArrayaCommonModuleForPlanningScripts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$GraphFallbackRequired
    )

    $commonModuleManifestPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'))
    if (-not (Test-Path -Path $commonModuleManifestPath)) {
        if ($GraphFallbackRequired) {
            Write-ImproveConsoleWarning "Snapshot plus live refresh was requested, but the common module manifest was not found: $commonModuleManifestPath"
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
        'Get-ArrayaTenantSnapshotMetricSet',
        'Import-ArrayaOffice365CustomLocal',
        'Import-ArrayaTenantSnapshotContext',
        'Import-ArrayaTenantSnapshot',
        'Resolve-ArrayaSnapshotOutputContext',
        'Convert-ArrayaSnapshotToLegacyTenantStatsHash',
        'Test-ArrayaTenantSnapshot'
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

    $repoRoot = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\..'))
    Import-ArrayaOffice365CustomLocal -RepoRoot $repoRoot -RequiredCommands @('Office365Custom\Get-GraphData') | Out-Null

    return $true
}

if (-not (Import-ArrayaCommonModuleForPlanningScripts -GraphFallbackRequired:$graphFallbackEnabled)) {
    $graphFallbackEnabled = $false
}

function Get-ArrayaGraphResource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [Parameter(Mandatory = $false)]
        [int]$PageSize = 999,
        [Parameter(Mandatory = $false)]
        [string]$Activity = 'Live refresh dataset fetch'
    )

    $resolvedUri = if ($Uri -match '^https?://') { $Uri } else { "https://graph.microsoft.com$Uri" }
    return Office365Custom\Get-GraphData -Uri $resolvedUri -PageSize $PageSize -Activity $Activity
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
        [string]$Activity = 'Live refresh dataset fetch',
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
            Write-ImproveConsoleWarning 'Snapshot plus live refresh was requested, but live Graph access is unavailable in the current session.'
            $script:graphFallbackUnavailableMessageShown = $true
        }
        return @()
    }

    try {
        $graphResult = Get-ArrayaGraphResource -Uri $GraphUri -Activity $Activity -PageSize 999
        return Convert-ArrayaObjectToArray $graphResult
    }
    catch {
        Write-ImproveConsoleWarning "Snapshot plus live refresh failed for '$Activity': $($_.Exception.Message)"
        return @()
    }
}

function Convert-ToArrayaBoolean {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { return [bool]$Value }

    switch -Regex (([string]$Value).Trim().ToLowerInvariant()) {
        '^(true|yes|enabled|on|1)$' { return $true }
        '^(false|no|disabled|off|0)$' { return $false }
        default { return $null }
    }
}

function Convert-ToArrayaStringList {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Value)) { return @() }
        return @(
            $Value -split '[,;]' |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )
    }

    return @(
        (Convert-ArrayaObjectToArray $Value) |
            ForEach-Object { [string]$_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Convert-ToArrayaDisplayText {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [Parameter(Mandatory = $false)]
        [string]$Default = 'N/A'
    )

    if ($null -eq $Value) { return $Default }
    if ($Value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
        return $Value.Trim()
    }
    if ($Value -is [System.Collections.IDictionary]) {
        if ($Value.Count -eq 0) { return $Default }
        return (($Value.Keys | Sort-Object | ForEach-Object { '{0}={1}' -f $_, $Value[$_] }) -join '; ')
    }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @(
            (Convert-ArrayaObjectToArray $Value) |
                ForEach-Object { Convert-ToArrayaDisplayText -Value $_ -Default '' } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )
        if ($items.Count -eq 0) { return $Default }
        return ($items -join '; ')
    }

    return [string]$Value
}

function Convert-ToArrayaMarkdownText {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    $text = Convert-ToArrayaDisplayText -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return 'N/A' }
    return (($text -replace '\|', '\|') -replace "(`r`n|`n|`r)", '<br/>')
}

function Convert-ToArrayaHtmlEncodedText {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [Parameter(Mandatory = $false)]
        [string]$Default = 'N/A'
    )

    $text = Convert-ToArrayaDisplayText -Value $Value -Default $Default
    return [System.Net.WebUtility]::HtmlEncode($text)
}

function Convert-ToArrayaHtmlFragment {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [Parameter(Mandatory = $false)]
        [string]$Default = 'N/A'
    )

    $encoded = Convert-ToArrayaHtmlEncodedText -Value $Value -Default $Default
    return ($encoded -replace "(`r`n|`n|`r)", '<br/>')
}

function Get-NormalizedSeverity {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [AllowNull()]$Priority
    )

    $normalized = if ($null -eq $Value) { '' } else { ([string]$Value).Trim().ToLowerInvariant() }
    switch ($normalized) {
        'critical' { return 'Critical' }
        'high' { return 'High' }
        'risk' { return 'High' }
        'warning' { return 'Medium' }
        'medium' { return 'Medium' }
        'low' { return 'Low' }
        'healthy' { return 'Info' }
        'success' { return 'Info' }
        'ok' { return 'Info' }
        'ready' { return 'Info' }
        'info' { return 'Info' }
    }

    $priorityText = if ($null -eq $Priority) { '' } else { ([string]$Priority).Trim().ToLowerInvariant() }
    switch ($priorityText) {
        'critical' { return 'Critical' }
        'p0' { return 'Critical' }
        'urgent' { return 'Critical' }
        'high' { return 'High' }
        'p1' { return 'High' }
        'medium' { return 'Medium' }
        'p2' { return 'Medium' }
        'low' { return 'Low' }
        'p3' { return 'Low' }
    }

    return 'Medium'
}

function Get-PriorityBand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Severity,
        [AllowNull()]$Priority
    )

    $priorityText = if ($null -eq $Priority) { '' } else { ([string]$Priority).Trim().ToLowerInvariant() }
    if ($priorityText -in @('immediate', 'urgent', 'p0')) { return 'Immediate' }
    if ($priorityText -in @('near term', 'near-term', 'p1', 'high')) { return 'Near Term' }
    if ($priorityText -in @('planned', 'p2', 'medium')) { return 'Planned' }
    if ($priorityText -in @('monitor', 'watch', 'p3', 'low')) { return 'Monitor' }

    switch ($Severity) {
        'Critical' { return 'Immediate' }
        'High' { return 'Near Term' }
        'Medium' { return 'Planned' }
        default { return 'Monitor' }
    }
}

function Get-OwnerTeam {
    [CmdletBinding()]
    param(
        [AllowNull()]$Area,
        [AllowNull()]$Category
    )

    $lookup = ('{0} {1}' -f [string]$Area, [string]$Category).ToLowerInvariant()
    if ($lookup -match 'identity|conditional access|mfa|admin|privileged|entra|authentication|passwordless|guest') { return 'Identity' }
    if ($lookup -match 'exchange|mail|smtp|domain|public folder|connector|spam') { return 'Messaging' }
    if ($lookup -match 'sharepoint|onedrive|teams|group|ownership|collaboration|stewardship|site') { return 'Collaboration' }
    if ($lookup -match 'license|sku') { return 'Licensing' }
    if ($lookup -match 'device|endpoint|intune|compliance') { return 'Endpoint' }
    if ($lookup -match 'secure score|security|defender|purview|zero trust') { return 'Security' }
    return 'Governance'
}

function Get-RoadmapPhase {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$PriorityBand)

    switch ($PriorityBand) {
        'Immediate' { return 'Immediate' }
        'Near Term' { return 'Near Term' }
        'Planned' { return 'Planned' }
        default { return 'Monitor' }
    }
}

function New-CustomerSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Finding,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Recommendation,
        [AllowNull()]$Value
    )

    $valueText = Convert-ToArrayaDisplayText -Value $Value
    if ($valueText -eq 'N/A') {
        return "$Finding Next step: $Recommendation"
    }

    return "$Finding Current state: $valueText. Next step: $Recommendation"
}

function Test-GenericActionText {
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $true }

    $normalizedText = $Text.Trim().ToLowerInvariant()
    return $normalizedText -in @(
        'review this recommendation and assign an implementation owner.',
        'review this finding and assign an implementation owner.',
        'n/a',
        'na'
    )
}

function New-EngineerNotes {
    [CmdletBinding()]
    param(
        [AllowNull()]$Source,
        [AllowNull()]$RelatedWorksheet,
        [AllowNull()]$RelatedSection,
        [AllowNull()]$AdditionalNotes
    )

    $notes = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace([string]$Source)) { $notes.Add("Source=$Source") | Out-Null }
    if (-not [string]::IsNullOrWhiteSpace([string]$RelatedWorksheet)) { $notes.Add("Worksheet=$RelatedWorksheet") | Out-Null }
    if (-not [string]::IsNullOrWhiteSpace([string]$RelatedSection)) { $notes.Add("Section=$RelatedSection") | Out-Null }
    if (-not [string]::IsNullOrWhiteSpace([string]$AdditionalNotes)) { $notes.Add([string]$AdditionalNotes) | Out-Null }

    if ($notes.Count -eq 0) {
        return 'Review the underlying snapshot row before implementing changes.'
    }

    return ($notes -join '; ')
}

function Get-DefaultWhyFlagged {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Finding,
        [AllowNull()]$CurrentValue
    )

    $valueText = Convert-ToArrayaDisplayText -Value $CurrentValue
    if ($valueText -ne 'N/A') {
        return "Flagged because the assessment observed: $valueText"
    }

    return $Finding
}

function Get-DefaultExampleAction {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$RuleId,
        [AllowNull()][string]$Area,
        [AllowNull()][string]$Category,
        [AllowNull()][string]$Finding
    )

    $lookup = ('{0} {1} {2} {3}' -f [string]$RuleId, [string]$Area, [string]$Category, [string]$Finding).ToLowerInvariant()
    if ($lookup -match 'enterprise app|service principal|application permission') {
        return 'Example: review the top flagged enterprise apps, confirm the business owner, and remove broad delegated or application permissions that are no longer required.'
    }
    if ($lookup -match 'onedrive|ownership|stewardship') {
        return 'Example: confirm who now owns the content, document the delegated steward, and decide whether the site should be retained, transferred, or removed.'
    }
    if ($lookup -match 'team|group owner|ownerless') {
        return 'Example: assign an active owner to the Team or group, confirm the service purpose, and archive the workspace if it is no longer needed.'
    }
    if ($lookup -match 'license|sku|capacity') {
        return 'Example: identify the affected SKU, reclaim unused assignments, and verify whether additional licensing needs to be purchased.'
    }
    if ($lookup -match 'device|endpoint|intune|compliance') {
        return 'Example: export the affected devices, assign remediation ownership to endpoint operations, and validate whether they should be enrolled, upgraded, or retired.'
    }
    if ($lookup -match 'guest|break-glass|global administrator|privileged|conditional access|mfa') {
        return 'Example: review the impacted identities or policies, confirm owner approval, and update the tenant baseline with documented exceptions only.'
    }
    if ($lookup -match 'exchange|mailbox|forward|connector|smtp|public folder') {
        return 'Example: validate the business use case, remove unapproved or unsupported configuration, and document any approved exceptions that must remain.'
    }

    return 'Example: review the supporting worksheet, confirm the current state with the service owner, and update the configuration or governance record to close the finding.'
}

function Test-IsTenantSpecificExampleAction {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Text,
        [AllowNull()][string]$CurrentValue
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }

    $normalizedText = $Text.Trim()
    if ($normalizedText -match '^(?i)example:') {
        return $false
    }

    $currentValueText = Convert-ToArrayaDisplayText -Value $CurrentValue -Default ''
    if (-not [string]::IsNullOrWhiteSpace($currentValueText) -and $normalizedText.Contains($currentValueText)) {
        return $true
    }

    return $true
}

function Get-EvidenceLocation {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$RelatedWorksheet,
        [AllowNull()][string]$RelatedSection,
        [AllowNull()][string]$Source
    )

    $parts = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($RelatedWorksheet) -and $RelatedWorksheet -ne 'N/A') {
        $parts.Add("Worksheet: $RelatedWorksheet") | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($RelatedSection) -and $RelatedSection -ne 'N/A') {
        $parts.Add("Section: $RelatedSection") | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($Source) -and $Source -ne 'N/A') {
        $parts.Add("Source: $Source") | Out-Null
    }

    if ($parts.Count -eq 0) {
        return 'Snapshot-derived evidence; review the generated improvement-plan JSON for the normalized record.'
    }

    return ($parts -join '; ')
}

function Get-DefaultBusinessValue {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$OwnerTeam,
        [AllowNull()][string]$Area,
        [AllowNull()][string]$Category,
        [AllowNull()][string]$Finding
    )

    $lookup = ('{0} {1} {2} {3}' -f [string]$OwnerTeam, [string]$Area, [string]$Category, [string]$Finding).ToLowerInvariant()
    if ($lookup -match 'identity|admin|privileged|conditional access|mfa|security') {
        return 'Reduces identity compromise risk, strengthens access control, and improves the tenant security baseline.'
    }
    if ($lookup -match 'licens|sku|capacity') {
        return 'Improves cost control, avoids licensing blockers, and makes future growth easier to plan.'
    }
    if ($lookup -match 'device|endpoint|intune|compliance') {
        return 'Improves policy enforcement, reduces unmanaged access risk, and strengthens endpoint visibility.'
    }
    if ($lookup -match 'exchange|mailbox|forward|connector|smtp|domain') {
        return 'Reduces data-loss and mail-flow risk while improving operational control over messaging.'
    }
    if ($lookup -match 'sharepoint|onedrive|team|group|collaboration|owner|steward') {
        return 'Improves ownership, governance, and lifecycle control for collaboration spaces and business content.'
    }

    return 'Improves governance clarity and reduces avoidable operational or security risk in the tenant.'
}

function Get-ExchangeForwardingPolicyEvidence {
    [CmdletBinding()]
    param(
        [AllowNull()]$ForwardingPolicySummary,
        [Parameter(Mandatory = $false)][object[]]$RemoteDomains = @()
    )

    $summary = if ($ForwardingPolicySummary) { Get-ArrayaObjectValue -Object $ForwardingPolicySummary -Names @('Summary') } else { $null }

    $collectionState = if ($summary) { Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $summary -Names @('PolicyCollectionState')) -Default '' } else { '' }
    $allowingPolicies = if ($summary) { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $summary -Names @('PoliciesExplicitlyAllowingAutoForwarding')) } else { $null }
    $restrictingPolicies = if ($summary) { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $summary -Names @('PoliciesRestrictingAutoForwarding')) } else { $null }
    $remoteDomainsAllow = if ($summary) { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $summary -Names @('RemoteDomainsAllowingAutoForwarding')) } else { $null }
    $defaultRemoteDomainAllows = if ($summary) { Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $summary -Names @('DefaultRemoteDomainAllowsAutoForwarding')) } else { $null }
    $modes = if ($summary) { Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $summary -Names @('PolicyAutoForwardingModes')) -Default '' } else { '' }

    if ($RemoteDomains.Count -gt 0 -and $null -eq $remoteDomainsAllow) {
        $remoteDomainsAllow = @(
            $RemoteDomains |
                Where-Object {
                    $_ -and
                    (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AutoForwardEnabled'))) -eq $true
                }
        ).Count
    }
    if ($RemoteDomains.Count -gt 0 -and $null -eq $defaultRemoteDomainAllows) {
        $defaultRemoteDomainRecord = @(
            $RemoteDomains |
                Where-Object {
                    $identity = [string](Get-ArrayaObjectValue -Object $_ -Names @('Identity'))
                    $domainName = [string](Get-ArrayaObjectValue -Object $_ -Names @('DomainName'))
                    $identity -match '^default$' -or $domainName -eq '*'
                }
        ) | Select-Object -First 1
        if ($defaultRemoteDomainRecord) {
            $defaultRemoteDomainAllows = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $defaultRemoteDomainRecord -Names @('AutoForwardEnabled'))
        }
    }

    $parts = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($collectionState)) { $parts.Add("Hosted outbound policy lookup=$collectionState") | Out-Null }
    if ($null -ne $allowingPolicies) { $parts.Add("$allowingPolicies hosted outbound policy/policies explicitly allow auto-forwarding") | Out-Null }
    if ($null -ne $restrictingPolicies) { $parts.Add("$restrictingPolicies hosted outbound policy/policies restrict auto-forwarding") | Out-Null }
    if ($null -ne $remoteDomainsAllow) { $parts.Add("$remoteDomainsAllow remote domain(s) have AutoForwardEnabled") | Out-Null }
    if ($null -ne $defaultRemoteDomainAllows) { $parts.Add("Default remote domain AutoForwardEnabled=$(if ($defaultRemoteDomainAllows) { 'True' } else { 'False' })") | Out-Null }
    if (-not [string]::IsNullOrWhiteSpace($modes)) { $parts.Add("Observed policy modes: $modes") | Out-Null }

    if ($parts.Count -eq 0) { return $null }
    return ($parts -join '; ')
}

function Get-FindingPresentationOverrides {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RuleId,
        [Parameter(Mandatory = $true)][string]$Area,
        [Parameter(Mandatory = $true)][string]$Category,
        [Parameter(Mandatory = $true)][string]$Finding,
        [Parameter(Mandatory = $true)][string]$Recommendation,
        [Parameter(Mandatory = $true)][string]$CurrentValue,
        [Parameter(Mandatory = $true)][string]$TargetValue
    )

    $override = @{}
    $lookup = ('{0} {1} {2} {3}' -f $RuleId, $Area, $Category, $Finding).ToLowerInvariant()

    switch -Regex ($RuleId) {
        '^COL-001$' {
            $override['WhyFlagged'] = 'The collaboration governance data shows ownerless assets, unmanaged objects, or unresolved stewardship gaps that leave content without clear accountability.'
            $override['Recommendation'] = 'Review each flagged collaboration asset, assign an accountable business owner or steward, and remove unmanaged spaces that no longer have a valid purpose. Confirm that ownership is documented in the operating model before closing the finding.'
            $override['TargetValue'] = 'Each in-scope collaboration asset has a documented owner or steward and unmanaged spaces are dispositioned'
        }
        '^COL-002$' {
            $override['Finding'] = 'OneDrive delegated ownership review is required.'
            $override['WhyFlagged'] = 'The assessment found OneDrive sites where the original owner and the current steward do not appear to align, which often happens after an employee departure or delegated takeover.'
            $override['Recommendation'] = 'Review each flagged OneDrive to confirm who now owns the business content, whether delegated stewardship is intentional, and whether lifecycle handling is documented. Update the steward, custodian, or retention decision before closing the finding.'
            $override['TargetValue'] = 'Each flagged OneDrive has a documented custodian, delegated steward, and lifecycle decision'
        }
        '^TM-001$' {
            $override['WhyFlagged'] = 'Teams without owners cannot be governed effectively and are more likely to drift without accountability for membership, sharing, or lifecycle decisions.'
            $override['Recommendation'] = 'Assign at least one active owner to every flagged Team, confirm the business purpose, and archive Teams that no longer need to remain active. Validate that the linked Microsoft 365 group ownership is aligned before closing the finding.'
            $override['TargetValue'] = 'Every in-scope Team has at least one active owner'
        }
        '^ID-005$' {
            $override['WhyFlagged'] = 'Guest sign-in activity shows external accounts that have not been active for more than 90 days and may no longer be required.'
            $override['Recommendation'] = 'Review inactive guests with the business sponsor, confirm whether each guest still needs access, and remove or disable accounts that no longer support an approved collaboration scenario. Document any exceptions that must remain.'
            $override['TargetValue'] = 'Inactive guest accounts are reviewed, dispositioned, and only approved exceptions remain'
        }
        '^ID-007$' {
            $override['WhyFlagged'] = 'The enterprise app inventory shows one or more service principals with broad delegated or application permissions that matched the high-privilege pattern set used by the assessment.'
            $override['Recommendation'] = 'Review each flagged enterprise application, confirm the business owner, validate whether the granted delegated or application permissions are still required, and reduce or remove broad consent where possible. Tighten the app-governance process so future high-privilege apps require documented approval.'
            $override['TargetValue'] = 'Each flagged enterprise app has an approved owner, documented business purpose, and validated permission scope'
        }
        '^DEV-002$' {
            $override['WhyFlagged'] = 'The device inventory includes devices marked non-compliant, which means they are failing the current endpoint compliance baseline or are operating under undocumented exceptions.'
            $override['Recommendation'] = 'Review each non-compliant device with endpoint operations, validate whether it should be remediated, excluded, or removed from access, and resolve undocumented exceptions. Re-check compliance after remediation before closing the finding.'
        }
        '^DEV-003$' {
            $override['WhyFlagged'] = 'The device inventory includes devices without a management signal, which indicates they may be accessing resources outside the expected endpoint-management model.'
            $override['Recommendation'] = 'Review unmanaged devices, confirm whether they are in scope for enrollment, and prevent unmanaged access where policy requires managed endpoints. Align enrollment expectations and exception handling with the access baseline.'
        }
        '^DEV-005$' {
            $override['WhyFlagged'] = 'The device-management summary shows a meaningful percentage of devices without a management signal, which weakens policy enforcement and endpoint visibility.'
            $override['Recommendation'] = 'Confirm the intended managed-device scope, review the unmanaged-device population with endpoint operations, and close enrollment gaps for devices that should access protected resources. Document any approved unmanaged exceptions.'
        }
        '^ADMIN-001$' {
            $override['WhyFlagged'] = 'The tenant has more standing Global Administrator accounts than the recommended operating threshold, increasing privileged access exposure.'
            $override['Recommendation'] = 'Review every standing Global Administrator assignment, move day-to-day administration to lower-privilege roles where possible, and retain only the minimum approved permanent admins. Document break-glass coverage separately from routine admin access.'
        }
        '^LIC-001$' {
            $override['WhyFlagged'] = 'One or more paid SKUs are at or near capacity, which can block onboarding and often indicates reclaimable or mismatched assignments.'
            $override['Recommendation'] = 'Review the constrained SKUs, reclaim unused or stale assignments, and confirm whether additional licensing must be purchased for planned demand. Close the gap between purchased capacity and assigned usage before considering the finding resolved.'
            $override['TargetValue'] = 'Affected paid SKUs are reconciled to purchased capacity with approved headroom'
        }
    }

    if ($lookup -match 'missing owner|owners and require ownership assignment|ownerless') {
        $override['WhyFlagged'] = 'The assessment found collaboration assets without a clear accountable owner, which creates gaps in stewardship, lifecycle handling, and access governance.'
        $override['Recommendation'] = 'Assign an active accountable owner or documented steward to each flagged asset, confirm the business purpose, and retire spaces that no longer have a sponsor. Record the ownership decision in the operating model before closing the finding.'
        $override['TargetValue'] = 'Each flagged collaboration asset has an active owner or documented steward'
    }
    elseif ($lookup -match 'over capacity') {
        $override['WhyFlagged'] = 'The licensing output shows at least one SKU assigned beyond purchased capacity, which creates an immediate licensing governance issue.'
        $override['Recommendation'] = 'Review the affected SKU assignments immediately, reclaim licenses from inactive or ineligible accounts, and purchase additional capacity if the assignments are valid and still required.'
        $override['TargetValue'] = 'Affected SKU assignments are brought back within purchased capacity'
    }
    elseif ($lookup -match 'global administrator') {
        $override['WhyFlagged'] = 'The tenant has more standing Global Administrator assignments than the recommended operating threshold, increasing privileged access exposure.'
        $override['Recommendation'] = 'Review each standing Global Administrator assignment, remove routine admin users from the role, and retain only the minimum approved permanent admins plus documented emergency access accounts. Validate whether lower-privilege roles or eligible access can replace standing assignments.'
        $override['TargetValue'] = 'Standing Global Administrator assignments reduced to the approved operating threshold'
    }
    elseif ($lookup -match 'not verified|unverified domain') {
        $override['WhyFlagged'] = 'The domain inventory shows one or more domains that are still not verified, which can indicate stale configuration, incomplete onboarding, or unsupported migration prerequisites.'
        $override['Recommendation'] = 'Review every unverified domain, confirm whether it is still required, complete DNS verification for domains that remain in scope, and remove stale entries that no longer serve a business purpose.'
        $override['TargetValue'] = 'All required domains are verified and stale domains are removed'
    }
    elseif ($lookup -match 'device compliance') {
        $override['WhyFlagged'] = 'The device posture output shows compliance results below the expected baseline, leaving managed access policies less effective.'
        $override['Recommendation'] = 'Review why the compliance baseline is being missed, remediate the highest-volume failure conditions, and tighten exception handling for devices that should not remain non-compliant.'
        $override['TargetValue'] = 'Device compliance meets the approved endpoint baseline'
    }
    elseif ($lookup -match 'appear unowned|ownerless team|without owners') {
        $override['WhyFlagged'] = 'The collaboration data shows workspaces without an accountable owner, which creates gaps in access control, lifecycle handling, and business accountability.'
        $override['Recommendation'] = 'Assign an accountable owner to each flagged workspace, confirm the business purpose, and retire any workspace that no longer has a sponsor.'
        $override['TargetValue'] = 'Each flagged workspace has an active owner or has been retired'
    }

    return $override
}

function New-Finding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RuleId,
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Area,
        [Parameter(Mandatory = $true)][string]$Category,
        [Parameter(Mandatory = $true)][string]$Severity,
        [Parameter(Mandatory = $true)][string]$Finding,
        [Parameter(Mandatory = $true)][string]$Recommendation,
        [AllowNull()]$CurrentValue,
        [AllowNull()]$TargetValue,
        [AllowNull()]$Priority,
        [Parameter(Mandatory = $false)][string]$OwnerTeam,
        [Parameter(Mandatory = $false)][string]$RoadmapPhase,
        [Parameter(Mandatory = $false)][string]$CustomerSummary,
        [Parameter(Mandatory = $false)][string]$EngineerNotes,
        [Parameter(Mandatory = $false)][string]$RelatedWorksheet,
        [Parameter(Mandatory = $false)][string]$RelatedSection,
        [Parameter(Mandatory = $false)][string]$WhyFlagged,
        [Parameter(Mandatory = $false)][string]$ExampleAction,
        [Parameter(Mandatory = $false)][string]$EvidenceLocation,
        [Parameter(Mandatory = $false)][string]$TechnicalRemediation,
        [Parameter(Mandatory = $false)][string]$BusinessValue
    )

    if ([string]::IsNullOrWhiteSpace($Source)) { $Source = 'Heuristic' }
    if ([string]::IsNullOrWhiteSpace($Area)) { $Area = 'General' }
    if ([string]::IsNullOrWhiteSpace($Category)) { $Category = $Area }
    if ([string]::IsNullOrWhiteSpace($Finding)) { $Finding = $RuleId }
    if ([string]::IsNullOrWhiteSpace($Recommendation)) { $Recommendation = 'Review this finding and assign an implementation owner.' }

    if ([string]::IsNullOrWhiteSpace($Severity)) { $Severity = 'Medium' }
    $normalizedSeverity = Get-NormalizedSeverity -Value $Severity -Priority $Priority
    $priorityBand = Get-PriorityBand -Severity $normalizedSeverity -Priority $Priority
    $valueText = Convert-ToArrayaDisplayText -Value $CurrentValue
    $targetText = Convert-ToArrayaDisplayText -Value $TargetValue
    $presentationOverrides = Get-FindingPresentationOverrides -RuleId $RuleId -Area $Area -Category $Category -Finding $Finding -Recommendation $Recommendation -CurrentValue $valueText -TargetValue $targetText

    if ($presentationOverrides.ContainsKey('Finding')) { $Finding = [string]$presentationOverrides['Finding'] }
    if ($presentationOverrides.ContainsKey('Recommendation')) { $Recommendation = [string]$presentationOverrides['Recommendation'] }
    if ($presentationOverrides.ContainsKey('TargetValue') -and ([string]::IsNullOrWhiteSpace($targetText) -or $targetText -eq 'N/A' -or $targetText -eq 'Reduce open findings in this workstream' -or $RuleId -eq 'COL-002')) {
        $targetText = Convert-ToArrayaDisplayText -Value $presentationOverrides['TargetValue']
    }

    if ([string]::IsNullOrWhiteSpace($OwnerTeam)) { $OwnerTeam = Get-OwnerTeam -Area $Area -Category $Category }
    if ([string]::IsNullOrWhiteSpace($RoadmapPhase)) { $RoadmapPhase = Get-RoadmapPhase -PriorityBand $priorityBand }

    if ([string]::IsNullOrWhiteSpace($WhyFlagged)) {
        if ($presentationOverrides.ContainsKey('WhyFlagged')) {
            $WhyFlagged = [string]$presentationOverrides['WhyFlagged']
        }
        else {
            $WhyFlagged = Get-DefaultWhyFlagged -Finding $Finding -CurrentValue $valueText
        }
    }

    if ([string]::IsNullOrWhiteSpace($ExampleAction)) {
        if ($presentationOverrides.ContainsKey('ExampleAction')) {
            $ExampleAction = [string]$presentationOverrides['ExampleAction']
        }
        else {
            $ExampleAction = Get-DefaultExampleAction -RuleId $RuleId -Area $Area -Category $Category -Finding $Finding
        }
    }
    if (-not (Test-IsTenantSpecificExampleAction -Text $ExampleAction -CurrentValue $valueText)) {
        $ExampleAction = $null
    }

    if ([string]::IsNullOrWhiteSpace($CustomerSummary)) { $CustomerSummary = New-CustomerSummary -Finding $Finding -Recommendation $Recommendation -Value $valueText }
    if ([string]::IsNullOrWhiteSpace($EngineerNotes)) { $EngineerNotes = New-EngineerNotes -Source $Source -RelatedWorksheet $RelatedWorksheet -RelatedSection $RelatedSection }
    if ([string]::IsNullOrWhiteSpace($EvidenceLocation)) { $EvidenceLocation = Get-EvidenceLocation -RelatedWorksheet $RelatedWorksheet -RelatedSection $RelatedSection -Source $Source }
    if ([string]::IsNullOrWhiteSpace($TechnicalRemediation)) { $TechnicalRemediation = $Recommendation }
    if ([string]::IsNullOrWhiteSpace($BusinessValue)) { $BusinessValue = Get-DefaultBusinessValue -OwnerTeam $OwnerTeam -Area $Area -Category $Category -Finding $Finding }

    return [PSCustomObject]@{
        RuleId           = $RuleId
        Source           = $Source
        Area             = $Area
        Category         = $Category
        Severity         = $normalizedSeverity
        PriorityBand     = $priorityBand
        Finding          = $Finding
        Recommendation   = $Recommendation
        CurrentValue     = $valueText
        TargetValue      = $targetText
        Value            = $valueText
        Target           = $targetText
        OwnerTeam        = $OwnerTeam
        RoadmapPhase     = $RoadmapPhase
        CustomerSummary  = $CustomerSummary
        EngineerNotes    = $EngineerNotes
        WhyFlagged       = $WhyFlagged
        ExampleAction    = $ExampleAction
        EvidenceLocation = $EvidenceLocation
        TechnicalRemediation = $TechnicalRemediation
        BusinessValue    = $BusinessValue
        RelatedWorksheet = $(if ([string]::IsNullOrWhiteSpace($RelatedWorksheet)) { 'N/A' } else { $RelatedWorksheet })
        RelatedSection   = $(if ([string]::IsNullOrWhiteSpace($RelatedSection)) { 'N/A' } else { $RelatedSection })
    }
}

function Get-FindingSourceRank {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Source)

    switch -Regex ($Source) {
        '^Derived' { return 30 }
        '^Summary' { return 25 }
        '^Hybrid' { return 20 }
        '^Heuristic' { return 10 }
        default { return 0 }
    }
}

function Get-SeverityWeight {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Severity)

    switch ($Severity) {
        'Critical' { return 5 }
        'High' { return 4 }
        'Medium' { return 3 }
        'Low' { return 2 }
        default { return 1 }
    }
}

function Get-FindingKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Area,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Category,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Finding
    )

    if ([string]::IsNullOrWhiteSpace($Area)) { $Area = 'General' }
    if ([string]::IsNullOrWhiteSpace($Category)) { $Category = 'General' }
    if ([string]::IsNullOrWhiteSpace($Finding)) { $Finding = 'General' }

    $parts = @($Area, $Category, $Finding) | ForEach-Object {
        ([string]$_ -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
    }
    return ($parts -join '|')
}

function Add-ImprovementFinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Store,
        [Parameter(Mandatory = $true)][pscustomobject]$Finding,
        [Parameter(Mandatory = $false)][string[]]$AdditionalKeys = @()
    )

    $safeArea = if ([string]::IsNullOrWhiteSpace([string]$Finding.Area)) { 'General' } else { [string]$Finding.Area }
    $safeCategory = if ([string]::IsNullOrWhiteSpace([string]$Finding.Category)) { 'General' } else { [string]$Finding.Category }
    $safeFinding = if ([string]::IsNullOrWhiteSpace([string]$Finding.Finding)) { $Finding.RuleId } else { [string]$Finding.Finding }

    $keys = New-Object System.Collections.Generic.List[string]
    $keys.Add((Get-FindingKey -Area $safeArea -Category $safeCategory -Finding $safeFinding)) | Out-Null
    foreach ($additionalKey in @($AdditionalKeys)) {
        if (-not [string]::IsNullOrWhiteSpace($additionalKey)) {
            $keys.Add($additionalKey) | Out-Null
        }
    }

    $existing = $null
    foreach ($key in $keys) {
        if ($Store.ContainsKey($key)) {
            $existing = $Store[$key]
            break
        }
    }

    if ($existing) {
        $existingRank = Get-FindingSourceRank -Source $existing.Source
        $incomingRank = Get-FindingSourceRank -Source $Finding.Source
        $existingSeverity = Get-SeverityWeight -Severity $existing.Severity
        $incomingSeverity = Get-SeverityWeight -Severity $Finding.Severity

        if ($incomingRank -lt $existingRank) { return $existing }
        if ($incomingRank -eq $existingRank -and $incomingSeverity -lt $existingSeverity) { return $existing }
    }

    foreach ($key in $keys) {
        $Store[$key] = $Finding
    }

    return $Finding
}

function Test-DerivedCoverage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string[]]$Tags,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$DerivedFindings
    )

    if ($DerivedFindings.Count -eq 0 -or $Tags.Count -eq 0) { return $false }

    foreach ($finding in $DerivedFindings) {
        $haystack = ('{0} {1} {2} {3}' -f $finding.Area, $finding.Category, $finding.Finding, $finding.Recommendation).ToLowerInvariant()
        foreach ($tag in $Tags) {
            if (-not [string]::IsNullOrWhiteSpace($tag) -and $haystack.Contains($tag.ToLowerInvariant())) {
                return $true
            }
        }
    }

    return $false
}

function Resolve-FindingMessage {
    [CmdletBinding()]
    param([AllowNull()]$Row)

    foreach ($name in @('Message', 'Issue', 'Finding', 'CurrentState', 'PrimaryFinding', 'Title', 'Name')) {
        $value = Get-ArrayaObjectValue -Object $Row -Names @($name)
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
            return [string]$value
        }
    }

    return $null
}

function Resolve-FindingRecommendation {
    [CmdletBinding()]
    param([AllowNull()]$Row)

    foreach ($name in @('RecommendedAction', 'Recommendation', 'NextAction', 'RoadmapStep', 'Notes')) {
        $value = Get-ArrayaObjectValue -Object $Row -Names @($name)
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
            return [string]$value
        }
    }

    return 'Review this recommendation and assign an implementation owner.'
}

function New-DerivedFindingFromRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][string]$SourceLabel
    )

    $findingText = Resolve-FindingMessage -Row $Row
    if ([string]::IsNullOrWhiteSpace($findingText)) { return $null }

    $area = [string](Get-ArrayaObjectValue -Object $Row -Names @('Area', 'Category', 'AssessmentArea'))
    if ([string]::IsNullOrWhiteSpace($area)) { $area = 'Best Practices' }

    $category = [string](Get-ArrayaObjectValue -Object $Row -Names @('Category', 'Area'))
    if ([string]::IsNullOrWhiteSpace($category)) { $category = $area }

    $severity = Get-NormalizedSeverity -Value (Get-ArrayaObjectValue -Object $Row -Names @('Severity', 'Status')) -Priority (Get-ArrayaObjectValue -Object $Row -Names @('Priority'))
    $recommendation = Resolve-FindingRecommendation -Row $Row
    $ruleId = [string](Get-ArrayaObjectValue -Object $Row -Names @('RecommendationId', 'RuleId'))
    if ([string]::IsNullOrWhiteSpace($ruleId)) {
        $prefix = ($area -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($prefix)) { $prefix = 'DERIVED' }
        $ruleId = "$prefix-$([Math]::Abs($findingText.GetHashCode()))"
    }

    $currentValue = Get-ArrayaObjectValue -Object $Row -Names @('CurrentValue', 'CurrentState', 'Status', 'Message')
    $targetValue = Get-ArrayaObjectValue -Object $Row -Names @('TargetValue', 'BestPracticeTarget', 'TargetState')
    $priority = Get-ArrayaObjectValue -Object $Row -Names @('Priority')
    $relatedWorksheet = [string](Get-ArrayaObjectValue -Object $Row -Names @('RelatedWorksheet', 'Worksheet'))
    $relatedSection = [string](Get-ArrayaObjectValue -Object $Row -Names @('RelatedSection', 'Section'))
    $sourceType = [string](Get-ArrayaObjectValue -Object $Row -Names @('SourceType'))

    return New-Finding `
        -RuleId $ruleId `
        -Source $SourceLabel `
        -Area $area `
        -Category $category `
        -Severity $severity `
        -Finding $findingText `
        -Recommendation $recommendation `
        -CurrentValue $currentValue `
        -TargetValue $targetValue `
        -Priority $priority `
        -RelatedWorksheet $relatedWorksheet `
        -RelatedSection $relatedSection `
        -EngineerNotes (New-EngineerNotes -Source $SourceLabel -RelatedWorksheet $relatedWorksheet -RelatedSection $relatedSection -AdditionalNotes $sourceType)
}

function Add-HeuristicFinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Store,
        [Parameter(Mandatory = $true)][string]$RuleId,
        [Parameter(Mandatory = $true)][string]$Area,
        [Parameter(Mandatory = $true)][string]$Category,
        [Parameter(Mandatory = $true)][string]$Severity,
        [Parameter(Mandatory = $true)][string]$Finding,
        [Parameter(Mandatory = $true)][string]$Recommendation,
        [AllowNull()]$CurrentValue,
        [AllowNull()]$TargetValue,
        [Parameter(Mandatory = $false)][string]$Source = 'Heuristic',
        [Parameter(Mandatory = $false)][string]$RelatedWorksheet = 'N/A',
        [Parameter(Mandatory = $false)][string]$RelatedSection = 'N/A',
        [Parameter(Mandatory = $false)][string]$WhyFlagged,
        [Parameter(Mandatory = $false)][string]$ExampleAction,
        [Parameter(Mandatory = $false)][string[]]$AdditionalKeys = @()
    )

    $findingRecord = @(
        New-Finding `
            -RuleId $RuleId `
            -Source $Source `
            -Area $Area `
            -Category $Category `
            -Severity $Severity `
            -Finding $Finding `
            -Recommendation $Recommendation `
            -CurrentValue $CurrentValue `
            -TargetValue $TargetValue `
            -RelatedWorksheet $RelatedWorksheet `
            -RelatedSection $RelatedSection `
            -WhyFlagged $WhyFlagged `
            -ExampleAction $ExampleAction
    ) | Where-Object { $_ -and $_.PSObject -and ($_.PSObject.Properties.Name -contains 'RuleId') } | Select-Object -Last 1

    if ($null -eq $findingRecord) {
        return
    }

    if (-not $script:ArrayaImproveHeuristicFindings) {
        $script:ArrayaImproveHeuristicFindings = New-Object System.Collections.Generic.List[object]
    }
    $script:ArrayaImproveHeuristicFindings.Add($findingRecord) | Out-Null

    Add-ImprovementFinding -Store $Store -Finding $findingRecord -AdditionalKeys $AdditionalKeys | Out-Null
}

function Get-SortedFindings {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Store)

    return @(
        $Store.Values |
            Where-Object { $_ -and $_.PSObject -and ($_.PSObject.Properties.Name -contains 'RuleId') } |
            Sort-Object -Property RuleId -Unique |
            Sort-Object `
                @{ Expression = { Get-SeverityWeight -Severity $_.Severity }; Descending = $true }, `
                @{ Expression = { switch ($_.PriorityBand) { 'Immediate' { 1 } 'Near Term' { 2 } 'Planned' { 3 } default { 4 } } } }, `
                OwnerTeam, Area, RuleId
    )
}

function Get-MergedSortedFindings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Store,
        [Parameter(Mandatory = $false)][object[]]$SupplementalFindings = @()
    )

    $mergedByRuleId = @{}
    $candidates = New-Object System.Collections.Generic.List[object]
    foreach ($candidate in @($Store.Values)) {
        if ($candidate -and $candidate.PSObject -and ($candidate.PSObject.Properties.Name -contains 'RuleId')) {
            $candidates.Add($candidate) | Out-Null
        }
    }
    foreach ($candidate in @($SupplementalFindings)) {
        if ($candidate -and $candidate.PSObject -and ($candidate.PSObject.Properties.Name -contains 'RuleId')) {
            $candidates.Add($candidate) | Out-Null
        }
    }

    foreach ($candidate in $candidates.ToArray()) {
        $ruleId = [string]$candidate.RuleId
        if ([string]::IsNullOrWhiteSpace($ruleId)) { continue }

        if (-not $mergedByRuleId.ContainsKey($ruleId)) {
            $mergedByRuleId[$ruleId] = $candidate
            continue
        }

        $existing = $mergedByRuleId[$ruleId]
        $existingRank = Get-FindingSourceRank -Source $existing.Source
        $candidateRank = Get-FindingSourceRank -Source $candidate.Source
        $existingSeverity = Get-SeverityWeight -Severity $existing.Severity
        $candidateSeverity = Get-SeverityWeight -Severity $candidate.Severity

        if ($candidateRank -gt $existingRank -or ($candidateRank -eq $existingRank -and $candidateSeverity -gt $existingSeverity)) {
            $mergedByRuleId[$ruleId] = $candidate
        }
    }

    return @(
        $mergedByRuleId.Values |
            Sort-Object `
                @{ Expression = { Get-SeverityWeight -Severity $_.Severity }; Descending = $true }, `
                @{ Expression = { switch ($_.PriorityBand) { 'Immediate' { 1 } 'Near Term' { 2 } 'Planned' { 3 } default { 4 } } } }, `
                OwnerTeam, Area, RuleId
    )
}

function Get-TenantDisplayName {
    [CmdletBinding()]
    param(
        [AllowNull()]$TenantInfoSummary,
        [AllowNull()]$LegacyData,
        [AllowNull()][string]$OutputPrefix
    )

    foreach ($value in @(
        (Get-ArrayaObjectValue -Object $TenantInfoSummary -Names @('TenantName', 'DisplayName', 'CompanyName')),
        (Get-ArrayaObjectValue -Object (Get-ArrayaObjectValue -Object $LegacyData -Names @('TenantInfo')) -Names @('DisplayName', 'Display Name', 'CompanyName', 'TenantName'))
    )) {
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
            return [string]$value
        }
    }

    return $OutputPrefix
}

function Get-TopFindings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Findings,
        [Parameter(Mandatory = $false)][int]$Count = 5
    )

    return @($Findings | Select-Object -First $Count)
}

function Resolve-WorkstreamTopSignals {
    [CmdletBinding()]
    param([AllowNull()]$Row)

    $directValue = Get-ArrayaObjectValue -Object $Row -Names @('TopSignals', 'TopSignalSummary', 'TopFindings')
    $directText = Convert-ToArrayaDisplayText -Value $directValue
    if ($directText -ne 'N/A') {
        return $directText
    }

    $message = [string](Resolve-FindingMessage -Row $Row)
    if ($message -match 'Top signals:\s*(.+)$') {
        return $matches[1].Trim().TrimEnd('.')
    }

    return $null
}

function New-WorkstreamSummaryFromRow {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Row)

    $area = [string](Get-ArrayaObjectValue -Object $Row -Names @('Area', 'Category', 'AssessmentArea'))
    if ([string]::IsNullOrWhiteSpace($area)) { $area = 'Best Practices' }

    $severity = Get-NormalizedSeverity -Value (Get-ArrayaObjectValue -Object $Row -Names @('Severity', 'Status')) -Priority (Get-ArrayaObjectValue -Object $Row -Names @('Priority'))
    $openFindings = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $Row -Names @('TotalFindings'))
    $criticalCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $Row -Names @('CriticalFindings'))
    $warningCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $Row -Names @('WarningFindings'))
    $infoCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $Row -Names @('InfoFindings'))

    return [PSCustomObject]@{
        Workstream    = Get-OwnerTeam -Area $area -Category 'Area Summary'
        Area          = $area
        Severity      = $severity
        OpenFindings  = $(if ($null -eq $openFindings) { 0 } else { $openFindings })
        CriticalCount = $(if ($null -eq $criticalCount) { 0 } else { $criticalCount })
        WarningCount  = $(if ($null -eq $warningCount) { 0 } else { $warningCount })
        InfoCount     = $(if ($null -eq $infoCount) { 0 } else { $infoCount })
        TopSignals    = Resolve-WorkstreamTopSignals -Row $Row
    }
}

function Get-SortedWorkstreamSummaries {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Summaries)

    return @(
        $Summaries |
            Where-Object { $_ } |
            Sort-Object `
                @{ Expression = { Get-SeverityWeight -Severity $_.Severity }; Descending = $true }, `
                @{ Expression = { [int]$_.OpenFindings }; Descending = $true }, `
                Workstream, Area
    )
}

function Get-DerivedWorkstreamTopSignals {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Summary,
        [Parameter(Mandatory = $true)][object[]]$Findings
    )

    $candidateRows = @(
        $Findings |
            Where-Object {
                $_ -and (
                    ([string]$_.OwnerTeam -eq [string]$Summary.Workstream) -or
                    ([string]$_.Area -eq [string]$Summary.Area) -or
                    ([string]$_.Category -eq [string]$Summary.Area)
                )
            }
    )

    if ($candidateRows.Count -eq 0) {
        return $null
    }

    $signals = @(
        $candidateRows |
            ForEach-Object {
                $label = if (-not [string]::IsNullOrWhiteSpace([string]$_.Area) -and [string]$_.Area -ne [string]$Summary.Area) {
                    [string]$_.Area
                }
                elseif (-not [string]::IsNullOrWhiteSpace([string]$_.Category) -and [string]$_.Category -ne [string]$Summary.Area) {
                    [string]$_.Category
                }
                else {
                    [string]$_.RuleId
                }

                if (-not [string]::IsNullOrWhiteSpace($label)) { $label }
            } |
            Group-Object |
            Sort-Object `
                @{ Expression = { $_.Count }; Descending = $true }, `
                @{ Expression = { $_.Name }; Descending = $false } |
            Select-Object -First 3 |
            ForEach-Object { '{0}: {1}' -f $_.Name, $_.Count }
    )

    if ($signals.Count -eq 0) {
        return $null
    }

    return ($signals -join '; ')
}

function Resolve-VisibleWorkstreamSummaries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Summaries,
        [Parameter(Mandatory = $true)][object[]]$Findings
    )

    $visibleSummaries = New-Object System.Collections.Generic.List[object]
    foreach ($summary in @($Summaries)) {
        if (-not $summary) { continue }

        $openFindings = [int](Convert-ArrayaToNumber $summary.OpenFindings)
        $topSignals = Convert-ToArrayaDisplayText -Value $summary.TopSignals -Default ''
        if ([string]::IsNullOrWhiteSpace($topSignals) -or $topSignals -eq 'N/A') {
            $topSignals = Get-DerivedWorkstreamTopSignals -Summary $summary -Findings $Findings
        }

        if ($openFindings -le 0 -and [string]::IsNullOrWhiteSpace($topSignals)) {
            continue
        }

        $visibleSummaries.Add([pscustomobject]@{
            Workstream    = $summary.Workstream
            Area          = $summary.Area
            Severity      = $summary.Severity
            OpenFindings  = $openFindings
            CriticalCount = [int](Convert-ArrayaToNumber $summary.CriticalCount)
            WarningCount  = [int](Convert-ArrayaToNumber $summary.WarningCount)
            InfoCount     = [int](Convert-ArrayaToNumber $summary.InfoCount)
            TopSignals    = $(if ([string]::IsNullOrWhiteSpace($topSignals)) { 'Review detailed findings in this workstream' } else { $topSignals })
        }) | Out-Null
    }

    return @($visibleSummaries.ToArray())
}

function Get-CustomerWorkstreamThemes {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object[]]$Findings)

    $themes = New-Object System.Collections.Generic.List[object]
    $groupedFindings = @(
        $Findings |
            Group-Object OwnerTeam |
            Sort-Object `
                @{ Expression = { $_.Count }; Descending = $true }, `
                @{ Expression = { $_.Name }; Descending = $false }
    )
    foreach ($ownerGroup in $groupedFindings) {
        $seenKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $orderedRows = @(
            $ownerGroup.Group |
                Sort-Object `
                    @{ Expression = { Get-SeverityWeight -Severity $_.Severity }; Descending = $true }, `
                    @{ Expression = { switch ($_.PriorityBand) { 'Immediate' { 1 } 'Near Term' { 2 } 'Planned' { 3 } default { 4 } } } }, `
                    Area, RuleId
        )

        foreach ($finding in $orderedRows) {
            $dedupeKey = ('{0}|{1}|{2}' -f ([string]$finding.Area).Trim().ToLowerInvariant(), ([string]$finding.Recommendation).Trim().ToLowerInvariant(), ([string]$finding.WhyFlagged).Trim().ToLowerInvariant())
            if (-not $seenKeys.Add($dedupeKey)) {
                continue
            }

            $themes.Add([pscustomobject]@{
                Workstream          = $ownerGroup.Name
                Area                = $finding.Area
                Severity            = $finding.Severity
                PriorityBand        = $finding.PriorityBand
                WhatNeedsAttention  = $finding.Finding
                WhyItMatters        = $finding.WhyFlagged
                RecommendedNextStep = $finding.Recommendation
                BusinessValue       = $finding.BusinessValue
            }) | Out-Null

            if ((@($themes | Where-Object { $_.Workstream -eq $ownerGroup.Name })).Count -ge 3) {
                break
            }
        }
    }

    return @($themes.ToArray())
}

function New-CustomerRemediationReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Findings,
        [Parameter(Mandatory = $false)][object[]]$WorkstreamSummaries = @(),
        [Parameter(Mandatory = $true)][string]$TenantName,
        [Parameter(Mandatory = $true)][string]$AssessmentJsonPath,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt
    )

    $severityCounts = $Findings | Group-Object Severity | Sort-Object Name
    $ownerGroups = $Findings | Group-Object OwnerTeam | Sort-Object Count -Descending
    $themeRows = Get-CustomerWorkstreamThemes -Findings $Findings
    $topFindings = Get-TopFindings -Findings @($Findings | Where-Object { $_.Severity -in @('Critical', 'High') }) -Count 6
    if ($topFindings.Count -eq 0) { $topFindings = Get-TopFindings -Findings $Findings -Count 6 }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("# $TenantName Microsoft 365 Remediation Report") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("Generated: $($GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss'))") | Out-Null
    $lines.Add("Assessment Snapshot: $AssessmentJsonPath") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('## Tenant Overview') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('This report turns the tenant assessment into a customer-ready remediation roadmap. It emphasizes business-impacting controls, governance gaps, and the workstreams needed to improve tenant posture.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('## Executive Summary') | Out-Null
    $lines.Add('') | Out-Null
    foreach ($group in $severityCounts) {
        $lines.Add("- $($group.Name): $($group.Count) finding(s)") | Out-Null
    }
    $lines.Add('') | Out-Null
    if ($WorkstreamSummaries.Count -gt 0) {
        $lines.Add('## Workstream Summary') | Out-Null
        $lines.Add('') | Out-Null
        foreach ($summary in $WorkstreamSummaries) {
            $lines.Add("- $($summary.Workstream) / $($summary.Area): $($summary.OpenFindings) open finding(s) with top signals $($summary.TopSignals).") | Out-Null
        }
        $lines.Add('') | Out-Null
    }
    if ($ownerGroups.Count -gt 0) {
        $lines.Add("Primary workstreams: $((@($ownerGroups | Select-Object -First 4 | ForEach-Object { '{0} ({1})' -f $_.Name, $_.Count }) -join '; '))") | Out-Null
        $lines.Add('') | Out-Null
    }
    $lines.Add('## Top Priority Risks') | Out-Null
    $lines.Add('') | Out-Null
    foreach ($finding in $topFindings) {
        $lines.Add("### $($finding.RuleId) - $($finding.Area)") | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add("- Severity: $($finding.Severity)") | Out-Null
        $lines.Add("- Priority: $($finding.PriorityBand)") | Out-Null
        $lines.Add("- Workstream: $($finding.OwnerTeam)") | Out-Null
        $lines.Add("- Why flagged: $($finding.WhyFlagged)") | Out-Null
        $lines.Add("- Recommended action: $($finding.Recommendation)") | Out-Null
        $lines.Add("- Success criteria: $($finding.TargetValue)") | Out-Null
        $lines.Add('') | Out-Null
    }

    $lines.Add('## Phased Roadmap') | Out-Null
    $lines.Add('') | Out-Null
    foreach ($phaseName in @('Immediate', 'Near Term', 'Planned', 'Monitor')) {
        $phaseRows = @($Findings | Where-Object { $_.RoadmapPhase -eq $phaseName })
        if ($phaseRows.Count -eq 0) { continue }
        $lines.Add("### $phaseName") | Out-Null
        $lines.Add('') | Out-Null
        foreach ($finding in $phaseRows) {
            $lines.Add("- [$($finding.OwnerTeam)] $($finding.Finding)") | Out-Null
        }
        $lines.Add('') | Out-Null
    }

    $lines.Add('## Remediation Themes By Workstream') | Out-Null
    $lines.Add('') | Out-Null
    foreach ($ownerGroup in $ownerGroups) {
        $workstreamThemes = @($themeRows | Where-Object { $_.Workstream -eq $ownerGroup.Name })
        if ($workstreamThemes.Count -eq 0) { continue }
        $lines.Add("### $($ownerGroup.Name)") | Out-Null
        $lines.Add('') | Out-Null
        foreach ($theme in $workstreamThemes) {
            $lines.Add("- $($theme.Area) - What needs attention: $($theme.WhatNeedsAttention) Why it matters: $($theme.WhyItMatters) Recommended next step: $($theme.RecommendedNextStep) Business value: $($theme.BusinessValue)") | Out-Null
        }
        $lines.Add('') | Out-Null
    }

    return ($lines -join [Environment]::NewLine)
}

function New-CustomerRemediationReportHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Findings,
        [Parameter(Mandatory = $false)][object[]]$WorkstreamSummaries = @(),
        [Parameter(Mandatory = $true)][string]$TenantName,
        [Parameter(Mandatory = $true)][string]$AssessmentJsonPath,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt
    )

    $severityCounts = $Findings | Group-Object Severity | Sort-Object Name
    $ownerGroups = $Findings | Group-Object OwnerTeam | Sort-Object Count -Descending
    $themeRows = Get-CustomerWorkstreamThemes -Findings $Findings
    $topFindings = Get-TopFindings -Findings @($Findings | Where-Object { $_.Severity -in @('Critical', 'High') }) -Count 6
    if ($topFindings.Count -eq 0) { $topFindings = Get-TopFindings -Findings $Findings -Count 6 }

    $severitySummaryHtml = if ($severityCounts.Count -gt 0) {
        (($severityCounts | ForEach-Object {
            "<li><strong>{0}</strong>: {1} finding(s)</li>" -f (Convert-ToArrayaHtmlEncodedText $_.Name), $_.Count
        }) -join [Environment]::NewLine)
    } else {
        '<li><strong>Info</strong>: 0 finding(s)</li>'
    }

    $workstreamSummaryHtml = if ($WorkstreamSummaries.Count -gt 0) {
        (($WorkstreamSummaries | ForEach-Object {
            @"
<tr>
  <td>{0}</td>
  <td>{1}</td>
  <td>{2}</td>
  <td>{3}</td>
  <td>{4}</td>
  <td>{5}</td>
  <td>{6}</td>
  <td>{7}</td>
</tr>
"@ -f `
                (Convert-ToArrayaHtmlEncodedText $_.Severity),
                (Convert-ToArrayaHtmlEncodedText $_.Workstream),
                (Convert-ToArrayaHtmlEncodedText $_.Area),
                $_.OpenFindings,
                $_.CriticalCount,
                $_.WarningCount,
                $_.InfoCount,
                (Convert-ToArrayaHtmlFragment $_.TopSignals)
        }) -join [Environment]::NewLine)
    } else {
        '<tr><td colspan="8">No workstream summary rows were generated.</td></tr>'
    }

    $topPriorityHtml = if ($topFindings.Count -gt 0) {
        (($topFindings | ForEach-Object {
            @"
<article class="risk-card">
  <div class="risk-meta">
    <span class="badge severity-{0}">{1}</span>
    <span class="badge phase">{2}</span>
    <span class="muted">{3}</span>
  </div>
  <h3>{4} - {5}</h3>
  <p><strong>Why flagged:</strong> {6}</p>
  <p><strong>Recommended action:</strong> {7}</p>
  <p><strong>Success criteria:</strong> {8}</p>
</article>
"@ -f `
                ([string]$_.Severity).ToLowerInvariant(),
                (Convert-ToArrayaHtmlEncodedText $_.Severity),
                (Convert-ToArrayaHtmlEncodedText $_.PriorityBand),
                (Convert-ToArrayaHtmlEncodedText $_.OwnerTeam),
                (Convert-ToArrayaHtmlEncodedText $_.RuleId),
                (Convert-ToArrayaHtmlEncodedText $_.Area),
                (Convert-ToArrayaHtmlFragment $_.WhyFlagged),
                (Convert-ToArrayaHtmlFragment $_.Recommendation),
                (Convert-ToArrayaHtmlFragment $_.TargetValue)
        }) -join [Environment]::NewLine)
    } else {
        '<p>No high-priority findings were generated.</p>'
    }

    $phaseSectionsHtml = @()
    foreach ($phaseName in @('Immediate', 'Near Term', 'Planned', 'Monitor')) {
        $phaseRows = @($Findings | Where-Object { $_.RoadmapPhase -eq $phaseName })
        if ($phaseRows.Count -eq 0) { continue }

        $phaseItems = (($phaseRows | ForEach-Object {
            "<li><strong>{0}</strong> <span class=""muted"">[{1}]</span><br/>{2}</li>" -f `
                (Convert-ToArrayaHtmlEncodedText $_.Finding),
                (Convert-ToArrayaHtmlEncodedText $_.OwnerTeam),
                (Convert-ToArrayaHtmlFragment $_.Recommendation)
        }) -join [Environment]::NewLine)

        $phaseSectionsHtml += @"
<section class="phase-block">
  <h3>$([System.Net.WebUtility]::HtmlEncode($phaseName))</h3>
  <ul>
$phaseItems
  </ul>
</section>
"@
    }
    if ($phaseSectionsHtml.Count -eq 0) {
        $phaseSectionsHtml = @('<p>No phased remediation items were generated.</p>')
    }

    $workstreamThemeHtml = if ($ownerGroups.Count -gt 0) {
        (($ownerGroups | ForEach-Object {
            $ownerGroup = $_
            $rows = @($themeRows | Where-Object { $_.Workstream -eq $ownerGroup.Name })
            if ($rows.Count -eq 0) { return $null }
            $items = (($rows | ForEach-Object {
                @"
<li class="theme-item">
  <h4>{0}</h4>
  <p><strong>What needs attention:</strong> {1}</p>
  <p><strong>Why it matters:</strong> {2}</p>
  <p><strong>Recommended next step:</strong> {3}</p>
  <p><strong>Business value:</strong> {4}</p>
</li>
"@ -f `
                    (Convert-ToArrayaHtmlEncodedText $_.Area),
                    (Convert-ToArrayaHtmlFragment $_.WhatNeedsAttention),
                    (Convert-ToArrayaHtmlFragment $_.WhyItMatters),
                    (Convert-ToArrayaHtmlFragment $_.RecommendedNextStep),
                    (Convert-ToArrayaHtmlFragment $_.BusinessValue)
            }) -join [Environment]::NewLine)
            @"
<section class="theme-block">
  <h3>{0}</h3>
  <ul class="theme-list">
{1}
  </ul>
</section>
"@ -f (Convert-ToArrayaHtmlEncodedText $ownerGroup.Name), $items
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [Environment]::NewLine)
    } else {
        '<p>No workstream themes were generated.</p>'
    }

    $primaryWorkstreamsText = if ($ownerGroups.Count -gt 0) {
        ((@($ownerGroups | Select-Object -First 4 | ForEach-Object { '{0} ({1})' -f $_.Name, $_.Count }) -join '; '))
    } else {
        'No dominant workstreams identified'
    }

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>$([System.Net.WebUtility]::HtmlEncode($TenantName)) Customer Remediation Report | Arraya Solutions</title>
  <style>
    :root {
      --bg: #f3f0e8;
      --surface: #fffdf8;
      --ink: #1f2933;
      --muted: #52606d;
      --line: #d9d2c3;
      --accent: #12343b;
      --accent-strong: #1e5160;
      --accent-soft: #e4eef1;
      --critical: #a61b29;
      --high: #b45309;
      --medium: #0f766e;
      --low: #2563eb;
      --info: #475569;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Segoe UI", "Aptos", Tahoma, sans-serif;
      color: var(--ink);
      background: linear-gradient(180deg, #e9e3d6 0%, var(--bg) 24%, #f7f4ed 100%);
      line-height: 1.55;
    }
    main {
      max-width: 1180px;
      margin: 0 auto;
      padding: 32px 20px 56px;
    }
    header, section {
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 24px;
      margin-bottom: 20px;
      box-shadow: 0 12px 30px rgba(31, 41, 51, 0.06);
    }
    h1, h2, h3 { margin-top: 0; color: #173042; }
    h1 { font-size: 2rem; margin-bottom: 0.3rem; }
    h2 { font-size: 1.25rem; margin-bottom: 0.9rem; }
    .lede { color: var(--muted); max-width: 72ch; }
    .meta { color: var(--muted); font-size: 0.95rem; margin-top: 12px; }
    .summary-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 16px;
      margin-top: 18px;
    }
    .summary-card {
      background: #fbf8f1;
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 16px;
    }
    .summary-card ul { margin: 0; padding-left: 18px; }
    .eyebrow {
      color: var(--accent);
      font: 700 11px/1.4 "Segoe UI", "Aptos", Tahoma, sans-serif;
      letter-spacing: 0.14em;
      text-transform: uppercase;
      margin-bottom: 10px;
    }
    .risk-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 16px;
    }
    .risk-card, .phase-block, .theme-block {
      background: #fbf8f1;
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 16px;
    }
    .risk-meta {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-bottom: 10px;
    }
    .badge {
      display: inline-block;
      border-radius: 999px;
      padding: 4px 10px;
      font-size: 0.82rem;
      font-weight: 600;
    }
    .severity-critical { background: rgba(166, 27, 41, 0.12); color: var(--critical); }
    .severity-high { background: rgba(180, 83, 9, 0.12); color: var(--high); }
    .severity-medium { background: rgba(15, 118, 110, 0.12); color: var(--medium); }
    .severity-low { background: rgba(37, 99, 235, 0.12); color: var(--low); }
    .severity-info { background: rgba(71, 85, 105, 0.12); color: var(--info); }
    .phase { background: var(--accent-soft); color: var(--accent); }
    .muted { color: var(--muted); }
    .theme-list {
      list-style: none;
      padding: 0;
      margin: 0;
      display: grid;
      gap: 12px;
    }
    .theme-item {
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 14px;
      background: #fffdf8;
    }
    .theme-item h4 {
      margin: 0 0 10px;
      color: var(--accent-strong);
    }
    .theme-item p {
      margin: 6px 0 0;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.95rem;
    }
    th, td {
      text-align: left;
      vertical-align: top;
      border-bottom: 1px solid var(--line);
      padding: 10px 8px;
    }
    th {
      background: #f3ede2;
      color: #173042;
      position: sticky;
      top: 0;
    }
    .table-wrap { overflow-x: auto; }
    ul { margin-top: 0.4rem; margin-bottom: 0; }
    @media print {
      body { background: #ffffff; }
      header, section, .risk-card, .phase-block, .theme-block, .summary-card {
        box-shadow: none;
        break-inside: avoid;
      }
      th { position: static; }
    }
  </style>
</head>
<body>
  <main>
    <header>
      <div class="eyebrow">Prepared by Arraya Solutions</div>
      <h1>$([System.Net.WebUtility]::HtmlEncode($TenantName)) Microsoft 365 Remediation Report</h1>
      <p class="lede">This report turns the tenant assessment into a remediation-first customer deliverable. It focuses on business-impacting risks, governance gaps, and the workstreams needed to improve tenant posture.</p>
      <div class="meta">
        <div><strong>Generated:</strong> $([System.Net.WebUtility]::HtmlEncode($GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss')))</div>
        <div><strong>Assessment Snapshot:</strong> $([System.Net.WebUtility]::HtmlEncode($AssessmentJsonPath))</div>
        <div><strong>Prepared by:</strong> Arraya Solutions</div>
        <div><strong>Artifact:</strong> Customer Remediation Report</div>
      </div>
    </header>

    <section>
      <h2>Executive Summary</h2>
      <div class="summary-grid">
        <div class="summary-card">
          <h3>Severity Mix</h3>
          <ul>
$severitySummaryHtml
          </ul>
        </div>
        <div class="summary-card">
          <h3>Primary Workstreams</h3>
          <p>$([System.Net.WebUtility]::HtmlEncode($primaryWorkstreamsText))</p>
        </div>
      </div>
    </section>

    <section>
      <h2>Workstream Summary</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Severity</th>
              <th>Workstream</th>
              <th>Area</th>
              <th>Open Findings</th>
              <th>Critical</th>
              <th>Warning</th>
              <th>Info</th>
              <th>Top Signals</th>
            </tr>
          </thead>
          <tbody>
$workstreamSummaryHtml
          </tbody>
        </table>
      </div>
    </section>

    <section>
      <h2>Top Priority Risks</h2>
      <div class="risk-grid">
$topPriorityHtml
      </div>
    </section>

    <section>
      <h2>Phased Roadmap</h2>
      $($phaseSectionsHtml -join [Environment]::NewLine)
    </section>

    <section>
      <h2>Remediation Themes By Workstream</h2>
      $workstreamThemeHtml
    </section>
    <section>
      <h2>Prepared By</h2>
      <p>Arraya Solutions prepared this remediation summary from the collected Microsoft 365 assessment snapshot for customer review and execution planning.</p>
    </section>
  </main>
</body>
</html>
"@
}

function New-EngineerActionPack {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Findings,
        [Parameter(Mandatory = $false)][object[]]$WorkstreamSummaries = @(),
        [Parameter(Mandatory = $true)][string]$TenantName,
        [Parameter(Mandatory = $true)][string]$AssessmentJsonPath,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt,
        [Parameter(Mandatory = $true)][string]$JsonOutPath,
        [Parameter(Mandatory = $false)][string]$CsvOutPath,
        [Parameter(Mandatory = $true)][string]$SnippetOutPath,
        [Parameter(Mandatory = $true)][string]$SupportFolderPath
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("# $TenantName Engineer Action Pack") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("Generated: $($GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss'))") | Out-Null
    $lines.Add("Assessment Snapshot: $AssessmentJsonPath") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('## Technical Summary') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("- Total findings: $($Findings.Count)") | Out-Null
    $lines.Add("- Immediate: $((@($Findings | Where-Object { $_.RoadmapPhase -eq 'Immediate' })).Count)") | Out-Null
    $lines.Add("- Near Term: $((@($Findings | Where-Object { $_.RoadmapPhase -eq 'Near Term' })).Count)") | Out-Null
    $lines.Add("- Planned: $((@($Findings | Where-Object { $_.RoadmapPhase -eq 'Planned' })).Count)") | Out-Null
    $lines.Add("- Monitor: $((@($Findings | Where-Object { $_.RoadmapPhase -eq 'Monitor' })).Count)") | Out-Null
    $lines.Add('') | Out-Null
    if ($WorkstreamSummaries.Count -gt 0) {
        $lines.Add('## Workstream Summary') | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add('| Severity | Workstream | Area | Open Findings | Top Signals |') | Out-Null
        $lines.Add('|---|---|---|---|---|') | Out-Null
        foreach ($summary in $WorkstreamSummaries) {
            $lines.Add("| $($summary.Severity) | $($summary.Workstream) | $($summary.Area) | $($summary.OpenFindings) | $(Convert-ToArrayaMarkdownText $summary.TopSignals) |") | Out-Null
        }
        $lines.Add('') | Out-Null
    }
    $lines.Add('## Dependencies And Prerequisites') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('- Validate tenant admin roles, Graph scopes, Exchange connectivity, and any pilot exclusions before enforcement changes.') | Out-Null
    $lines.Add('- Use the support artifacts for backlog import, validation context, and change-record attachment as needed.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('## Normalized Findings') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('| Severity | Phase | Workstream | Rule | Finding | Why Flagged | Technical Remediation | Evidence | Evidence Location | Success Criteria |') | Out-Null
    $lines.Add('|---|---|---|---|---|---|---|---|---|---|') | Out-Null
    foreach ($finding in $Findings) {
        $lines.Add("| $($finding.Severity) | $($finding.RoadmapPhase) | $($finding.OwnerTeam) | $($finding.RuleId) | $(Convert-ToArrayaMarkdownText $finding.Finding) | $(Convert-ToArrayaMarkdownText $finding.WhyFlagged) | $(Convert-ToArrayaMarkdownText $finding.TechnicalRemediation) | $(Convert-ToArrayaMarkdownText $finding.CurrentValue) | $(Convert-ToArrayaMarkdownText $finding.EvidenceLocation) | $(Convert-ToArrayaMarkdownText $finding.TargetValue) |") | Out-Null
    }
    $lines.Add('') | Out-Null
    $lines.Add('## Validation Steps') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('- Confirm each recommendation against the latest tenant state before making changes.') | Out-Null
    $lines.Add('- Pilot security and access-policy changes with a limited scope before broad enforcement.') | Out-Null
    $lines.Add('- Re-run `M365Collect` and `Improve` after remediation milestones to measure delta and retire closed findings.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('## Command References') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add(('- Support folder: `{0}`' -f $SupportFolderPath)) | Out-Null
    $lines.Add(('- Improvement plan JSON: `{0}`' -f $JsonOutPath)) | Out-Null
    if (-not [string]::IsNullOrWhiteSpace($CsvOutPath)) {
        $lines.Add(('- CSV output: `{0}`' -f $CsvOutPath)) | Out-Null
    }
    $lines.Add(('- Remediation snippets: `{0}`' -f $SnippetOutPath)) | Out-Null
    $lines.Add('') | Out-Null
    return ($lines -join [Environment]::NewLine)
}

$snapshotContext = Import-ArrayaTenantSnapshotContext -Path $AssessmentJsonPath -Purpose ImprovementPlan
$AssessmentJsonPath = $snapshotContext.Path
$tenantData = $snapshotContext.LegacyData
$snapshotDerived = $snapshotContext.Derived
$snapshotDiagnostics = $snapshotContext.Diagnostics

$outputContext = Resolve-ArrayaSnapshotOutputContext -PrimaryInputPath $AssessmentJsonPath -OutputFolder $OutputFolder -OutputPrefix $OutputPrefix
$OutputFolder = $outputContext.OutputFolder
$OutputPrefix = $outputContext.OutputPrefix

$generatedAt = Get-Date
$findingStore = @{}
$workstreamSummaries = New-Object System.Collections.Generic.List[object]
$derivedFindings = @()
$script:ArrayaImproveHeuristicFindings = New-Object System.Collections.Generic.List[object]

foreach ($derivedKey in @('BestPracticeFindings', 'Findings')) {
    if (-not $snapshotDerived.ContainsKey($derivedKey)) { continue }
    foreach ($row in (Convert-ArrayaObjectToArray $snapshotDerived[$derivedKey])) {
        $normalizedFinding = New-DerivedFindingFromRow -Row $row -SourceLabel "Derived/$derivedKey"
        if ($null -eq $normalizedFinding) { continue }

        Add-ImprovementFinding -Store $findingStore -Finding $normalizedFinding | Out-Null
        $derivedFindings += $normalizedFinding
    }
}

if ($snapshotDerived.ContainsKey('BestPractices')) {
    foreach ($row in (Convert-ArrayaObjectToArray $snapshotDerived['BestPractices'])) {
        $summaryRecord = New-WorkstreamSummaryFromRow -Row $row
        if ($summaryRecord) {
            $workstreamSummaries.Add($summaryRecord) | Out-Null
        }
    }
}

if ($snapshotDiagnostics.Count -gt 0) {
    $diagWarningCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $snapshotDiagnostics -Names @('WarningCount'))
    $diagErrorCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $snapshotDiagnostics -Names @('ErrorCount'))
    if ($null -ne $diagWarningCount -or $null -ne $diagErrorCount) {
        Add-HeuristicFinding `
            -Store $findingStore `
            -RuleId 'DIAG-001' `
            -Area 'Collection Diagnostics' `
            -Category 'Collection Diagnostics' `
            -Severity 'Info' `
            -Finding 'Snapshot contains collection diagnostics that should be reviewed before using the assessment as a remediation baseline.' `
            -Recommendation 'Review collector warnings and errors, then re-run collection if any important workload data was incomplete.' `
            -CurrentValue "Warnings=$diagWarningCount; Errors=$diagErrorCount" `
            -TargetValue 'Warnings=0; Errors=0' `
            -Source 'Summary/Diagnostics' `
            -RelatedWorksheet 'Diagnostics' `
            -RelatedSection 'Collection Diagnostics'
    }
}

$secureScoreRows = Get-ImprovementPlanDataset -DataRoot $tenantData -Names @('SecuritySecureScore', 'SecureScore') -GraphUri '/v1.0/security/secureScores?$top=10' -Activity 'Secure Score fallback' -UseGraphFallback:$graphFallbackEnabled
$caPolicies = Get-ImprovementPlanDataset -DataRoot $tenantData -Names @('ConditionalAccessPolicies', 'ConditionalAccess') -GraphUri '/v1.0/identity/conditionalAccess/policies' -Activity 'Conditional Access policy fallback' -UseGraphFallback:$graphFallbackEnabled
$adminRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('AllOffice365Admins', 'Office365Admins', 'Admins'))
$domainRows = Get-ImprovementPlanDataset -DataRoot $tenantData -Names @('Domains') -GraphUri '/v1.0/domains' -Activity 'Domain fallback' -UseGraphFallback:$graphFallbackEnabled
$licenseRows = Get-ImprovementPlanDataset -DataRoot $tenantData -Names @('LicenseSKUs', 'Licenses') -GraphUri '/v1.0/subscribedSkus' -Activity 'Subscribed SKU fallback' -UseGraphFallback:$graphFallbackEnabled
$deviceRows = Get-ImprovementPlanDataset -DataRoot $tenantData -Names @('DeviceDetails', 'Devices') -GraphUri '/v1.0/devices?$select=id,displayName,approximateLastSignInDateTime' -Activity 'Device fallback' -UseGraphFallback:$graphFallbackEnabled
$userRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('Users', 'UserFullDetails'))
$mailboxRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('AllMailboxes', 'MailboxFullDetails'))
$connectorRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('MailFlowConnectors'))
$remoteDomainRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('RemoteDomains'))
$publicFolderRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('PublicFolderDetails'))
$sharePointRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('SharePoint'))
$oneDriveRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('OneDrive'))
$teamRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('AllTeams'))
$unifiedGroupRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('UnifiedGroups', 'EntraIDGroups'))
$smtpRelaySummary = Get-ArrayaObjectValue -Object $tenantData -Names @('SMTPRelaySummary')
$authConfig = Get-ArrayaObjectValue -Object $tenantData -Names @('AuthenticationConfig')
$mfaSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('MfaRegistrationSummary', 'MFARegistrationSummary', 'MfaRegistration', 'MFARegistration')
$ownershipSummary = Get-ArrayaObjectValue -Object $snapshotDerived -Names @('OwnershipGovernanceSummary')
$unmanagedObjects = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $snapshotDerived -Names @('UnmanagedObjects'))
$oneDriveOwnerMismatches = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $snapshotDerived -Names @('OneDriveOwnerMismatches'))
$tenantInfoSummary = Get-ArrayaObjectValue -Object $snapshotDerived -Names @('TenantInfoSummary')
$authSummary = Get-ArrayaObjectValue -Object $snapshotDerived -Names @('AuthenticationConfigSummary')
$mfaDerivedSummary = Get-ArrayaObjectValue -Object $snapshotDerived -Names @('MfaRegistrationSummary')
$conditionalAccessSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('ConditionalAccessPolicySummary')
$securityDefaultsPolicy = Get-ArrayaObjectValue -Object $tenantData -Names @('SecurityDefaultsPolicy')
$enterpriseApplications = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('EnterpriseApplications', 'AuthenticationSSOApplications'))
$enterpriseApplicationSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('EnterpriseApplicationSummary')
$guestSignInSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('GuestSignInSummary')
$privilegedAccessSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('PrivilegedAccessSummary')
$inboxRulesExternalForwarding = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('InboxRulesExternalForwarding'))
$inboxRuleForwardingSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('InboxRuleForwardingSummary')
$forwardingPolicySummary = Get-ArrayaObjectValue -Object $tenantData -Names @('ForwardingPolicySummary')
$sharedMailboxGovernanceSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('SharedMailboxGovernanceSummary')
$sharePointSharingSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('SharePointSharingSummary')
$collaborationActivitySummary = Get-ArrayaObjectValue -Object $tenantData -Names @('CollaborationActivitySummary')
$deviceManagementSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('DeviceManagementSummary')

$snapshotMetrics = Get-ArrayaTenantSnapshotMetricSet `
    -SecureScoreRows $secureScoreRows `
    -ConditionalAccessRows $caPolicies `
    -AdminRows $adminRows `
    -DomainRows $domainRows `
    -LicenseRows $licenseRows `
    -DeviceRows $deviceRows `
    -StaleDeviceDays $StaleDeviceDays

if (-not (Test-DerivedCoverage -Tags @('secure score') -DerivedFindings $derivedFindings)) {
    if ($snapshotMetrics.SecureScoreRows.Count -gt 0 -and $null -ne $snapshotMetrics.SecureScorePercent) {
        $pct = $snapshotMetrics.SecureScorePercent
        if ($pct -lt 40) {
            Add-HeuristicFinding -Store $findingStore -RuleId 'SEC-001' -Area 'Secure Score' -Category 'Security Configuration' -Severity 'Critical' -Finding 'Secure Score is below 40%.' -Recommendation 'Prioritize the highest-impact Secure Score improvement actions across identity, email, collaboration, and endpoint controls.' -CurrentValue "$pct%" -TargetValue '>= 60%' -RelatedWorksheet 'SecuritySecureScore' -RelatedSection 'Secure Score'
        } elseif ($pct -lt 60) {
            Add-HeuristicFinding -Store $findingStore -RuleId 'SEC-001' -Area 'Secure Score' -Category 'Security Configuration' -Severity 'High' -Finding 'Secure Score is below 60%.' -Recommendation 'Implement the highest-value Secure Score improvements and sequence them into the near-term roadmap.' -CurrentValue "$pct%" -TargetValue '>= 60%' -RelatedWorksheet 'SecuritySecureScore' -RelatedSection 'Secure Score'
        }
    }
}

if (-not (Test-DerivedCoverage -Tags @('conditional access', 'mfa') -DerivedFindings $derivedFindings)) {
    if ($snapshotMetrics.ConditionalAccessPolicyCount -eq 0) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-001' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Critical' -Finding 'No Conditional Access policies were found.' -Recommendation 'Create baseline Conditional Access policies for admins, MFA, legacy authentication blocking, and risk-based access control.' -CurrentValue '0 policies' -TargetValue '>= 3 enabled baseline policies' -RelatedWorksheet 'ConditionalAccessPolicies' -RelatedSection 'Conditional Access'
    } elseif ($snapshotMetrics.EnabledConditionalAccessCount -eq 0) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-001' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'High' -Finding 'Conditional Access policies exist but none appear enabled.' -Recommendation 'Progress validated policies from report-only to enforced mode in a staged rollout.' -CurrentValue "$($snapshotMetrics.ConditionalAccessPolicyCount) policies; 0 enabled" -TargetValue 'Baseline policies enforced' -RelatedWorksheet 'ConditionalAccessPolicies' -RelatedSection 'Conditional Access'
    }
}

if (-not (Test-DerivedCoverage -Tags @('mfa registration') -DerivedFindings $derivedFindings)) {
    $registeredUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaSummary -Names @('RegisteredUsers', 'RegisteredUserCount', 'MfaRegisteredUsers'))
    $totalUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaSummary -Names @('TotalUsers', 'UserCount', 'TotalUserCount'))
    $registrationPct = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaSummary -Names @('RegistrationPercent', 'RegisteredPercent', 'MfaRegistrationPercent'))
    if ($null -eq $registrationPct -and $null -ne $registeredUsers -and $null -ne $totalUsers -and $totalUsers -gt 0) {
        $registrationPct = [math]::Round(($registeredUsers / $totalUsers) * 100, 2)
    }
    if ($null -ne $registrationPct) {
        if ($registrationPct -lt 70) {
            Add-HeuristicFinding -Store $findingStore -RuleId 'MFA-001' -Area 'Identity Governance' -Category 'Identity Governance' -Severity 'High' -Finding 'MFA registration appears low.' -Recommendation 'Run a registration campaign, validate methods, and enforce MFA through Conditional Access.' -CurrentValue "$registrationPct%" -TargetValue '>= 90%' -RelatedWorksheet 'MfaRegistrationSummary' -RelatedSection 'MFA Registration'
        } elseif ($registrationPct -lt 90) {
            Add-HeuristicFinding -Store $findingStore -RuleId 'MFA-001' -Area 'Identity Governance' -Category 'Identity Governance' -Severity 'Medium' -Finding 'MFA registration is below the target adoption threshold.' -Recommendation 'Drive remaining users through registration completion and verify policy scope gaps.' -CurrentValue "$registrationPct%" -TargetValue '>= 90%' -RelatedWorksheet 'MfaRegistrationSummary' -RelatedSection 'MFA Registration'
        }
    }
}

if (-not (Test-DerivedCoverage -Tags @('global administrator', 'privileged') -DerivedFindings $derivedFindings) -and $snapshotMetrics.GlobalAdminCount -gt $MaxGlobalAdmins) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'ADMIN-001' -Area 'Identity Governance' -Category 'Privileged Access' -Severity 'High' -Finding 'Global administrator count exceeds the recommended threshold.' -Recommendation 'Reduce standing Global Administrator access and move privileged tasks to least privilege or eligible access where possible.' -CurrentValue "$($snapshotMetrics.GlobalAdminCount) accounts" -TargetValue "<= $MaxGlobalAdmins accounts" -RelatedWorksheet 'Admins' -RelatedSection 'Privileged Access'
}

if (-not (Test-DerivedCoverage -Tags @('domain') -DerivedFindings $derivedFindings) -and $snapshotMetrics.UnverifiedDomainCount -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'DOMAIN-001' -Area 'Exchange Hygiene' -Category 'Domain Hygiene' -Severity 'Medium' -Finding 'Unverified domains were detected.' -Recommendation 'Verify required domains and retire any stale or unused domain entries.' -CurrentValue "$($snapshotMetrics.UnverifiedDomainCount) unverified" -TargetValue '0 unverified' -RelatedWorksheet 'Domains' -RelatedSection 'Domains'
}

if (-not (Test-DerivedCoverage -Tags @('license', 'sku') -DerivedFindings $derivedFindings) -and $snapshotMetrics.HighUtilizationSkus.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'LIC-001' -Area 'Licensing Optimization' -Category 'Licensing Optimization' -Severity 'Medium' -Finding 'One or more license SKUs are near or at capacity.' -Recommendation 'Reclaim unused assignments, reconcile dormant users, and procure additional seats for constrained SKUs.' -CurrentValue ($snapshotMetrics.HighUtilizationSkus -join '; ') -TargetValue '< 95% utilization per paid SKU' -RelatedWorksheet 'LicenseSKUs' -RelatedSection 'Licensing'
}

if (-not (Test-DerivedCoverage -Tags @('stale device', 'device') -DerivedFindings $derivedFindings) -and $snapshotMetrics.DeviceCount -gt 0 -and $null -ne $snapshotMetrics.StaleDevicePercent -and $snapshotMetrics.StaleDevicePercent -ge 15) {
    $severity = if ($snapshotMetrics.StaleDevicePercent -ge 30) { 'High' } else { 'Medium' }
    Add-HeuristicFinding -Store $findingStore -RuleId 'DEV-001' -Area 'Device / Endpoint Posture' -Category 'Device / Endpoint Posture' -Severity $severity -Finding 'Device stale rate is above the target threshold.' -Recommendation 'Review stale device inventory, remove or disable obsolete devices, and tighten lifecycle cleanup controls.' -CurrentValue "$($snapshotMetrics.StaleDevicePercent)% stale ($($snapshotMetrics.StaleDeviceCount)/$($snapshotMetrics.DeviceCount))" -TargetValue '< 15% stale' -RelatedWorksheet 'DeviceDetails' -RelatedSection 'Devices'
}

$enabledCloudOnlyGlobalAdmins = @(
    $adminRows |
        Where-Object {
            $roleText = [string](Get-ArrayaObjectValue -Object $_ -Names @('Role', 'DirectoryRole', 'AdminRole', 'RoleName'))
            $enabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))
            $synced = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('OnPremisesSyncEnabled', 'DirSyncEnabled', 'IsSynced'))
            $roleText -match 'Global Administrator' -and $enabled -ne $false -and $synced -ne $true
        }
)
if ($enabledCloudOnlyGlobalAdmins.Count -eq 0 -and -not (Test-DerivedCoverage -Tags @('emergency access', 'break-glass') -DerivedFindings $derivedFindings)) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'ID-001' -Area 'Identity Governance' -Category 'Identity Governance' -Severity 'High' -Finding 'No enabled cloud-only break-glass Global Administrator accounts were identified.' -Recommendation 'Maintain at least two cloud-only emergency access accounts and validate that Conditional Access exclusions are documented and tested.' -CurrentValue '0 cloud-only emergency access accounts' -TargetValue '>= 2 documented break-glass accounts' -RelatedWorksheet 'Admins' -RelatedSection 'Identity & Admins'
}

$guestUsers = @($userRows | Where-Object { ([string](Get-ArrayaObjectValue -Object $_ -Names @('UserType'))).ToLowerInvariant() -eq 'guest' })
$inactiveGuestUsers = @(
    $guestUsers |
        Where-Object {
            $lastSignIn = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastSignInDateTime', 'LastSuccessfulSignInDateTime', 'SignInActivityLastSignInDateTime'))
            $enabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))
            ($enabled -ne $false) -and ($null -eq $lastSignIn -or $lastSignIn -lt (Get-Date).AddDays(-90))
        }
)
if ($guestUsers.Count -gt 25 -or $inactiveGuestUsers.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'ID-002' -Area 'Identity Governance' -Category 'Identity Governance' -Severity $(if ($inactiveGuestUsers.Count -gt 10) { 'High' } else { 'Medium' }) -Finding 'Guest lifecycle hygiene needs review.' -Recommendation 'Review inactive guests, validate sponsor ownership, and tighten guest access governance for external collaboration.' -CurrentValue "$($guestUsers.Count) guests; $($inactiveGuestUsers.Count) inactive/stale" -TargetValue 'Inactive guests reviewed and guest access kept within approved collaboration scope' -Source 'Hybrid/Users' -RelatedWorksheet 'Users' -RelatedSection 'Guests'
}

$stalePrivilegedAdmins = @(
    $adminRows |
        Where-Object {
            $enabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))
            $lastSignIn = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastSignInDateTime', 'LastSuccessfulSignInDateTime'))
            ($enabled -ne $false) -and $lastSignIn -and $lastSignIn -lt (Get-Date).AddDays(-90)
        }
)
if ($stalePrivilegedAdmins.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'ID-003' -Area 'Identity Governance' -Category 'Identity Governance' -Severity 'Medium' -Finding 'Enabled privileged accounts appear stale based on sign-in recency.' -Recommendation 'Validate whether stale privileged accounts are still needed and remove or downgrade unused standing admin roles.' -CurrentValue "$($stalePrivilegedAdmins.Count) stale privileged account(s)" -TargetValue '0 stale privileged accounts' -Source 'Hybrid/Admins' -RelatedWorksheet 'Admins' -RelatedSection 'Privileged Access'
}

$authConfigSummaryRecord = if ($authSummary) { Get-ArrayaObjectValue -Object $authSummary -Names @('Summary') } else { $null }
$passwordlessMethods = @(Convert-ToArrayaStringList (Get-ArrayaObjectValue -Object $authConfigSummaryRecord -Names @('PasswordlessMethods')))
if ($passwordlessMethods.Count -eq 0) {
    $passwordlessMethods = @(Convert-ToArrayaStringList (Get-ArrayaObjectValue -Object $authConfig -Names @('PasswordlessMethods')))
}
$registrationPercent = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $(if ($mfaDerivedSummary) { Get-ArrayaObjectValue -Object $mfaDerivedSummary -Names @('Summary') } else { $mfaSummary }) -Names @('RegistrationPercent', 'RegisteredPercent'))
if ($null -ne $registrationPercent -and $registrationPercent -ge 60 -and $passwordlessMethods.Count -eq 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'ID-004' -Area 'Identity Governance' -Category 'Identity Governance' -Severity 'Low' -Finding 'MFA adoption is present, but passwordless methods do not appear enabled in the assessment summary.' -Recommendation 'Review passwordless readiness and decide whether Windows Hello for Business, FIDO2, or Microsoft Authenticator passwordless should be added to the roadmap.' -CurrentValue 'Passwordless methods not detected' -TargetValue 'At least one approved passwordless method evaluated or enabled' -Source 'Hybrid/AuthenticationConfig' -RelatedWorksheet 'AuthenticationConfig' -RelatedSection 'Authentication'
}

$reportOnlyPolicies = @($caPolicies | Where-Object { ([string](Get-ArrayaObjectValue -Object $_ -Names @('State'))).ToLowerInvariant() -match 'report' })
if ($reportOnlyPolicies.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'CA-002' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Medium' -Finding 'Conditional Access policies remain in report-only mode.' -Recommendation 'Review report-only results, document expected impact, and progress approved policies to enabled enforcement.' -CurrentValue "$($reportOnlyPolicies.Count) report-only policy/policies" -TargetValue 'Only intentional pilot policies left in report-only mode' -Source 'Hybrid/ConditionalAccessPolicies' -RelatedWorksheet 'ConditionalAccessPolicies' -RelatedSection 'Conditional Access'
}

$caNameText = @($caPolicies | ForEach-Object { [string](Get-ArrayaObjectValue -Object $_ -Names @('DisplayName', 'Name')) })
$caNamesJoined = ($caNameText -join ' || ').ToLowerInvariant()
if ($caPolicies.Count -gt 0) {
    if ($caNamesJoined -notmatch 'legacy') {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-003' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'High' -Finding 'No Conditional Access policy was obviously scoped to block legacy authentication.' -Recommendation 'Add or validate a policy that explicitly blocks legacy authentication client types.' -CurrentValue 'Legacy authentication protection not clearly identified' -TargetValue 'Legacy authentication blocked by an enabled policy' -Source 'Hybrid/ConditionalAccessPolicies' -RelatedWorksheet 'ConditionalAccessPolicies' -RelatedSection 'Conditional Access'
    }
    if ($caNamesJoined -notmatch 'risk') {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-004' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Medium' -Finding 'Risk-based Conditional Access coverage was not obvious from policy names.' -Recommendation 'Review user risk and sign-in risk protection for high-value users and high-risk access scenarios.' -CurrentValue 'Risk-based coverage not clearly identified' -TargetValue 'Risk-based Conditional Access reviewed and implemented where appropriate' -Source 'Hybrid/ConditionalAccessPolicies' -RelatedWorksheet 'ConditionalAccessPolicies' -RelatedSection 'Conditional Access'
    }
    if ($caNamesJoined -notmatch 'admin|privileged') {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-005' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Medium' -Finding 'Dedicated privileged-role Conditional Access coverage was not obvious from policy names.' -Recommendation 'Validate dedicated admin or privileged role Conditional Access coverage, including stronger controls and tighter exclusions.' -CurrentValue 'Privileged-role coverage not clearly identified' -TargetValue 'Privileged roles protected by dedicated Conditional Access policies' -Source 'Hybrid/ConditionalAccessPolicies' -RelatedWorksheet 'ConditionalAccessPolicies' -RelatedSection 'Conditional Access'
    }
    if ($caNamesJoined -notmatch 'guest|external') {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-006' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Low' -Finding 'Guest or external-user Conditional Access coverage was not obvious from policy names.' -Recommendation 'Review external-user controls and confirm guest access is covered by dedicated Conditional Access policies where required.' -CurrentValue 'Guest coverage not clearly identified' -TargetValue 'Guest/external access reviewed and controlled' -Source 'Hybrid/ConditionalAccessPolicies' -RelatedWorksheet 'ConditionalAccessPolicies' -RelatedSection 'Conditional Access'
    }
}

$excludedPolicies = @(
    $caPolicies |
        Where-Object {
            $json = ($_ | ConvertTo-Json -Depth 6 -Compress)
            $json -match '"exclude'
        }
)
if ($excludedPolicies.Count -ge 3) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'CA-007' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Medium' -Finding 'Several Conditional Access policies appear to use exclusions.' -Recommendation 'Review exclusions for break-glass necessity, business justification, and drift from the intended protection model.' -CurrentValue "$($excludedPolicies.Count) policies with exclusions detected" -TargetValue 'Exclusions minimized and documented' -Source 'Hybrid/ConditionalAccessPolicies' -RelatedWorksheet 'ConditionalAccessPolicies' -RelatedSection 'Conditional Access'
}

$forwardingMailboxes = @(
    $mailboxRows |
        Where-Object {
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DeliverToMailboxAndForward'))) -eq $true -or
            -not [string]::IsNullOrWhiteSpace([string](Get-ArrayaObjectValue -Object $_ -Names @('ForwardingAddress', 'ForwardingSmtpAddress')))
        }
)
if ($forwardingMailboxes.Count -gt 0) {
    $forwardingPolicyEvidence = Get-ExchangeForwardingPolicyEvidence -ForwardingPolicySummary $forwardingPolicySummary -RemoteDomains $remoteDomainRows
    $ex001CurrentValue = "$($forwardingMailboxes.Count) mailbox(es) with forwarding configured"
    if (-not [string]::IsNullOrWhiteSpace($forwardingPolicyEvidence)) {
        $ex001CurrentValue = "$ex001CurrentValue; $forwardingPolicyEvidence"
    }
    Add-HeuristicFinding -Store $findingStore -RuleId 'EX-001' -Area 'Exchange Hygiene' -Category 'Exchange Hygiene' -Severity 'High' -Finding 'Mailbox forwarding is enabled for one or more mailboxes.' -Recommendation 'Review forwarding use cases, validate external destinations, and compare the mailbox list against the tenant outbound auto-forwarding posture and remote-domain settings. Remove or formally approve forwarding configurations that are still required, and document that this reflects tenant-level policy context rather than a per-mailbox authorization verdict.' -CurrentValue $ex001CurrentValue -TargetValue 'All mailbox forwarding configurations reviewed and either approved or removed, with tenant forwarding policy aligned to the approved baseline' -Source 'Hybrid/Exchange' -RelatedWorksheet 'AllMailboxes' -RelatedSection 'Mailbox Forwarding' -WhyFlagged 'Mailbox forwarding was detected on one or more objects, and the assessment also reviewed tenant-level outbound auto-forwarding posture and remote-domain forwarding settings to show whether external forwarding appears broadly permitted.' -ExampleAction 'Example: export the forwarding mailbox list, review the current outbound spam filter auto-forwarding modes and remote domains with AutoForwardEnabled, then disable or document any forwarding path that is no longer required.'
}

$riskyConnectors = @(
    $connectorRows |
        Where-Object {
            $enabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('Enabled'))
            $requireTls = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('RequireTls'))
            $internal = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('TreatMessagesAsInternal'))
            $senderIps = Convert-ToArrayaStringList (Get-ArrayaObjectValue -Object $_ -Names @('SenderIPAddresses'))
            ($enabled -ne $false) -and ($requireTls -eq $false -or $internal -eq $true -or $senderIps.Count -gt 0)
        }
)
if ($riskyConnectors.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'EX-002' -Area 'Exchange Hygiene' -Category 'Exchange Hygiene' -Severity 'Medium' -Finding 'Mail flow connectors with higher-trust or relay-like characteristics were detected.' -Recommendation 'Review connector purpose, tighten IP restrictions, require TLS where appropriate, and remove unneeded trusted relay behavior.' -CurrentValue "$($riskyConnectors.Count) connector(s) flagged for review" -TargetValue 'Only documented secure connectors remain' -Source 'Hybrid/Exchange' -RelatedWorksheet 'MailFlowConnectors' -RelatedSection 'Mail Flow Connectors'
}

$smtpAuthEnabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $smtpRelaySummary -Names @('SMTPAuthEnabled'))
$smtpAuthUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $smtpRelaySummary -Names @('SMTPAuthUsers'))
$smtpClientAuthDisabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $smtpRelaySummary -Names @('SmtpClientAuthenticationDisabled'))
if ($smtpAuthEnabled -eq $true -or ($null -ne $smtpAuthUsers -and $smtpAuthUsers -gt 0) -or $smtpClientAuthDisabled -eq $false) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'EX-003' -Area 'Exchange Hygiene' -Category 'Exchange Hygiene' -Severity 'Medium' -Finding 'SMTP relay or client-auth style mail submission still appears enabled.' -Recommendation 'Validate relay dependencies, move applications to modern authenticated patterns where possible, and retire broad SMTP relay exposure.' -CurrentValue "SMTPAuthEnabled=$(Convert-ToArrayaDisplayText $smtpAuthEnabled); SMTPAuthUsers=$(Convert-ToArrayaDisplayText $smtpAuthUsers)" -TargetValue 'SMTP relay minimized and documented' -Source 'Hybrid/Exchange' -RelatedWorksheet 'SMTPRelaySummary' -RelatedSection 'SMTP Relay'
}

if ($publicFolderRows.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'EX-004' -Area 'Exchange Hygiene' -Category 'Exchange Hygiene' -Severity 'Medium' -Finding 'Public folders are still present in the tenant.' -Recommendation 'Confirm whether public folders remain in scope, then migrate or retire them as part of the messaging roadmap.' -CurrentValue "$($publicFolderRows.Count) public folder object(s)" -TargetValue 'Public folder strategy documented and remediation path assigned' -Source 'Hybrid/Exchange' -RelatedWorksheet 'PublicFolderDetails' -RelatedSection 'Public Folders'
}

$sharedMailboxRows = @($mailboxRows | Where-Object { ([string](Get-ArrayaObjectValue -Object $_ -Names @('RecipientTypeDetails', 'MailboxType'))).ToLowerInvariant() -match 'shared' })
$ownerlessSharedMailboxes = @($sharedMailboxRows | Where-Object { [string]::IsNullOrWhiteSpace([string](Get-ArrayaObjectValue -Object $_ -Names @('Manager', 'Owner', 'PrimarySmtpAddressOwner'))) })
$oversizedSharedMailboxes = @(
    $sharedMailboxRows |
        Where-Object {
            $size = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('TotalItemSizeGB', 'StorageUsedGB', 'SizeGB'))
            $null -ne $size -and $size -gt 50
        }
)
if ($ownerlessSharedMailboxes.Count -gt 0 -or $oversizedSharedMailboxes.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'EX-005' -Area 'Exchange Hygiene' -Category 'Exchange Hygiene' -Severity 'Low' -Finding 'Shared mailbox governance requires review for ownership and growth control.' -Recommendation 'Assign clear mailbox owners, validate business purpose, and monitor oversized shared mailboxes for cleanup or archival.' -CurrentValue "$($ownerlessSharedMailboxes.Count) ownerless; $($oversizedSharedMailboxes.Count) oversized" -TargetValue 'Shared mailbox ownership documented and oversized growth reviewed' -Source 'Hybrid/Exchange' -RelatedWorksheet 'AllMailboxes' -RelatedSection 'Shared Mailboxes'
}

$ownershipSummaryRecord = if ($ownershipSummary) { Get-ArrayaObjectValue -Object $ownershipSummary -Names @('Summary') } else { $null }
$missingOwnerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ownershipSummaryRecord -Names @('MissingOwnerCount'))
$oneDriveMismatchCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ownershipSummaryRecord -Names @('OneDriveOwnerMismatchCount'))
if (($null -ne $missingOwnerCount -and $missingOwnerCount -gt 0) -or ($null -ne $oneDriveMismatchCount -and $oneDriveMismatchCount -gt 0) -or $unmanagedObjects.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'COL-001' -Area 'SharePoint / OneDrive Governance' -Category 'SharePoint / OneDrive Governance' -Severity 'High' -Finding 'Ownerless or unmanaged collaboration objects were identified.' -Recommendation 'Assign accountable owners to collaboration spaces and review unmanaged objects as part of collaboration governance.' -CurrentValue "$($missingOwnerCount) missing owners; $($unmanagedObjects.Count) unmanaged objects; $($oneDriveMismatchCount) OneDrive mismatches" -TargetValue 'Owner assignment completed for in-scope collaboration assets' -Source 'Hybrid/Ownership' -RelatedWorksheet 'OwnershipGovernanceSummary' -RelatedSection 'Ownership Governance'
}

if ($oneDriveOwnerMismatches.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'COL-002' -Area 'SharePoint / OneDrive Governance' -Category 'SharePoint / OneDrive Governance' -Severity 'Medium' -Finding 'OneDrive ownership mismatches were detected.' -Recommendation 'Review OneDrive ownership mapping and resolve orphaned or incorrectly attributed personal sites.' -CurrentValue "$($oneDriveOwnerMismatches.Count) mismatch(es)" -TargetValue '0 unresolved OneDrive ownership mismatches' -Source 'Hybrid/Ownership' -RelatedWorksheet 'OneDriveOwnerMismatches' -RelatedSection 'Ownership Governance'
}

$staleSharePointSites = @(
    $sharePointRows |
        Where-Object {
            $lastModified = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastContentModifiedDate'))
            $lastModified -and $lastModified -lt (Get-Date).AddDays(-180)
        }
)
$staleOneDrives = @(
    $oneDriveRows |
        Where-Object {
            $lastModified = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastContentModifiedDate'))
            $lastModified -and $lastModified -lt (Get-Date).AddDays(-180)
        }
)
if ($staleSharePointSites.Count -gt 0 -or $staleOneDrives.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'COL-003' -Area 'SharePoint / OneDrive Governance' -Category 'SharePoint / OneDrive Governance' -Severity 'Low' -Finding 'Inactive SharePoint or OneDrive locations were identified.' -Recommendation 'Review stale sites and personal storage locations for retention, ownership, and lifecycle actions.' -CurrentValue "$($staleSharePointSites.Count) stale SharePoint site(s); $($staleOneDrives.Count) stale OneDrive site(s)" -TargetValue 'Inactive locations reviewed and dispositioned' -Source 'Hybrid/Collaboration' -RelatedWorksheet 'SharePoint' -RelatedSection 'Storage and Activity'
}

$storageHotSpots = @(
    @(@($sharePointRows) + @($oneDriveRows)) |
        Where-Object {
            $storageGb = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('StorageUsedGB', 'TotalItemSizeGB', 'StorageGB'))
            $null -ne $storageGb -and $storageGb -gt 1024
        }
)
if ($storageHotSpots.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'COL-004' -Area 'SharePoint / OneDrive Governance' -Category 'SharePoint / OneDrive Governance' -Severity 'Medium' -Finding 'Large collaboration storage hot spots were detected.' -Recommendation 'Review large sites for ownership, lifecycle, archival planning, and storage governance.' -CurrentValue "$($storageHotSpots.Count) location(s) over 1 TB" -TargetValue 'Storage hot spots reviewed and governed' -Source 'Hybrid/Collaboration' -RelatedWorksheet 'SharePoint' -RelatedSection 'Storage and Activity'
}

$ownerlessTeams = @($teamRows | Where-Object {
    $ownerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OwnerCount'))
    $null -ne $ownerCount -and $ownerCount -eq 0
})
$guestHeavyTeams = @($teamRows | Where-Object {
    $guestCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('GuestCount', 'Guests'))
    $memberCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('MemberCount', 'Members'))
    $null -ne $guestCount -and $null -ne $memberCount -and $memberCount -gt 0 -and (($guestCount / $memberCount) -ge 0.4)
})
$channelSprawlTeams = @($teamRows | Where-Object {
    $privateCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('PrivateChannelCount'))
    $sharedCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('SharedChannelCount'))
    (($null -ne $privateCount -and $privateCount -ge 5) -or ($null -ne $sharedCount -and $sharedCount -ge 5))
})
if ($ownerlessTeams.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'TM-001' -Area 'Teams / M365 Groups Governance' -Category 'Teams / M365 Groups Governance' -Severity 'High' -Finding 'One or more Teams do not have any owners.' -Recommendation 'Assign at least one accountable owner to every Team and review the underlying group governance process.' -CurrentValue "$($ownerlessTeams.Count) ownerless Team(s)" -TargetValue '0 ownerless Teams' -Source 'Hybrid/Teams' -RelatedWorksheet 'AllTeams' -RelatedSection 'Teams Governance'
}
if ($guestHeavyTeams.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'TM-002' -Area 'Teams / M365 Groups Governance' -Category 'Teams / M365 Groups Governance' -Severity 'Low' -Finding 'Guest-heavy Teams were detected.' -Recommendation 'Review guest-heavy Teams for business justification, data sensitivity, and owner oversight.' -CurrentValue "$($guestHeavyTeams.Count) guest-heavy Team(s)" -TargetValue 'Guest-heavy Teams reviewed for governance fit' -Source 'Hybrid/Teams' -RelatedWorksheet 'AllTeams' -RelatedSection 'Teams Governance'
}
if ($channelSprawlTeams.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'TM-003' -Area 'Teams / M365 Groups Governance' -Category 'Teams / M365 Groups Governance' -Severity 'Medium' -Finding 'Private or shared channel sprawl was detected in Teams.' -Recommendation 'Review channel lifecycle, naming, access control, and eDiscovery coverage for Teams with high channel counts.' -CurrentValue "$($channelSprawlTeams.Count) Team(s) with elevated private/shared channel counts" -TargetValue 'High-channel Teams governed with clear lifecycle controls' -Source 'Hybrid/Teams' -RelatedWorksheet 'AllTeams' -RelatedSection 'Teams Governance'
}

$ownerlessGroups = @($unifiedGroupRows | Where-Object {
    $ownerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OwnerCount'))
    $null -ne $ownerCount -and $ownerCount -eq 0
})
if ($ownerlessGroups.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'TM-004' -Area 'Teams / M365 Groups Governance' -Category 'Teams / M365 Groups Governance' -Severity 'Medium' -Finding 'Microsoft 365 groups without owners were detected.' -Recommendation 'Review ownerless groups, assign accountable owners, and retire dormant groups that no longer serve a collaboration purpose.' -CurrentValue "$($ownerlessGroups.Count) ownerless group(s)" -TargetValue '0 ownerless groups' -Source 'Hybrid/Groups' -RelatedWorksheet 'UnifiedGroups' -RelatedSection 'Group Governance'
}

$inactiveLicensedUsers = @(
    $userRows |
        Where-Object {
            $assignedLicenses = Convert-ToArrayaStringList (Get-ArrayaObjectValue -Object $_ -Names @('AssignedLicensesFriendly', 'AssignedLicenses'))
            $enabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))
            $lastSignIn = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastSignInDateTime', 'LastSuccessfulSignInDateTime'))
            $assignedLicenses.Count -gt 0 -and (($enabled -eq $false) -or ($lastSignIn -and $lastSignIn -lt (Get-Date).AddDays(-90)))
        }
)
if ($inactiveLicensedUsers.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'LIC-002' -Area 'Licensing Optimization' -Category 'Licensing Optimization' -Severity 'Medium' -Finding 'Disabled or inactive users still appear to hold paid licenses.' -Recommendation 'Review paid license assignments for disabled or inactive users and reclaim unused SKUs where appropriate.' -CurrentValue "$($inactiveLicensedUsers.Count) inactive/disabled licensed user(s)" -TargetValue 'Inactive paid-license assignments reviewed and reclaimed' -Source 'Hybrid/Licensing' -RelatedWorksheet 'Users' -RelatedSection 'Licensing'
}

$nonCompliantDevices = @($deviceRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('IsCompliant', 'Compliant'))) -eq $false })
$unmanagedDevices = @(
    $deviceRows |
        Where-Object {
            $mdm = [string](Get-ArrayaObjectValue -Object $_ -Names @('MDMSolution', 'ManagementAgent', 'ManagedBy'))
            [string]::IsNullOrWhiteSpace($mdm)
        }
)
$intuneDevices = @(
    $deviceRows |
        Where-Object {
            $mdm = [string](Get-ArrayaObjectValue -Object $_ -Names @('MDMSolution', 'ManagementAgent', 'ManagedBy'))
            $mdm -match 'intune'
        }
)
if ($nonCompliantDevices.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'DEV-002' -Area 'Device / Endpoint Posture' -Category 'Device / Endpoint Posture' -Severity 'Medium' -Finding 'Non-compliant devices were identified in the assessment snapshot.' -Recommendation 'Review non-compliant devices, validate policy exceptions, and align remediation ownership with endpoint operations.' -CurrentValue "$($nonCompliantDevices.Count) non-compliant device(s)" -TargetValue 'Non-compliant devices reviewed and remediated' -Source 'Hybrid/Devices' -RelatedWorksheet 'DeviceDetails' -RelatedSection 'Devices'
}
if ($deviceRows.Count -gt 0 -and $unmanagedDevices.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'DEV-003' -Area 'Device / Endpoint Posture' -Category 'Device / Endpoint Posture' -Severity 'Medium' -Finding 'Unmanaged devices were detected in the tenant inventory.' -Recommendation 'Review unmanaged devices, determine whether they should be enrolled, and tighten enrollment expectations for corporate access.' -CurrentValue "$($unmanagedDevices.Count) unmanaged device(s) out of $($deviceRows.Count)" -TargetValue 'Managed-device coverage aligned to access policy requirements' -Source 'Hybrid/Devices' -RelatedWorksheet 'DeviceDetails' -RelatedSection 'Devices'
}
if ($deviceRows.Count -gt 0) {
    $intuneCoveragePct = [math]::Round(($intuneDevices.Count / $deviceRows.Count) * 100, 2)
    if ($intuneCoveragePct -lt 75) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'DEV-004' -Area 'Device / Endpoint Posture' -Category 'Device / Endpoint Posture' -Severity 'Low' -Finding 'Intune enrollment coverage appears limited relative to discovered devices.' -Recommendation 'Confirm the intended endpoint-management model and close enrollment gaps for in-scope devices.' -CurrentValue "$intuneCoveragePct% of discovered devices show Intune management" -TargetValue 'Intune coverage aligned to endpoint-management policy' -Source 'Hybrid/Devices' -RelatedWorksheet 'DeviceDetails' -RelatedSection 'Devices'
    }
}

$adminConsentWorkflowEnabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $authConfigSummaryRecord -Names @('AdminConsentWorkflowEnabled'))
if ($null -eq $adminConsentWorkflowEnabled) {
    $adminConsentWorkflowEnabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $authConfig -Names @('AdminConsentWorkflowEnabled'))
}
if ($adminConsentWorkflowEnabled -eq $false) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'SEC-002' -Area 'Security Configuration' -Category 'Security Configuration' -Severity 'Medium' -Finding 'Admin consent workflow appears disabled in the authentication configuration summary.' -Recommendation 'Review application consent governance and enable or replace the consent review workflow where the operating model requires it.' -CurrentValue 'Admin consent workflow disabled' -TargetValue 'Consent governance process documented and enforced' -Source 'Hybrid/AuthenticationConfig' -RelatedWorksheet 'AuthenticationConfig' -RelatedSection 'Consent Governance'
}

$permissionGrantPolicies = Convert-ToArrayaStringList (Get-ArrayaObjectValue -Object $authConfigSummaryRecord -Names @('PermissionGrantPolicies'))
if ($permissionGrantPolicies.Count -eq 0) {
    $permissionGrantPolicies = Convert-ToArrayaStringList (Get-ArrayaObjectValue -Object $authConfig -Names @('PermissionGrantPoliciesAssigned', 'PermissionGrantPolicies'))
}
if ($permissionGrantPolicies.Count -gt 0) {
    $grantPoliciesText = ($permissionGrantPolicies -join ', ').ToLowerInvariant()
    if ($grantPoliciesText -match 'managepermissiongrantsforownedresource|default') {
        Add-HeuristicFinding -Store $findingStore -RuleId 'SEC-003' -Area 'Security Configuration' -Category 'Security Configuration' -Severity 'Low' -Finding 'Permission grant policy settings may allow broader user consent than desired.' -Recommendation 'Review application consent policy assignments and align them to the tenant app-governance model.' -CurrentValue ($permissionGrantPolicies -join '; ') -TargetValue 'Consent grant policies aligned to approved governance model' -Source 'Hybrid/AuthenticationConfig' -RelatedWorksheet 'AuthenticationConfig' -RelatedSection 'Consent Governance'
    }
}

$guestSummaryRecord = if ($guestSignInSummary) { Get-ArrayaObjectValue -Object $guestSignInSummary -Names @('Summary') } else { $null }
$inactiveGuests90Days = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $guestSummaryRecord -Names @('InactiveGuests90Days'))
if ($null -ne $inactiveGuests90Days -and $inactiveGuests90Days -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'ID-005' -Area 'Identity Governance' -Category 'Identity Governance' -Severity $(if ($inactiveGuests90Days -ge 10) { 'High' } else { 'Medium' }) -Finding 'Inactive guest accounts were identified by the Tier B sign-in summary.' -Recommendation 'Review stale guest identities, validate sponsor ownership, and remove or disable guests that are no longer required.' -CurrentValue "$inactiveGuests90Days inactive guest account(s) over 90 days" -TargetValue 'Inactive guest accounts reviewed and dispositioned' -Source 'Summary/GuestSignIn' -RelatedWorksheet 'GuestSignInSummary' -RelatedSection 'Guest Access'
}

$privilegedSummaryRecord = if ($privilegedAccessSummary) { Get-ArrayaObjectValue -Object $privilegedAccessSummary -Names @('Summary') } else { $null }
$stalePrivilegedSummary = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $privilegedSummaryRecord -Names @('StalePrivilegedAccounts90Days'))
if ($null -ne $stalePrivilegedSummary -and $stalePrivilegedSummary -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'ID-006' -Area 'Identity Governance' -Category 'Identity Governance' -Severity 'Medium' -Finding 'Privileged identities with stale sign-in activity were identified by the Tier B summary.' -Recommendation 'Review stale privileged accounts, remove unused role assignments, and validate emergency access documentation.' -CurrentValue "$stalePrivilegedSummary stale privileged account(s) over 90 days" -TargetValue '0 stale privileged accounts' -Source 'Summary/PrivilegedAccess' -RelatedWorksheet 'PrivilegedAccessSummary' -RelatedSection 'Privileged Access'
}

$enterpriseAppSummaryRecord = if ($enterpriseApplicationSummary) { Get-ArrayaObjectValue -Object $enterpriseApplicationSummary -Names @('Summary') } else { $null }
$highPrivilegeApps = @($enterpriseApplications | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('HighPrivilegePermissionCount'))) -gt 0 })
if ($highPrivilegeApps.Count -gt 0) {
    $highPrivilegePermissionExamples = @(
        $highPrivilegeApps |
            ForEach-Object { Convert-ToArrayaStringList (Get-ArrayaObjectValue -Object $_ -Names @('HighPrivilegePermissions')) } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique -First 5
    )
    $appsWithApplicationPermissions = @(
        $highPrivilegeApps |
            Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('ApplicationPermissionCount'))) -gt 0 }
    ).Count
    $appsWithDelegatedGrants = @(
        $highPrivilegeApps |
            Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('DelegatedPermissionGrantCount'))) -gt 0 }
    ).Count
    $highPrivilegeCurrentValue = "{0} high-privilege enterprise application(s); {1} with application permissions; {2} with delegated grants" -f $highPrivilegeApps.Count, $appsWithApplicationPermissions, $appsWithDelegatedGrants
    if ($highPrivilegePermissionExamples.Count -gt 0) {
        $highPrivilegeCurrentValue = "{0}; matched permissions: {1}" -f $highPrivilegeCurrentValue, ($highPrivilegePermissionExamples -join ', ')
    }

    Add-HeuristicFinding -Store $findingStore -RuleId 'ID-007' -Area 'Identity Governance' -Category 'Identity Governance' -Severity 'High' -Finding 'Enterprise applications with broad or high-privilege permissions were detected.' -Recommendation 'Review high-privilege enterprise apps, validate business need, and tighten app consent governance for broad delegated or application permissions.' -CurrentValue $highPrivilegeCurrentValue -TargetValue 'High-privilege enterprise apps reviewed and governed' -Source 'Summary/EnterpriseApplications' -RelatedWorksheet 'EnterpriseApplications' -RelatedSection 'Enterprise Apps'
}

$securityDefaultsRecord = if ($securityDefaultsPolicy) { Get-ArrayaObjectValue -Object $securityDefaultsPolicy -Names @('Configuration', 'Summary') } else { $null }
$securityDefaultsEnabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $securityDefaultsRecord -Names @('IsEnabled'))
$caSummaryRecord = if ($conditionalAccessSummary) { Get-ArrayaObjectValue -Object $conditionalAccessSummary -Names @('Summary') } else { $null }
if ($securityDefaultsEnabled -eq $true -and $snapshotMetrics.ConditionalAccessPolicyCount -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'SEC-004' -Area 'Security Configuration' -Category 'Security Configuration' -Severity 'Low' -Finding 'Security Defaults appear enabled while Conditional Access policies are also present.' -Recommendation 'Confirm the intended identity-control model and retire overlapping baseline protections where Conditional Access has replaced Security Defaults.' -CurrentValue 'Security Defaults enabled with Conditional Access policies present' -TargetValue 'One clear baseline identity protection model in place' -Source 'Summary/SecurityDefaults' -RelatedWorksheet 'SecurityDefaultsPolicy' -RelatedSection 'Security Defaults'
}
if ($securityDefaultsEnabled -eq $false -and $caSummaryRecord) {
    $hasGuestCoverage = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $caSummaryRecord -Names @('HasGuestCoverage'))
    $hasPrivilegedCoverage = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $caSummaryRecord -Names @('HasPrivilegedRoleCoverage'))
    $hasCompliantDeviceRequirement = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $caSummaryRecord -Names @('HasCompliantDeviceRequirement'))
    $hasRiskCoverage = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $caSummaryRecord -Names @('HasRiskBasedCoverage'))
    $policiesWithExclusions = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $caSummaryRecord -Names @('PoliciesWithExclusions'))

    if ($hasGuestCoverage -eq $false) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-008' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Medium' -Finding 'Conditional Access summary does not show guest or external-user coverage.' -Recommendation 'Validate guest/external-user policy coverage now that Security Defaults are not the fallback protection model.' -CurrentValue 'Guest/external-user coverage not detected in summary' -TargetValue 'Guest/external-user coverage documented and enabled' -Source 'Summary/ConditionalAccess' -RelatedWorksheet 'ConditionalAccessPolicySummary' -RelatedSection 'Conditional Access'
    }
    if ($hasPrivilegedCoverage -eq $false) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-009' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Medium' -Finding 'Conditional Access summary does not show privileged-role coverage.' -Recommendation 'Add or validate dedicated privileged-role Conditional Access protection for administrative identities.' -CurrentValue 'Privileged-role coverage not detected in summary' -TargetValue 'Privileged-role coverage documented and enabled' -Source 'Summary/ConditionalAccess' -RelatedWorksheet 'ConditionalAccessPolicySummary' -RelatedSection 'Conditional Access'
    }
    if ($hasCompliantDeviceRequirement -eq $false) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-010' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Low' -Finding 'Conditional Access summary does not show a compliant-device requirement.' -Recommendation 'Review whether device-compliance requirements should be part of the access-control baseline for managed access scenarios.' -CurrentValue 'Compliant-device requirement not detected in summary' -TargetValue 'Compliant-device requirement reviewed and documented' -Source 'Summary/ConditionalAccess' -RelatedWorksheet 'ConditionalAccessPolicySummary' -RelatedSection 'Conditional Access'
    }
    if ($hasRiskCoverage -eq $false) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-011' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Low' -Finding 'Conditional Access summary does not show risk-based controls.' -Recommendation 'Review sign-in risk and user risk controls for higher-risk access scenarios.' -CurrentValue 'Risk-based Conditional Access not detected in summary' -TargetValue 'Risk-based Conditional Access reviewed and documented' -Source 'Summary/ConditionalAccess' -RelatedWorksheet 'ConditionalAccessPolicySummary' -RelatedSection 'Conditional Access'
    }
    if ($null -ne $policiesWithExclusions -and $policiesWithExclusions -ge 5) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-012' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Medium' -Finding 'Conditional Access summary shows a higher number of policies with exclusions.' -Recommendation 'Review exclusion sprawl and reduce broad bypass patterns where they are not required for emergency access or documented exceptions.' -CurrentValue "$policiesWithExclusions policy/policies with exclusions" -TargetValue 'Exclusions minimized and documented' -Source 'Summary/ConditionalAccess' -RelatedWorksheet 'ConditionalAccessPolicySummary' -RelatedSection 'Conditional Access'
    }
}

$externalInboxSummary = if ($inboxRuleForwardingSummary) { Get-ArrayaObjectValue -Object $inboxRuleForwardingSummary -Names @('Summary') } else { $null }
$externalForwardRuleCount = if ($inboxRulesExternalForwarding.Count -gt 0) { $inboxRulesExternalForwarding.Count } else { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $externalInboxSummary -Names @('ExternalForwardingRuleCount')) }
if ($null -ne $externalForwardRuleCount -and $externalForwardRuleCount -gt 0) {
    $forwardingPolicyEvidence = Get-ExchangeForwardingPolicyEvidence -ForwardingPolicySummary $forwardingPolicySummary -RemoteDomains $remoteDomainRows
    $ex006CurrentValue = "$externalForwardRuleCount external-forwarding inbox rule(s)"
    if (-not [string]::IsNullOrWhiteSpace($forwardingPolicyEvidence)) {
        $ex006CurrentValue = "$ex006CurrentValue; $forwardingPolicyEvidence"
    }
    Add-HeuristicFinding -Store $findingStore -RuleId 'EX-006' -Area 'Exchange Hygiene' -Category 'Exchange Hygiene' -Severity 'High' -Finding 'Inbox rules with external forwarding targets were detected.' -Recommendation 'Review inbox-rule forwarding behavior, confirm business justification, and compare the rule list against the tenant outbound auto-forwarding posture and remote-domain settings. Remove or formally approve external forwarding paths that must remain, and document that the policy context is tenant-level rather than a per-rule authorization verdict.' -CurrentValue $ex006CurrentValue -TargetValue 'All external inbox-rule forwarding paths reviewed and either approved or removed, with tenant forwarding policy aligned to the approved baseline' -Source 'Summary/InboxRules' -RelatedWorksheet 'InboxRulesExternalForwarding' -RelatedSection 'Inbox Rules' -WhyFlagged 'The assessment found inbox rules that forward to domains outside the tenant accepted-domain list, and it also inspected tenant-level outbound auto-forwarding posture and remote-domain settings to show whether external forwarding appears broadly permitted.' -ExampleAction 'Example: export the external inbox-rule list, review the target domains against approved forwarding use cases, and compare the result to hosted outbound spam filter auto-forwarding modes plus remote domains where AutoForwardEnabled is still true.'
}

$sharedMailboxSummaryRecord = if ($sharedMailboxGovernanceSummary) { Get-ArrayaObjectValue -Object $sharedMailboxGovernanceSummary -Names @('Summary') } else { $null }
$oversizedSharedMailboxCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $sharedMailboxSummaryRecord -Names @('OversizedSharedMailboxes'))
$ownerlessSharedMailboxCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $sharedMailboxSummaryRecord -Names @('SharedMailboxesWithoutOwnerSignal'))
if (($null -ne $oversizedSharedMailboxCount -and $oversizedSharedMailboxCount -gt 0) -or ($null -ne $ownerlessSharedMailboxCount -and $ownerlessSharedMailboxCount -gt 0)) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'EX-007' -Area 'Exchange Hygiene' -Category 'Exchange Hygiene' -Severity 'Low' -Finding 'Shared mailbox governance summary shows ownership or growth gaps.' -Recommendation 'Review oversized shared mailboxes and assign accountable ownership signals for long-lived shared mailbox workloads.' -CurrentValue "$oversizedSharedMailboxCount oversized; $ownerlessSharedMailboxCount lacking owner signal" -TargetValue 'Shared mailbox ownership and growth governance documented' -Source 'Summary/SharedMailboxGovernance' -RelatedWorksheet 'SharedMailboxGovernanceSummary' -RelatedSection 'Shared Mailboxes'
}

$sharePointSharingRecord = if ($sharePointSharingSummary) { Get-ArrayaObjectValue -Object $sharePointSharingSummary -Names @('Summary') } else { $null }
$tenantSharingCapability = [string](Get-ArrayaObjectValue -Object $sharePointSharingRecord -Names @('TenantSharingCapability'))
$defaultSharingLinkType = [string](Get-ArrayaObjectValue -Object $sharePointSharingRecord -Names @('DefaultSharingLinkType'))
if ($tenantSharingCapability -match 'ExternalUserAndGuestSharing|ExternalUserSharingOnly') {
    Add-HeuristicFinding -Store $findingStore -RuleId 'COL-005' -Area 'SharePoint / OneDrive Governance' -Category 'SharePoint / OneDrive Governance' -Severity 'Medium' -Finding 'SharePoint sharing is configured to allow external sharing at the tenant level.' -Recommendation 'Validate whether tenant-wide external sharing and default link settings match the intended collaboration governance model.' -CurrentValue "SharingCapability=$tenantSharingCapability; DefaultSharingLinkType=$defaultSharingLinkType" -TargetValue 'Tenant sharing posture documented and aligned to policy' -Source 'Summary/SharePointSharing' -RelatedWorksheet 'SharePointSharingSummary' -RelatedSection 'Sharing'
}
if ($defaultSharingLinkType -match 'AnonymousAccess') {
    Add-HeuristicFinding -Store $findingStore -RuleId 'COL-006' -Area 'SharePoint / OneDrive Governance' -Category 'SharePoint / OneDrive Governance' -Severity 'Medium' -Finding 'Anonymous links appear to remain the default sharing link type.' -Recommendation 'Review anonymous link defaults, external sharing use cases, and whether named-user links should be the default posture.' -CurrentValue "DefaultSharingLinkType=$defaultSharingLinkType" -TargetValue 'Default sharing link type aligned to policy' -Source 'Summary/SharePointSharing' -RelatedWorksheet 'SharePointSharingSummary' -RelatedSection 'Sharing'
}

$collaborationSummaryRecord = if ($collaborationActivitySummary) { Get-ArrayaObjectValue -Object $collaborationActivitySummary -Names @('Summary') } else { $null }
$dormantGroups = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $collaborationSummaryRecord -Names @('DormantGroups'))
if ($null -ne $dormantGroups -and $dormantGroups -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'TM-005' -Area 'Teams / M365 Groups Governance' -Category 'Teams / M365 Groups Governance' -Severity 'Low' -Finding 'Dormant Microsoft 365 groups were identified in the collaboration activity summary.' -Recommendation 'Review dormant groups for archival, ownership confirmation, or retirement as part of collaboration lifecycle governance.' -CurrentValue "$dormantGroups dormant group(s) observed in activity reporting" -TargetValue 'Dormant groups reviewed and lifecycle action assigned' -Source 'Summary/CollaborationActivity' -RelatedWorksheet 'CollaborationActivitySummary' -RelatedSection 'Group Activity'
}
$dormantTeams = @(
    $teamRows | Where-Object {
        $lastActivity = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastActivityDate'))
        $lastActivity -and $lastActivity -lt (Get-Date).AddDays(-90)
    }
)
if ($dormantTeams.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'TM-006' -Area 'Teams / M365 Groups Governance' -Category 'Teams / M365 Groups Governance' -Severity 'Low' -Finding 'Teams with older recorded activity dates were detected.' -Recommendation 'Review dormant Teams for archival, owner confirmation, and retention/lifecycle handling.' -CurrentValue "$($dormantTeams.Count) Team(s) with activity older than 90 days" -TargetValue 'Dormant Teams reviewed and lifecycle action assigned' -Source 'Hybrid/Teams' -RelatedWorksheet 'AllTeams' -RelatedSection 'Teams Governance'
}

$guestHeavyTeamsTierB = @(
    $teamRows | Where-Object {
        $guestCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('GuestCount'))
        $memberCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('MemberCount'))
        $null -ne $guestCount -and $null -ne $memberCount -and $memberCount -gt 0 -and (($guestCount / $memberCount) -ge 0.4)
    }
)
if ($guestHeavyTeamsTierB.Count -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'TM-007' -Area 'Teams / M365 Groups Governance' -Category 'Teams / M365 Groups Governance' -Severity 'Medium' -Finding 'Teams with a high guest-member ratio were detected.' -Recommendation 'Review guest-heavy Teams for data sensitivity, owner oversight, and whether guest access still matches the intended collaboration model.' -CurrentValue "$($guestHeavyTeamsTierB.Count) guest-heavy Team(s)" -TargetValue 'Guest-heavy Teams reviewed and governed' -Source 'Hybrid/Teams' -RelatedWorksheet 'AllTeams' -RelatedSection 'Teams Governance'
}

$deviceManagementSummaryRecord = if ($deviceManagementSummary) { Get-ArrayaObjectValue -Object $deviceManagementSummary -Names @('Summary') } else { $null }
$unmanagedDevicesSummary = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $deviceManagementSummaryRecord -Names @('UnmanagedDevices'))
$totalDevicesSummary = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $deviceManagementSummaryRecord -Names @('TotalDevices'))
$unsupportedOsDevices = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $deviceManagementSummaryRecord -Names @('UnsupportedOsDevices'))
if ($null -ne $unmanagedDevicesSummary -and $null -ne $totalDevicesSummary -and $totalDevicesSummary -gt 0) {
    $unmanagedRatio = [math]::Round(($unmanagedDevicesSummary / $totalDevicesSummary) * 100, 2)
    if ($unmanagedRatio -ge 20) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'DEV-005' -Area 'Device / Endpoint Posture' -Category 'Device / Endpoint Posture' -Severity 'Medium' -Finding 'The device-management summary shows a meaningful unmanaged-device population.' -Recommendation 'Confirm endpoint-management scope and close unmanaged-device gaps for devices expected to access protected resources.' -CurrentValue "$unmanagedRatio% unmanaged ($unmanagedDevicesSummary/$totalDevicesSummary)" -TargetValue 'Managed-device coverage aligned to policy' -Source 'Summary/DeviceManagement' -RelatedWorksheet 'DeviceManagementSummary' -RelatedSection 'Devices'
    }
}
if ($null -ne $unsupportedOsDevices -and $unsupportedOsDevices -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'DEV-006' -Area 'Device / Endpoint Posture' -Category 'Device / Endpoint Posture' -Severity 'Medium' -Finding 'Unsupported or older operating-system versions were detected in the device-management summary.' -Recommendation 'Review unsupported endpoint versions, align them to the supported-device policy, and retire or upgrade obsolete devices.' -CurrentValue "$unsupportedOsDevices unsupported device(s)" -TargetValue 'Unsupported device count reduced to approved exceptions only' -Source 'Summary/DeviceManagement' -RelatedWorksheet 'DeviceManagementSummary' -RelatedSection 'Devices'
}

if ($findingStore.Count -eq 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'GEN-000' -Area 'General' -Category 'General' -Severity 'Info' -Finding 'No actionable findings were generated from the detected snapshot fields.' -Recommendation 'Validate snapshot completeness and rerun the assessment in a fuller collection mode if you expected additional workloads.' -CurrentValue 'N/A' -TargetValue 'N/A'
}

$sortedFindings = Get-SortedFindings -Store $findingStore
$sortedFindingMap = @{}
foreach ($finding in @($sortedFindings)) {
    if (-not $finding -or -not $finding.PSObject -or -not ($finding.PSObject.Properties.Name -contains 'RuleId')) { continue }
    $ruleId = [string]$finding.RuleId
    if ([string]::IsNullOrWhiteSpace($ruleId)) { continue }
    $sortedFindingMap[$ruleId] = $finding
}
foreach ($finding in $script:ArrayaImproveHeuristicFindings.ToArray()) {
    if (-not $finding -or -not $finding.PSObject -or -not ($finding.PSObject.Properties.Name -contains 'RuleId')) { continue }
    $ruleId = [string]$finding.RuleId
    if ([string]::IsNullOrWhiteSpace($ruleId)) { continue }

    if (-not $sortedFindingMap.ContainsKey($ruleId)) {
        $sortedFindingMap[$ruleId] = $finding
        continue
    }

    $existing = $sortedFindingMap[$ruleId]
    $existingRank = Get-FindingSourceRank -Source $existing.Source
    $incomingRank = Get-FindingSourceRank -Source $finding.Source
    $existingSeverity = Get-SeverityWeight -Severity $existing.Severity
    $incomingSeverity = Get-SeverityWeight -Severity $finding.Severity

    if ($incomingRank -gt $existingRank -or ($incomingRank -eq $existingRank -and $incomingSeverity -gt $existingSeverity)) {
        $sortedFindingMap[$ruleId] = $finding
    }
}
$sortedFindings = @(
    $sortedFindingMap.Values |
        Sort-Object `
            @{ Expression = { Get-SeverityWeight -Severity $_.Severity }; Descending = $true }, `
            @{ Expression = { switch ($_.PriorityBand) { 'Immediate' { 1 } 'Near Term' { 2 } 'Planned' { 3 } default { 4 } } } }, `
            OwnerTeam, Area, RuleId
)
foreach ($finding in $sortedFindings) {
    if (-not ($finding.PSObject.Properties.Name -contains 'Severity')) { continue }
    if ([string]::IsNullOrWhiteSpace([string]$finding.Severity)) { $finding.Severity = 'Info' }
    if ([string]::IsNullOrWhiteSpace([string]$finding.PriorityBand)) { $finding.PriorityBand = 'Monitor' }
    if ([string]::IsNullOrWhiteSpace([string]$finding.RoadmapPhase)) { $finding.RoadmapPhase = 'Monitor' }
}
$sortedWorkstreamSummaries = Get-SortedWorkstreamSummaries -Summaries (Resolve-VisibleWorkstreamSummaries -Summaries $workstreamSummaries.ToArray() -Findings $sortedFindings)
$tenantName = Get-TenantDisplayName -TenantInfoSummary $(if ($tenantInfoSummary) { Get-ArrayaObjectValue -Object $tenantInfoSummary -Names @('Summary') } else { $null }) -LegacyData $tenantData -OutputPrefix $OutputPrefix

$supportFolder = Join-Path -Path $OutputFolder -ChildPath 'Support'
if (-not (Test-Path -Path $supportFolder)) {
    $null = New-Item -ItemType Directory -Path $supportFolder -Force
}

$jsonOutPath = Join-Path -Path $supportFolder -ChildPath "$OutputPrefix-ImprovementPlan.json"
$csvOutPath = if ($IncludeLegacyArtifacts) { Join-Path -Path $supportFolder -ChildPath "$OutputPrefix-ImprovementPlan.csv" } else { $null }
$mdOutPath = if ($IncludeLegacyArtifacts) { Join-Path -Path $supportFolder -ChildPath "$OutputPrefix-ImprovementPlan.md" } else { $null }
$customerHtmlOutPath = Join-Path -Path $OutputFolder -ChildPath "$OutputPrefix-CustomerRemediationReport.html"
$customerMdOutPath = if ($IncludeLegacyArtifacts) { Join-Path -Path $supportFolder -ChildPath "$OutputPrefix-CustomerRemediationReport.md" } else { $null }
$engineerMdOutPath = Join-Path -Path $OutputFolder -ChildPath "$OutputPrefix-EngineerActionPack.md"
$snippetOutPath = Join-Path -Path $supportFolder -ChildPath "$OutputPrefix-RemediationSnippets.ps1"

$deliverables = [ordered]@{
    ImprovementPlanJson       = $jsonOutPath
    CustomerRemediationReport = $customerHtmlOutPath
    EngineerActionPack        = $engineerMdOutPath
    RemediationSnippets       = $snippetOutPath
    SupportFolder             = $supportFolder
}
if ($IncludeLegacyArtifacts) {
    $deliverables['ImprovementPlanCsv'] = $csvOutPath
    $deliverables['ImprovementPlanMarkdown'] = $mdOutPath
    $deliverables['CustomerRemediationReportMarkdown'] = $customerMdOutPath
}

$payload = [PSCustomObject]@{
    GeneratedAt = $generatedAt.ToString('o')
    SourceFile  = (Resolve-Path -Path $AssessmentJsonPath).Path
    Thresholds  = [PSCustomObject]@{
        StaleDeviceDays = $StaleDeviceDays
        MaxGlobalAdmins = $MaxGlobalAdmins
    }
    Deliverables = [PSCustomObject]$deliverables
    WorkstreamSummaries = $sortedWorkstreamSummaries
    Findings           = $sortedFindings
}

$payload | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonOutPath -Encoding UTF8
if ($IncludeLegacyArtifacts) {
    $sortedFindings | Export-Csv -Path $csvOutPath -NoTypeInformation -Encoding UTF8

    $severityCounts = $sortedFindings | Group-Object Severity | Sort-Object Name
    $summaryLines = @(
        '# Microsoft 365 Tenant Improvement Plan',
        '',
        "Generated: $($generatedAt.ToString('yyyy-MM-dd HH:mm:ss'))",
        "Source: $AssessmentJsonPath",
        '',
        '## Severity Summary'
    )
    foreach ($group in $severityCounts) {
        $summaryLines += "- $($group.Name): $($group.Count)"
    }

    if ($sortedWorkstreamSummaries.Count -gt 0) {
        $summaryLines += @(
            '',
            '## Workstream Summary',
            '',
            '| Severity | Workstream | Area | Open Findings | Critical | Warning | Info | Top Signals |',
            '|---|---|---|---|---|---|---|---|'
        )
        foreach ($summary in $sortedWorkstreamSummaries) {
            $summaryLines += "| $($summary.Severity) | $($summary.Workstream) | $($summary.Area) | $($summary.OpenFindings) | $($summary.CriticalCount) | $($summary.WarningCount) | $($summary.InfoCount) | $(Convert-ToArrayaMarkdownText $summary.TopSignals) |"
        }
    }

    $summaryLines += @(
        '',
        '## Findings',
        '',
        '| Severity | Priority | Workstream | Rule | Finding | Why Flagged | Recommended Action | Success Criteria |',
        '|---|---|---|---|---|---|---|---|'
    )
    foreach ($finding in $sortedFindings) {
        $summaryLines += "| $($finding.Severity) | $($finding.PriorityBand) | $($finding.OwnerTeam) | $($finding.RuleId) | $(Convert-ToArrayaMarkdownText $finding.Finding) | $(Convert-ToArrayaMarkdownText $finding.WhyFlagged) | $(Convert-ToArrayaMarkdownText $finding.Recommendation) | $(Convert-ToArrayaMarkdownText $finding.TargetValue) |"
    }
    Set-Content -Path $mdOutPath -Value ($summaryLines -join [Environment]::NewLine) -Encoding UTF8
}

$customerReportHtml = New-CustomerRemediationReportHtml -Findings $sortedFindings -WorkstreamSummaries $sortedWorkstreamSummaries -TenantName $tenantName -AssessmentJsonPath $AssessmentJsonPath -GeneratedAt $generatedAt
Set-Content -Path $customerHtmlOutPath -Value $customerReportHtml -Encoding UTF8

if ($IncludeLegacyArtifacts) {
    $customerReportMarkdown = New-CustomerRemediationReport -Findings $sortedFindings -WorkstreamSummaries $sortedWorkstreamSummaries -TenantName $tenantName -AssessmentJsonPath $AssessmentJsonPath -GeneratedAt $generatedAt
    Set-Content -Path $customerMdOutPath -Value $customerReportMarkdown -Encoding UTF8
}

$engineerActionPackMarkdown = New-EngineerActionPack -Findings $sortedFindings -WorkstreamSummaries $sortedWorkstreamSummaries -TenantName $tenantName -AssessmentJsonPath $AssessmentJsonPath -GeneratedAt $generatedAt -JsonOutPath $jsonOutPath -CsvOutPath $csvOutPath -SnippetOutPath $snippetOutPath -SupportFolderPath $supportFolder
Set-Content -Path $engineerMdOutPath -Value $engineerActionPackMarkdown -Encoding UTF8

$snippetLibrary = @{
    'SEC-001' = @'
# Secure Score review
Connect-MgGraph -Scopes 'SecurityEvents.Read.All','SecurityEvents.ReadWrite.All'
Get-MgSecuritySecureScore -Top 1 | Format-List *
'@
    'CA-001' = @'
# Conditional Access baseline validation
Connect-MgGraph -Scopes 'Policy.Read.All','Policy.ReadWrite.ConditionalAccess'
Get-MgIdentityConditionalAccessPolicy | Select-Object DisplayName, State
'@
    'CA-002' = @'
# Conditional Access report-only review
Connect-MgGraph -Scopes 'Policy.Read.All'
Get-MgIdentityConditionalAccessPolicy |
    Where-Object { $_.State -match 'report' } |
    Select-Object DisplayName, State
'@
    'MFA-001' = @'
# MFA registration review
Connect-MgGraph -Scopes 'Reports.Read.All'
Get-MgReportAuthenticationMethodUserRegistrationDetail -All |
    Select-Object UserPrincipalName, IsMfaRegistered
'@
    'ADMIN-001' = @'
# Global Administrator review
Connect-MgGraph -Scopes 'RoleManagement.Read.Directory','Directory.Read.All'
$gaRole = Get-MgDirectoryRole -Filter "displayName eq 'Global Administrator'"
Get-MgDirectoryRoleMember -DirectoryRoleId $gaRole.Id -All
'@
    'ID-001' = @'
# Break-glass account review
Connect-MgGraph -Scopes 'Directory.Read.All','Policy.Read.All'
Get-MgUser -All -Property DisplayName,UserPrincipalName,OnPremisesSyncEnabled,AccountEnabled |
    Where-Object { $_.AccountEnabled -and -not $_.OnPremisesSyncEnabled }
'@
    'DOMAIN-001' = @'
# Domain verification review
Connect-MgGraph -Scopes 'Domain.Read.All'
Get-MgDomain | Select-Object Id, IsVerified, IsDefault
'@
    'LIC-001' = @'
# License utilization review
Connect-MgGraph -Scopes 'Organization.Read.All'
Get-MgSubscribedSku | Select-Object SkuPartNumber, ConsumedUnits, @{N='Enabled';E={$_.PrepaidUnits.Enabled}}
'@
    'LIC-002' = @'
# Inactive licensed user review
Connect-MgGraph -Scopes 'User.Read.All','Directory.Read.All'
Get-MgUser -All -Property DisplayName,AccountEnabled,AssignedLicenses,SignInActivity |
    Where-Object { $_.AssignedLicenses.Count -gt 0 }
'@
    'DEV-001' = @'
# Stale device review
Connect-MgGraph -Scopes 'Device.Read.All'
$cutoff = (Get-Date).AddDays(-180)
Get-MgDevice -All -Property Id,DisplayName,ApproximateLastSignInDateTime |
    Where-Object { $_.ApproximateLastSignInDateTime -and $_.ApproximateLastSignInDateTime -lt $cutoff } |
    Select-Object DisplayName, ApproximateLastSignInDateTime
'@
    'DEV-002' = @'
# Non-compliant device review
Connect-MgGraph -Scopes 'DeviceManagementManagedDevices.Read.All'
Get-MgDeviceManagementManagedDevice -All |
    Select-Object DeviceName, ComplianceState, OperatingSystem
'@
    'EX-001' = @'
# Mailbox forwarding review
Connect-ExchangeOnline
Get-EXOMailbox -ResultSize Unlimited -Properties DeliverToMailboxAndForward,ForwardingAddress,ForwardingSmtpAddress |
    Where-Object { $_.DeliverToMailboxAndForward -or $_.ForwardingAddress -or $_.ForwardingSmtpAddress } |
    Select-Object DisplayName, PrimarySmtpAddress, DeliverToMailboxAndForward, ForwardingAddress, ForwardingSmtpAddress
'@
    'EX-002' = @'
# Mail flow connector review
Connect-ExchangeOnline
Get-InboundConnector | Format-List Name,Enabled,RequireTls,SenderIPAddresses,TreatMessagesAsInternal
Get-OutboundConnector | Format-List Name,Enabled,ConnectorType,SmartHosts
'@
    'EX-004' = @'
# Public folder review
Connect-ExchangeOnline
Get-PublicFolder -Recurse -ResultSize Unlimited | Select-Object Name, Identity
'@
}

$presentRuleIds = New-Object System.Collections.Generic.HashSet[string]
foreach ($ruleId in ($sortedFindings.RuleId | Sort-Object -Unique)) {
    [void]$presentRuleIds.Add([string]$ruleId)
}
foreach ($finding in $sortedFindings) {
    $topicText = ('{0} {1} {2} {3}' -f $finding.Area, $finding.Category, $finding.Finding, $finding.Recommendation).ToLowerInvariant()
    if ($topicText -match 'secure score') { [void]$presentRuleIds.Add('SEC-001') }
    if ($topicText -match 'conditional access') { [void]$presentRuleIds.Add('CA-001') }
    if ($topicText -match 'report-only') { [void]$presentRuleIds.Add('CA-002') }
    if ($topicText -match 'mfa') { [void]$presentRuleIds.Add('MFA-001') }
    if ($topicText -match 'global admin|administrator|privileged') { [void]$presentRuleIds.Add('ADMIN-001') }
    if ($topicText -match 'break-glass|emergency access') { [void]$presentRuleIds.Add('ID-001') }
    if ($topicText -match 'domain') { [void]$presentRuleIds.Add('DOMAIN-001') }
    if ($topicText -match 'license|sku') { [void]$presentRuleIds.Add('LIC-001') }
    if ($topicText -match 'inactive.*license|disabled.*license') { [void]$presentRuleIds.Add('LIC-002') }
    if ($topicText -match 'device|stale device') { [void]$presentRuleIds.Add('DEV-001') }
    if ($topicText -match 'non-compliant') { [void]$presentRuleIds.Add('DEV-002') }
    if ($topicText -match 'forward') { [void]$presentRuleIds.Add('EX-001') }
    if ($topicText -match 'connector|relay') { [void]$presentRuleIds.Add('EX-002') }
    if ($topicText -match 'public folder') { [void]$presentRuleIds.Add('EX-004') }
}
$snippetBlocks = @(
    '<#',
    'Remediation snippets generated by New-M365TenantImprovementPlan.ps1.',
    'Review and test each command in a non-production scope before enforcing changes.',
    '#>',
    ''
)
foreach ($ruleId in ($presentRuleIds | Sort-Object)) {
    if ($snippetLibrary.ContainsKey($ruleId)) {
        $snippetBlocks += '##############################'
        $snippetBlocks += "# Rule: $ruleId"
        $snippetBlocks += '##############################'
        $snippetBlocks += $snippetLibrary[$ruleId].Trim()
        $snippetBlocks += ''
    }
}
Set-Content -Path $snippetOutPath -Value ($snippetBlocks -join [Environment]::NewLine) -Encoding UTF8

$outputSummary = [PSCustomObject]@{
    FindingsCount                 = $sortedFindings.Count
    SupportFolderPath             = $supportFolder
    JsonPath                      = $jsonOutPath
    CsvPath                       = $csvOutPath
    MarkdownPath                  = $mdOutPath
    CustomerRemediationReportPath = $customerHtmlOutPath
    CustomerRemediationReportMarkdownPath = $customerMdOutPath
    EngineerActionPackPath        = $engineerMdOutPath
    RemediationPs1Path            = $snippetOutPath
}

Write-Host 'Improvement plan generated.'
Write-Host "  Customer report : $customerHtmlOutPath" -ForegroundColor Green
Write-Host "  Engineer pack   : $engineerMdOutPath" -ForegroundColor Cyan
Write-Host "  Support folder  : $supportFolder" -ForegroundColor DarkGray

if ($PassThru) {
    $outputSummary
}
