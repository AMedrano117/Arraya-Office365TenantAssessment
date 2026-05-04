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
    [switch]$PassThru,
    [Parameter(Mandatory = $false)]
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
$graphFallbackEnabled = $UseGraphFallback.IsPresent
$graphFallbackUnavailableMessageShown = $false

$customerAssessmentDocxHelperPath = Join-Path -Path $PSScriptRoot -ChildPath 'Private\CustomerAssessmentDocx.ps1'
if (-not (Test-Path -Path $customerAssessmentDocxHelperPath -PathType Leaf)) {
    throw "Customer assessment DOCX helper script was not found: $customerAssessmentDocxHelperPath"
}
. $customerAssessmentDocxHelperPath

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
        Import-Module -Name $resolvedCommonManifestPath -Force -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop
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

function Convert-ToImprovementCollectionRows {
    [CmdletBinding()]
    param(
        [AllowNull()]$Value,
        [string[]]$MarkerNames = @()
    )

    if ($null -eq $Value) {
        return @()
    }

    $looksLikeRow = $false
    if ($MarkerNames.Count -gt 0) {
        $looksLikeRow = $null -ne (Get-ArrayaObjectValue -Object $Value -Names $MarkerNames)
    }

    if ($looksLikeRow) {
        if ($Value -is [System.Collections.IDictionary]) {
            return [pscustomobject]$Value
        }

        return $Value
    }

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($item in @($Value.Values)) {
            if ($null -eq $item) {
                continue
            }

            if ($MarkerNames.Count -gt 0 -and $null -eq (Get-ArrayaObjectValue -Object $item -Names $MarkerNames)) {
                continue
            }

            if ($item -is [System.Collections.IDictionary]) {
                [pscustomobject]$item
            }
            else {
                $item
            }
        }

        return
    }

    foreach ($item in (Convert-ArrayaObjectToArray $Value)) {
        if ($null -eq $item) {
            continue
        }

        if ($MarkerNames.Count -gt 0 -and $null -eq (Get-ArrayaObjectValue -Object $item -Names $MarkerNames)) {
            continue
        }

        if ($item -is [System.Collections.IDictionary]) {
            [pscustomobject]$item
        }
        else {
            $item
        }
    }
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

function Get-RoadmapPhaseRank {
    [CmdletBinding()]
    param([AllowNull()][string]$Phase)

    switch (([string]$Phase).Trim().ToLowerInvariant()) {
        'immediate' { return 1 }
        'near term' { return 2 }
        'near-term' { return 2 }
        'planned' { return 3 }
        'monitor' { return 4 }
        default { return 5 }
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
        '^ID-009$' {
            $override['WhyFlagged'] = 'The enterprise app inventory includes third-party applications with application permissions, which can access tenant data or workloads without a signed-in user context.'
            $override['Recommendation'] = 'Review each third-party app with application permissions, confirm the vendor, owner, and business purpose, and remove or reduce permissions that are no longer required. Require a documented approval path before new third-party application permissions are granted.'
            $override['TargetValue'] = 'Each third-party app with application permissions has a documented owner, approved purpose, and validated least-privilege scope'
        }
        '^ID-010$' {
            $override['WhyFlagged'] = 'The assessment found first-party app registrations without an owner signal, which increases the chance that applications become orphaned and permissions are not reviewed over time.'
            $override['Recommendation'] = 'Assign an accountable owner to each flagged first-party app registration, confirm its business purpose, and retire applications that no longer have an approved sponsor. Include backup ownership or stewardship in the operating model where appropriate.'
            $override['TargetValue'] = 'Each flagged first-party app registration has an accountable owner and approved business purpose'
        }
        '^ID-011$' {
            $override['WhyFlagged'] = 'The application inventory includes redirect URI patterns that warrant security review, such as localhost, wildcard, HTTP, or short-link style redirects.'
            $override['Recommendation'] = 'Review each flagged redirect URI, confirm whether it is still required, and remove or tighten patterns that expand the authentication attack surface. Validate that production applications use only approved redirect endpoints.'
            $override['TargetValue'] = 'Redirect URIs are reduced to approved, documented endpoints with review exceptions recorded'
        }
        '^ID-012$' {
            $override['WhyFlagged'] = 'The app inventory includes enterprise applications without a recent activity signal, which can indicate stale integrations, abandoned test apps, or permissions that are no longer required.'
            $override['Recommendation'] = 'Review each inactive application with the business owner, confirm whether it should remain deployed, and remove or disable applications that no longer support an approved scenario. Document exceptions for dormant but still-required integrations.'
            $override['TargetValue'] = 'Inactive enterprise apps are reviewed, dispositioned, and only approved exceptions remain'
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
            $override['Recommendation'] = 'Review every standing Global Administrator assignment, remove the role entirely from inactive administrators, move infrequent administrators to lower-privilege roles such as Global Reader where possible, and retain only the minimum approved permanent admins. Where full tenant-wide access is only occasionally required, use Global Administrator as an eligible role that requires approval before activation. Document break-glass coverage separately from routine admin access.'
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
        $override['Recommendation'] = 'Review each standing Global Administrator assignment, remove the role entirely from inactive or stale administrators, move infrequent admins to lower-privilege roles such as Global Reader where possible, and retain only the minimum approved permanent admins plus documented emergency access accounts. Validate whether Global Administrator should remain an eligible role that requires approval before activation instead of a standing assignment.'
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

function Get-ImprovementOutputPrefix {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$OutputPrefix
    )

    $normalized = if ([string]::IsNullOrWhiteSpace($OutputPrefix)) { 'assessment' } else { $OutputPrefix.Trim() }

    foreach ($suffix in @('-AssessmentSnapshot', '-Snapshot')) {
        if ($normalized.EndsWith($suffix, [System.StringComparison]::Ordinal)) {
            $normalized = $normalized.Substring(0, $normalized.Length - $suffix.Length)
            break
        }
    }

    $normalized = $normalized.Replace(' Tenant Discovery Report-', '-')

    $replacements = [ordered]@{
        '-TenantToTenantMigration_' = '-T2T_'
        '-SolutionsEngineer_'       = '-SE_'
        '-ExecutiveLevel_'          = '-Exec_'
        '-TenantToTenantMigration'  = '-T2T'
        '-SolutionsEngineer'        = '-SE'
        '-ExecutiveLevel'           = '-Exec'
    }

    foreach ($entry in $replacements.GetEnumerator()) {
        $normalized = $normalized.Replace([string]$entry.Key, [string]$entry.Value)
    }

    return $normalized
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
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Summaries,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Findings
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

function Get-ArrayaLeadSentence {
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    $normalized = Convert-ToArrayaDisplayText -Value $Text -Default ''
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    if ($normalized -match '^(.*?[\.!?])(?:\s|$)') {
        return $matches[1].Trim()
    }

    return $normalized.Trim()
}

function Get-CustomerThemeProfile {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Finding)

    $ownerTeam = [string]$Finding.OwnerTeam
    $area = [string]$Finding.Area
    $category = [string]$Finding.Category
    $ruleId = [string]$Finding.RuleId
    $combinedText = (@($ownerTeam, $area, $category, $ruleId, [string]$Finding.Finding) -join ' ').ToLowerInvariant()

    if ($ownerTeam -eq 'Collaboration' -or $combinedText -match 'ownership|steward|sharepoint|onedrive|team|group|workspace') {
        return [pscustomobject]@{
            Key         = 'collaboration-governance'
            Theme       = 'Collaboration ownership and lifecycle governance'
            ActionTitle = 'Establish accountable ownership for collaboration spaces'
        }
    }

    if ($ownerTeam -eq 'Identity' -or $combinedText -match 'global administrator|privileged|guest|mfa|conditional access|enterprise app|break-glass|passwordless|consent') {
        return [pscustomobject]@{
            Key         = 'identity-governance'
            Theme       = 'Identity and privileged access governance'
            ActionTitle = 'Reduce privileged access and strengthen identity controls'
        }
    }

    if ($ownerTeam -eq 'Endpoint' -or $combinedText -match 'device|endpoint|intune|compliance|managed') {
        return [pscustomobject]@{
            Key         = 'endpoint-posture'
            Theme       = 'Endpoint compliance and device control'
            ActionTitle = 'Improve device compliance and managed endpoint coverage'
        }
    }

    if ($combinedText -match 'spf|dkim|dmarc|spoof|dns|mail-auth|unverified domain|domain verification') {
        return [pscustomobject]@{
            Key         = 'domain-authentication'
            Theme       = 'Domain authentication and DNS posture'
            ActionTitle = 'Strengthen domain and anti-spoofing controls'
        }
    }

    if ($ownerTeam -eq 'Governance' -or $combinedText -match 'licens|sku|sync|ad connect|capacity') {
        return [pscustomobject]@{
            Key         = 'governance-capacity'
            Theme       = 'Licensing, capacity, and tenant governance'
            ActionTitle = 'Reconcile license capacity and tenant governance gaps'
        }
    }

    if ($ownerTeam -eq 'Messaging' -or $combinedText -match 'forward|inbox rule|mail flow|connector|smtp|relay') {
        return [pscustomobject]@{
            Key         = 'messaging-exposure'
            Theme       = 'External forwarding and mail flow exposure'
            ActionTitle = 'Review external forwarding and mail flow exposure'
        }
    }

    if ($ownerTeam -eq 'Security' -or $combinedText -match 'secure score|security|zero trust|risk-based') {
        return [pscustomobject]@{
            Key         = 'security-baseline'
            Theme       = 'Security baseline and zero-trust readiness'
            ActionTitle = 'Strengthen baseline security and access protections'
        }
    }

    return [pscustomobject]@{
        Key         = ('misc-{0}' -f (($ownerTeam -replace '[^a-zA-Z0-9]+', '-').Trim('-').ToLowerInvariant()))
        Theme       = $(if ([string]::IsNullOrWhiteSpace($ownerTeam)) { 'General tenant priorities' } else { "$ownerTeam priorities" })
        ActionTitle = $(if ([string]::IsNullOrWhiteSpace($ownerTeam)) { 'Review remaining tenant priorities' } else { "Address $ownerTeam priorities" })
    }
}

function Get-CustomerExecutiveThemes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Findings,
        [Parameter(Mandatory = $false)][int]$Count = 5
    )

    $priorityRows = @($Findings | Where-Object { $_.Severity -in @('Critical', 'High', 'Medium') })
    if ($priorityRows.Count -eq 0) { $priorityRows = @($Findings) }

    $grouped = @(
        $priorityRows |
            Group-Object { (Get-CustomerThemeProfile -Finding $_).Key } |
            Sort-Object `
                @{ Expression = { ($_.Group | ForEach-Object { Get-SeverityWeight -Severity $_.Severity } | Measure-Object -Maximum).Maximum }; Descending = $true }, `
                @{ Expression = { $_.Count }; Descending = $true }, `
                Name
    )

    $themes = New-Object System.Collections.Generic.List[object]
    foreach ($group in $grouped | Select-Object -First $Count) {
        $orderedRows = @(
            $group.Group |
                Sort-Object `
                    @{ Expression = { Get-SeverityWeight -Severity $_.Severity }; Descending = $true }, `
                    @{ Expression = { switch ($_.PriorityBand) { 'Immediate' { 1 } 'Near Term' { 2 } 'Planned' { 3 } default { 4 } } } }, `
                    RuleId
        )
        if ($orderedRows.Count -eq 0) { continue }

        $leadFinding = $orderedRows[0]
        $profile = Get-CustomerThemeProfile -Finding $leadFinding
        $areas = @(
            $orderedRows |
                ForEach-Object { [string]$_.Area } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -Unique |
                Select-Object -First 3
        )
        $workstreams = @(
            $orderedRows |
                ForEach-Object { [string]$_.OwnerTeam } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -Unique |
                Select-Object -First 2
        )

        $whatThisMeans = if ($areas.Count -gt 0) {
            'The assessment identified {0} related findings concentrated across {1}.' -f $orderedRows.Count, (Join-ArrayaReadableList -Items $areas)
        } else {
            'The assessment identified {0} related findings in this priority grouping.' -f $orderedRows.Count
        }

        $themes.Add([pscustomobject]@{
            Theme               = $profile.Theme
            Severity            = $leadFinding.Severity
            PriorityBand        = $leadFinding.PriorityBand
            WorkstreamLabel     = $(if ($workstreams.Count -gt 0) { $workstreams -join ' / ' } else { [string]$leadFinding.OwnerTeam })
            FindingCount        = $orderedRows.Count
            WhatThisMeans       = $whatThisMeans
            StandoutReason      = Get-CustomerStandoutSentence -LeadFinding $leadFinding -FindingCount $orderedRows.Count -Areas $areas
            ExampleText         = Get-CustomerFindingExampleText -Finding $leadFinding -MaxItems 2
            WhyItMatters        = $(if ([string]::IsNullOrWhiteSpace([string](Convert-ToArrayaSentenceFragment -Text $leadFinding.WhyFlagged))) { 'This indicates that the observed control pattern is not isolated and is affecting a core part of the tenant.' } else { 'This indicates that ' + (Convert-ToArrayaSentenceFragment -Text $leadFinding.WhyFlagged) + '.' })
            RecommendedNextStep = Get-ArrayaLeadSentence -Text $leadFinding.Recommendation
            CustomerValue       = $(if ([string]::IsNullOrWhiteSpace([string](Convert-ToArrayaOutcomeSentence -Text $leadFinding.BusinessValue))) { 'Left unaddressed, this creates additional operational drag because the same risk pattern remains active in the tenant.' } else { (Convert-ToArrayaOutcomeSentence -Text $leadFinding.BusinessValue) + '.' })
        }) | Out-Null
    }

    return @($themes.ToArray())
}

function Get-CustomerActionPrimaryOwnerLabel {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Finding)

    $ownerTeam = Convert-ToArrayaDisplayText -Value $Finding.OwnerTeam -Default ''
    switch ($ownerTeam.ToLowerInvariant()) {
        'identity' { return 'Identity and access administration' }
        'messaging' { return 'Messaging and email administration' }
        'collaboration' { return 'Collaboration service owner / workspace sponsor' }
        'endpoint' { return 'Endpoint engineering / device administration' }
        'governance' { return 'Tenant governance and platform ownership' }
        'security' { return 'Security operations / control owner' }
        default {
            if ([string]::IsNullOrWhiteSpace($ownerTeam)) {
                return 'Platform owner not clearly surfaced in the current source'
            }

            return "$ownerTeam owner"
        }
    }
}

function Get-CustomerActionReferenceSection {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$LeadFinding)

    $relatedSection = Convert-ToArrayaDisplayText -Value $LeadFinding.RelatedSection -Default ''
    if (-not [string]::IsNullOrWhiteSpace($relatedSection) -and $relatedSection -ne 'N/A') {
        return $relatedSection
    }

    $area = Convert-ToArrayaDisplayText -Value $LeadFinding.Area -Default ''
    if (-not [string]::IsNullOrWhiteSpace($area) -and $area -ne 'N/A') {
        return $area
    }

    return $null
}

function Get-CustomerActionFirstValidationStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$LeadFinding,
        [Parameter(Mandatory = $false)][string[]]$Areas = @()
    )

    $themeProfile = Get-CustomerThemeProfile -Finding $LeadFinding
    switch ([string]$themeProfile.ActionTitle) {
        'Reduce privileged access and strengthen identity controls' {
            return 'Confirm the target identity protection baseline and decide which privileged, guest, and core user populations should be brought under it first.'
        }
        'Establish accountable ownership for collaboration spaces' {
            return 'Confirm which collaboration spaces need an owner, steward, transfer, or retirement decision first.'
        }
        'Strengthen domain and anti-spoofing controls' {
            return 'Confirm the approved domain trust and mail-authentication baseline before cleanup begins.'
        }
        'Improve device compliance and managed endpoint coverage' {
            return 'Confirm which devices are expected to retain access and which endpoint exceptions are still justified.'
        }
        'Review external forwarding and mail flow exposure' {
            return 'Confirm which forwarding and mail-flow patterns are approved business exceptions and which should be removed.'
        }
        'Reconcile license capacity and tenant governance gaps' {
            return 'Confirm the target licensing, governance, and external-access baseline before cleanup work begins.'
        }
        'Strengthen baseline security and access protections' {
            return 'Confirm the target protection baseline and sequence the first rollout wave around the highest-value control gaps.'
        }
        default {
            return 'Confirm the desired operating baseline, scope, and exception path before detailed remediation begins.'
        }
    }
}

function Get-CustomerActionSuccessCheck {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$LeadFinding)

    $themeProfile = Get-CustomerThemeProfile -Finding $LeadFinding
    switch ([string]$themeProfile.ActionTitle) {
        'Reduce privileged access and strengthen identity controls' {
            return 'Privileged and external access follow one approved protection model, with only documented exceptions remaining.'
        }
        'Establish accountable ownership for collaboration spaces' {
            return 'Each in-scope collaboration space has an accountable owner or a documented lifecycle decision.'
        }
        'Strengthen domain and anti-spoofing controls' {
            return 'Required domains and mail-authentication controls align to the approved baseline.'
        }
        'Improve device compliance and managed endpoint coverage' {
            return 'Protected access is limited to approved device states, with documented exceptions only.'
        }
        'Review external forwarding and mail flow exposure' {
            return 'Only approved forwarding and mail-flow exceptions remain, and they are documented.'
        }
        'Reconcile license capacity and tenant governance gaps' {
            return 'Capacity, ownership, and governance decisions are aligned to the approved operating baseline.'
        }
        'Strengthen baseline security and access protections' {
            return 'The agreed protection baseline is active for the intended population and no longer relies on broad temporary exceptions.'
        }
        default {
            return 'The agreed baseline is in place, validated, and supported by documented exceptions only.'
        }
    }
}

function Get-CustomerGuestMfaExperienceSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$ExternalIdentityRestrictionsRecord,
        [Parameter(Mandatory = $false)]$GuestAccessConfigurationRecord,
        [Parameter(Mandatory = $false)]$MfaEnforcementSummaryRecord
    )

    $allowInvitesFrom = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $GuestAccessConfigurationRecord -Names @('GuestInvitationControl', 'AllowInvitesFrom')) -Default ''
    if ([string]::IsNullOrWhiteSpace($allowInvitesFrom)) {
        $allowInvitesFrom = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $ExternalIdentityRestrictionsRecord -Names @('AllowInvitesFrom')) -Default 'Not validated from the reviewed data'
    }

    $hasCrossTenantAccessPolicy = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ExternalIdentityRestrictionsRecord -Names @('HasCrossTenantAccessPolicy'))
    $defaultInboundMfaTrustRaw = Get-ArrayaObjectValue -Object $ExternalIdentityRestrictionsRecord -Names @('DefaultInboundMfaTrust')
    $defaultOutboundMfaTrustRaw = Get-ArrayaObjectValue -Object $ExternalIdentityRestrictionsRecord -Names @('DefaultOutboundMfaTrust')
    $defaultInboundMfaTrust = Convert-ToArrayaDisplayText -Value $defaultInboundMfaTrustRaw -Default 'Not validated from the reviewed data'
    $defaultOutboundMfaTrust = Convert-ToArrayaDisplayText -Value $defaultOutboundMfaTrustRaw -Default 'Not validated from the reviewed data'
    $guestCoverageDetected = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $MfaEnforcementSummaryRecord -Names @('GuestOrExternalCoverage'))
    $enabledPoliciesRequiringMfa = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $MfaEnforcementSummaryRecord -Names @('EnabledPoliciesRequiringMfa'))
    $guestUserCoveragePercent = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $MfaEnforcementSummaryRecord -Names @('GuestUserCoveragePercent'))
    $guestCoverageLow = ($null -ne $guestUserCoveragePercent -and $guestUserCoveragePercent -lt 80)
    $guestCoverageNotClear = ($guestCoverageDetected -eq $false -or $null -eq $guestUserCoveragePercent)

    $inboundTrustBoolean = Convert-ToArrayaBoolean $defaultInboundMfaTrustRaw
    $trustValidated = -not ($defaultInboundMfaTrust -match '(?i)^not validated|^not available')
    $trustConfigured = ($inboundTrustBoolean -eq $true)
    $trustNotConfigured = (($inboundTrustBoolean -eq $false) -or ($defaultInboundMfaTrust -match '(?i)^not configured$'))

    $guestInvitationContext = if ($allowInvitesFrom -match '(?i)everyone|allmembers') {
        'The current invitation model is broad enough that guest MFA should be treated as a baseline control, not an exception path.'
    }
    elseif ($allowInvitesFrom -match '(?i)admins') {
        'Even with a more restricted invitation model, guest MFA still matters for every external identity that receives access to this tenant.'
    }
    else {
        'Guest MFA should be reviewed alongside the current invitation model so external access does not rely on sponsorship alone.'
    }

    $currentCoverageText = if ($null -ne $guestUserCoveragePercent) {
        "The reviewed data currently estimates guest-user Conditional Access MFA coverage at $guestUserCoveragePercent%."
    }
    elseif ($guestCoverageDetected -eq $true) {
        'Guest or external-user Conditional Access coverage was detected in the reviewed baseline.'
    }
    elseif ($guestCoverageDetected -eq $false) {
        'Guest or external-user Conditional Access coverage was not clearly detected in the reviewed baseline.'
    }
    else {
        'The reviewed data did not clearly confirm guest or external-user Conditional Access coverage.'
    }

    $trustStateText = if ($trustConfigured -and ($hasCrossTenantAccessPolicy -ne $false)) {
        'Home-tenant MFA trust appears available for the default guest access model.'
    }
    elseif ($trustNotConfigured) {
        'Home-tenant MFA trust does not appear configured in the reviewed default cross-tenant settings.'
    }
    else {
        'The reviewed data did not confirm that home-tenant MFA trust is configured.'
    }

    $scopeInterpretationText = 'Some Conditional Access policies naturally target employee, admin, or workload-specific populations, so the guest risk story should be read from guest coverage, guest-specific targeting, and explicit exclusions rather than from every raw include and exclude row.'

    $userExperienceText = if ($trustConfigured -and ($hasCrossTenantAccessPolicy -ne $false)) {
        'When guest MFA is enforced and inbound MFA trust is configured, many guests can satisfy the requirement with the MFA they already complete in their home tenant instead of registering separately in this tenant. That experience is usually best reserved for trusted partner tenants with meaningful ongoing collaboration.'
    }
    elseif ($trustNotConfigured) {
        'When guest MFA is enforced without trusted home-tenant MFA, guest users may see additional verification prompts and a more disruptive sign-in experience depending on the collaboration flow and identity type. That is why selective trust for validated partner tenants is usually a better first step than broad trust for every external tenant.'
    }
    else {
        'The reviewed data did not confirm that home-tenant MFA trust is configured, so guest-user experience may be more disruptive until that design is validated. A safer default is to trust high-collaboration partner tenants first instead of enabling broad trust immediately.'
    }

    $recommendationStrategyText = if ($guestCoverageLow -and $trustConfigured -and ($hasCrossTenantAccessPolicy -ne $false)) {
        'Recommended approach: keep or create a guest-specific Conditional Access policy requiring MFA for the uncovered population, then use inbound MFA trust first for the partner tenants with the most collaboration and a validated trust relationship instead of broad trust for every external tenant.'
    }
    elseif ($trustConfigured -and ($hasCrossTenantAccessPolicy -ne $false)) {
        'Recommended approach: keep a guest-specific Conditional Access policy requiring MFA in place, and trust home-tenant MFA first for the partner tenants with the most collaboration and a validated trust relationship rather than enabling broad trust for every external tenant.'
    }
    elseif ($trustNotConfigured -and ($guestCoverageLow -or $guestCoverageNotClear)) {
        'Recommended approach: create or validate a guest-specific Conditional Access policy requiring MFA, then enable inbound MFA trust first for the external tenants with the most collaboration and a validated trust relationship. Keep other guest access on stricter resource-tenant enforcement until trust is deliberately approved.'
    }
    elseif ($trustNotConfigured) {
        'Recommended approach: keep guest MFA requirements explicit in Conditional Access, then evaluate inbound MFA trust first for the external tenants with the most collaboration before expanding trust more broadly.'
    }
    else {
        'Recommended approach: start with a guest-specific Conditional Access policy requiring MFA, then validate which partner tenants justify inbound MFA trust before expanding that trust beyond a small strategic set.'
    }

    $recommendationImpactText = if ($trustConfigured -and ($hasCrossTenantAccessPolicy -ne $false)) {
        'Guest users should expect MFA to be required for protected access, but trusted strategic partner tenants can often satisfy that requirement with the MFA already completed in the tenant they belong to instead of registering again here.'
    }
    elseif ($trustNotConfigured) {
        'Guest users should expect stronger sign-in requirements, and some may see additional verification or separate registration friction until trusted home-tenant MFA is configured for the partner tenants that justify it.'
    }
    else {
        'Guest users should expect stronger sign-in requirements. The reviewed data did not confirm home-tenant MFA trust, so the access experience should be validated before broad enforcement or broad trust decisions are made.'
    }

    return [pscustomobject]@{
        AllowInvitesFrom                = $allowInvitesFrom
        HasCrossTenantAccessPolicy      = $hasCrossTenantAccessPolicy
        DefaultInboundMfaTrust          = $defaultInboundMfaTrust
        DefaultOutboundMfaTrust         = $defaultOutboundMfaTrust
        GuestOrExternalCoverage         = $guestCoverageDetected
        EnabledPoliciesRequiringMfa     = $enabledPoliciesRequiringMfa
        GuestUserCoveragePercent        = $guestUserCoveragePercent
        WhyThisMatters                  = "Guest MFA enforcement matters because guest identities are external accounts with access into this tenant's resources. Without active MFA requirements, guest collaboration can bypass the same identity assurance expected for internal users. $guestInvitationContext"
        CurrentStateSummary             = "$currentCoverageText $trustStateText $scopeInterpretationText"
        UserExperience                  = $userExperienceText
        DesiredState                    = 'The desired baseline is to require strong guest authentication through a guest-specific Conditional Access policy and to trust the guest home-tenant MFA where supported and approved, rather than asking every guest to register separately in the resource tenant or broadly trusting every external tenant by default.'
        RecommendationStrategy          = $recommendationStrategyText
        RecommendationImpactText        = $recommendationImpactText
    }
}

function Get-CustomerActionUserImpactExperience {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Action,
        [Parameter(Mandatory = $false)]$GuestMfaExperienceSummary
    )

    $actionTitle = Convert-ToArrayaDisplayText -Value $Action.ActionTitle -Default ''
    switch ($actionTitle) {
        'Reduce privileged access and strengthen identity controls' {
            if ($null -ne $GuestMfaExperienceSummary) {
                return "Users should expect more consistent sign-in verification for sensitive access. $($GuestMfaExperienceSummary.RecommendationImpactText)"
            }

            return 'Users should expect more consistent sign-in verification for sensitive access, with limited disruption once pilot exclusions and break-glass paths are confirmed.'
        }
        'Strengthen baseline security and access protections' {
            if ($null -ne $GuestMfaExperienceSummary) {
                return "Users should expect short-term sign-in prompt increases as baseline access protections move from observation into enforcement. $($GuestMfaExperienceSummary.RecommendationImpactText)"
            }

            return 'Users should expect short-term sign-in prompt increases during rollout until the stronger access baseline is fully established.'
        }
        'Improve device compliance and managed endpoint coverage' {
            return 'Users on unmanaged or non-compliant devices may temporarily lose access to protected resources until their device is brought into policy or granted a documented exception.'
        }
        'Establish accountable ownership for collaboration spaces' {
            return 'Minimal day-to-day end-user disruption. Most change is administrative and affects workspace owners, sponsors, and lifecycle decisions more than routine collaboration.'
        }
        'Strengthen domain and anti-spoofing controls' {
            return 'Minimal end-user disruption. The main impact is on administrators and approved senders who may need to update mail-authentication records or documented relay exceptions.'
        }
        'Review external forwarding and mail flow exposure' {
            return 'Little impact for most users. Unapproved forwarding paths, legacy relay patterns, or inbox rules may stop working once exceptions are reviewed and removed.'
        }
        'Reconcile license capacity and tenant governance gaps' {
            return 'Primarily administrative change with little day-to-day end-user disruption once assignments, ownership, and license headroom are reconciled.'
        }
        default {
            return 'Primarily administrative change with low end-user impact once the current state is validated and approved exceptions are documented.'
        }
    }
}

function Get-CustomerActionImplementationExperienceNote {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Action,
        [Parameter(Mandatory = $false)]$GuestMfaExperienceSummary
    )

    $actionTitle = Convert-ToArrayaDisplayText -Value $Action.ActionTitle -Default ''
    switch ($actionTitle) {
        'Reduce privileged access and strengthen identity controls' {
            return 'Pilot identity-enforcement changes with admins, guests, and a small user group before broader rollout.'
        }
        'Strengthen baseline security and access protections' {
            return 'Expect a staged rollout with report-only review, pilot enforcement, and an exception review before broad enablement.'
        }
        'Improve device compliance and managed endpoint coverage' {
            return 'Validate exception handling first so unmanaged or unsupported devices do not create avoidable access outages.'
        }
        'Review external forwarding and mail flow exposure' {
            return 'Review approved business exceptions before removing forwarding paths or tightening relay settings.'
        }
        default {
            return 'Validate the affected user population, exceptions, and rollback path before broad enforcement.'
        }
    }
}

function Add-CustomerRoadmapActionExperienceNotes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$RoadmapActions = @(),
        [Parameter(Mandatory = $false)]$GuestMfaExperienceSummary
    )

    return @(
        foreach ($action in @($RoadmapActions)) {
            $actionProperties = [ordered]@{}
            foreach ($property in @($action.PSObject.Properties)) {
                $actionProperties[$property.Name] = $property.Value
            }

            $actionProperties['UserImpactExperience'] = Get-CustomerActionUserImpactExperience -Action $action -GuestMfaExperienceSummary $GuestMfaExperienceSummary
            $actionProperties['ImplementationExperienceNote'] = Get-CustomerActionImplementationExperienceNote -Action $action -GuestMfaExperienceSummary $GuestMfaExperienceSummary
            [pscustomobject]$actionProperties
        }
    )
}

function Get-CustomerRoadmapActions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Findings,
        [Parameter(Mandatory = $false)][int]$MaxPerPhase = 5
    )

    $actions = New-Object System.Collections.Generic.List[object]
    $themeGroups = @(
        $Findings |
            Group-Object { (Get-CustomerThemeProfile -Finding $_).Key } |
            Sort-Object `
                @{ Expression = { (@($_.Group | ForEach-Object { Get-RoadmapPhaseRank -Phase ([string]$_.RoadmapPhase) }) | Measure-Object -Minimum).Minimum }; Descending = $false }, `
                @{ Expression = { ($_.Group | ForEach-Object { Get-SeverityWeight -Severity $_.Severity } | Measure-Object -Maximum).Maximum }; Descending = $true }, `
                @{ Expression = { $_.Count }; Descending = $true }, `
                Name
    )

    foreach ($group in @($themeGroups)) {
        $orderedRows = @(
            $group.Group |
                Sort-Object `
                    @{ Expression = { Get-SeverityWeight -Severity $_.Severity }; Descending = $true }, `
                    @{ Expression = { Get-RoadmapPhaseRank -Phase ([string]$_.RoadmapPhase) }; Descending = $false }, `
                    RuleId
        )
        if ($orderedRows.Count -eq 0) { continue }

        $leadFinding = $orderedRows[0]
        $profile = Get-CustomerThemeProfile -Finding $leadFinding
        $areas = @(
            $orderedRows |
                ForEach-Object { [string]$_.Area } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -Unique |
                Select-Object -First 4
        )
        $workstreams = @(
            $orderedRows |
                ForEach-Object { [string]$_.OwnerTeam } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -Unique
        )
        $roadmapPhase = [string](@(
            $orderedRows |
                Sort-Object @{ Expression = { Get-RoadmapPhaseRank -Phase ([string]$_.RoadmapPhase) }; Descending = $false } |
                Select-Object -First 1
        )[0].RoadmapPhase)
        $highestSeverity = [string](@(
            $orderedRows |
                Sort-Object @{ Expression = { Get-SeverityWeight -Severity $_.Severity }; Descending = $true } |
                Select-Object -First 1
        )[0].Severity)

        $actions.Add([pscustomobject]@{
            RoadmapPhase        = $roadmapPhase
            ActionTitle         = $profile.ActionTitle
            Theme               = $profile.Theme
            Workstream          = $(if ($workstreams.Count -gt 0) { $workstreams[0] } else { [string]$leadFinding.OwnerTeam })
            HighestSeverity     = $highestSeverity
            FindingCount        = $orderedRows.Count
            WhatThisAddresses   = $(if ($areas.Count -gt 0) { 'This work item addresses {0} related findings across {1}.' -f $orderedRows.Count, (Join-ArrayaReadableList -Items $areas) } else { 'This work item addresses {0} related findings in the same control area.' -f $orderedRows.Count })
            StandoutReason      = Get-CustomerStandoutSentence -LeadFinding $leadFinding -FindingCount $orderedRows.Count -Areas $areas
            ExampleText         = Get-CustomerFindingExampleText -Finding $leadFinding -MaxItems 2
            WhyItMatters        = $(if ([string]::IsNullOrWhiteSpace([string](Convert-ToArrayaSentenceFragment -Text $leadFinding.WhyFlagged))) { 'This indicates that the tenant is carrying a repeat control pattern that will remain visible until this workstream is addressed.' } else { 'This matters because ' + (Convert-ToArrayaSentenceFragment -Text $leadFinding.WhyFlagged) + '.' })
            RecommendedNextStep = Get-ArrayaLeadSentence -Text $leadFinding.Recommendation
            BusinessValue       = $(if ([string]::IsNullOrWhiteSpace([string](Convert-ToArrayaOutcomeSentence -Text $leadFinding.BusinessValue))) { 'If it remains open, the tenant will continue to carry the same exposure and operational friction reflected in the current findings.' } else { (Convert-ToArrayaOutcomeSentence -Text $leadFinding.BusinessValue) + '.' })
            PrimaryOwner        = Get-CustomerActionPrimaryOwnerLabel -Finding $leadFinding
            FirstValidationStep = Get-CustomerActionFirstValidationStep -LeadFinding $leadFinding -Areas $areas
            SuccessCheck        = Get-CustomerActionSuccessCheck -LeadFinding $leadFinding
            RelatedSection      = Get-CustomerActionReferenceSection -LeadFinding $leadFinding
        }) | Out-Null
    }

    return @(
        $actions.ToArray() |
            Sort-Object `
                @{ Expression = { Get-RoadmapPhaseRank -Phase ([string]$_.RoadmapPhase) }; Descending = $false }, `
                @{ Expression = { Get-SeverityWeight -Severity ([string]$_.HighestSeverity) }; Descending = $true }, `
                @{ Expression = { [int]$_.FindingCount }; Descending = $true }, `
                ActionTitle
    )
}

function Get-CustomerExecutiveSummaryNarrative {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Findings,
        [Parameter(Mandatory = $true)][object[]]$ExecutiveThemes,
        [Parameter(Mandatory = $false)][object[]]$OwnerGroups = @()
    )

    $topOwnerGroups = if (@($OwnerGroups).Count -gt 0) {
        @($OwnerGroups | Select-Object -First 3)
    }
    else {
        @(
            $Findings |
                Group-Object OwnerTeam |
                Sort-Object Count -Descending |
                Select-Object -First 3
        )
    }
    $themeNames = @($ExecutiveThemes | Select-Object -First 3 | ForEach-Object { $_.Theme })

    $workstreamText = if ($topOwnerGroups.Count -gt 0) {
        Join-ArrayaReadableList -Items @($topOwnerGroups | ForEach-Object { $_.Name })
    }
    else {
        'the tenant overall'
    }
    $themeText = if ($themeNames.Count -gt 0) { $themeNames -join '; ' } else { 'the highest-risk Microsoft 365 control areas' }

    return "The current review shows the highest risk concentration in $workstreamText. The leading themes of $themeText indicate that the tenant is carrying repeated control drift in the same operating areas that already carry the most day-to-day administration load. The tables below focus on the risk pattern, why it matters now, and which leadership decisions would remove the biggest blockers to a cleaner operating baseline."
}

function Join-ArrayaReadableList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Items = @()
    )

    $values = @(
        $Items |
            ForEach-Object { Convert-ToArrayaDisplayText -Value $_ -Default '' } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )

    switch ($values.Count) {
        0 { return '' }
        1 { return $values[0] }
        2 { return ($values -join ' and ') }
        default {
            return ('{0}, and {1}' -f (($values[0..($values.Count - 2)]) -join ', '), $values[-1])
        }
    }
}

function Get-CustomerCondensedListText {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Text,
        [Parameter(Mandatory = $false)][int]$MaxItems = 3
    )

    $normalized = Convert-ToArrayaDisplayText -Value $Text -Default ''
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    $items = @(
        $normalized -split '\s*;\s*' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First $MaxItems
    )

    if ($items.Count -eq 0) {
        return $null
    }

    return (Join-ArrayaReadableList -Items $items)
}

function Get-CustomerListItemsFromText {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Text,
        [Parameter(Mandatory = $false)][int]$MaxItems = 3
    )

    $normalized = Convert-ToArrayaDisplayText -Value $Text -Default ''
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return @()
    }

    return @(
        $normalized -split '\s*;\s*' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First $MaxItems
    )
}

function Get-CustomerFindingExampleText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Finding,
        [Parameter(Mandatory = $false)][int]$MaxItems = 2
    )

    $currentValue = Convert-ToArrayaDisplayText -Value $Finding.CurrentValue -Default ''
    if ([string]::IsNullOrWhiteSpace($currentValue) -or $currentValue -eq 'N/A') {
        return $null
    }

    $items = @(
        $currentValue -split '\s*;\s*' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First $MaxItems
    )

    if ($items.Count -eq 0) {
        return $null
    }

    return (Join-ArrayaReadableList -Items $items)
}

function Convert-ToArrayaOutcomeSentence {
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    $normalized = Get-ArrayaLeadSentence -Text $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    $trimmed = $normalized.Trim().TrimEnd('.', '!', '?')
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return $null
    }

    if ($trimmed -match '^(Improves|Reduces|Strengthens|Avoids|Makes|Supports)\b\s*(.*)$') {
        $verb = switch ($matches[1].ToLowerInvariant()) {
            'improves' { 'improve' }
            'reduces' { 'reduce' }
            'strengthens' { 'strengthen' }
            'avoids' { 'avoid' }
            'makes' { 'make' }
            'supports' { 'support' }
            default { $matches[1].ToLowerInvariant() }
        }
        $remainder = $matches[2].Trim()
        if ([string]::IsNullOrWhiteSpace($remainder)) {
            return "Addressing this helps $verb the current condition"
        }

        return "Addressing this helps $verb $remainder"
    }

    return $trimmed
}

function Convert-ToArrayaSentenceFragment {
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    $normalized = Get-ArrayaLeadSentence -Text $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    $trimmed = $normalized.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return $null
    }

    $trimmed = $trimmed -replace '^(?i)(flagged because|because)\s+', ''
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return $null
    }

    if ($trimmed -match '^(?i)the assessment observed:\s*(.+)$') {
        $trimmed = 'the tenant currently shows ' + $matches[1].Trim()
    }

    return "$([char]::ToLowerInvariant($trimmed[0]))$($trimmed.Substring(1))"
}

function Get-CustomerStandoutSentence {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$LeadFinding,
        [Parameter(Mandatory = $true)][int]$FindingCount,
        [Parameter(Mandatory = $false)][string[]]$Areas = @()
    )

    $currentValue = Convert-ToArrayaDisplayText -Value $LeadFinding.CurrentValue -Default ''
    $condensedCurrentValue = Get-CustomerCondensedListText -Text $currentValue -MaxItems 2
    if (-not [string]::IsNullOrWhiteSpace($condensedCurrentValue)) {
        $currentValue = $condensedCurrentValue
    }
    $areaText = Join-ArrayaReadableList -Items $Areas
    if (-not [string]::IsNullOrWhiteSpace($currentValue)) {
        if (-not [string]::IsNullOrWhiteSpace($areaText)) {
            return "This pattern stood out during review because the current source shows $currentValue across $areaText, which is a visible concentration rather than a one-off exception."
        }

        return "This pattern stood out during review because the current source shows $currentValue, which is material in the context of these $FindingCount related findings."
    }

    $whyFlagged = Convert-ToArrayaSentenceFragment -Text $LeadFinding.WhyFlagged
    if (-not [string]::IsNullOrWhiteSpace($whyFlagged)) {
        return "This pattern stood out during review because $whyFlagged."
    }

    return "This pattern stood out during review because $FindingCount related findings are concentrated in the same part of the tenant."
}

function New-CustomerTechnicalObservationSection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$SectionNumber,
        [Parameter(Mandatory = $true)][string]$SectionTitle,
        [Parameter(Mandatory = $false)][object[]]$ConfigurationRows = @(),
        [Parameter(Mandatory = $false)][string]$ObservedNarrative,
        [Parameter(Mandatory = $false)][string]$PositiveNarrative,
        [Parameter(Mandatory = $true)][string]$WhyItMatters,
        [Parameter(Mandatory = $false)][string]$WhatWasReviewed,
        [Parameter(Mandatory = $false)][string]$WhatWasObserved
    )

    if (@($ConfigurationRows).Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($WhatWasReviewed)) {
        $ConfigurationRows = @(
            New-CustomerConfigurationRow -Signal 'Reviewed configuration scope' -State $WhatWasReviewed
        )
    }

    if ([string]::IsNullOrWhiteSpace($ObservedNarrative)) {
        $ObservedNarrative = $WhatWasObserved
    }

    if ([string]::IsNullOrWhiteSpace($PositiveNarrative)) {
        $PositiveNarrative = 'The current source still provides enough signal in this area to distinguish where controls are present and where follow-through is still needed.'
    }

    return [pscustomobject]@{
        SectionNumber     = $SectionNumber
        SectionTitle      = $SectionTitle
        ConfigurationRows = @($ConfigurationRows)
        ObservedNarrative = $ObservedNarrative
        PositiveNarrative = $PositiveNarrative
        WhyItMatters      = $WhyItMatters
    }
}

function New-CustomerConfigurationRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Signal,
        [Parameter(Mandatory = $false)][AllowNull()]$State
    )

    $displayState = Convert-ToArrayaDisplayText -Value $State -Default 'Not surfaced in current source'
    return [pscustomobject]@{
        Signal = $Signal
        State  = $displayState
    }
}

function New-CustomerConsultativeSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $false)][object[]]$SnapshotRows = @(),
        [Parameter(Mandatory = $false)][object[]]$SecondaryRows = @(),
        [Parameter(Mandatory = $false)][object[]]$TertiaryRows = @(),
        [Parameter(Mandatory = $false)][string]$Narrative,
        [Parameter(Mandatory = $false)][string]$RecommendationSupport
    )

    return [pscustomobject]@{
        Title                 = $Title
        SnapshotRows          = @($SnapshotRows)
        SecondaryRows         = @($SecondaryRows)
        TertiaryRows          = @($TertiaryRows)
        Narrative             = $Narrative
        RecommendationSupport = $RecommendationSupport
    }
}

function Convert-ToArrayaCountBreakdownText {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [string]) {
        $normalized = Convert-ToArrayaDisplayText -Value $Value -Default ''
        return $(if ([string]::IsNullOrWhiteSpace($normalized)) { $null } else { $normalized })
    }

    $rows = New-Object System.Collections.Generic.List[object]
    if ($Value -is [hashtable] -or $Value -is [System.Collections.Specialized.OrderedDictionary]) {
        foreach ($entry in $Value.GetEnumerator()) {
            $rows.Add([pscustomobject]@{
                Name  = [string]$entry.Key
                Count = Convert-ArrayaToNumber $entry.Value
            }) | Out-Null
        }
    }
    elseif ($Value.PSObject -and $Value.PSObject.Properties.Count -gt 0) {
        foreach ($property in $Value.PSObject.Properties) {
            $rows.Add([pscustomobject]@{
                Name  = [string]$property.Name
                Count = Convert-ArrayaToNumber $property.Value
            }) | Out-Null
        }
    }

    if ($rows.Count -eq 0) {
        return $null
    }

    return (
        @(
            $rows |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) } |
                Sort-Object @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Name } } |
                ForEach-Object { '{0}={1}' -f $_.Name, $(if ($null -eq $_.Count) { 0 } else { $_.Count }) }
        ) -join '; '
    )
}

function Get-CustomerFriendlyMfaMethodName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$MethodName
    )

    $rawValue = [string]$MethodName
    if ([string]::IsNullOrWhiteSpace($rawValue)) {
        return $null
    }

    $normalized = ($rawValue -replace '\s+', '').Trim().ToLowerInvariant()
    switch -Regex ($normalized) {
        '^(email|emailotp)$' { return 'Email one-time passcode' }
        '^(sms|mobilephone|alternatemobilephone|voice.*|officephone|phoneappnotificationlesssecure|phone)$' {
            return $(if ($normalized -match 'voice|officephone') { 'Phone call' } else { 'SMS / phone' })
        }
        '^(microsoftauthenticator.*|authenticatorapp|phoneappnotification|passwordlessphonesignin)$' { return 'Microsoft Authenticator' }
        '^(softwareonetimepasscode|softwareotp)$' { return 'Software one-time passcode' }
        '^(softwareoath.*|oath.*)$' { return 'Software OATH token' }
        '^(temporaryaccesspass|tap)$' { return 'Temporary Access Pass' }
        '^(fido2.*|passkey.*)$' { return 'FIDO2 security key / passkey' }
        '^(windowshelloforbusiness|windowshello.*)$' { return 'Windows Hello for Business' }
        default {
            if ($rawValue -cmatch '[a-z][A-Z]') {
                return ([regex]::Replace($rawValue, '(?<=[a-z])(?=[A-Z])', ' ')).Trim()
            }

            return $rawValue.Trim()
        }
    }
}

function Convert-ToCustomerMfaMethodBreakdownText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace([string]$Text)) {
        return $null
    }

    $segments = @(
        [string]$Text -split ';' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($segments.Count -eq 0) {
        return $null
    }

    return (
        @(
            foreach ($segment in $segments) {
                if ($segment -match '^(?<label>[^=]+?)\s*=\s*(?<count>.+)$') {
                    '{0}={1}' -f (Get-CustomerFriendlyMfaMethodName -MethodName $Matches.label), $Matches.count.Trim()
                    continue
                }

                Get-CustomerFriendlyMfaMethodName -MethodName $segment
            }
        ) -join '; '
    )
}

function Get-CustomerMfaEnrollmentSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$ExistingSummary,
        [Parameter(Mandatory = $false)]$MfaSummary,
        [Parameter(Mandatory = $false)][object[]]$MfaRegistrationDetails = @()
    )

    $source = if ($ExistingSummary) { $ExistingSummary } else { $MfaSummary }
    $detailRows = @($MfaRegistrationDetails | Where-Object { $null -ne $_ })
    if (-not $source -and $detailRows.Count -eq 0) {
        return $null
    }

    $detailMethodCounts = @{}
    $detailWeakMethodCounts = @{}
    $detailStrongMethodCounts = @{}
    $detailPhishingResistantMethodCounts = @{}
    $detailDefaultMethodCounts = @{}
    $detailTotalUsers = $detailRows.Count
    $detailRegisteredUsers = 0
    $detailUsersWithWeakMethodsOnly = 0
    $detailUsersWithWeakDefaultMethod = 0
    $detailUsersWithStrongMethods = 0
    $detailUsersWithPhishingResistantMethods = 0
    foreach ($detailRow in $detailRows) {
        $isRegistered = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $detailRow -Names @('IsMfaRegistered'))
        if ($isRegistered -eq $true) {
            $detailRegisteredUsers++
        }

        $methodsRegisteredSource = if ($detailRow.PSObject -and ($detailRow.PSObject.Properties.Name -contains 'MethodsRegistered')) {
            $detailRow.MethodsRegistered
        }
        else {
            Get-ArrayaObjectValue -Object $detailRow -Names @('MethodsRegistered')
        }
        $methodsRegistered = @()
        if ($methodsRegisteredSource -is [System.Collections.IEnumerable] -and -not ($methodsRegisteredSource -is [string]) -and -not ($methodsRegisteredSource -is [System.Collections.IDictionary])) {
            $methodsRegistered = @(
                @($methodsRegisteredSource) |
                    ForEach-Object { Convert-ToArrayaDisplayText -Value $_ -Default '' } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
        }
        else {
            $methodsRegistered = Convert-ToArrayaStringList $methodsRegisteredSource
        }
        foreach ($methodName in @($methodsRegistered | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })) {
            if (-not $detailMethodCounts.ContainsKey($methodName)) { $detailMethodCounts[$methodName] = 0 }
            $detailMethodCounts[$methodName]++
        }

        $defaultMethodSource = if ($detailRow.PSObject -and ($detailRow.PSObject.Properties.Name -contains 'DefaultMfaMethod')) {
            $detailRow.DefaultMfaMethod
        }
        else {
            Get-ArrayaObjectValue -Object $detailRow -Names @('DefaultMfaMethod')
        }
        $defaultMethod = Convert-ToArrayaDisplayText -Value $defaultMethodSource -Default ''
        if (-not [string]::IsNullOrWhiteSpace($defaultMethod)) {
            if (-not $detailDefaultMethodCounts.ContainsKey($defaultMethod)) { $detailDefaultMethodCounts[$defaultMethod] = 0 }
            $detailDefaultMethodCounts[$defaultMethod]++
        }

        $hasWeakMethod = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $detailRow -Names @('HasWeakMethod'))
        $hasStrongMethod = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $detailRow -Names @('HasStrongMethod'))
        $hasPhishingResistantMethod = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $detailRow -Names @('HasPhishingResistantMethod'))
        if ($isRegistered -eq $true -and $hasWeakMethod -eq $true -and $hasStrongMethod -ne $true) {
            $detailUsersWithWeakMethodsOnly++
        }
        if ($isRegistered -eq $true -and $hasStrongMethod -eq $true) {
            $detailUsersWithStrongMethods++
            foreach ($methodName in @($methodsRegistered | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })) {
                if (-not $detailStrongMethodCounts.ContainsKey($methodName)) { $detailStrongMethodCounts[$methodName] = 0 }
                $detailStrongMethodCounts[$methodName]++
            }
        }
        if ($isRegistered -eq $true -and $hasPhishingResistantMethod -eq $true) {
            $detailUsersWithPhishingResistantMethods++
            foreach ($methodName in @($methodsRegistered | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })) {
                if (-not $detailPhishingResistantMethodCounts.ContainsKey($methodName)) { $detailPhishingResistantMethodCounts[$methodName] = 0 }
                $detailPhishingResistantMethodCounts[$methodName]++
            }
        }
        if ($isRegistered -eq $true -and $hasWeakMethod -eq $true -and -not [string]::IsNullOrWhiteSpace($defaultMethod)) {
            if (-not $detailWeakMethodCounts.ContainsKey($defaultMethod)) { $detailWeakMethodCounts[$defaultMethod] = 0 }
            $detailWeakMethodCounts[$defaultMethod]++
            $detailUsersWithWeakDefaultMethod++
        }
    }

    $totalUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $source -Names @('TotalUsers', 'UserCount', 'TotalUserCount'))
    $registeredUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $source -Names @('RegisteredUsers', 'RegisteredUserCount', 'MfaRegisteredUsers'))
    $notRegisteredUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $source -Names @('NotRegisteredUsers'))
    $registrationPercent = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $source -Names @('RegistrationPercent', 'RegisteredPercent', 'MfaRegistrationPercent'))
    if ($null -eq $totalUsers -and $detailTotalUsers -gt 0) {
        $totalUsers = $detailTotalUsers
    }
    if ($null -eq $registeredUsers -and $detailTotalUsers -gt 0) {
        $registeredUsers = $detailRegisteredUsers
    }
    if ($null -eq $notRegisteredUsers -and $detailTotalUsers -gt 0) {
        $notRegisteredUsers = ($detailTotalUsers - $detailRegisteredUsers)
    }
    if ($null -eq $registrationPercent -and $null -ne $registeredUsers -and $null -ne $totalUsers -and $totalUsers -gt 0) {
        $registrationPercent = [math]::Round(($registeredUsers / $totalUsers) * 100, 1)
    }

    $registeredMethodBreakdownText = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $source -Names @('RegisteredMethodBreakdown')) -Default ''
    $weakMethodBreakdownText = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $source -Names @('WeakMethodBreakdown')) -Default ''
    $strongMethodBreakdownText = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $source -Names @('StrongMethodBreakdown')) -Default ''
    $phishingResistantMethodBreakdownText = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $source -Names @('PhishingResistantMethodBreakdown')) -Default ''
    $defaultMethodBreakdownText = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $source -Names @('DefaultMethodBreakdown')) -Default ''

    $methodCounts = Get-ArrayaObjectValue -Object $source -Names @('MethodCounts')
    if ($detailMethodCounts.Count -gt 0 -and [string]::IsNullOrWhiteSpace($registeredMethodBreakdownText)) { $methodCounts = $detailMethodCounts }
    elseif ($null -eq $methodCounts -and $detailMethodCounts.Count -gt 0) { $methodCounts = $detailMethodCounts }

    $weakMethodCounts = Get-ArrayaObjectValue -Object $source -Names @('WeakMethodCounts')
    if ($detailWeakMethodCounts.Count -gt 0 -and [string]::IsNullOrWhiteSpace($weakMethodBreakdownText)) { $weakMethodCounts = $detailWeakMethodCounts }
    elseif ($null -eq $weakMethodCounts -and $detailWeakMethodCounts.Count -gt 0) { $weakMethodCounts = $detailWeakMethodCounts }

    $strongMethodCounts = Get-ArrayaObjectValue -Object $source -Names @('StrongMethodCounts')
    if ($detailStrongMethodCounts.Count -gt 0 -and [string]::IsNullOrWhiteSpace($strongMethodBreakdownText)) { $strongMethodCounts = $detailStrongMethodCounts }
    elseif ($null -eq $strongMethodCounts -and $detailStrongMethodCounts.Count -gt 0) { $strongMethodCounts = $detailStrongMethodCounts }

    $phishingResistantMethodCounts = Get-ArrayaObjectValue -Object $source -Names @('PhishingResistantMethodCounts')
    if ($detailPhishingResistantMethodCounts.Count -gt 0 -and [string]::IsNullOrWhiteSpace($phishingResistantMethodBreakdownText)) { $phishingResistantMethodCounts = $detailPhishingResistantMethodCounts }
    elseif ($null -eq $phishingResistantMethodCounts -and $detailPhishingResistantMethodCounts.Count -gt 0) { $phishingResistantMethodCounts = $detailPhishingResistantMethodCounts }

    $defaultMethodCounts = Get-ArrayaObjectValue -Object $source -Names @('DefaultMethodCounts')
    if ($detailDefaultMethodCounts.Count -gt 0 -and [string]::IsNullOrWhiteSpace($defaultMethodBreakdownText)) { $defaultMethodCounts = $detailDefaultMethodCounts }
    elseif ($null -eq $defaultMethodCounts -and $detailDefaultMethodCounts.Count -gt 0) { $defaultMethodCounts = $detailDefaultMethodCounts }
    $usersWithWeakMethodsOnly = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $source -Names @('UsersWithWeakMethodsOnly'))
    if ($null -eq $usersWithWeakMethodsOnly -and $detailTotalUsers -gt 0) { $usersWithWeakMethodsOnly = $detailUsersWithWeakMethodsOnly }
    $usersWithWeakDefaultMethod = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $source -Names @('UsersWithWeakDefaultMethod'))
    if ($null -eq $usersWithWeakDefaultMethod -and $detailTotalUsers -gt 0) { $usersWithWeakDefaultMethod = $detailUsersWithWeakDefaultMethod }
    $usersWithStrongMethods = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $source -Names @('UsersWithStrongMethods'))
    if ($null -eq $usersWithStrongMethods -and $detailTotalUsers -gt 0) { $usersWithStrongMethods = $detailUsersWithStrongMethods }
    $usersWithPhishingResistantMethods = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $source -Names @('UsersWithPhishingResistantMethods'))
    if ($null -eq $usersWithPhishingResistantMethods -and $detailTotalUsers -gt 0) { $usersWithPhishingResistantMethods = $detailUsersWithPhishingResistantMethods }
    $collectionState = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $source -Names @('CollectionState')) -Default ''
    if ([string]::IsNullOrWhiteSpace($collectionState) -or $collectionState -eq 'N/A') {
        $collectionState = if ($detailTotalUsers -gt 0) { 'Derived from MFA registration details' } else { 'Not validated from the reviewed data' }
    }

    return [pscustomobject]@{
        TotalUsers                        = $totalUsers
        RegisteredUsers                   = $registeredUsers
        NotRegisteredUsers                = $(if ($null -ne $notRegisteredUsers) { $notRegisteredUsers } elseif ($null -ne $totalUsers -and $null -ne $registeredUsers) { $totalUsers - $registeredUsers } else { $null })
        RegistrationPercent               = $registrationPercent
        RegisteredMethodBreakdown         = Convert-ToCustomerMfaMethodBreakdownText $(if (-not [string]::IsNullOrWhiteSpace($registeredMethodBreakdownText)) { $registeredMethodBreakdownText } else { Convert-ToArrayaCountBreakdownText $methodCounts })
        WeakMethodBreakdown               = Convert-ToCustomerMfaMethodBreakdownText $(if (-not [string]::IsNullOrWhiteSpace($weakMethodBreakdownText)) { $weakMethodBreakdownText } else { Convert-ToArrayaCountBreakdownText $weakMethodCounts })
        StrongMethodBreakdown             = Convert-ToCustomerMfaMethodBreakdownText $(if (-not [string]::IsNullOrWhiteSpace($strongMethodBreakdownText)) { $strongMethodBreakdownText } else { Convert-ToArrayaCountBreakdownText $strongMethodCounts })
        PhishingResistantMethodBreakdown  = Convert-ToCustomerMfaMethodBreakdownText $(if (-not [string]::IsNullOrWhiteSpace($phishingResistantMethodBreakdownText)) { $phishingResistantMethodBreakdownText } else { Convert-ToArrayaCountBreakdownText $phishingResistantMethodCounts })
        DefaultMethodBreakdown            = Convert-ToCustomerMfaMethodBreakdownText $(if (-not [string]::IsNullOrWhiteSpace($defaultMethodBreakdownText)) { $defaultMethodBreakdownText } else { Convert-ToArrayaCountBreakdownText $defaultMethodCounts })
        UsersWithWeakMethodsOnly          = $usersWithWeakMethodsOnly
        UsersWithWeakDefaultMethod        = $usersWithWeakDefaultMethod
        UsersWithStrongMethods            = $usersWithStrongMethods
        UsersWithPhishingResistantMethods = $usersWithPhishingResistantMethods
        CollectionState                   = $collectionState
    }
}

function Get-CustomerGuestUserEnforcementState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$ExistingSummary,
        [Parameter(Mandatory = $false)][object[]]$ConditionalAccessPolicies = @(),
        [Parameter(Mandatory = $false)]$ConditionalAccessSummary
    )

    $existingState = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('GuestUserEnforcementState')) -Default ''
    if (-not [string]::IsNullOrWhiteSpace($existingState) -and $existingState -ne 'N/A') {
        return $existingState
    }

    $guestCoverage = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('GuestOrExternalCoverage'))
    if ($null -eq $guestCoverage -and $ConditionalAccessSummary) {
        $guestCoverage = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object (Get-ArrayaObjectValue -Object $ConditionalAccessSummary -Names @('Summary')) -Names @('HasGuestCoverage'))
    }
    $enabledPoliciesRequiringMfa = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('EnabledPoliciesRequiringMfa'))
    $reportOnlyPoliciesRequiringMfa = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('ReportOnlyPoliciesRequiringMfa'))
    $enabledGuestUsersReviewed = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('EnabledGuestUsersReviewed'))
    $guestUsersCovered = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('GuestUsersCoveredByEnabledMfaPolicies'))
    $guestCoveragePercent = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('GuestUserCoveragePercent'))

    if ($null -ne $enabledGuestUsersReviewed -and $enabledGuestUsersReviewed -eq 0) {
        return 'No enabled guest users were available for guest MFA coverage review.'
    }
    if ($null -ne $enabledGuestUsersReviewed -and $enabledGuestUsersReviewed -gt 0 -and $null -ne $guestUsersCovered -and $guestUsersCovered -ge $enabledGuestUsersReviewed) {
        return 'Reviewed enabled guest users appear covered by the active MFA enforcement baseline.'
    }
    if ($null -ne $guestCoveragePercent -and $guestCoveragePercent -gt 0) {
        return 'Guest users are partially covered by the active MFA enforcement baseline.'
    }
    if ($guestCoverage -eq $true -and (($null -ne $enabledPoliciesRequiringMfa -and $enabledPoliciesRequiringMfa -gt 0) -or @($ConditionalAccessPolicies).Count -gt 0)) {
        return 'Guest users are partially covered by the active MFA enforcement baseline.'
    }
    if ($guestCoverage -eq $true -and $null -ne $reportOnlyPoliciesRequiringMfa -and $reportOnlyPoliciesRequiringMfa -gt 0) {
        return 'Guest MFA appears staged in report-only Conditional Access, but active guest enforcement was not clearly detected.'
    }

    return 'Guest MFA enforcement was not clearly detected in the reviewed policy baseline.'
}

function Test-ArrayaMeaningfulNestedValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return $false
    }

    if ($Value -is [string]) {
        $textValue = $Value.Trim()
        if ([string]::IsNullOrWhiteSpace($textValue) -or $textValue -in @('null', '{}', '[]')) {
            return $false
        }

        if (
            ($textValue.StartsWith('{') -and $textValue.EndsWith('}')) -or
            ($textValue.StartsWith('[') -and $textValue.EndsWith(']'))
        ) {
            try {
                return (Test-ArrayaMeaningfulNestedValue -Value ($textValue | ConvertFrom-Json -Depth 10))
            }
            catch {
                return $true
            }
        }

        return $true
    }

    if ($Value -is [bool] -or $Value -is [ValueType]) {
        return $true
    }

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($entry in $Value.GetEnumerator()) {
            if (Test-ArrayaMeaningfulNestedValue -Value $entry.Value) {
                return $true
            }
        }

        return $false
    }

    if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])) {
        foreach ($item in @($Value)) {
            if (Test-ArrayaMeaningfulNestedValue -Value $item) {
                return $true
            }
        }

        return $false
    }

    $propertyBag = @($Value.PSObject.Properties | Where-Object { $_.MemberType -like '*Property' })
    if ($propertyBag.Count -gt 0) {
        foreach ($property in $propertyBag) {
            if (Test-ArrayaMeaningfulNestedValue -Value $property.Value) {
                return $true
            }
        }

        return $false
    }

    return $true
}

function Test-ArrayaConditionalAccessRequiresMfa {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $Policy
    )

    if ($null -eq $Policy) {
        return $false
    }

    $builtInControls = @()
    $builtInControlValue = Get-ArrayaObjectValue -Object $Policy -Names @('GrantControls_BuiltInControls')
    if ($builtInControlValue -is [string]) {
        $builtInControls = @(
            $builtInControlValue -split ',' |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )
    }
    elseif (($builtInControlValue -is [System.Collections.IEnumerable]) -and -not ($builtInControlValue -is [string])) {
        $builtInControls = @(
            @($builtInControlValue) |
                ForEach-Object { [string]$_ } |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )
    }
    if (@($builtInControls | Where-Object { $_ -match '^(?i)mfa$' }).Count -gt 0) {
        return $true
    }

    return (Test-ArrayaMeaningfulNestedValue -Value (Get-ArrayaObjectValue -Object $Policy -Names @('GrantControls_AuthenticationStrength', 'AuthenticationStrength')))
}

function Get-CustomerMfaEnforcementSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]$ExistingSummary,
        [Parameter(Mandatory = $false)][object[]]$ConditionalAccessPolicies = @(),
        [Parameter(Mandatory = $false)]$ConditionalAccessSummary,
        [Parameter(Mandatory = $false)]$SecurityDefaultsPolicy
    )

    $conditionalAccessSummaryRecord = if ($ConditionalAccessSummary) { Get-ArrayaObjectValue -Object $ConditionalAccessSummary -Names @('Summary') } else { $null }
    $securityDefaultsRecord = if ($SecurityDefaultsPolicy) { Get-ArrayaObjectValue -Object $SecurityDefaultsPolicy -Names @('Configuration') } else { $null }

    $summaryTotalPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $conditionalAccessSummaryRecord -Names @('TotalPolicies'))
    $summaryEnabledPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $conditionalAccessSummaryRecord -Names @('EnabledPolicies'))
    $summaryReportOnlyPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $conditionalAccessSummaryRecord -Names @('ReportOnlyPolicies'))
    $summaryPoliciesWithExclusions = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $conditionalAccessSummaryRecord -Names @('PoliciesWithExclusions'))
    $calculatedMfaPolicies = @(
        $ConditionalAccessPolicies |
            Where-Object { Test-ArrayaConditionalAccessRequiresMfa -Policy $_ }
    )
    $calculatedEnabledMfaPolicies = @($calculatedMfaPolicies | Where-Object { ([string](Get-ArrayaObjectValue -Object $_ -Names @('State'))).ToLowerInvariant() -eq 'enabled' })
    $calculatedReportOnlyMfaPolicies = @($calculatedMfaPolicies | Where-Object { ([string](Get-ArrayaObjectValue -Object $_ -Names @('State'))) -match '(?i)report' })

    if ($ExistingSummary) {
        $reviewedPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('ConditionalAccessPoliciesReviewed'))
        if ($null -ne $summaryTotalPolicies) {
            $reviewedPolicies = $summaryTotalPolicies
        }
        elseif (($null -eq $reviewedPolicies) -and $ConditionalAccessPolicies.Count -gt 0) {
            $reviewedPolicies = $ConditionalAccessPolicies.Count
        }
        $guestUserEnforcementState = Get-CustomerGuestUserEnforcementState -ExistingSummary $ExistingSummary -ConditionalAccessPolicies $ConditionalAccessPolicies -ConditionalAccessSummary $ConditionalAccessSummary

        return [pscustomobject]@{
            ConditionalAccessPoliciesReviewed   = $reviewedPolicies
            EnabledConditionalAccessPolicies    = $(if ($null -ne $summaryEnabledPolicies) { $summaryEnabledPolicies } else { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('EnabledConditionalAccessPolicies')) })
            ReportOnlyConditionalAccessPolicies = $(if ($null -ne $summaryReportOnlyPolicies) { $summaryReportOnlyPolicies } else { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('ReportOnlyConditionalAccessPolicies')) })
            PoliciesWithExclusions              = $(if ($null -ne $summaryPoliciesWithExclusions) { $summaryPoliciesWithExclusions } else { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('PoliciesWithExclusions')) })
            PoliciesRequiringMfa                = $(if ($calculatedMfaPolicies.Count -gt 0) { $calculatedMfaPolicies.Count } else { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('PoliciesRequiringMfa')) })
            EnabledPoliciesRequiringMfa         = $(if ($calculatedEnabledMfaPolicies.Count -gt 0) { $calculatedEnabledMfaPolicies.Count } else { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('EnabledPoliciesRequiringMfa')) })
            ReportOnlyPoliciesRequiringMfa      = $(if ($calculatedReportOnlyMfaPolicies.Count -gt 0) { $calculatedReportOnlyMfaPolicies.Count } else { Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('ReportOnlyPoliciesRequiringMfa')) })
            EnabledUsersReviewed                = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('EnabledUsersReviewed'))
            EnabledMemberUsersReviewed          = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('EnabledMemberUsersReviewed'))
            EnabledGuestUsersReviewed           = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('EnabledGuestUsersReviewed'))
            UsersCoveredByEnabledMfaPolicies    = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('UsersCoveredByEnabledMfaPolicies'))
            UserCoveragePercent                 = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('UserCoveragePercent'))
            MemberUsersCoveredByEnabledMfaPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('MemberUsersCoveredByEnabledMfaPolicies'))
            MemberUserCoveragePercent           = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('MemberUserCoveragePercent'))
            GuestUsersCoveredByEnabledMfaPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('GuestUsersCoveredByEnabledMfaPolicies'))
            GuestUserCoveragePercent            = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('GuestUserCoveragePercent'))
            GroupTargetedPolicyCount            = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('GroupTargetedPolicyCount'))
            RoleTargetedPolicyCount             = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('RoleTargetedPolicyCount'))
            CoverageCalculationNote             = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('CoverageCalculationNote')) -Default 'Not validated from the reviewed data'
            SecurityDefaultsEnabled             = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('SecurityDefaultsEnabled'))
            GuestOrExternalCoverage             = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('GuestOrExternalCoverage'))
            PrivilegedRoleCoverage              = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('PrivilegedRoleCoverage'))
            RiskBasedCoverage                   = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('RiskBasedCoverage'))
            CompliantDeviceRequirement          = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('CompliantDeviceRequirement'))
            GuestUserEnforcementState           = $guestUserEnforcementState
            EnforcementState                    = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('EnforcementState')) -Default 'Not validated from the reviewed data'
            CollectionState                     = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $ExistingSummary -Names @('CollectionState')) -Default 'Not validated from the reviewed data'
        }
    }

    $mfaPolicies = @($calculatedMfaPolicies)
    $enabledMfaPolicies = @($calculatedEnabledMfaPolicies)
    $reportOnlyMfaPolicies = @($calculatedReportOnlyMfaPolicies)
    $securityDefaultsEnabled = $null
    if ($securityDefaultsRecord) {
        $securityDefaultsEnabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $securityDefaultsRecord -Names @('IsEnabled', 'Enabled'))
    }

    $enforcementState = if ($enabledMfaPolicies.Count -gt 0 -and $securityDefaultsEnabled -eq $true) {
        'Conditional Access and Security Defaults both enforce MFA; confirm that the overlapping baseline is intentional.'
    }
    elseif ($enabledMfaPolicies.Count -gt 0) {
        'MFA enforcement is active through enabled Conditional Access policies.'
    }
    elseif ($securityDefaultsEnabled -eq $true) {
        'MFA enforcement is active through Security Defaults.'
    }
    elseif ($reportOnlyMfaPolicies.Count -gt 0) {
        'MFA appears staged in report-only Conditional Access, but active enforcement was not clearly detected.'
    }
    else {
        'Active MFA enforcement was not clearly detected in the reviewed policy baseline.'
    }
    $guestUserEnforcementState = Get-CustomerGuestUserEnforcementState -ExistingSummary ([pscustomobject]@{
        GuestOrExternalCoverage        = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $conditionalAccessSummaryRecord -Names @('HasGuestCoverage'))
        EnabledPoliciesRequiringMfa    = $enabledMfaPolicies.Count
        ReportOnlyPoliciesRequiringMfa = $reportOnlyMfaPolicies.Count
    }) -ConditionalAccessPolicies $ConditionalAccessPolicies -ConditionalAccessSummary $ConditionalAccessSummary

    return [pscustomobject]@{
        ConditionalAccessPoliciesReviewed   = $(if ($null -ne $summaryTotalPolicies) { $summaryTotalPolicies } else { $ConditionalAccessPolicies.Count })
        EnabledConditionalAccessPolicies    = $summaryEnabledPolicies
        ReportOnlyConditionalAccessPolicies = $summaryReportOnlyPolicies
        PoliciesWithExclusions              = $summaryPoliciesWithExclusions
        PoliciesRequiringMfa                = $mfaPolicies.Count
        EnabledPoliciesRequiringMfa         = $enabledMfaPolicies.Count
        ReportOnlyPoliciesRequiringMfa      = $reportOnlyMfaPolicies.Count
        EnabledUsersReviewed                = $null
        EnabledMemberUsersReviewed          = $null
        EnabledGuestUsersReviewed           = $null
        UsersCoveredByEnabledMfaPolicies    = $null
        UserCoveragePercent                 = $null
        MemberUsersCoveredByEnabledMfaPolicies = $null
        MemberUserCoveragePercent           = $null
        GuestUsersCoveredByEnabledMfaPolicies = $null
        GuestUserCoveragePercent            = $null
        GroupTargetedPolicyCount            = $null
        RoleTargetedPolicyCount             = $null
        CoverageCalculationNote             = 'Not validated from the reviewed data'
        SecurityDefaultsEnabled             = $securityDefaultsEnabled
        GuestOrExternalCoverage             = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $conditionalAccessSummaryRecord -Names @('HasGuestCoverage'))
        PrivilegedRoleCoverage              = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $conditionalAccessSummaryRecord -Names @('HasPrivilegedRoleCoverage'))
        RiskBasedCoverage                   = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $conditionalAccessSummaryRecord -Names @('HasRiskBasedCoverage'))
        CompliantDeviceRequirement          = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $conditionalAccessSummaryRecord -Names @('HasCompliantDeviceRequirement'))
        GuestUserEnforcementState           = $guestUserEnforcementState
        EnforcementState                    = $enforcementState
        CollectionState                     = 'Derived'
    }
}

function New-CustomerDocumentationReference {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$WhyItIsRelevant
    )

    return [pscustomobject]@{
        Title           = $Title
        Url             = $Url
        WhyItIsRelevant = $WhyItIsRelevant
    }
}

function Get-CustomerTechnicalSectionTitleMap {
    [CmdletBinding()]
    param()

    return @(
        'Identity & Access (Entra ID)',
        'Devices & Endpoint Management',
        'Messaging (Exchange Online)',
        'Collaboration (Teams, SharePoint, OneDrive)',
        'Data Protection & Governance',
        'Offboarding & Lifecycle Management'
    )
}

function Get-CustomerDocumentationAppendixSections {
    [CmdletBinding()]
    param()

    $sections = @()

    $identityRefs = @()
    $reference = New-CustomerDocumentationReference -Title 'Conditional Access overview' -Url 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview' -WhyItIsRelevant 'Supports the access-control observations around policy coverage, exclusions, and enforcement maturity.'
    $identityRefs += $reference
    $reference = New-CustomerDocumentationReference -Title 'Plan a Microsoft Entra multifactor authentication deployment' -Url 'https://learn.microsoft.com/en-us/entra/identity/authentication/howto-mfa-getstarted' -WhyItIsRelevant 'Provides Microsoft guidance for improving MFA coverage and rollout maturity.'
    $identityRefs += $reference
    $reference = New-CustomerDocumentationReference -Title 'Privileged Identity Management overview' -Url 'https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-configure' -WhyItIsRelevant 'Supports the recommendations related to privileged access hygiene, role control, and reducing standing access.'
    $identityRefs += $reference
    $section = [pscustomobject]@{
        Title      = 'Identity & Access'
        Intro      = 'The following Microsoft documentation supports the identity observations and recommended control changes discussed in this report.'
        References = $identityRefs
    }
    $sections += $section

    $deviceRefs = @()
    $reference = New-CustomerDocumentationReference -Title 'Get started with device compliance policies in Microsoft Intune' -Url 'https://learn.microsoft.com/en-us/intune/intune-service/protect/device-compliance-get-started' -WhyItIsRelevant 'Supports the discussion around compliance baselines and how endpoint posture is evaluated.'
    $deviceRefs += $reference
    $reference = New-CustomerDocumentationReference -Title 'Add actions for noncompliance to device compliance policies' -Url 'https://learn.microsoft.com/en-us/intune/intune-service/protect/actions-for-noncompliance' -WhyItIsRelevant 'Relevant to the observed gaps between compliant and non-compliant devices and how exceptions are handled.'
    $deviceRefs += $reference
    $section = [pscustomobject]@{
        Title      = 'Devices & Endpoint Management'
        Intro      = 'These Microsoft references support the device-compliance and endpoint-management observations in the current-state review.'
        References = $deviceRefs
    }
    $sections += $section

    $messagingRefs = @()
    $reference = New-CustomerDocumentationReference -Title 'Control automatic external email forwarding in Microsoft 365' -Url 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/outbound-spam-policies-external-email-forwarding' -WhyItIsRelevant 'Supports the observations around outbound forwarding policy, remote domains, and forwarding exposure.'
    $messagingRefs += $reference
    $reference = New-CustomerDocumentationReference -Title 'Set up SPF in Microsoft 365 to help prevent spoofing' -Url 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/set-up-spf-in-office-365-to-help-prevent-spoofing' -WhyItIsRelevant 'Relevant to the domain and anti-spoofing observations surfaced in the messaging and governance sections.'
    $messagingRefs += $reference
    $reference = New-CustomerDocumentationReference -Title 'Use DKIM to validate outbound email sent from your custom domain' -Url 'https://learn.microsoft.com/en-us/defender-office-365/email-authentication-dkim-configure' -WhyItIsRelevant 'Supports the recommendations tied to mail-authentication maturity and custom-domain protection.'
    $messagingRefs += $reference
    $section = [pscustomobject]@{
        Title      = 'Messaging'
        Intro      = 'These references support the messaging findings related to forwarding, mail hygiene, and domain-based protection controls.'
        References = $messagingRefs
    }
    $sections += $section

    $collaborationRefs = @()
    $reference = New-CustomerDocumentationReference -Title 'Overview of external sharing in SharePoint and OneDrive' -Url 'https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview' -WhyItIsRelevant 'Supports the observations around sharing defaults, external access posture, and collaboration exposure.'
    $collaborationRefs += $reference
    $reference = New-CustomerDocumentationReference -Title 'Set expiration for Microsoft 365 groups' -Url 'https://learn.microsoft.com/en-us/entra/identity/users/groups-lifecycle' -WhyItIsRelevant 'Relevant to ownership, inactivity, and lifecycle controls for Teams-connected groups and collaboration spaces.'
    $collaborationRefs += $reference
    $section = [pscustomobject]@{
        Title      = 'Collaboration'
        Intro      = 'The following guidance supports the collaboration observations related to sharing, ownership, and workspace lifecycle.'
        References = $collaborationRefs
    }
    $sections += $section

    $governanceRefs = @()
    $reference = New-CustomerDocumentationReference -Title 'Learn about retention policies and retention labels' -Url 'https://learn.microsoft.com/en-us/purview/retention' -WhyItIsRelevant 'Supports the governance discussion around retention coverage and the absence or presence of lifecycle controls for retained content.'
    $governanceRefs += $reference
    $reference = New-CustomerDocumentationReference -Title 'Microsoft Secure Score' -Url 'https://learn.microsoft.com/en-us/defender-xdr/microsoft-secure-score' -WhyItIsRelevant 'Provides Microsoft baseline context for interpreting the current Secure Score posture captured in the assessment.'
    $governanceRefs += $reference
    $reference = New-CustomerDocumentationReference -Title 'Admin consent workflow' -Url 'https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-admin-consent-workflow' -WhyItIsRelevant 'Supports the governance observations around application consent, approval workflow, and control maturity.'
    $governanceRefs += $reference
    $section = [pscustomobject]@{
        Title      = 'Data Protection & Governance'
        Intro      = 'These references provide Microsoft guidance for the governance, secure-score, and retention-oriented observations captured in the current-state review.'
        References = $governanceRefs
    }
    $sections += $section

    $lifecycleRefs = @()
    $reference = New-CustomerDocumentationReference -Title 'Remove a former employee and secure data - Step 5: Give access to another employee to OneDrive and Outlook data' -Url 'https://learn.microsoft.com/en-us/microsoft-365/admin/add-users/remove-former-employee-step-5?view=o365-worldwide' -WhyItIsRelevant 'Supports the lifecycle recommendations for former-user content, mailbox ownership, and OneDrive handoff.'
    $lifecycleRefs += $reference
    $reference = New-CustomerDocumentationReference -Title 'Retention and deletion in OneDrive and SharePoint' -Url 'https://learn.microsoft.com/en-us/sharepoint/retention-and-deletion' -WhyItIsRelevant 'Relevant to stale content locations, retained data, and lifecycle handling for dormant collaboration sites.'
    $lifecycleRefs += $reference
    $reference = New-CustomerDocumentationReference -Title 'Convert a mailbox to a shared mailbox' -Url 'https://learn.microsoft.com/en-us/exchange/recipients-in-exchange-online/manage-user-mailboxes/convert-a-mailbox' -WhyItIsRelevant 'Supports mailbox lifecycle handling where access must be preserved but direct user ownership has ended.'
    $lifecycleRefs += $reference
    $section = [pscustomobject]@{
        Title      = 'Offboarding & Lifecycle Management'
        Intro      = 'These Microsoft references support the lifecycle observations around stale accounts, former-user content, and workload cleanup.'
        References = $lifecycleRefs
    }
    $sections += $section

    return @($sections)
}

function Get-CustomerTypeBreakdownText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Rows = @(),
        [Parameter(Mandatory = $false)][string]$PropertyName = 'RecipientTypeDetails',
        [Parameter(Mandatory = $false)][int]$Top = 4
    )

    $groupedItems = @()
    $groupedItems += $Rows |
        ForEach-Object { Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @($PropertyName)) -Default '' } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Group-Object
    $items = @(
        $groupedItems |
            Sort-Object -Property Count, Name -Descending |
            Select-Object -First $Top |
            ForEach-Object { '{0}: {1}' -f $_.Name, $_.Count }
    )

    if ($items.Count -eq 0) {
        return 'Not surfaced in current source'
    }

    return ($items -join '; ')
}

function Get-CustomerRecipientDomainBreakdownText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$DomainRows = @(),
        [Parameter(Mandatory = $false)][object[]]$RecipientRows = @(),
        [Parameter(Mandatory = $false)][int]$Top = 3
    )

    if (@($RecipientRows).Count -gt 0) {
        $domainSummary = @{}
        foreach ($recipientRow in @($RecipientRows)) {
            if ($null -eq $recipientRow) {
                continue
            }

            $primarySmtp = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $recipientRow -Names @('PrimarySmtpAddress', 'WindowsEmailAddress', 'UserPrincipalName')) -Default ''
            if (-not [string]::IsNullOrWhiteSpace($primarySmtp) -and $primarySmtp -match '@(?<domain>[^@\s>]+)$') {
                $domainName = $Matches['domain'].ToLowerInvariant()
                if (-not $domainSummary.ContainsKey($domainName)) {
                    $domainSummary[$domainName] = [ordered]@{ Domain = $domainName; Primary = 0; AliasOnly = 0 }
                }
                $domainSummary[$domainName].Primary++
            }

            $emailAddresses = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $recipientRow -Names @('EmailAddresses'))
            foreach ($emailAddress in @($emailAddresses)) {
                $addressText = Convert-ToArrayaDisplayText -Value $emailAddress -Default ''
                if ([string]::IsNullOrWhiteSpace($addressText)) {
                    continue
                }

                if ($addressText -match '^[A-Za-z]+:(?<value>.+)$') {
                    $addressText = $Matches['value']
                }
                if ($addressText -notmatch '@(?<domain>[^@\s>]+)$') {
                    continue
                }

                $domainName = $Matches['domain'].ToLowerInvariant()
                if (-not $domainSummary.ContainsKey($domainName)) {
                    $domainSummary[$domainName] = [ordered]@{ Domain = $domainName; Primary = 0; AliasOnly = 0 }
                }

                if (-not [string]::IsNullOrWhiteSpace($primarySmtp) -and [string]::Equals($addressText, $primarySmtp, [System.StringComparison]::OrdinalIgnoreCase)) {
                    continue
                }

                $domainSummary[$domainName].AliasOnly++
            }
        }

        $items = @(
            @($domainSummary.Values) |
                ForEach-Object {
                    [pscustomobject]@{
                        Domain    = $_.Domain
                        Total     = ($_.Primary + $_.AliasOnly)
                        Primary   = $_.Primary
                        AliasOnly = $_.AliasOnly
                    }
                } |
                Sort-Object -Property Total, Primary -Descending |
                Select-Object -First $Top |
                ForEach-Object {
                    '{0}: {1} total ({2} primary, {3} alias-only)' -f $_.Domain, $_.Total, $_.Primary, $_.AliasOnly
                }
        )

        if ($items.Count -eq 0) {
            return 'Not surfaced in current source'
        }

        return ($items -join '; ')
    }

    $domainRowsWithSortValue = @()
    foreach ($row in @($DomainRows)) {
        $totalRecipientsRaw = Get-ArrayaObjectValue -Object $row -Names 'TotalDomainRecipients'
        $domainNameRaw = Get-ArrayaObjectValue -Object $row -Names 'Domain'
        $totalRecipients = Convert-ArrayaToNumber $totalRecipientsRaw
        $domainName = Convert-ToArrayaDisplayText -Value $domainNameRaw -Default 'Unknown domain'
        $domainRowsWithSortValue += [pscustomobject]@{
            DomainRow       = $row
            TotalRecipients = $totalRecipients
            DomainName      = $domainName
        }
    }
    $items = @(
        $domainRowsWithSortValue |
            Sort-Object -Property TotalRecipients, DomainName -Descending |
            Select-Object -First $Top |
            ForEach-Object {
                $domainName = $_.DomainName
                $total = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_.DomainRow -Names 'TotalDomainRecipients')
                $primary = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_.DomainRow -Names 'PrimarySMTPRecipients')
                $aliasOnly = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_.DomainRow -Names 'AliasOnlyRecipients')
                $totalDisplay = if ($null -eq $total) { 0 } else { $total }
                $primaryDisplay = if ($null -eq $primary) { 0 } else { $primary }
                $aliasDisplay = if ($null -eq $aliasOnly) { 0 } else { $aliasOnly }
                '{0}: {1} total ({2} primary, {3} alias-only)' -f $domainName, $totalDisplay, $primaryDisplay, $aliasDisplay
            }
    )

    if ($items.Count -eq 0) {
        return 'Not surfaced in current source'
    }

    return ($items -join '; ')
}

function Get-CustomerTopLicensePressureText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$LicenseRows = @(),
        [Parameter(Mandatory = $false)][int]$Top = 3
    )

    $licensePressureRows = @()
    foreach ($license in @($LicenseRows)) {
        $consumedRaw = Get-ArrayaObjectValue -Object $license -Names 'ConsumedUnits'
        $purchasedRaw = Get-ArrayaObjectValue -Object $license -Names @('PurchasedUnits', 'TotalLicenses')
        $skuNameRaw = Get-ArrayaObjectValue -Object $license -Names @('SkuFriendlyName', 'SkuPartNumber')
        $consumed = Convert-ArrayaToNumber $consumedRaw
        $purchased = Convert-ArrayaToNumber $purchasedRaw
        if ($null -eq $consumed -or $null -eq $purchased -or $purchased -le 0) { continue }

        $skuName = Convert-ToArrayaDisplayText -Value $skuNameRaw -Default 'Unknown SKU'
        $percent = [math]::Round(($consumed / $purchased) * 100, 2)
        $licensePressureRows += [pscustomobject]@{
            Sku     = $skuName
            Percent = $percent
            State   = $("{0}: {1}/{2} ({3}%)" -f $skuName, $consumed, $purchased, $percent)
        }
    }

    $items = @(
        $licensePressureRows |
            Sort-Object -Property Percent, Sku -Descending |
            Select-Object -First $Top |
            ForEach-Object { $_.State }
    )

    if ($items.Count -eq 0) {
        return 'Not surfaced in current source'
    }

    return ($items -join '; ')
}

function Get-CustomerExampleText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object[]]$Rows = @(),
        [Parameter(Mandatory = $true)][scriptblock]$Project,
        [Parameter(Mandatory = $false)][int]$Top = 2,
        [Parameter(Mandatory = $false)][string]$Default = 'No specific examples were surfaced in the current source.'
    )

    $items = @(
        foreach ($row in @($Rows | Select-Object -First $Top)) {
            $value = & $Project $row
            $text = Convert-ToArrayaDisplayText -Value $value -Default ''
            if (-not [string]::IsNullOrWhiteSpace($text)) { $text }
        }
    )

    if ($items.Count -eq 0) {
        return $Default
    }

    return ($items -join '; ')
}

function Get-CustomerTechnicalObservations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Signals,
        [Parameter(Mandatory = $true)][object[]]$Findings,
        [Parameter(Mandatory = $false)][object[]]$WorkstreamSummaries = @()
    )

    $observations = New-Object System.Collections.Generic.List[object]

    $adminRows = Convert-ArrayaObjectToArray $Signals.Admins
    $caPolicies = Convert-ArrayaObjectToArray $Signals.ConditionalAccessPolicies
    $caSummaryRecord = if ($Signals.ConditionalAccessSummary) { Get-ArrayaObjectValue -Object $Signals.ConditionalAccessSummary -Names @('Summary') } else { $null }
    $authConfig = $Signals.AuthenticationConfig
    $mfaSummary = $Signals.MfaRegistrationSummary
    $mfaEnrollmentSummaryRecord = Get-CustomerMfaEnrollmentSummary -ExistingSummary $Signals.MfaEnrollmentSummary -MfaSummary $mfaSummary -MfaRegistrationDetails (Convert-ArrayaObjectToArray $Signals.MfaRegistrationDetails)
    $mfaEnforcementSummaryRecord = Get-CustomerMfaEnforcementSummary -ExistingSummary $Signals.MfaEnforcementSummary -ConditionalAccessPolicies $caPolicies -ConditionalAccessSummary $Signals.ConditionalAccessSummary -SecurityDefaultsPolicy $Signals.SecurityDefaultsPolicy
    $enterpriseApps = Convert-ArrayaObjectToArray $Signals.EnterpriseApplications
    $enterpriseAppSummaryRecord = if ($Signals.EnterpriseApplicationSummary) { Get-ArrayaObjectValue -Object $Signals.EnterpriseApplicationSummary -Names @('Summary') } else { $null }
    $guestSummaryRecord = if ($Signals.GuestSignInSummary) { Get-ArrayaObjectValue -Object $Signals.GuestSignInSummary -Names @('Summary') } else { $null }
    $externalIdentityRestrictionsRecord = if ($Signals.ExternalIdentityRestrictions) { Get-ArrayaObjectValue -Object $Signals.ExternalIdentityRestrictions -Names @('Summary') } else { $null }
    $guestAccessConfigurationRecord = if ($Signals.GuestAccessConfiguration) { Get-ArrayaObjectValue -Object $Signals.GuestAccessConfiguration -Names @('Summary') } else { $null }
    $privilegedSummaryRecord = if ($Signals.PrivilegedAccessSummary) { Get-ArrayaObjectValue -Object $Signals.PrivilegedAccessSummary -Names @('Summary') } else { $null }
    $recipientRows = Convert-ArrayaObjectToArray $Signals.AllRecipients
    $retentionPolicyRows = Convert-ArrayaObjectToArray $Signals.RetentionPolicies
    $adminConsentWorkflowEnabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $authConfig -Names @('AdminConsentWorkflowEnabled'))

    $reportOnlyPolicies = @($caPolicies | Where-Object { [string]$_.State -match 'report' }).Count
    $inactiveGuests90Days = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $guestSummaryRecord -Names @('InactiveGuests90Days'))
    $stalePrivileged90Days = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $privilegedSummaryRecord -Names @('StalePrivilegedAccounts90Days'))
    $mfaPercent = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaSummary -Names @('RegistrationPercent'))
    if ($null -eq $mfaPercent) {
        $mfaPercent = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('RegistrationPercent'))
    }
    $usersWithWeakMethodsOnly = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('UsersWithWeakMethodsOnly'))
    $usersWithWeakDefaultMethod = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('UsersWithWeakDefaultMethod'))
    $usersWithPhishingResistantMethods = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('UsersWithPhishingResistantMethods'))
    $enabledMfaPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('EnabledPoliciesRequiringMfa'))
    $reportOnlyMfaPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('ReportOnlyPoliciesRequiringMfa'))
    $mfaEnforcementState = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('EnforcementState')) -Default ''
    $highPrivilegeApps = @($enterpriseApps | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('HighPrivilegePermissionCount'))) -gt 0 }).Count
    $policiesWithExclusions = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $caSummaryRecord -Names @('PoliciesWithExclusions'))
    $policiesUsingRiskSignals = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $caSummaryRecord -Names @('PoliciesUsingRiskSignals', 'RiskBasedPolicyCount'))
    $guestInvitationControl = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $guestAccessConfigurationRecord -Names @('GuestInvitationControl')) -Default ''
    $crossTenantPartnerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('CrossTenantPartnerCount'))
    $defaultInboundMfaTrust = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('DefaultInboundMfaTrust')) -Default ''
    if ($null -eq $policiesUsingRiskSignals) {
        $policiesUsingRiskSignals = @(
            $caPolicies |
                Where-Object {
                    (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('UsesRiskSignals', 'HasRiskSignals'))) -eq $true
                }
        ).Count
    }
    $whatReviewed = 'Privileged administrator assignments, Conditional Access policy state, authentication configuration, MFA registration, enterprise application permissions, and guest and privileged sign-in signals were reviewed.'
    $observedIdentity = @()
    if ($adminRows.Count -gt 0) { $observedIdentity += "$($adminRows.Count) administrator record(s) were present in scope" }
    if ($caPolicies.Count -gt 0) { $observedIdentity += "$($caPolicies.Count) Conditional Access policy/policies were in scope, including $reportOnlyPolicies in report-only mode" }
    if ($null -ne $inactiveGuests90Days) { $observedIdentity += "$inactiveGuests90Days inactive guest account(s) over 90 days were identified" }
    if ($null -ne $stalePrivileged90Days) { $observedIdentity += "$stalePrivileged90Days privileged account(s) showed stale sign-in activity over 90 days" }
    if ($null -ne $mfaPercent) { $observedIdentity += "MFA enrollment was $mfaPercent%" }
    if ($null -ne $usersWithWeakMethodsOnly) { $observedIdentity += "$usersWithWeakMethodsOnly registered user(s) relied only on weaker MFA methods" }
    if ($null -ne $enabledMfaPolicies) { $observedIdentity += "$enabledMfaPolicies enabled Conditional Access policy/policies required MFA" }
    if (-not [string]::IsNullOrWhiteSpace($guestInvitationControl)) { $observedIdentity += "guest invitation control was set to $guestInvitationControl" }
    if ($null -ne $crossTenantPartnerCount) { $observedIdentity += "$crossTenantPartnerCount cross-tenant partner configuration(s) were present" }
    if (-not [string]::IsNullOrWhiteSpace($defaultInboundMfaTrust)) { $observedIdentity += "default inbound MFA trust was set to $defaultInboundMfaTrust" }
    if ($highPrivilegeApps -gt 0) { $observedIdentity += "$highPrivilegeApps enterprise application(s) carried high-privilege permissions" }
    if ($null -ne $policiesWithExclusions) { $observedIdentity += "$policiesWithExclusions Conditional Access policy/policies included exclusions" }
    if ($observedIdentity.Count -eq 0) { $observedIdentity += 'the current source did not show a large identity control concentration, but it did include enough identity signals to evaluate privileged access, policy coverage, and authentication posture' }
    $whyIdentity = 'This matters in this tenant because identity findings are one of the largest concentrations in the report, which indicates that privileged access hygiene and policy enforcement are not keeping pace with the exposure shown in the current administrative and guest-access footprint.'
    $observations.Add((New-CustomerTechnicalObservationSection -SectionTitle 'Identity & Access (Entra ID)' -WhatWasReviewed $whatReviewed -WhatWasObserved ((Join-ArrayaReadableList -Items $observedIdentity) + '.') -WhyItMatters $whyIdentity)) | Out-Null

    $deviceRows = Convert-ArrayaObjectToArray $Signals.DeviceDetails
    $deviceManagementSummaryRecord = if ($Signals.DeviceManagementSummary) { Get-ArrayaObjectValue -Object $Signals.DeviceManagementSummary -Names @('Summary') } else { $null }
    $totalDevices = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $deviceManagementSummaryRecord -Names @('TotalDevices'))
    $unmanagedDevices = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $deviceManagementSummaryRecord -Names @('UnmanagedDevices'))
    $unsupportedOsDevices = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $deviceManagementSummaryRecord -Names @('UnsupportedOsDevices'))
    $nonCompliantDevices = @($deviceRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('IsCompliant', 'Compliant'))) -eq $false }).Count
    $staleDeviceCutoff = (Get-Date).AddDays(-180)
    $staleDevices = @($deviceRows | Where-Object { $lastSeen = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('ApproximateLastSignInDateTime')); $lastSeen -and $lastSeen -lt $staleDeviceCutoff }).Count
    $intuneManaged = @($deviceRows | Where-Object { ([string](Get-ArrayaObjectValue -Object $_ -Names @('MDMSolution', 'ManagementAgent', 'ManagedBy'))) -match 'intune' }).Count
    $intuneCoverage = if ($deviceRows.Count -gt 0) { [math]::Round(($intuneManaged / $deviceRows.Count) * 100, 2) } else { $null }
    $deviceObserved = @()
    if ($null -ne $totalDevices) { $deviceObserved += "$totalDevices device(s) were represented in the management summary" }
    if ($nonCompliantDevices -gt 0) { $deviceObserved += "$nonCompliantDevices device(s) were marked non-compliant" }
    if ($null -ne $unmanagedDevices) { $deviceObserved += "$unmanagedDevices device(s) were shown as unmanaged" }
    if ($staleDevices -gt 0) { $deviceObserved += "$staleDevices device(s) had not signed in within the 180-day stale threshold used in this review" }
    if ($null -ne $unsupportedOsDevices) { $deviceObserved += "$unsupportedOsDevices device(s) were tagged with unsupported operating-system status" }
    if ($null -ne $intuneCoverage) { $deviceObserved += "Intune management coverage was $intuneCoverage%" }
    if ($deviceObserved.Count -eq 0) { $deviceObserved += 'device signals were present, but the current source did not show a large concentration in endpoint control gaps' }
    $whyDevice = 'This matters in this tenant because access control strength depends on devices being both managed and compliant; where the observed device mix is fragmented, policy enforcement becomes less predictable and exception handling grows.'
    $observations.Add((New-CustomerTechnicalObservationSection -SectionTitle 'Devices & Endpoint Management' -WhatWasReviewed 'Device inventory, compliance state, management-source indicators, stale-device activity, and the device management summary were reviewed.' -WhatWasObserved ((Join-ArrayaReadableList -Items $deviceObserved) + '.') -WhyItMatters $whyDevice)) | Out-Null

    $mailboxRows = Convert-ArrayaObjectToArray $Signals.AllMailboxes
    $inboxRulesExternalForwarding = Convert-ArrayaObjectToArray $Signals.InboxRulesExternalForwarding
    $inboxRuleForwardingSummaryRecord = if ($Signals.InboxRuleForwardingSummary) { Get-ArrayaObjectValue -Object $Signals.InboxRuleForwardingSummary -Names @('Summary') } else { $null }
    $forwardingPolicySummaryRecord = if ($Signals.ForwardingPolicySummary) { Get-ArrayaObjectValue -Object $Signals.ForwardingPolicySummary -Names @('Summary') } else { $null }
    $connectorRows = Convert-ArrayaObjectToArray $Signals.MailFlowConnectors
    $remoteDomainRows = Convert-ArrayaObjectToArray $Signals.RemoteDomains
    $sharedMailboxGovernanceSummaryRecord = if ($Signals.SharedMailboxGovernanceSummary) { Get-ArrayaObjectValue -Object $Signals.SharedMailboxGovernanceSummary -Names @('Summary') } else { $null }
    $publicFolderRows = Convert-ArrayaObjectToArray $Signals.PublicFolderDetails
    $forwardedMailboxCount = @($mailboxRows | Where-Object {
        $deliverAndForward = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DeliverToMailboxAndForward'))
        $forwardSmtp = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ForwardingSmtpAddress')) -Default ''
        $forwardAddress = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ForwardingAddress')) -Default ''
        $deliverAndForward -or -not [string]::IsNullOrWhiteSpace($forwardSmtp) -or -not [string]::IsNullOrWhiteSpace($forwardAddress)
    }).Count
    $externalForwardRuleCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $inboxRuleForwardingSummaryRecord -Names @('ExternalForwardingRuleCount'))
    if ($null -eq $externalForwardRuleCount) { $externalForwardRuleCount = $inboxRulesExternalForwarding.Count }
    $policiesAllowingAutoForwarding = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $forwardingPolicySummaryRecord -Names @('PoliciesExplicitlyAllowingAutoForwarding'))
    $remoteDomainsAllowingAutoForwarding = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $forwardingPolicySummaryRecord -Names @('RemoteDomainsAllowingAutoForwarding'))
    $sharedMailboxesWithoutOwnerSignal = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $sharedMailboxGovernanceSummaryRecord -Names @('SharedMailboxesWithoutOwnerSignal'))
    $messagingObserved = @()
    if ($forwardedMailboxCount -gt 0) { $messagingObserved += "$forwardedMailboxCount mailbox(es) showed forwarding configuration" }
    if ($null -ne $externalForwardRuleCount) { $messagingObserved += "$externalForwardRuleCount inbox rule(s) were identified with external forwarding targets" }
    if ($connectorRows.Count -gt 0) { $messagingObserved += "$($connectorRows.Count) mail-flow connector(s) were present in scope" }
    if ($null -ne $policiesAllowingAutoForwarding) { $messagingObserved += "$policiesAllowingAutoForwarding outbound policy/policies explicitly allowed auto-forwarding" }
    if ($null -ne $remoteDomainsAllowingAutoForwarding) { $messagingObserved += "$remoteDomainsAllowingAutoForwarding remote domain(s) still allowed auto-forwarding" }
    if ($null -ne $sharedMailboxesWithoutOwnerSignal) { $messagingObserved += "$sharedMailboxesWithoutOwnerSignal shared mailbox(es) lacked an ownership signal in the governance summary" }
    if ($publicFolderRows.Count -gt 0) { $messagingObserved += "$($publicFolderRows.Count) public folder object(s) remained in the environment" }
    if ($messagingObserved.Count -eq 0) { $messagingObserved += 'the messaging review did not show a large concentration of forwarding or transport exceptions, but mailbox, rule, and connector signals were still present for review' }
    $whyMessaging = 'This matters in this tenant because the review is showing messaging risk at more than one layer: object-level forwarding, tenant-level forwarding policy, and retained legacy objects all affect how difficult it is to control mail flow and data movement.'
    $observations.Add((New-CustomerTechnicalObservationSection -SectionTitle 'Messaging (Exchange Online)' -WhatWasReviewed 'Mailbox forwarding state, inbox-rule forwarding, outbound forwarding policy, mail-flow connectors, remote domains, shared mailbox governance, and public folder inventory were reviewed.' -WhatWasObserved ((Join-ArrayaReadableList -Items $messagingObserved) + '.') -WhyItMatters $whyMessaging)) | Out-Null

    $sharePointRows = Convert-ArrayaObjectToArray $Signals.SharePoint
    $oneDriveRows = Convert-ArrayaObjectToArray $Signals.OneDrive
    $teamRows = Convert-ArrayaObjectToArray $Signals.AllTeams
    $unifiedGroupRows = Convert-ArrayaObjectToArray $Signals.UnifiedGroups
    $sharePointSharingSummaryRecord = if ($Signals.SharePointSharingSummary) { Get-ArrayaObjectValue -Object $Signals.SharePointSharingSummary -Names @('Summary') } else { $null }
    $externalSharingSummaryRecord = if ($Signals.ExternalSharingSummary) { Get-ArrayaObjectValue -Object $Signals.ExternalSharingSummary -Names @('Summary') } else { $null }
    $externalSharingSiteOverrideRows = Convert-ArrayaObjectToArray $Signals.ExternalSharingSiteOverrides
    $collaborationActivitySummaryRecord = if ($Signals.CollaborationActivitySummary) { Get-ArrayaObjectValue -Object $Signals.CollaborationActivitySummary -Names @('Summary') } else { $null }
    $oneDriveOwnerMismatches = Convert-ArrayaObjectToArray $Signals.OneDriveOwnerMismatches
    $ownerlessTeams = @($teamRows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OwnerCount'))) -eq 0 }).Count
    $ownerlessGroups = @($unifiedGroupRows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OwnerCount'))) -eq 0 }).Count
    $guestHeavyTeams = @($teamRows | Where-Object {
        $guestCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('GuestCount'))
        $memberCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('MemberCount'))
        $null -ne $guestCount -and $null -ne $memberCount -and $memberCount -gt 0 -and (($guestCount / $memberCount) -ge 0.4)
    }).Count
    $channelSprawlTeams = @($teamRows | Where-Object {
        $privateCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('PrivateChannelCount'))
        $sharedCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('SharedChannelCount'))
        (($null -ne $privateCount -and $privateCount -ge 5) -or ($null -ne $sharedCount -and $sharedCount -ge 5))
    }).Count
    $staleSharePointSites = @($sharePointRows | Where-Object { $lastModified = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastContentModifiedDate')); $lastModified -and $lastModified -lt (Get-Date).AddDays(-180) }).Count
    $staleOneDrives = @($oneDriveRows | Where-Object { $lastModified = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastContentModifiedDate')); $lastModified -and $lastModified -lt (Get-Date).AddDays(-180) }).Count
    $tenantSharingCapability = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('TenantSharingCapability')) -Default ''
    if ([string]::IsNullOrWhiteSpace($tenantSharingCapability)) {
        $tenantSharingCapability = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('TenantSharingCapability')) -Default ''
    }
    $defaultSharingLinkType = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('DefaultSharingLinkType')) -Default ''
    if ([string]::IsNullOrWhiteSpace($defaultSharingLinkType)) {
        $defaultSharingLinkType = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('DefaultSharingLinkType')) -Default ''
    }
    $sharingDomainRestrictionMode = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('SharingDomainRestrictionMode')) -Default ''
    $sitesWithExternalSharingEnabled = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('SitesWithExternalSharingEnabled'))
    $siteOverrideCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('SiteOverrideCount'))
    if ($null -eq $siteOverrideCount) { $siteOverrideCount = $externalSharingSiteOverrideRows.Count }
    $dormantGroups = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $collaborationActivitySummaryRecord -Names @('DormantGroups'))
    $collabObserved = @(
        "$($sharePointRows.Count) SharePoint site(s) were reviewed",
        "$($oneDriveRows.Count) OneDrive site(s) were reviewed",
        "$($teamRows.Count) Team(s) were reviewed",
        "$($unifiedGroupRows.Count) Microsoft 365 group(s) were reviewed"
    )
    if ($ownerlessTeams -gt 0) { $collabObserved += "$ownerlessTeams Team(s) did not have an owner" }
    if ($ownerlessGroups -gt 0) { $collabObserved += "$ownerlessGroups Microsoft 365 group(s) did not have an owner" }
    if ($oneDriveOwnerMismatches.Count -gt 0) { $collabObserved += "$($oneDriveOwnerMismatches.Count) OneDrive ownership mismatch(es) were identified" }
    if ($staleSharePointSites -gt 0 -or $staleOneDrives -gt 0) { $collabObserved += "$staleSharePointSites stale SharePoint site(s) and $staleOneDrives stale OneDrive site(s) were detected" }
    if ($guestHeavyTeams -gt 0) { $collabObserved += "$guestHeavyTeams Team(s) had a high guest-member ratio" }
    if ($channelSprawlTeams -gt 0) { $collabObserved += "$channelSprawlTeams Team(s) showed elevated private or shared channel counts" }
    if (-not [string]::IsNullOrWhiteSpace($tenantSharingCapability)) { $collabObserved += "tenant sharing was set to $tenantSharingCapability with a default link type of $defaultSharingLinkType" }
    if (-not [string]::IsNullOrWhiteSpace($sharingDomainRestrictionMode)) { $collabObserved += "sharing domain restriction mode was $sharingDomainRestrictionMode" }
    if ($null -ne $sitesWithExternalSharingEnabled) { $collabObserved += "$sitesWithExternalSharingEnabled reviewed site(s) supported external sharing" }
    if ($siteOverrideCount -gt 0) { $collabObserved += "$siteOverrideCount site-level sharing override(s) were identified" }
    if ($null -ne $dormantGroups) { $collabObserved += "$dormantGroups dormant Microsoft 365 group(s) appeared in the collaboration activity summary" }
    $whyCollab = 'This matters in this tenant because collaboration risk is showing up as a mix of ownership gaps, stale content locations, and broad sharing posture. That combination makes lifecycle decisions harder and increases the chance that content remains accessible after accountability has faded.'
    $observations.Add((New-CustomerTechnicalObservationSection -SectionTitle 'Collaboration (Teams, SharePoint, OneDrive)' -WhatWasReviewed 'Teams inventory, Microsoft 365 groups, SharePoint and OneDrive site inventories, sharing configuration, collaboration activity, and ownership-governance signals were reviewed.' -WhatWasObserved ((Join-ArrayaReadableList -Items $collabObserved) + '.') -WhyItMatters $whyCollab)) | Out-Null

    $domainRows = Convert-ArrayaObjectToArray $Signals.Domains
    $licenseRows = Convert-ArrayaObjectToArray $Signals.LicenseSKUs
    $tenantInfoSummaryRecord = if ($Signals.TenantInfoSummary) { Get-ArrayaObjectValue -Object $Signals.TenantInfoSummary -Names @('Summary') } else { $null }
    $secureScoreRows = Convert-ArrayaObjectToArray $Signals.SecuritySecureScore
    $smtpRelaySummary = $Signals.SMTPRelaySummary
    $authConfigSummaryRecord = if ($Signals.AuthenticationConfigSummary) { Get-ArrayaObjectValue -Object $Signals.AuthenticationConfigSummary -Names @('Summary') } else { $null }
    $unverifiedDomains = @($domainRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('IsVerified'))) -eq $false }).Count
    $secureScoreLatest = $secureScoreRows | Sort-Object { Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('CreatedDateTime', 'createdDateTime')) } -Descending | Select-Object -First 1
    $secureScoreCurrent = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $secureScoreLatest -Names @('CurrentScore', 'currentScore'))
    $secureScoreMax = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $secureScoreLatest -Names @('MaxScore', 'maxScore'))
    $secureScorePercent = if ($null -ne $secureScoreCurrent -and $null -ne $secureScoreMax -and $secureScoreMax -gt 0) { [math]::Round(($secureScoreCurrent / $secureScoreMax) * 100, 2) } else { $null }
    $dirSyncEnabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $tenantInfoSummaryRecord -Names @('DirSyncEnabled', 'DirectorySynchronizationEnabled'))
    $smtpAuthUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $smtpRelaySummary -Names @('SMTPAuthUsers'))
    $licenseStress = @()
    foreach ($license in $licenseRows) {
        $consumed = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $license -Names @('ConsumedUnits'))
        $active = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $license -Names @('ActiveUnits', 'EnabledUnits'))
        if ($null -ne $consumed -and $null -ne $active -and $active -gt 0) {
            $pct = [math]::Round(($consumed / $active) * 100, 2)
            if ($pct -ge 95) {
                $licenseStress += ("{0} at {1}%" -f (Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $license -Names @('SkuPartNumber')) -Default 'Unknown SKU'), $pct)
            }
        }
    }
    $governanceObserved = @()
    if ($domainRows.Count -gt 0) { $governanceObserved += "$($domainRows.Count) domain(s) were reviewed, including $unverifiedDomains that were not verified" }
    if ($null -ne $secureScoreCurrent -and $null -ne $secureScoreMax) { $governanceObserved += "Microsoft Secure Score was $secureScoreCurrent out of $secureScoreMax" }
    if ($licenseStress.Count -gt 0) { $governanceObserved += ("license capacity pressure was visible on " + (Join-ArrayaReadableList -Items $licenseStress)) }
    if ($null -ne $smtpAuthUsers) { $governanceObserved += "$smtpAuthUsers account(s) were shown with SMTP authentication usage" }
    if ($null -ne $dirSyncEnabled) { $governanceObserved += ("directory synchronization was " + $(if ($dirSyncEnabled) { 'enabled' } else { 'disabled' })) }
    if ($governanceObserved.Count -eq 0) { $governanceObserved += 'domains, licensing, secure score, and tenant-governance signals were reviewed, but the current source did not show a single dominant concentration in this section' }
    $whyGovernance = 'This matters in this tenant because governance weakness is showing up in foundational controls such as domain hygiene, licensing headroom, and baseline security posture. When those controls drift together, the tenant becomes harder to administer cleanly and more likely to carry avoidable operational risk.'
    $observations.Add((New-CustomerTechnicalObservationSection -SectionTitle 'Data Protection & Governance' -WhatWasReviewed 'Domain inventory, license capacity, secure score, tenant synchronization state, SMTP authentication signals, and governance-related summary data were reviewed.' -WhatWasObserved ((Join-ArrayaReadableList -Items $governanceObserved) + '.') -WhyItMatters $whyGovernance)) | Out-Null

    $userRows = Convert-ArrayaObjectToArray $Signals.Users
    $inactiveLicensedUsers = @(
        $userRows |
            Where-Object {
                $assignedLicenses = Convert-ToArrayaStringList (Get-ArrayaObjectValue -Object $_ -Names @('AssignedLicensesFriendly', 'AssignedLicenses'))
                $enabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))
                $lastSignIn = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastSignInDateTime', 'LastSuccessfulSignInDateTime'))
                $assignedLicenses.Count -gt 0 -and (($enabled -eq $false) -or ($lastSignIn -and $lastSignIn -lt (Get-Date).AddDays(-90)))
            }
    ).Count
    $dormantTeams = @(
        $teamRows | Where-Object {
            $lastActivity = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastActivityDate'))
            $lastActivity -and $lastActivity -lt (Get-Date).AddDays(-90)
        }
    ).Count
    $sharedMailboxesWithoutOwnerSignalLifecycle = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $sharedMailboxGovernanceSummaryRecord -Names @('SharedMailboxesWithoutOwnerSignal'))
    $offboardingObserved = @()
    if ($oneDriveOwnerMismatches.Count -gt 0) { $offboardingObserved += "$($oneDriveOwnerMismatches.Count) OneDrive ownership mismatch(es) suggested content remained after ownership changed" }
    if ($staleSharePointSites -gt 0 -or $staleOneDrives -gt 0) { $offboardingObserved += "$staleSharePointSites stale SharePoint site(s) and $staleOneDrives stale OneDrive site(s) remained in scope" }
    if ($null -ne $inactiveGuests90Days) { $offboardingObserved += "$inactiveGuests90Days guest account(s) were inactive for more than 90 days" }
    if ($null -ne $stalePrivileged90Days) { $offboardingObserved += "$stalePrivileged90Days privileged account(s) showed stale sign-in activity over 90 days" }
    if ($inactiveLicensedUsers -gt 0) { $offboardingObserved += "$inactiveLicensedUsers inactive or disabled user(s) still held paid licenses" }
    if ($dormantTeams -gt 0) { $offboardingObserved += "$dormantTeams Team(s) showed older activity dates consistent with dormancy" }
    if ($null -ne $dormantGroups) { $offboardingObserved += "$dormantGroups dormant Microsoft 365 group(s) appeared in the activity summary" }
    if ($null -ne $sharedMailboxesWithoutOwnerSignalLifecycle) { $offboardingObserved += "$sharedMailboxesWithoutOwnerSignalLifecycle shared mailbox(es) lacked an ownership signal" }
    if ($offboardingObserved.Count -eq 0) { $offboardingObserved += 'identity, collaboration, and mailbox lifecycle signals were reviewed, and the current source did not show a large offboarding backlog, but ownership and activity indicators were still evaluated' }
    $whyOffboarding = 'This matters in this tenant because the same lifecycle pattern appears across identities, collaboration content, and shared workloads: access or data can remain in place after clear day-to-day ownership has faded. That creates both security risk and operational drag when later cleanup depends on reconstructing who is responsible.'
    $observations.Add((New-CustomerTechnicalObservationSection -SectionTitle 'Offboarding & Lifecycle Management' -WhatWasReviewed 'Ownership mismatches, stale collaboration locations, inactive guest and privileged identities, dormant Teams and groups, inactive licensed users, and shared mailbox ownership signals were reviewed.' -WhatWasObserved ((Join-ArrayaReadableList -Items $offboardingObserved) + '.') -WhyItMatters $whyOffboarding)) | Out-Null

    $staleAdminRows = @(
        $adminRows |
            Where-Object {
                $lastSignIn = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastSignInDateTime'))
                $lastSignIn -and $lastSignIn -lt (Get-Date).AddDays(-180)
            } |
            Sort-Object { Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastSignInDateTime')) }
    )
    $globalAdminCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $privilegedSummaryRecord -Names @('GlobalAdministratorCount'))
    if ($null -eq $globalAdminCount) {
        $globalAdminCount = @($adminRows | Where-Object { ([string](Get-ArrayaObjectValue -Object $_ -Names @('Role', 'RolesAssigned'))) -match 'global administrator' }).Count
    }
    $totalPrivilegedIdentities = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $privilegedSummaryRecord -Names @('TotalPrivilegedIdentities'))
    $recipientRowsForMessaging = if ($recipientRows.Count -gt 0) { $recipientRows } else { $mailboxRows }
    $recipientBreakdown = Get-CustomerTypeBreakdownText -Rows $recipientRowsForMessaging -PropertyName 'RecipientTypeDetails' -Top 6
    $domainBreakdown = Get-CustomerRecipientDomainBreakdownText -DomainRows $domainRows -RecipientRows $recipientRowsForMessaging -Top 3
    $recipientBreakdownShort = Get-CustomerCondensedListText -Text $recipientBreakdown -MaxItems 3
    $domainBreakdownShort = Get-CustomerCondensedListText -Text $domainBreakdown -MaxItems 2
    $recipientBreakdownItems = Get-CustomerListItemsFromText -Text $recipientBreakdown -MaxItems 3
    $domainBreakdownItems = Get-CustomerListItemsFromText -Text $domainBreakdown -MaxItems 3
    $verifiedDomains = @($domainRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('IsVerified', 'Verified'))) -eq $true }).Count
    $licensePressureText = Get-CustomerTopLicensePressureText -LicenseRows $licenseRows -Top 3
    $licenseTopTwoText = Get-CustomerTopLicensePressureText -LicenseRows $licenseRows -Top 2
    $retentionCount = if ($retentionPolicyRows.Count -gt 0) { $retentionPolicyRows.Count } else { $null }
    $policyNames = Convert-ToArrayaStringList (Get-ArrayaObjectValue -Object $authConfigSummaryRecord -Names @('PermissionGrantPolicies'))
    if ($policyNames.Count -eq 0) {
        $policyNames = Convert-ToArrayaStringList (Get-ArrayaObjectValue -Object $authConfig -Names @('PermissionGrantPoliciesAssigned', 'PermissionGrantPolicies'))
    }
    $identityExamples = Get-CustomerExampleText -Rows $staleAdminRows -Top 2 -Default 'No specific stale admin examples were surfaced in the current source.' -Project {
        param($row)
        $displayName = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $row -Names @('DisplayName')) -Default 'Unnamed admin'
        $upn = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $row -Names @('UserPrincipalName', 'Mail')) -Default 'UPN not surfaced'
        $lastSignIn = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $row -Names @('LastSignInDateTime'))
        '{0} ({1}, last sign-in {2})' -f $displayName, $upn, $(if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'not surfaced' })
    }
    $forwardedMailboxExamples = Get-CustomerExampleText -Rows @($mailboxRows | Where-Object {
        (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DeliverToMailboxAndForward'))) -or
        -not [string]::IsNullOrWhiteSpace((Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ForwardingSmtpAddress')) -Default ''))
    }) -Top 2 -Default 'No individual forwarded mailbox examples were surfaced in the current source.' -Project {
        param($row)
        $name = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $row -Names @('DisplayName')) -Default 'Unnamed mailbox'
        $address = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $row -Names @('PrimarySmtpAddress')) -Default 'address not surfaced'
        '{0} ({1})' -f $name, $address
    }
    $teamExamples = Get-CustomerExampleText -Rows @($teamRows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OwnerCount'))) -eq 0 }) -Top 2 -Default 'No specific ownerless Team examples were surfaced in the current source.' -Project {
        param($row)
        $displayName = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $row -Names @('DisplayName')) -Default 'Unnamed Team'
        $guestCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $row -Names @('GuestCount'))
        '{0} ({1} guests)' -f $displayName, $(if ($null -eq $guestCount) { 0 } else { $guestCount })
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
    $inactiveLicensedExamples = Get-CustomerExampleText -Rows $inactiveLicensedUsers -Top 2 -Default 'No inactive licensed-user examples were surfaced in the current source.' -Project {
        param($row)
        $name = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $row -Names @('DisplayName', 'UserPrincipalName')) -Default 'Unnamed user'
        $licenses = Convert-ToArrayaStringList (Get-ArrayaObjectValue -Object $row -Names @('AssignedLicensesFriendly', 'AssignedLicenses'))
        '{0} ({1})' -f $name, $(if ($licenses.Count -gt 0) { ($licenses -join ', ') } else { 'license names not surfaced' })
    }

    foreach ($observation in @($observations.ToArray())) {
        switch ($observation.SectionTitle) {
            'Identity & Access (Entra ID)' {
                $observation.ConfigurationRows = @(
                    New-CustomerConfigurationRow -Signal 'Global Administrator count' -State $globalAdminCount
                    New-CustomerConfigurationRow -Signal 'Privileged identities reviewed' -State $(if ($null -eq $totalPrivilegedIdentities) { $adminRows.Count } else { $totalPrivilegedIdentities })
                    New-CustomerConfigurationRow -Signal 'Stale privileged admins (>180 days)' -State $(if ($null -eq $stalePrivileged90Days) { $staleAdminRows.Count } else { $stalePrivileged90Days })
                    New-CustomerConfigurationRow -Signal 'Conditional Access policies' -State $caPolicies.Count
                    New-CustomerConfigurationRow -Signal 'Report-only Conditional Access policies' -State $reportOnlyPolicies
                    New-CustomerConfigurationRow -Signal 'Policies with exclusions' -State $policiesWithExclusions
                    New-CustomerConfigurationRow -Signal 'MFA enrollment rate' -State $(if ($null -ne $mfaPercent) { "$mfaPercent%" } else { $null })
                    New-CustomerConfigurationRow -Signal 'Users with weak MFA methods only' -State $usersWithWeakMethodsOnly
                    New-CustomerConfigurationRow -Signal 'Users with weak default MFA method' -State $usersWithWeakDefaultMethod
                    New-CustomerConfigurationRow -Signal 'Users with phishing-resistant MFA methods' -State $usersWithPhishingResistantMethods
                    New-CustomerConfigurationRow -Signal 'Enabled MFA enforcement policies' -State $enabledMfaPolicies
                    New-CustomerConfigurationRow -Signal 'Report-only MFA enforcement policies' -State $reportOnlyMfaPolicies
                    New-CustomerConfigurationRow -Signal 'MFA enforcement state' -State $mfaEnforcementState
                    New-CustomerConfigurationRow -Signal 'Inactive guest accounts (>90 days)' -State $inactiveGuests90Days
                    New-CustomerConfigurationRow -Signal 'Guest invitation control' -State $guestInvitationControl
                    New-CustomerConfigurationRow -Signal 'Cross-tenant partner count' -State $crossTenantPartnerCount
                    New-CustomerConfigurationRow -Signal 'Default inbound MFA trust' -State $defaultInboundMfaTrust
                    New-CustomerConfigurationRow -Signal 'Admin consent workflow' -State $adminConsentWorkflowEnabled
                    New-CustomerConfigurationRow -Signal 'Permission-grant policies' -State $(if ($policyNames.Count -gt 0) { ($policyNames | Select-Object -First 2) -join '; ' } else { $null })
                    New-CustomerConfigurationRow -Signal 'Example stale privileged identities' -State $identityExamples
                )
                $observation.ObservedNarrative = "Identity stood out because the tenant is carrying $globalAdminCount Global Administrators while also showing $reportOnlyPolicies report-only Conditional Access policies and $policiesWithExclusions policies with exclusions. MFA enrollment is currently $mfaPercent%, but enforcement should be read separately: the reviewed policy baseline shows $enabledMfaPolicies enabled MFA enforcement policies, with $reportOnlyMfaPolicies still in report-only mode. Weak-method usage remains visible in the registration data, which means enrollment quality should be reviewed alongside coverage. Guest invitation control is currently shown as $guestInvitationControl, and cross-tenant partner count is $crossTenantPartnerCount, which means external identity exposure should be reviewed alongside privileged access rather than as a separate track. The stale-admin examples in the table show that some long-lived privileged access remains in place well past what would normally be expected in a tighter operating model."
                $observation.PositiveNarrative = if (($policiesUsingRiskSignals -gt 0) -or ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $authConfig -Names @('MFAEnabled'))) -eq $true)) { "What is working well is that the tenant already has risk-aware Conditional Access coverage in place, MFA is enabled, and the current source contains enough identity telemetry to target cleanup precisely." } else { "What is working well is that this source still surfaces administrator, guest, and policy telemetry clearly enough to support evidence-based identity cleanup." }
            }
            'Devices & Endpoint Management' {
                $observation.ConfigurationRows = @(
                    New-CustomerConfigurationRow -Signal 'Devices represented in current source' -State $(if ($null -ne $totalDevices) { $totalDevices } else { $deviceRows.Count })
                    New-CustomerConfigurationRow -Signal 'Non-compliant devices' -State $nonCompliantDevices
                    New-CustomerConfigurationRow -Signal 'Unmanaged devices' -State $unmanagedDevices
                    New-CustomerConfigurationRow -Signal 'Stale devices (>180 days)' -State $staleDevices
                    New-CustomerConfigurationRow -Signal 'Unsupported operating-system devices' -State $unsupportedOsDevices
                    New-CustomerConfigurationRow -Signal 'Intune management coverage' -State $(if ($null -ne $intuneCoverage) { "$intuneCoverage%" } else { $null })
                )
                $observation.ObservedNarrative = "Endpoint posture is uneven in this tenant. The current source shows $nonCompliantDevices non-compliant devices, $unmanagedDevices unmanaged devices, and $staleDevices stale devices, which indicates the tenant has device visibility but not yet the managed coverage needed for consistent compliance-based access control."
                $observation.PositiveNarrative = if ($deviceRows.Count -gt 0) { "What is working well is that the tenant does have a usable device inventory and management summary. That makes it possible to separate compliance gaps from simple data gaps." } else { "What is working well is that endpoint visibility is still represented in the source, even though detailed device state is limited." }
            }
            'Messaging (Exchange Online)' {
                $observation.ConfigurationRows = @(
                    New-CustomerConfigurationRow -Signal 'Recipients in current source' -State $recipientRowsForMessaging.Count
                    New-CustomerConfigurationRow -Signal 'Mailboxes with forwarding configured' -State $forwardedMailboxCount
                    New-CustomerConfigurationRow -Signal 'Inbox rules forwarding externally' -State $externalForwardRuleCount
                    New-CustomerConfigurationRow -Signal 'Mail-flow connectors' -State $connectorRows.Count
                    New-CustomerConfigurationRow -Signal 'Remote domains allowing auto-forwarding' -State $remoteDomainsAllowingAutoForwarding
                    New-CustomerConfigurationRow -Signal 'Shared mailboxes without ownership signal' -State $sharedMailboxesWithoutOwnerSignal
                    New-CustomerConfigurationRow -Signal 'Public folders still present' -State $publicFolderRows.Count
                    New-CustomerConfigurationRow -Signal 'Example forwarded mailboxes' -State $forwardedMailboxExamples
                )
                $typeIndex = 0
                foreach ($item in $recipientBreakdownItems) {
                    $typeIndex++
                    $observation.ConfigurationRows += New-CustomerConfigurationRow -Signal ("Top recipient type {0}" -f $typeIndex) -State $item
                }
                $domainIndex = 0
                foreach ($item in $domainBreakdownItems) {
                    $domainIndex++
                    $observation.ConfigurationRows += New-CustomerConfigurationRow -Signal ("Top recipient domain {0}" -f $domainIndex) -State $item
                }
                $observation.ObservedNarrative = "Messaging is carrying both scale and control complexity. The tenant has $($recipientRowsForMessaging.Count) recipients in scope, and the largest share of that population sits in $recipientBreakdownShort. The domain footprint is centered on $domainBreakdownShort, which means mailbox-governance decisions affect a concentrated part of the namespace rather than a small edge case. Against that backdrop, $forwardedMailboxCount mailbox(es) still show forwarding and $sharedMailboxesWithoutOwnerSignal shared mailboxes do not yet show an ownership signal."
                $observation.PositiveNarrative = if ($externalForwardRuleCount -eq 0) { "What is working well is that the current source does not show inbox-rule-based external forwarding. The messaging inventory is also detailed enough to separate recipient, transport, and mailbox-governance concerns." } else { "What is working well is that the tenant messaging inventory is detailed enough to distinguish mailbox behavior from transport-level forwarding controls." }
            }
            'Collaboration (Teams, SharePoint, OneDrive)' {
                $observation.ConfigurationRows = @(
                    New-CustomerConfigurationRow -Signal 'SharePoint sites reviewed' -State $sharePointRows.Count
                    New-CustomerConfigurationRow -Signal 'OneDrive sites reviewed' -State $oneDriveRows.Count
                    New-CustomerConfigurationRow -Signal 'Teams reviewed' -State $teamRows.Count
                    New-CustomerConfigurationRow -Signal 'Microsoft 365 groups reviewed' -State $unifiedGroupRows.Count
                    New-CustomerConfigurationRow -Signal 'Ownerless Teams' -State $ownerlessTeams
                    New-CustomerConfigurationRow -Signal 'Ownerless Microsoft 365 groups' -State $ownerlessGroups
                    New-CustomerConfigurationRow -Signal 'Stale SharePoint locations' -State $staleSharePointSites
                    New-CustomerConfigurationRow -Signal 'Stale OneDrive locations' -State $staleOneDrives
                    New-CustomerConfigurationRow -Signal 'Guest-heavy Teams' -State $guestHeavyTeams
                    New-CustomerConfigurationRow -Signal 'Sharing posture' -State "$tenantSharingCapability; default link type $defaultSharingLinkType"
                    New-CustomerConfigurationRow -Signal 'Sharing domain restriction mode' -State $sharingDomainRestrictionMode
                    New-CustomerConfigurationRow -Signal 'Sites with external sharing enabled' -State $sitesWithExternalSharingEnabled
                    New-CustomerConfigurationRow -Signal 'Site-level sharing overrides' -State $siteOverrideCount
                    New-CustomerConfigurationRow -Signal 'Dormant Microsoft 365 groups' -State $dormantGroups
                    New-CustomerConfigurationRow -Signal 'Example ownerless Teams' -State $teamExamples
                )
                $observation.ObservedNarrative = "Collaboration is carrying a clear ownership-and-lifecycle imbalance. The current source shows $($sharePointRows.Count) SharePoint sites, $($oneDriveRows.Count) OneDrive sites, $($teamRows.Count) Teams, and $($unifiedGroupRows.Count) Microsoft 365 groups. Within that footprint, $ownerlessTeams Team(s) and $ownerlessGroups group(s) do not show ownership, while $staleSharePointSites SharePoint sites and $staleOneDrives OneDrive locations already appear stale. External sharing is not just enabled at the tenant level; $sitesWithExternalSharingEnabled reviewed site(s) surfaced external-sharing capability and $siteOverrideCount site-level sharing override(s) were identified, which shows that exposure is being shaped at both the organization and workload layers. The examples in the table make that drift more tangible by showing specific workspaces where ownership has not kept pace with collaboration growth."
                $observation.PositiveNarrative = if ($oneDriveOwnerMismatches.Count -eq 0 -and ($null -ne $dormantGroups -and $dormantGroups -eq 0)) { "What is working well is that the current source does not show OneDrive ownership mismatches or a large dormant-group backlog. That suggests collaboration telemetry is healthy even where governance follow-through is uneven." } else { "What is working well is that this source separates ownership, activity, and sharing posture clearly enough to show where collaboration drift is concentrated." }
            }
            'Data Protection & Governance' {
                $observation.ConfigurationRows = @(
                    New-CustomerConfigurationRow -Signal 'Domains reviewed' -State $domainRows.Count
                    New-CustomerConfigurationRow -Signal 'Verified / unverified domains' -State "$verifiedDomains verified; $unverifiedDomains unverified"
                    New-CustomerConfigurationRow -Signal 'Secure Score' -State $(if ($null -ne $secureScoreCurrent -and $null -ne $secureScoreMax) { "$secureScoreCurrent / $secureScoreMax ($secureScorePercent%)" } else { $null })
                    New-CustomerConfigurationRow -Signal 'Highest license utilization' -State $licensePressureText
                    New-CustomerConfigurationRow -Signal 'Directory synchronization' -State $(if ($null -eq $dirSyncEnabled) { $null } elseif ($dirSyncEnabled) { 'Enabled' } else { 'Disabled' })
                    New-CustomerConfigurationRow -Signal 'SMTP-authenticated accounts' -State $smtpAuthUsers
                    New-CustomerConfigurationRow -Signal 'Retention policies surfaced' -State $retentionCount
                    New-CustomerConfigurationRow -Signal 'Admin consent workflow' -State $adminConsentWorkflowEnabled
                )
                $observation.ObservedNarrative = "Governance signals show a mixed baseline rather than an empty one. The current source shows $($domainRows.Count) domains, Secure Score at $(if ($null -ne $secureScorePercent) { "$secureScorePercent%" } else { 'not surfaced' }), and directory synchronization $(if ($null -eq $dirSyncEnabled) { 'not surfaced' } elseif ($dirSyncEnabled) { 'enabled' } else { 'disabled' }). At the same time, license saturation is visible on $licenseTopTwoText, and retention policies are $(if ($null -eq $retentionCount) { 'not surfaced in this source' } else { "surfaced as $retentionCount policy objects" })."
                $observation.PositiveNarrative = if ($unverifiedDomains -eq 0) { "What is working well is that the current source does not show unverified domains. Secure Score telemetry and sync-state visibility are also present, which gives the tenant a stronger governance baseline than a blind environment." } else { "What is working well is that foundational governance telemetry is present across domains, licenses, sync, and Secure Score, which supports targeted remediation." }
            }
            'Offboarding & Lifecycle Management' {
                $observation.ConfigurationRows = @(
                    New-CustomerConfigurationRow -Signal 'Inactive guest accounts (>90 days)' -State $inactiveGuests90Days
                    New-CustomerConfigurationRow -Signal 'Stale privileged accounts (>90 days)' -State $stalePrivileged90Days
                    New-CustomerConfigurationRow -Signal 'Inactive or disabled licensed users' -State $inactiveLicensedUsers.Count
                    New-CustomerConfigurationRow -Signal 'Stale SharePoint / OneDrive locations' -State "$staleSharePointSites SharePoint; $staleOneDrives OneDrive"
                    New-CustomerConfigurationRow -Signal 'Dormant Teams / groups' -State "$dormantTeams Teams; $(if ($null -eq $dormantGroups) { 'Not surfaced in current source' } else { $dormantGroups }) groups"
                    New-CustomerConfigurationRow -Signal 'Shared mailboxes without ownership signal' -State $sharedMailboxesWithoutOwnerSignalLifecycle
                    New-CustomerConfigurationRow -Signal 'OneDrive ownership mismatches' -State $oneDriveOwnerMismatches.Count
                    New-CustomerConfigurationRow -Signal 'Example inactive licensed users' -State $inactiveLicensedExamples
                )
                $observation.ObservedNarrative = "Lifecycle signals show that stale access and stale content are accumulating in parallel. The current source shows $inactiveGuests90Days inactive guests, $stalePrivileged90Days stale privileged identities, $($inactiveLicensedUsers.Count) inactive or disabled licensed users, and $sharedMailboxesWithoutOwnerSignalLifecycle shared mailboxes without ownership signal. The examples in the table show that this is not just a count issue; it is affecting named identities and workloads that now need an ownership decision."
                $observation.PositiveNarrative = if ($oneDriveOwnerMismatches.Count -eq 0 -and ($null -ne $dormantGroups -and $dormantGroups -eq 0)) { "What is working well is that the current source does not show OneDrive ownership mismatches and does not indicate a large dormant Microsoft 365 group backlog. Not every lifecycle indicator is drifting at the same rate." } else { "What is working well is that stale identities, stale content, and shared-mailbox ownership are all surfaced separately, which gives cleanup efforts a usable starting point." }
            }
        }
    }

    return @($observations.ToArray())
}

function Get-CustomerModelLeadershipDecisionText {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Action)

    switch ([string]$Action.ActionTitle) {
        'Reduce privileged access and strengthen identity controls' { return 'Approve the reduction of standing privileged access and the move from observation-only identity controls into the supported baseline.' }
        'Establish accountable ownership for collaboration spaces' { return 'Approve an ownership model for collaboration spaces, including when stale workspaces should be retained, transferred, or retired.' }
        'Strengthen domain and anti-spoofing controls' { return 'Approve the domain and mail-authentication baseline so verification, SPF, DKIM, and DMARC cleanup can be completed consistently.' }
        'Improve device compliance and managed endpoint coverage' { return 'Approve the expected endpoint baseline and the exception path for devices that should not retain access while non-compliant.' }
        'Review external forwarding and mail flow exposure' { return 'Approve the external forwarding and mail-flow standard so exceptions can be formally validated or removed.' }
        'Reconcile license capacity and tenant governance gaps' { return 'Approve the capacity and governance cleanup path so constrained licensing and governance gaps can be resolved together.' }
        'Strengthen baseline security and access protections' { return 'Approve the security baseline changes required to close the highest-value protection gaps first.' }
        default { return 'Approve the operating model, accountable owner, and remediation sequence for this work item.' }
    }
}

function Get-CustomerExecutiveDecisionSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$ExecutiveThemes,
        [Parameter(Mandatory = $false)][object[]]$RoadmapActions = @(),
        [Parameter(Mandatory = $false)][object[]]$OwnerGroups = @()
    )

    $themeNames = @($ExecutiveThemes | Select-Object -First 3 | ForEach-Object { [string]$_.Theme })
    $ownerNames = if (@($OwnerGroups).Count -gt 0) {
        @($OwnerGroups | Select-Object -First 3 | ForEach-Object { [string]$_.Name })
    }
    else {
        @()
    }

    $riskRows = @(
        foreach ($theme in @($ExecutiveThemes | Select-Object -First 3)) {
            [pscustomobject]@{
                RiskCluster             = [string]$theme.Theme
                WhatStandsOut           = [string]$theme.WhatThisMeans
                WhyLeadershipShouldCare = [string]$theme.WhyItMatters
            }
        }
    )

    $decisionRows = @(
        foreach ($action in @($RoadmapActions | Select-Object -First 3)) {
            [pscustomobject]@{
                DecisionFocus        = Convert-ToArrayaDisplayText -Value $action.ActionTitle -Default 'Priority work item'
                WhatShouldHappenNext = Get-CustomerModelLeadershipDecisionText -Action $action
                WhyNow               = Convert-ToArrayaDisplayText -Value $action.WhyItMatters -Default 'The current review shows this as one of the highest-value decision points.'
            }
        }
    )

    if ($decisionRows.Count -eq 0) {
        $decisionRows = @(
            [pscustomobject]@{
                DecisionFocus        = 'Priority work item not clearly surfaced'
                WhatShouldHappenNext = 'Approve the remediation path that best matches the reviewed evidence.'
                WhyNow               = 'The current review did not include a clear roadmap ordering, so the grouped findings should guide the next approval.'
            }
        )
    }

    $themeText = if ($themeNames.Count -gt 0) { Join-ArrayaReadableList -Items $themeNames } else { 'the highest-risk Microsoft 365 control areas' }
    $ownerText = if ($ownerNames.Count -gt 0) { Join-ArrayaReadableList -Items $ownerNames } else { 'the tenant overall' }

    return [pscustomobject]@{
        Narrative    = "The report shows the clearest risk concentration in $ownerText. The leading risk clusters of $themeText indicate repeated control drift rather than a single isolated exception, so the brief below focuses on the decisions that remove the biggest blockers to a cleaner operating baseline."
        RiskRows     = @($riskRows)
        DecisionRows = @($decisionRows)
    }
}

function Get-CustomerConsultativeSummaries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Signals,
        [Parameter(Mandatory = $false)][object[]]$Findings = @(),
        [Parameter(Mandatory = $false)][object[]]$ExecutiveThemes = @(),
        [Parameter(Mandatory = $false)][object[]]$RoadmapActions = @(),
        [Parameter(Mandatory = $false)][object[]]$OwnerGroups = @()
    )

    $adminRows = Convert-ArrayaObjectToArray $Signals.Admins
    $userRows = Convert-ArrayaObjectToArray $Signals.Users
    $guestSummaryRecord = if ($Signals.GuestSignInSummary) { Get-ArrayaObjectValue -Object $Signals.GuestSignInSummary -Names @('Summary') } else { $null }
    $privilegedSummaryRecord = if ($Signals.PrivilegedAccessSummary) { Get-ArrayaObjectValue -Object $Signals.PrivilegedAccessSummary -Names @('Summary') } else { $null }
    $mfaSummary = $Signals.MfaRegistrationSummary
    $caPolicies = Convert-ArrayaObjectToArray $Signals.ConditionalAccessPolicies
    $mfaEnrollmentSummaryRecord = Get-CustomerMfaEnrollmentSummary -ExistingSummary $Signals.MfaEnrollmentSummary -MfaSummary $mfaSummary -MfaRegistrationDetails (Convert-ArrayaObjectToArray $Signals.MfaRegistrationDetails)
    $mfaEnforcementSummaryRecord = Get-CustomerMfaEnforcementSummary -ExistingSummary $Signals.MfaEnforcementSummary -ConditionalAccessPolicies $caPolicies -ConditionalAccessSummary $Signals.ConditionalAccessSummary -SecurityDefaultsPolicy $Signals.SecurityDefaultsPolicy
    $caSummaryRecord = if ($Signals.ConditionalAccessSummary) { Get-ArrayaObjectValue -Object $Signals.ConditionalAccessSummary -Names @('Summary') } else { $null }
    $recipientRows = Convert-ArrayaObjectToArray $Signals.AllRecipients
    $mailboxRows = Convert-ArrayaObjectToArray $Signals.AllMailboxes
    $archiveMailboxRows = Convert-ArrayaObjectToArray $Signals.ArchiveMailboxes
    $archiveMailboxStatsRows = Convert-ArrayaObjectToArray $Signals.ArchiveMailboxStats
    $inactiveMailboxRows = Convert-ArrayaObjectToArray $Signals.InactiveMailboxes
    $inboxRulesExternalForwarding = Convert-ArrayaObjectToArray $Signals.InboxRulesExternalForwarding
    $inboxRuleForwardingSummaryRecord = if ($Signals.InboxRuleForwardingSummary) { Get-ArrayaObjectValue -Object $Signals.InboxRuleForwardingSummary -Names @('Summary') } else { $null }
    $forwardingPolicySummaryRecord = if ($Signals.ForwardingPolicySummary) { Get-ArrayaObjectValue -Object $Signals.ForwardingPolicySummary -Names @('Summary') } else { $null }
    $connectorRows = Convert-ArrayaObjectToArray $Signals.MailFlowConnectors
    $sharedMailboxGovernanceSummaryRecord = if ($Signals.SharedMailboxGovernanceSummary) { Get-ArrayaObjectValue -Object $Signals.SharedMailboxGovernanceSummary -Names @('Summary') } else { $null }
    $publicFolderRows = Convert-ArrayaObjectToArray $Signals.PublicFolderDetails
    $sharePointRows = Convert-ArrayaObjectToArray $Signals.SharePoint
    $oneDriveRows = Convert-ArrayaObjectToArray $Signals.OneDrive
    $teamRows = Convert-ArrayaObjectToArray $Signals.AllTeams
    $unifiedGroupRows = Convert-ArrayaObjectToArray $Signals.UnifiedGroups
    $externalSharingSummaryRecord = if ($Signals.ExternalSharingSummary) { Get-ArrayaObjectValue -Object $Signals.ExternalSharingSummary -Names @('Summary') } else { $null }
    $sharePointSharingSummaryRecord = if ($Signals.SharePointSharingSummary) { Get-ArrayaObjectValue -Object $Signals.SharePointSharingSummary -Names @('Summary') } else { $null }
    $externalExposureFindings = Convert-ArrayaObjectToArray $Signals.ExternalExposureFindings
    $externalSharingSiteOverrideRows = Convert-ArrayaObjectToArray $Signals.ExternalSharingSiteOverrides
    $oneDriveOwnerMismatches = Convert-ArrayaObjectToArray $Signals.OneDriveOwnerMismatches
    $domainRows = Convert-ArrayaObjectToArray $Signals.Domains
    $licenseRows = Convert-ArrayaObjectToArray $Signals.LicenseSKUs
    $secureScoreRows = Convert-ArrayaObjectToArray $Signals.SecuritySecureScore
    $smtpRelaySummary = $Signals.SMTPRelaySummary
    $tenantInfoSummaryRecord = if ($Signals.TenantInfoSummary) { Get-ArrayaObjectValue -Object $Signals.TenantInfoSummary -Names @('Summary') } else { $null }

    $globalAdminCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $privilegedSummaryRecord -Names @('GlobalAdministratorCount'))
    if ($null -eq $globalAdminCount) {
        $globalAdminCount = @($adminRows | Where-Object { ([string](Get-ArrayaObjectValue -Object $_ -Names @('Role', 'RolesAssigned'))) -match 'global administrator' }).Count
    }
    $stalePrivileged90Days = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $privilegedSummaryRecord -Names @('StalePrivilegedAccounts90Days'))
    $inactiveGuests90Days = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $guestSummaryRecord -Names @('InactiveGuests90Days'))
    $mfaPercent = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('RegistrationPercent'))
    $policiesWithExclusions = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $caSummaryRecord -Names @('PoliciesWithExclusions'))
    $usersWithWeakMethodsOnly = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('UsersWithWeakMethodsOnly'))
    $usersWithWeakDefaultMethod = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('UsersWithWeakDefaultMethod'))
    $usersWithPhishingResistantMethods = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('UsersWithPhishingResistantMethods'))
    $enabledMfaPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('EnabledPoliciesRequiringMfa'))
    $reportOnlyMfaPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('ReportOnlyPoliciesRequiringMfa'))
    $mfaEnforcementState = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $mfaEnforcementSummaryRecord -Names @('EnforcementState')) -Default $null

    $forwardedMailboxCount = @($mailboxRows | Where-Object {
        $deliverAndForward = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DeliverToMailboxAndForward'))
        $forwardSmtp = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ForwardingSmtpAddress')) -Default ''
        $forwardAddress = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('ForwardingAddress')) -Default ''
        $deliverAndForward -or -not [string]::IsNullOrWhiteSpace($forwardSmtp) -or -not [string]::IsNullOrWhiteSpace($forwardAddress)
    }).Count
    $externalForwardRuleCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $inboxRuleForwardingSummaryRecord -Names @('ExternalForwardingRuleCount'))
    if ($null -eq $externalForwardRuleCount) { $externalForwardRuleCount = $inboxRulesExternalForwarding.Count }
    $policiesAllowingAutoForwarding = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $forwardingPolicySummaryRecord -Names @('PoliciesExplicitlyAllowingAutoForwarding'))
    $remoteDomainsAllowingAutoForwarding = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $forwardingPolicySummaryRecord -Names @('RemoteDomainsAllowingAutoForwarding'))
    $sharedMailboxesWithoutOwnerSignal = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $sharedMailboxGovernanceSummaryRecord -Names @('SharedMailboxesWithoutOwnerSignal'))
    $oversizedSharedMailboxes = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $sharedMailboxGovernanceSummaryRecord -Names @('OversizedSharedMailboxes'))
    $mailboxesWithHoldSignals = @(
        $mailboxRows | Where-Object {
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('LitigationHoldEnabled'))) -eq $true -or
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('RetentionHoldEnabled'))) -eq $true -or
            (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DelayHoldApplied'))) -eq $true
        }
    ).Count
    $archiveMailboxesOverFiftyGb = @(
        $archiveMailboxStatsRows | Where-Object {
            (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('TotalItemSizeBytes'))) -ge 53687091200
        }
    ).Count
    $recipientRowsForMessaging = if ($recipientRows.Count -gt 0) { $recipientRows } else { $mailboxRows }
    $recipientBreakdown = Get-CustomerTypeBreakdownText -Rows $recipientRowsForMessaging -PropertyName 'RecipientTypeDetails' -Top 3
    $domainBreakdown = Get-CustomerRecipientDomainBreakdownText -DomainRows $domainRows -RecipientRows $recipientRowsForMessaging -Top 2
    $smtpAuthUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $smtpRelaySummary -Names @('SMTPAuthUsers'))

    $ownerlessTeams = @($teamRows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OwnerCount'))) -eq 0 }).Count
    $ownerlessGroups = @($unifiedGroupRows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OwnerCount'))) -eq 0 }).Count
    $guestHeavyTeams = @($teamRows | Where-Object {
        $guestCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('GuestCount'))
        $memberCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('MemberCount'))
        $null -ne $guestCount -and $null -ne $memberCount -and $memberCount -gt 0 -and (($guestCount / $memberCount) -ge 0.4)
    }).Count
    $staleSharePointSites = @($sharePointRows | Where-Object { $lastModified = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastContentModifiedDate')); $lastModified -and $lastModified -lt (Get-Date).AddDays(-180) }).Count
    $staleOneDrives = @($oneDriveRows | Where-Object { $lastModified = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastContentModifiedDate')); $lastModified -and $lastModified -lt (Get-Date).AddDays(-180) }).Count
    $tenantSharingCapability = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('TenantSharingCapability')) -Default ''
    if ([string]::IsNullOrWhiteSpace($tenantSharingCapability)) {
        $tenantSharingCapability = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('TenantSharingCapability')) -Default ''
    }
    $defaultSharingLinkType = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('DefaultSharingLinkType')) -Default ''
    if ([string]::IsNullOrWhiteSpace($defaultSharingLinkType)) {
        $defaultSharingLinkType = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $sharePointSharingSummaryRecord -Names @('DefaultSharingLinkType')) -Default ''
    }
    $siteOverrideCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('SiteOverrideCount'))
    if ($null -eq $siteOverrideCount) { $siteOverrideCount = $externalSharingSiteOverrideRows.Count }
    $sharingDomainRestrictionMode = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('SharingDomainRestrictionMode')) -Default ''

    $secureScoreLatest = $secureScoreRows | Sort-Object { Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('CreatedDateTime', 'createdDateTime')) } -Descending | Select-Object -First 1
    $secureScoreCurrent = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $secureScoreLatest -Names @('CurrentScore', 'currentScore'))
    $secureScoreMax = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $secureScoreLatest -Names @('MaxScore', 'maxScore'))
    $secureScorePercent = if ($null -ne $secureScoreCurrent -and $null -ne $secureScoreMax -and $secureScoreMax -gt 0) { [math]::Round(($secureScoreCurrent / $secureScoreMax) * 100, 2) } else { $null }
    $unverifiedDomains = @($domainRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('IsVerified', 'Verified'))) -eq $false }).Count
    $dmarcMissingCount = @($domainRows | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DmarcConfigured'))) -ne $true }).Count
    $licensePressure = Get-CustomerTopLicensePressureText -LicenseRows $licenseRows -Top 3
    $dirSyncEnabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $tenantInfoSummaryRecord -Names @('DirSyncEnabled', 'DirectorySynchronizationEnabled'))
    $guestMfaExperienceSummary = Get-CustomerGuestMfaExperienceSummary `
        -ExternalIdentityRestrictionsRecord $externalIdentityRestrictionsRecord `
        -GuestAccessConfigurationRecord $guestAccessConfigurationRecord `
        -MfaEnforcementSummaryRecord $mfaEnforcementSummaryRecord

    $inactiveLicensedUsers = @(
        $userRows |
            Where-Object {
                $assignedLicenses = Convert-ToArrayaStringList (Get-ArrayaObjectValue -Object $_ -Names @('AssignedLicensesFriendly', 'AssignedLicenses'))
                $enabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('AccountEnabled', 'Enabled'))
                $lastSignIn = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastSignInDateTime', 'LastSuccessfulSignInDateTime'))
                $assignedLicenses.Count -gt 0 -and (($enabled -eq $false) -or ($lastSignIn -and $lastSignIn -lt (Get-Date).AddDays(-90)))
            }
    ).Count
    $dormantTeams = @(
        $teamRows | Where-Object {
            $lastActivity = Convert-ArrayaToDate (Get-ArrayaObjectValue -Object $_ -Names @('LastActivityDate'))
            $lastActivity -and $lastActivity -lt (Get-Date).AddDays(-90)
        }
    ).Count

    return [pscustomobject]@{
        ExecutiveDecisionSummary = Get-CustomerExecutiveDecisionSummary -ExecutiveThemes $ExecutiveThemes -RoadmapActions $RoadmapActions -OwnerGroups $OwnerGroups
        GuestMfaExperienceSummary = $guestMfaExperienceSummary
        IdentityConsultativeSummary = (New-CustomerConsultativeSummary -Title 'IdentityConsultativeSummary' -SnapshotRows @(
            (New-CustomerConfigurationRow -Signal 'Global Administrator count' -State $globalAdminCount),
            (New-CustomerConfigurationRow -Signal 'Stale privileged accounts (>90 days)' -State $stalePrivileged90Days),
            (New-CustomerConfigurationRow -Signal 'Inactive guest accounts (>90 days)' -State $inactiveGuests90Days),
            (New-CustomerConfigurationRow -Signal 'MFA enrollment rate' -State $(if ($null -ne $mfaPercent) { "$mfaPercent%" } else { $null })),
            (New-CustomerConfigurationRow -Signal 'Registered method mix' -State (Get-ArrayaObjectValue -Object $mfaEnrollmentSummaryRecord -Names @('RegisteredMethodBreakdown'))),
            (New-CustomerConfigurationRow -Signal 'Users with weak MFA methods only' -State $usersWithWeakMethodsOnly),
            (New-CustomerConfigurationRow -Signal 'Users with weak default MFA method' -State $usersWithWeakDefaultMethod),
            (New-CustomerConfigurationRow -Signal 'Users with phishing-resistant MFA methods' -State $usersWithPhishingResistantMethods),
            (New-CustomerConfigurationRow -Signal 'Conditional Access policies reviewed' -State $caPolicies.Count),
            (New-CustomerConfigurationRow -Signal 'Enabled MFA enforcement policies' -State $enabledMfaPolicies),
            (New-CustomerConfigurationRow -Signal 'Report-only MFA enforcement policies' -State $reportOnlyMfaPolicies),
            (New-CustomerConfigurationRow -Signal 'Policies with exclusions' -State $policiesWithExclusions),
            (New-CustomerConfigurationRow -Signal 'MFA enforcement state' -State $mfaEnforcementState)
        ) -Narrative 'Identity risk is clustering around privileged-access hygiene, guest lifecycle, uneven MFA enrollment quality, and policy enforcement that is not yet consistently active across the reviewed baseline.' -RecommendationSupport 'This section supports the identity and access recommendations in 4.0.')
        MessagingConsultativeSummary = (New-CustomerConsultativeSummary -Title 'MessagingConsultativeSummary' -SnapshotRows @(
            (New-CustomerConfigurationRow -Signal 'Recipients in scope' -State $(if ($recipientRowsForMessaging.Count -gt 0) { $recipientRowsForMessaging.Count } else { $null })),
            (New-CustomerConfigurationRow -Signal 'Recipient mix by type' -State $recipientBreakdown),
            (New-CustomerConfigurationRow -Signal 'Top recipient domains' -State $domainBreakdown),
            (New-CustomerConfigurationRow -Signal 'Mailboxes with forwarding' -State $forwardedMailboxCount),
            (New-CustomerConfigurationRow -Signal 'Inbox rules with external forwarding' -State $externalForwardRuleCount),
            (New-CustomerConfigurationRow -Signal 'Connectors in scope' -State $connectorRows.Count),
            (New-CustomerConfigurationRow -Signal 'Remote domains allowing auto-forwarding' -State $remoteDomainsAllowingAutoForwarding)
        ) -SecondaryRows @(
            (New-CustomerConfigurationRow -Signal 'Shared mailboxes without owner signal' -State $sharedMailboxesWithoutOwnerSignal),
            (New-CustomerConfigurationRow -Signal 'Oversized shared mailboxes' -State $oversizedSharedMailboxes),
            (New-CustomerConfigurationRow -Signal 'Inactive mailboxes' -State $inactiveMailboxRows.Count),
            (New-CustomerConfigurationRow -Signal 'Mailboxes with hold signals' -State $mailboxesWithHoldSignals),
            (New-CustomerConfigurationRow -Signal 'Archive-enabled mailboxes' -State $archiveMailboxRows.Count),
            (New-CustomerConfigurationRow -Signal 'Archive mailboxes over 50 GB' -State $archiveMailboxesOverFiftyGb)
        ) -TertiaryRows @(
            (New-CustomerConfigurationRow -Signal 'Outbound policies allowing auto-forwarding' -State $policiesAllowingAutoForwarding),
            (New-CustomerConfigurationRow -Signal 'Remote domains allowing auto-forwarding' -State $remoteDomainsAllowingAutoForwarding),
            (New-CustomerConfigurationRow -Signal 'SMTP-authenticated accounts' -State $smtpAuthUsers),
            (New-CustomerConfigurationRow -Signal 'Public folder objects' -State $publicFolderRows.Count),
            (New-CustomerConfigurationRow -Signal 'External forwarding exposure' -State ("{0} inbox rule(s); {1} mailbox(es) with forwarding" -f $(if ($null -eq $externalForwardRuleCount) { 0 } else { $externalForwardRuleCount }), $forwardedMailboxCount))
        ) -Narrative 'Messaging risk is most visible in forwarding pathways, long-lived shared workloads, and retained legacy transport objects rather than raw mailbox count alone.' -RecommendationSupport 'This section supports the messaging and anti-spoofing recommendations in 4.0.')
        CollaborationConsultativeSummary = (New-CustomerConsultativeSummary -Title 'CollaborationConsultativeSummary' -SnapshotRows @(
            (New-CustomerConfigurationRow -Signal 'Ownerless Teams' -State $ownerlessTeams),
            (New-CustomerConfigurationRow -Signal 'Ownerless Microsoft 365 groups' -State $ownerlessGroups),
            (New-CustomerConfigurationRow -Signal 'Stale SharePoint sites (>180 days)' -State $staleSharePointSites),
            (New-CustomerConfigurationRow -Signal 'Stale OneDrive locations (>180 days)' -State $staleOneDrives),
            (New-CustomerConfigurationRow -Signal 'Guest-heavy Teams' -State $guestHeavyTeams),
            (New-CustomerConfigurationRow -Signal 'Tenant external sharing posture' -State $tenantSharingCapability),
            (New-CustomerConfigurationRow -Signal 'Default sharing link type' -State $defaultSharingLinkType),
            (New-CustomerConfigurationRow -Signal 'Site-level sharing overrides' -State $siteOverrideCount),
            (New-CustomerConfigurationRow -Signal 'External exposure review rows' -State $externalExposureFindings.Count)
        ) -Narrative 'Collaboration risk is clustering where broad sharing posture, stale content locations, and weak ownership signals overlap.' -RecommendationSupport 'This section supports the collaboration ownership, lifecycle, and external-sharing recommendations in 4.0.')
        GovernanceConsultativeSummary = (New-CustomerConsultativeSummary -Title 'GovernanceConsultativeSummary' -SnapshotRows @(
            (New-CustomerConfigurationRow -Signal 'Microsoft Secure Score' -State $(if ($null -ne $secureScorePercent) { "$secureScorePercent% ($secureScoreCurrent/$secureScoreMax)" } else { $null })),
            (New-CustomerConfigurationRow -Signal 'Unverified domains' -State $unverifiedDomains),
            (New-CustomerConfigurationRow -Signal 'Domains without DMARC' -State $dmarcMissingCount),
            (New-CustomerConfigurationRow -Signal 'License capacity pressure' -State $licensePressure),
            (New-CustomerConfigurationRow -Signal 'SMTP-authenticated accounts' -State $smtpAuthUsers),
            (New-CustomerConfigurationRow -Signal 'Directory synchronization' -State $(if ($null -eq $dirSyncEnabled) { $null } elseif ($dirSyncEnabled) { 'Enabled' } else { 'Disabled' }))
        ) -Narrative 'Governance pressure is most visible where domain trust controls, licensing headroom, and baseline security posture are drifting together.' -RecommendationSupport 'This section supports the governance, domain, and security-baseline recommendations in 4.0.')
        LifecycleConsultativeSummary = (New-CustomerConsultativeSummary -Title 'LifecycleConsultativeSummary' -SnapshotRows @(
            (New-CustomerConfigurationRow -Signal 'Inactive guest accounts (>90 days)' -State $inactiveGuests90Days),
            (New-CustomerConfigurationRow -Signal 'Stale privileged accounts (>90 days)' -State $stalePrivileged90Days),
            (New-CustomerConfigurationRow -Signal 'Inactive or disabled licensed users' -State $inactiveLicensedUsers),
            (New-CustomerConfigurationRow -Signal 'Stale collaboration assets' -State ("{0} SharePoint site(s); {1} OneDrive location(s); {2} Team(s)" -f $staleSharePointSites, $staleOneDrives, $dormantTeams)),
            (New-CustomerConfigurationRow -Signal 'OneDrive ownership mismatches' -State $oneDriveOwnerMismatches.Count),
            (New-CustomerConfigurationRow -Signal 'Shared mailboxes without owner signal' -State $sharedMailboxesWithoutOwnerSignal)
        ) -Narrative 'Lifecycle drift is showing up across identities, collaboration assets, and shared workloads, which makes cleanup slower and increases the chance that stale access or stale data remains in place.' -RecommendationSupport 'This section supports the lifecycle and ownership-governance recommendations in 4.0.')
    }
}

function Get-CondensedWorkstreamSignalText {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Text,
        [Parameter(Mandatory = $false)][int]$MaxItems = 2
    )

    $normalized = Convert-ToArrayaDisplayText -Value $Text -Default ''
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    $items = @(
        $normalized -split '\s*;\s*' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First $MaxItems
    )

    if ($items.Count -eq 0) {
        return $normalized
    }

    return ($items -join '; ')
}

function Resolve-ImprovementPlanInputSources {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        throw "Assessment input path was not found: $Path"
    }

    $resolvedPath = (Resolve-Path -Path $Path).Path
    if (Test-Path -Path $resolvedPath -PathType Container) {
        $manifestFiles = @(
            Get-ChildItem -Path $resolvedPath -Recurse -Filter '*.manifest.json' -File -ErrorAction SilentlyContinue |
                Sort-Object FullName
        )

        if ($manifestFiles.Count -eq 0) {
            throw "No manifest.json files were found under: $resolvedPath"
        }

        return @(
            $manifestFiles | ForEach-Object {
                [pscustomobject]@{
                    InputPath   = $_.FullName
                    SourceType  = 'Manifest'
                    SourceLabel = [System.IO.Path]::GetFileNameWithoutExtension([System.IO.Path]::GetFileNameWithoutExtension($_.Name))
                }
            }
        )
    }

    $sourceType = if ($resolvedPath -match '\.manifest\.json$') { 'Manifest' } else { 'Snapshot' }
    $sourceLabel = if ($sourceType -eq 'Manifest') {
        [System.IO.Path]::GetFileNameWithoutExtension([System.IO.Path]::GetFileNameWithoutExtension([System.IO.Path]::GetFileName($resolvedPath)))
    }
    else {
        [System.IO.Path]::GetFileNameWithoutExtension([System.IO.Path]::GetFileName($resolvedPath))
    }

    return @(
        [pscustomobject]@{
            InputPath   = $resolvedPath
            SourceType  = $sourceType
            SourceLabel = $sourceLabel
        }
    )
}

function Convert-ToArrayaSafeFileComponent {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Value)

    $safe = $Value -replace '[^a-zA-Z0-9\-_\. ]+', '-' -replace '\s+', ' '
    $safe = $safe.Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return 'Source'
    }

    return $safe
}

function Get-TenantArtifactFilePrefix {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$TenantName,
        [AllowNull()][string]$Fallback = 'Tenant'
    )

    $candidate = if ([string]::IsNullOrWhiteSpace($TenantName)) { $Fallback } else { $TenantName }
    $safe = Convert-ToArrayaSafeFileComponent -Value $candidate
    $safe = ($safe -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return $Fallback
    }

    return $safe
}

function Get-CustomerFindingsLegendRows {
    [CmdletBinding()]
    param()

    return @(
        [pscustomobject]@{
            Term    = 'Open Findings'
            Meaning = 'Assessment items that still need remediation, validation, or an explicit decision to accept or retain the current state.'
        },
        [pscustomobject]@{
            Term    = 'Severity / Impact'
            Meaning = 'The highest current impact level associated with the findings grouped into that row.'
        },
        [pscustomobject]@{
            Term    = 'Critical'
            Meaning = 'A concentrated control gap that should be addressed first because the current exposure or operational impact is already significant.'
        },
        [pscustomobject]@{
            Term    = 'High'
            Meaning = 'A higher-priority gap or control weakness that should be reviewed early because the exposure or operational impact is concentrated.'
        },
        [pscustomobject]@{
            Term    = 'Medium'
            Meaning = 'A meaningful inconsistency or governance gap that belongs in the planned remediation sequence.'
        },
        [pscustomobject]@{
            Term    = 'Low'
            Meaning = 'A narrower-scope cleanup item or limited-surface issue that still deserves follow-through.'
        },
        [pscustomobject]@{
            Term    = 'Info'
            Meaning = 'Context or supporting evidence that helps explain current state, but is not urgent by itself.'
        },
        [pscustomobject]@{
            Term    = 'Full Itemized Reference'
            Meaning = 'Use the Engineer Pack for the detailed finding-by-finding evidence and validation path behind the grouped recommendations.'
        },
        [pscustomobject]@{
            Term    = 'How grouped rows map forward'
            Meaning = 'Each workstream row rolls related findings into one customer-readable summary. Use 4.0 Modern Workplace Recommendations for execution planning, the companion roadmap for leadership sequencing, and the Engineer Pack for the detailed crosswalk.'
        }
    )
}

function Get-CustomerSourceSummaryRows {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$SourceModel)

    return @(
        [pscustomobject]@{
            Signal = 'Scope reviewed'
            State  = 'Identity, messaging, collaboration, endpoint, security, and governance signals'
        },
        [pscustomobject]@{
            Signal = 'How to read this report'
            State  = 'Use 3.0 Executive Summary for the risk brief, 4.0 Modern Workplace Recommendations for the action sequence, and the later sections for the supporting evidence.'
        },
        [pscustomobject]@{
            Signal = 'What the tables show'
            State  = 'Each section starts with current-state evidence so the later recommendations can be traced back to the observed tenant condition.'
        },
        [pscustomobject]@{
            Signal = 'Where the action matrix appears'
            State  = '4.0 Modern Workplace Recommendations'
        },
        [pscustomobject]@{
            Signal = 'Where the leadership roadmap appears'
            State  = 'Companion remediation roadmap in the Support folder'
        },
        [pscustomobject]@{
            Signal = 'Where the detailed technical evidence appears'
            State  = 'Engineer Pack in the Support folder'
        }
    )
}

function Get-CustomerSourceSummaryText {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$SourceModel)

    return 'This report starts with an orientation snapshot of what was reviewed and where the detailed evidence appears later in the document.'
}

function Get-ArrayaCollectorScriptVersionInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$CollectorScriptPath = (Join-Path -Path $PSScriptRoot -ChildPath 'Get-FullTenantReportDetails.ps1'),
        [Parameter(Mandatory = $false)]
        [int]$MajorVersion = 3,
        [Parameter(Mandatory = $false)]
        [int]$FallbackMinorVersion = 42
    )

    $minorVersion = $FallbackMinorVersion
    $source = 'Fallback'

    try {
        $resolvedScriptPath = (Resolve-Path -Path $CollectorScriptPath -ErrorAction Stop).Path
        $gitCommand = Get-Command -Name 'git' -ErrorAction SilentlyContinue
        if ($gitCommand) {
            $scriptDirectory = Split-Path -Path $resolvedScriptPath -Parent
            $gitRoot = (& $gitCommand.Source -C $scriptDirectory rev-parse --show-toplevel 2>$null | Select-Object -First 1)
            if (-not [string]::IsNullOrWhiteSpace($gitRoot)) {
                $relativePath = $resolvedScriptPath.Substring($gitRoot.Length).TrimStart([char[]]@('\', '/'))
                $revisionCountRaw = (& $gitCommand.Source -C $gitRoot rev-list --count HEAD -- $relativePath 2>$null | Select-Object -First 1)
                if ($revisionCountRaw -match '^\d+$') {
                    $minorVersion = [int]$revisionCountRaw
                    $source = 'Git history'

                    & $gitCommand.Source -C $gitRoot diff --quiet -- $relativePath 2>$null
                    if ($LASTEXITCODE -eq 1) {
                        $minorVersion++
                        $source = 'Git history + working tree'
                    }
                }
            }
        }
    }
    catch {
        $minorVersion = $FallbackMinorVersion
        $source = 'Fallback'
    }

    return [pscustomobject]@{
        MajorVersion = $MajorVersion
        MinorVersion = $minorVersion
        Label        = ('{0}.{1}' -f $MajorVersion, $minorVersion)
        Source       = $source
    }
}

function Resolve-ArrayaAssessmentVersionLabel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $SnapshotContext
    )

    if ($null -ne $SnapshotContext -and $SnapshotContext.PSObject.Properties.Name -contains 'Snapshot') {
        $metadata = Get-ArrayaObjectValue -Object $SnapshotContext.Snapshot -Names @('Metadata')
        $metadataVersion = [string](Get-ArrayaObjectValue -Object $metadata -Names @('AssessmentVersion'))
        if (-not [string]::IsNullOrWhiteSpace($metadataVersion)) {
            return $metadataVersion.Trim()
        }
    }

    return (Get-ArrayaCollectorScriptVersionInfo).Label
}

function New-CustomerReportSourceModel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$TenantName,
        [Parameter(Mandatory = $true)][string]$AssessmentJsonPath,
        [Parameter(Mandatory = $true)][string]$SourceInputPath,
        [Parameter(Mandatory = $true)][string]$SourceType,
        [Parameter(Mandatory = $true)][string]$SourceLabel,
        [Parameter(Mandatory = $true)][object[]]$Findings,
        [Parameter(Mandatory = $false)][object[]]$WorkstreamSummaries = @(),
        [Parameter(Mandatory = $false)][object[]]$TechnicalObservations = @(),
        [Parameter(Mandatory = $false)]$ConsultativeSummaries = $null,
        [Parameter(Mandatory = $false)][string]$AssessmentVersion
    )

    $severityCounts = @($Findings | Group-Object Severity | Sort-Object Name)
    $ownerGroups = @($Findings | Group-Object OwnerTeam | Sort-Object Count -Descending)
    $executiveThemes = Get-CustomerExecutiveThemes -Findings $Findings -Count 5
    $roadmapActions = Get-CustomerRoadmapActions -Findings $Findings -MaxPerPhase 5
    $technicalObservations = if ($PSBoundParameters.ContainsKey('TechnicalObservations') -and @($TechnicalObservations).Count -gt 0) { @($TechnicalObservations) } else { @() }
    $executiveNarrative = Get-CustomerExecutiveSummaryNarrative -Findings $Findings -ExecutiveThemes $executiveThemes -OwnerGroups $ownerGroups
    $consultativeSummaries = if ($null -ne $ConsultativeSummaries) { $ConsultativeSummaries } else { [pscustomobject]@{} }
    $guestMfaExperienceSummary = $consultativeSummaries.GuestMfaExperienceSummary
    $roadmapActions = Add-CustomerRoadmapActionExperienceNotes -RoadmapActions $roadmapActions -GuestMfaExperienceSummary $guestMfaExperienceSummary
    $resolvedAssessmentVersion = if ([string]::IsNullOrWhiteSpace($AssessmentVersion)) { (Resolve-ArrayaAssessmentVersionLabel) } else { $AssessmentVersion.Trim() }

    $sourceModel = [pscustomobject]@{
        TenantName          = $TenantName
        AssessmentJsonPath  = $AssessmentJsonPath
        SourceInputPath     = $SourceInputPath
        SourceType          = $SourceType
        SourceLabel         = $SourceLabel
        AssessmentVersion   = $resolvedAssessmentVersion
        Findings            = @($Findings)
        WorkstreamSummaries = @($WorkstreamSummaries)
        SeverityCounts      = $severityCounts
        OwnerGroups         = $ownerGroups
        ExecutiveThemes     = @($executiveThemes)
        RoadmapActions      = @($roadmapActions)
        ExecutiveNarrative  = $executiveNarrative
        TechnicalObservations = $technicalObservations
    }

    $sourceModel | Add-Member -NotePropertyName SummaryText -NotePropertyValue (Get-CustomerSourceSummaryText -SourceModel $sourceModel)
    $sourceModel | Add-Member -NotePropertyName SummaryRows -NotePropertyValue (Get-CustomerSourceSummaryRows -SourceModel $sourceModel)
    $sourceModel | Add-Member -NotePropertyName FindingsLegendRows -NotePropertyValue (Get-CustomerFindingsLegendRows)
    $sourceModel | Add-Member -NotePropertyName ExecutiveDecisionSummary -NotePropertyValue $(if ($consultativeSummaries.PSObject.Properties.Name -contains 'ExecutiveDecisionSummary') { $consultativeSummaries.ExecutiveDecisionSummary } else { Get-CustomerExecutiveDecisionSummary -ExecutiveThemes $executiveThemes -RoadmapActions $roadmapActions -OwnerGroups $ownerGroups })
    $sourceModel | Add-Member -NotePropertyName GuestMfaExperienceSummary -NotePropertyValue $guestMfaExperienceSummary
    $sourceModel | Add-Member -NotePropertyName IdentityConsultativeSummary -NotePropertyValue $consultativeSummaries.IdentityConsultativeSummary
    $sourceModel | Add-Member -NotePropertyName MessagingConsultativeSummary -NotePropertyValue $consultativeSummaries.MessagingConsultativeSummary
    $sourceModel | Add-Member -NotePropertyName CollaborationConsultativeSummary -NotePropertyValue $consultativeSummaries.CollaborationConsultativeSummary
    $sourceModel | Add-Member -NotePropertyName GovernanceConsultativeSummary -NotePropertyValue $consultativeSummaries.GovernanceConsultativeSummary
    $sourceModel | Add-Member -NotePropertyName LifecycleConsultativeSummary -NotePropertyValue $consultativeSummaries.LifecycleConsultativeSummary
    return $sourceModel
}

function New-CustomerReportModel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$TenantName,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt,
        [Parameter(Mandatory = $true)][object[]]$Sources
    )

    $isMultiSource = (@($Sources).Count -gt 1)
    $overview = if ($isMultiSource) {
        "This Microsoft 365 assessment report reflects $(@($Sources).Count) separately prepared assessment sources. Each source is presented independently so the observed conditions remain attributable to the environment reviewed, rather than being blended into a generalized summary."
    }
    else {
        "This Microsoft 365 assessment report describes the current-state conditions identified in this tenant. The report is written to show where risk is concentrated, where control coverage is uneven, and where day-to-day ownership appears weaker than the exposure carried by the environment."
    }

    $methodology = if ($isMultiSource) {
        'Each assessment source was considered independently and is presented in a separate source section. Observations, recommendations, and technical walkthroughs remain grouped to that source so the report preserves where each condition was observed.'
    }
    else {
        'The report is based on the current tenant signals available in the assessment source and organizes those signals into executive observations, recommendations, and technical walkthroughs. The intent is to show what stood out in this tenant and why those patterns matter operationally.'
    }

    $appendixSections = Get-CustomerDocumentationAppendixSections

    $numberedSources = @()
    $sourceIndex = 0
    foreach ($source in @($Sources)) {
        $sourceIndex++
        $observationIndex = 0
        foreach ($technicalObservation in @($source.TechnicalObservations)) {
            $observationIndex++
            if ($technicalObservation.PSObject.Properties.Name -contains 'SectionNumber') {
                $technicalObservation.SectionNumber = "5.$sourceIndex.$observationIndex"
            }
            else {
                $technicalObservation | Add-Member -NotePropertyName SectionNumber -NotePropertyValue "5.$sourceIndex.$observationIndex"
            }
        }
        $numberedSources += $source
    }

    $appendixIndex = 0
    foreach ($section in @($appendixSections)) {
        $appendixIndex++
        if ($section.PSObject.Properties.Name -contains 'SectionNumber') {
            $section.SectionNumber = "6.$appendixIndex"
        }
        else {
            $section | Add-Member -NotePropertyName SectionNumber -NotePropertyValue "6.$appendixIndex"
        }
    }

    return [pscustomobject]@{
        TenantName      = $TenantName
        GeneratedAt     = $GeneratedAt
        Sources         = @($numberedSources)
        IsMultiSource   = $isMultiSource
        OverviewText    = $overview
        MethodologyText = $methodology
        AppendixSections = @($appendixSections)
    }
}

function New-CustomerRecommendationBlockMarkdown {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $false)][string]$ObservedText,
        [Parameter(Mandatory = $false)][string]$StandoutText,
        [Parameter(Mandatory = $false)][string]$WhyItMattersText,
        [Parameter(Mandatory = $false)][string]$ExampleText,
        [Parameter(Mandatory = $false)][string]$Action,
        [Parameter(Mandatory = $false)][string]$ValueText
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("**$Title**") | Out-Null
    if (-not [string]::IsNullOrWhiteSpace($ObservedText)) {
        $lines.Add("**What was observed:** $ObservedText") | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($StandoutText)) {
        $lines.Add("**Why this stood out:** $StandoutText") | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($WhyItMattersText)) {
        $lines.Add("**Why it matters:** $WhyItMattersText") | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($ExampleText)) {
        $lines.Add("**Example from this tenant:** $ExampleText") | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($Action)) {
        $lines.Add("**Recommended action:** $Action") | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($ValueText)) {
        $lines.Add("**Expected outcome:** $ValueText") | Out-Null
    }
    $lines.Add('') | Out-Null
    return ($lines -join [Environment]::NewLine)
}

function New-CustomerRecommendationBlockHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $false)][string]$ObservedText,
        [Parameter(Mandatory = $false)][string]$StandoutText,
        [Parameter(Mandatory = $false)][string]$WhyItMattersText,
        [Parameter(Mandatory = $false)][string]$ExampleText,
        [Parameter(Mandatory = $false)][string]$Action,
        [Parameter(Mandatory = $false)][string]$ValueText
    )

    $observedHtml = if ([string]::IsNullOrWhiteSpace($ObservedText)) { '' } else { "<p><strong>What was observed:</strong> $(Convert-ToArrayaHtmlFragment $ObservedText)</p>" }
    $standoutHtml = if ([string]::IsNullOrWhiteSpace($StandoutText)) { '' } else { "<p><strong>Why this stood out:</strong> $(Convert-ToArrayaHtmlFragment $StandoutText)</p>" }
    $whyHtml = if ([string]::IsNullOrWhiteSpace($WhyItMattersText)) { '' } else { "<p><strong>Why it matters:</strong> $(Convert-ToArrayaHtmlFragment $WhyItMattersText)</p>" }
    $exampleHtml = if ([string]::IsNullOrWhiteSpace($ExampleText)) { '' } else { "<p><strong>Example from this tenant:</strong> $(Convert-ToArrayaHtmlFragment $ExampleText)</p>" }
    $actionHtml = if ([string]::IsNullOrWhiteSpace($Action)) { '' } else { "<p><strong>Recommended action:</strong> $(Convert-ToArrayaHtmlFragment $Action)</p>" }
    $valueHtml = if ([string]::IsNullOrWhiteSpace($ValueText)) { '' } else { "<p><strong>Expected outcome:</strong> $(Convert-ToArrayaHtmlFragment $ValueText)</p>" }
    return @"
<div class="recommendation-block">
  <h4>$(Convert-ToArrayaHtmlEncodedText $Title)</h4>
  $observedHtml
  $standoutHtml
  $whyHtml
  $exampleHtml
  $actionHtml
  $valueHtml
</div>
"@
}

function New-CustomerRemediationReportMarkdownFromModel {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Model)

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("# $($Model.TenantName) Microsoft 365 Assessment Report") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("Generated: $($Model.GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss'))") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('## **1.0 Executive Summary**') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add($Model.OverviewText) | Out-Null
    $lines.Add('') | Out-Null
    if ($Model.IsMultiSource) {
        $lines.Add("This report contains **$(@($Model.Sources).Count)** separately presented assessment sources, and each one is addressed independently in the sections that follow.") | Out-Null
    }
    else {
        $lines.Add($Model.Sources[0].ExecutiveNarrative) | Out-Null
    }
    $lines.Add('') | Out-Null
    $lines.Add('## **2.0 Assessment Scope And Methodology**') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add($Model.MethodologyText) | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('The review considered the tenant signals available across identity, collaboration, messaging, endpoint, and governance. The intent of this report is to show what the environment is currently doing, where the strongest concentrations appear, and why those patterns matter in day-to-day operations.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('## **3.0 Key Observations**') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('The following observations highlight the patterns that stood out most clearly in this tenant. Each observation ties the current concentration back to the control area where it was seen so the narrative reads as an environment review, not a generalized issue list.') | Out-Null
    $lines.Add('') | Out-Null

    $sourceIndex = 0
    foreach ($source in @($Model.Sources)) {
        $sourceIndex++
        if ($Model.IsMultiSource) {
            $lines.Add("### **3.$sourceIndex Manifest Source: $($source.SourceInputPath)**") | Out-Null
            $lines.Add('') | Out-Null
            $lines.Add($source.SummaryText) | Out-Null
            $lines.Add('') | Out-Null
        }

        foreach ($theme in @($source.ExecutiveThemes)) {
            $lines.Add((New-CustomerRecommendationBlockMarkdown -Title $theme.Theme -ObservedText $theme.WhatThisMeans -StandoutText $theme.StandoutReason -WhyItMattersText $theme.WhyItMatters -ExampleText $theme.ExampleText -Action $theme.RecommendedNextStep -ValueText $theme.CustomerValue)) | Out-Null
        }
    }

    $lines.Add('## **4.0 Recommendations**') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('The recommendations below keep the existing remediation actions intact and place them against the conditions observed in the tenant. The explanatory text focuses on why each work item is rising to the surface now, given the current concentration of findings.') | Out-Null
    $lines.Add('') | Out-Null
    $sourceIndex = 0
    foreach ($source in @($Model.Sources)) {
        $sourceIndex++
        if ($Model.IsMultiSource) {
            $lines.Add("### **4.$sourceIndex Manifest Source: $($source.SourceInputPath)**") | Out-Null
            $lines.Add('') | Out-Null
        }

        foreach ($action in @($source.RoadmapActions)) {
            $title = '{0} ({1})' -f $action.ActionTitle, $action.RoadmapPhase
            $lines.Add((New-CustomerRecommendationBlockMarkdown -Title $title -ObservedText $action.WhatThisAddresses -StandoutText $action.StandoutReason -WhyItMattersText $action.WhyItMatters -ExampleText $action.ExampleText -Action $action.RecommendedNextStep -ValueText $action.BusinessValue)) | Out-Null
        }
    }

    $lines.Add('## **5.0 Workstream Detail**') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('The following workstream detail presents the current state of the tenant by workload. Each source begins with the workstream concentration table, followed by numbered technical sections that show the current configuration, what stood out, what is already working, and why those conditions matter.') | Out-Null
    $lines.Add('') | Out-Null
    $sourceIndex = 0
    foreach ($source in @($Model.Sources)) {
        $sourceIndex++
        $label = if ($source.SourceType -eq 'Manifest') { 'Manifest Source' } else { 'Assessment Snapshot Source' }
        $lines.Add("### **5.$sourceIndex ${label}: $($source.SourceInputPath)**") | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add($source.SummaryText) | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add('| Workstream | Area | Open Findings | Severity | Observation Summary |') | Out-Null
        $lines.Add('|---|---|---|---|---|') | Out-Null
        foreach ($summary in @($source.WorkstreamSummaries)) {
            $signalText = Get-CondensedWorkstreamSignalText -Text $summary.TopSignals -MaxItems 3
            $lines.Add("| $(Convert-ToArrayaMarkdownText $summary.Workstream) | $(Convert-ToArrayaMarkdownText $summary.Area) | $($summary.OpenFindings) | $(Convert-ToArrayaMarkdownText $summary.Severity) | $(Convert-ToArrayaMarkdownText $signalText) |") | Out-Null
        }
        $lines.Add('') | Out-Null
        $lines.Add('The following numbered sections provide the deeper current-state review for this source.') | Out-Null
        $lines.Add('') | Out-Null
        foreach ($technicalObservation in @($source.TechnicalObservations)) {
            $lines.Add("### **$($technicalObservation.SectionNumber) $($technicalObservation.SectionTitle)**") | Out-Null
            $lines.Add('') | Out-Null
            $lines.Add('| Configuration Signal | Current State |') | Out-Null
            $lines.Add('|---|---|') | Out-Null
            foreach ($configurationRow in @($technicalObservation.ConfigurationRows)) {
                $lines.Add("| $(Convert-ToArrayaMarkdownText $configurationRow.Signal) | $(Convert-ToArrayaMarkdownText $configurationRow.State) |") | Out-Null
            }
            $lines.Add('') | Out-Null
            $lines.Add('**Current-state observation**') | Out-Null
            $lines.Add($technicalObservation.ObservedNarrative) | Out-Null
            $lines.Add('') | Out-Null
            $lines.Add('**What is working well**') | Out-Null
            $lines.Add($technicalObservation.PositiveNarrative) | Out-Null
            $lines.Add('') | Out-Null
            $lines.Add('**Why it matters in this tenant**') | Out-Null
            $lines.Add($technicalObservation.WhyItMatters) | Out-Null
            $lines.Add('') | Out-Null
        }
    }

    $lines.Add('## **6.0 Appendix: Microsoft Documentation**') | Out-Null
    $lines.Add('') | Out-Null
    foreach ($appendixSection in @($Model.AppendixSections)) {
        $lines.Add("### **$($appendixSection.SectionNumber) $($appendixSection.Title)**") | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add($appendixSection.Intro) | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add('| Microsoft Guidance | Why It Is Relevant |') | Out-Null
        $lines.Add('|---|---|') | Out-Null
        foreach ($reference in @($appendixSection.References)) {
            $lines.Add("| [$($reference.Title)]($($reference.Url)) | $(Convert-ToArrayaMarkdownText $reference.WhyItIsRelevant) |") | Out-Null
        }
        $lines.Add('') | Out-Null
    }
    return ($lines -join [Environment]::NewLine)
}

function New-CustomerRemediationReportHtmlFromModel {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Model)

    $observationSections = New-Object System.Collections.Generic.List[string]
    $recommendationSections = New-Object System.Collections.Generic.List[string]
    $workstreamSections = New-Object System.Collections.Generic.List[string]
    $manifestList = New-Object System.Collections.Generic.List[string]

    $sourceIndex = 0
    foreach ($source in @($Model.Sources)) {
        $sourceIndex++
        $sourceKindLabel = if ($source.SourceType -eq 'Manifest') { 'Manifest Source' } else { 'Assessment Snapshot Source' }
        $sourceHeadingLabel = Convert-ToArrayaHtmlEncodedText ("3.$sourceIndex $sourceKindLabel")
        $sourceKindLabelHtml = Convert-ToArrayaHtmlEncodedText $sourceKindLabel
        $sourcePathHtml = Convert-ToArrayaHtmlEncodedText $source.SourceInputPath
        $sourceSummaryHtml = Convert-ToArrayaHtmlFragment $source.SummaryText
        $sourceHeading = if ($Model.IsMultiSource) {
            "<h3>$sourceHeadingLabel</h3><p class=""source-note""><strong>${sourceKindLabelHtml}:</strong> $sourcePathHtml</p><p>$sourceSummaryHtml</p>"
        }
        else {
            ''
        }

        $blocks = @($source.ExecutiveThemes | ForEach-Object {
            New-CustomerRecommendationBlockHtml -Title $_.Theme -ObservedText $_.WhatThisMeans -StandoutText $_.StandoutReason -WhyItMattersText $_.WhyItMatters -ExampleText $_.ExampleText -Action $_.RecommendedNextStep -ValueText $_.CustomerValue
        }) -join [Environment]::NewLine
        $observationSections.Add("<section class=""source-section"">$sourceHeading$blocks</section>") | Out-Null

        $recommendationHeading = if ($Model.IsMultiSource) {
            ('<h3>{0}</h3><p class="source-note"><strong>{1}:</strong> {2}</p>' -f (Convert-ToArrayaHtmlEncodedText ("4.$sourceIndex $sourceKindLabel")), $sourceKindLabelHtml, $sourcePathHtml)
        }
        else {
            ''
        }
        $recommendationBlocks = @($source.RoadmapActions | ForEach-Object {
            New-CustomerRecommendationBlockHtml -Title ('{0} ({1})' -f $_.ActionTitle, $_.RoadmapPhase) -ObservedText $_.WhatThisAddresses -StandoutText $_.StandoutReason -WhyItMattersText $_.WhyItMatters -ExampleText $_.ExampleText -Action $_.RecommendedNextStep -ValueText $_.BusinessValue
        }) -join [Environment]::NewLine
        $recommendationSections.Add("<section class=""source-section"">$recommendationHeading$recommendationBlocks</section>") | Out-Null

        $workstreamRows = @($source.WorkstreamSummaries | ForEach-Object {
@"
<tr>
  <td>$(Convert-ToArrayaHtmlEncodedText $_.Workstream)</td>
  <td>$(Convert-ToArrayaHtmlEncodedText $_.Area)</td>
  <td>$($_.OpenFindings)</td>
  <td>$(Convert-ToArrayaHtmlEncodedText $_.Severity)</td>
  <td>$(Convert-ToArrayaHtmlFragment (Get-CondensedWorkstreamSignalText -Text $_.TopSignals -MaxItems 3))</td>
</tr>
"@
        }) -join [Environment]::NewLine
        $label = $sourceKindLabel
        $technicalObservationHtml = @($source.TechnicalObservations | ForEach-Object {
            $configurationRowsHtml = @($_.ConfigurationRows | ForEach-Object {
@"
      <tr>
        <td>$(Convert-ToArrayaHtmlEncodedText $_.Signal)</td>
        <td>$(Convert-ToArrayaHtmlFragment $_.State)</td>
      </tr>
"@
            }) -join [Environment]::NewLine
@"
  <section class="technical-observation">
    <h3>$(Convert-ToArrayaHtmlEncodedText ("$($_.SectionNumber) $($_.SectionTitle)"))</h3>
    <table class="configuration-table">
      <thead>
        <tr>
          <th>Configuration Signal</th>
          <th>Current State</th>
        </tr>
      </thead>
      <tbody>
$configurationRowsHtml
      </tbody>
    </table>
    <p><strong>Current-state observation</strong><br/>$(Convert-ToArrayaHtmlFragment $_.ObservedNarrative)</p>
    <p><strong>What is working well</strong><br/>$(Convert-ToArrayaHtmlFragment $_.PositiveNarrative)</p>
    <p><strong>Why it matters in this tenant</strong><br/>$(Convert-ToArrayaHtmlFragment $_.WhyItMatters)</p>
  </section>
"@
        }) -join [Environment]::NewLine
        $workstreamSections.Add(@"
<section class="source-section">
  <h3>$(Convert-ToArrayaHtmlEncodedText ("5.$sourceIndex $label"))</h3>
  <p class="source-note"><strong>${sourceKindLabelHtml}:</strong> $sourcePathHtml</p>
  <p>$sourceSummaryHtml</p>
  <table>
    <thead>
      <tr>
        <th>Workstream</th>
        <th>Area</th>
        <th>Open Findings</th>
        <th>Severity</th>
        <th>Observation Summary</th>
      </tr>
    </thead>
    <tbody>
$workstreamRows
    </tbody>
  </table>
  <p>The following numbered sections provide the deeper current-state review for this source.</p>
$technicalObservationHtml
</section>
"@) | Out-Null

        $manifestList.Add("<li><strong>${sourceKindLabelHtml}:</strong> $sourcePathHtml</li>") | Out-Null
    }

    $appendixSectionsHtml = @($Model.AppendixSections | ForEach-Object {
        $referenceRows = @($_.References | ForEach-Object {
@"
      <tr>
        <td><a href="$(Convert-ToArrayaHtmlEncodedText $_.Url)">$(Convert-ToArrayaHtmlEncodedText $_.Title)</a></td>
        <td>$(Convert-ToArrayaHtmlFragment $_.WhyItIsRelevant)</td>
      </tr>
"@
        }) -join [Environment]::NewLine
@"
      <section class="source-section">
        <h3>$(Convert-ToArrayaHtmlEncodedText ("$($_.SectionNumber) $($_.Title)"))</h3>
        <p>$(Convert-ToArrayaHtmlFragment $_.Intro)</p>
        <table>
          <thead>
            <tr>
              <th>Microsoft Guidance</th>
              <th>Why It Is Relevant</th>
            </tr>
          </thead>
          <tbody>
$referenceRows
          </tbody>
        </table>
      </section>
"@
    }) -join [Environment]::NewLine

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>$([System.Net.WebUtility]::HtmlEncode($Model.TenantName)) Microsoft 365 Assessment Report</title>
  <style>
    body { font-family: "Segoe UI", Arial, sans-serif; margin: 0; background: #f7f8fa; color: #1f2933; }
    main { max-width: 1100px; margin: 0 auto; padding: 32px 24px 48px; }
    .hero, .panel { background: #ffffff; border: 1px solid #d8dee4; border-radius: 16px; padding: 24px; margin-bottom: 20px; }
    .hero { border-top: 6px solid #12343b; }
    h1, h2, h3, h4 { color: #12343b; margin-top: 0; }
    h2 { font-size: 1.45rem; }
    h3 { font-size: 1.1rem; margin-top: 24px; }
    h4 { margin-bottom: 8px; }
    p, li, td, th { line-height: 1.6; }
    .recommendation-block { border: 1px solid #d8dee4; border-left: 4px solid #1e5160; border-radius: 10px; padding: 16px; margin: 14px 0; background: #fcfcfd; }
    .technical-observation { border-top: 1px solid #d8dee4; margin-top: 18px; padding-top: 18px; }
    .configuration-table { margin-top: 8px; }
    table { width: 100%; border-collapse: collapse; margin-top: 14px; }
    th, td { border: 1px solid #d8dee4; padding: 10px 12px; text-align: left; vertical-align: top; }
    th { background: #eef3f5; }
    .source-note { color: #52606d; font-size: 0.95rem; }
  </style>
</head>
<body>
  <main>
    <section class="hero">
      <h1>$([System.Net.WebUtility]::HtmlEncode($Model.TenantName)) Microsoft 365 Assessment Report</h1>
      <p><strong>Generated:</strong> $($Model.GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss'))</p>
      <p>$(Convert-ToArrayaHtmlFragment $Model.OverviewText)</p>
    </section>

    <section class="panel">
      <h2>1.0 Executive Summary</h2>
      <p>$(Convert-ToArrayaHtmlFragment $Model.OverviewText)</p>
      <p>$(Convert-ToArrayaHtmlFragment $(if ($Model.IsMultiSource) { "This report contains $(@($Model.Sources).Count) separately presented assessment sources, each shown independently so the observed conditions remain attributable to the source reviewed." } else { $Model.Sources[0].ExecutiveNarrative }))</p>
    </section>

    <section class="panel">
      <h2>2.0 Assessment Scope And Methodology</h2>
      <p>$(Convert-ToArrayaHtmlFragment $Model.MethodologyText)</p>
      <p>The review considered the tenant signals available across identity, collaboration, messaging, endpoint, and governance. The intent of this report is to show what the environment is currently doing, where the strongest concentrations appear, and why those patterns matter in day-to-day operations.</p>
    </section>

    <section class="panel">
      <h2>3.0 Key Observations</h2>
      <p>The following observations highlight the patterns that stood out most clearly in this tenant. Each observation ties the current concentration back to the control area where it was seen so the narrative reads as an environment review, not a generalized issue list.</p>
      $($observationSections -join [Environment]::NewLine)
    </section>

    <section class="panel">
      <h2>4.0 Recommendations</h2>
      <p>The recommendations below keep the existing remediation actions intact and place them against the conditions observed in the tenant. The explanatory text focuses on why each work item is rising to the surface now, given the current concentration of findings.</p>
      $($recommendationSections -join [Environment]::NewLine)
    </section>

    <section class="panel">
      <h2>5.0 Workstream Detail</h2>
      <p>The following workstream detail presents the current state of the tenant by workload. Each source begins with the workstream concentration table, followed by numbered technical sections that show the current configuration, what stood out, what is already working, and why those conditions matter.</p>
      $($workstreamSections -join [Environment]::NewLine)
    </section>

    <section class="panel">
      <h2>6.0 Appendix: Microsoft Documentation</h2>
      $appendixSectionsHtml
    </section>
  </main>
</body>
</html>
"@
}

function New-CustomerRemediationReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Findings,
        [Parameter(Mandatory = $false)][object[]]$WorkstreamSummaries = @(),
        [Parameter(Mandatory = $false)][object[]]$TechnicalObservations = @(),
        [Parameter(Mandatory = $true)][string]$TenantName,
        [Parameter(Mandatory = $true)][string]$AssessmentJsonPath,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt
    )

    $sourceModel = New-CustomerReportSourceModel -TenantName $TenantName -AssessmentJsonPath $AssessmentJsonPath -SourceInputPath $AssessmentJsonPath -SourceType 'Snapshot' -SourceLabel ([System.IO.Path]::GetFileNameWithoutExtension($AssessmentJsonPath)) -Findings $Findings -WorkstreamSummaries $WorkstreamSummaries -TechnicalObservations $TechnicalObservations
    $model = New-CustomerReportModel -TenantName $TenantName -GeneratedAt $GeneratedAt -Sources @($sourceModel)
    return (New-CustomerRemediationReportMarkdownFromModel -Model $model)
}

function New-CustomerRemediationReportHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$Findings,
        [Parameter(Mandatory = $false)][object[]]$WorkstreamSummaries = @(),
        [Parameter(Mandatory = $false)][object[]]$TechnicalObservations = @(),
        [Parameter(Mandatory = $true)][string]$TenantName,
        [Parameter(Mandatory = $true)][string]$AssessmentJsonPath,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt
    )

    $sourceModel = New-CustomerReportSourceModel -TenantName $TenantName -AssessmentJsonPath $AssessmentJsonPath -SourceInputPath $AssessmentJsonPath -SourceType 'Snapshot' -SourceLabel ([System.IO.Path]::GetFileNameWithoutExtension($AssessmentJsonPath)) -Findings $Findings -WorkstreamSummaries $WorkstreamSummaries -TechnicalObservations $TechnicalObservations
    $model = New-CustomerReportModel -TenantName $TenantName -GeneratedAt $GeneratedAt -Sources @($sourceModel)
    return (New-CustomerRemediationReportHtmlFromModel -Model $model)
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
    $lines.Add("Source snapshot: $AssessmentJsonPath") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Use this pack to review the current tenant state, confirm each finding, and plan the remediation work in a way that fits your change process.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('## Engineering Summary') | Out-Null
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
    $lines.Add('## Before You Start') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('- Validate tenant admin roles, Graph scopes, Exchange connectivity, and any pilot exclusions before enforcement changes.') | Out-Null
    $lines.Add('- Use the support artifacts for backlog import, validation context, and change-record attachment as needed.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('## Findings To Work') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('| Severity | Phase | Workstream | Reference | What Needs Attention | Why It Matters | Technical Remediation | Current Evidence | Where To Verify | Done When |') | Out-Null
    $lines.Add('|---|---|---|---|---|---|---|---|---|---|') | Out-Null
    foreach ($finding in $Findings) {
        $lines.Add("| $($finding.Severity) | $($finding.RoadmapPhase) | $($finding.OwnerTeam) | $($finding.RuleId) | $(Convert-ToArrayaMarkdownText $finding.Finding) | $(Convert-ToArrayaMarkdownText $finding.WhyFlagged) | $(Convert-ToArrayaMarkdownText $finding.TechnicalRemediation) | $(Convert-ToArrayaMarkdownText $finding.CurrentValue) | $(Convert-ToArrayaMarkdownText $finding.EvidenceLocation) | $(Convert-ToArrayaMarkdownText $finding.TargetValue) |") | Out-Null
    }
    $lines.Add('') | Out-Null
    $lines.Add('## Verification Checklist') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('- Confirm each recommendation against the latest tenant state before making changes.') | Out-Null
    $lines.Add('- Pilot security and access-policy changes with a limited scope before broad enforcement.') | Out-Null
    $lines.Add('- Re-run `M365Collect` and `Improve` after remediation milestones to measure delta and retire closed findings.') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('## Supporting Files') | Out-Null
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

function New-MultiSourceEngineerActionPack {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$SourcePayloads,
        [Parameter(Mandatory = $true)][string]$TenantName,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt,
        [Parameter(Mandatory = $true)][string]$SupportFolderPath,
        [Parameter(Mandatory = $true)][string]$JsonOutPath,
        [Parameter(Mandatory = $true)][string]$SnippetOutPath
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("# $TenantName Engineer Action Pack") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add("Generated: $($GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss'))") | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('Use this pack to review the findings associated with each assessment source independently. Findings remain grouped by source so technical follow-up can preserve traceability.') | Out-Null
    $lines.Add('') | Out-Null

    $sourceIndex = 0
    foreach ($payload in @($SourcePayloads)) {
        $sourceIndex++
        $sourcePath = Convert-ToArrayaDisplayText -Value $payload.SourceInputPath -Default (Convert-ToArrayaDisplayText -Value $payload.SourceFile -Default 'Unknown source')
        $lines.Add("## Source $sourceIndex") | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add("Source path: $sourcePath") | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add('| Severity | Phase | Workstream | Reference | What Needs Attention | Why It Matters | Technical Remediation | Current Evidence | Where To Verify | Done When |') | Out-Null
        $lines.Add('|---|---|---|---|---|---|---|---|---|---|') | Out-Null
        foreach ($finding in (Convert-ArrayaObjectToArray $payload.Findings)) {
            $lines.Add("| $($finding.Severity) | $($finding.RoadmapPhase) | $($finding.OwnerTeam) | $($finding.RuleId) | $(Convert-ToArrayaMarkdownText $finding.Finding) | $(Convert-ToArrayaMarkdownText $finding.WhyFlagged) | $(Convert-ToArrayaMarkdownText $finding.TechnicalRemediation) | $(Convert-ToArrayaMarkdownText $finding.CurrentValue) | $(Convert-ToArrayaMarkdownText $finding.EvidenceLocation) | $(Convert-ToArrayaMarkdownText $finding.TargetValue) |") | Out-Null
        }
        $lines.Add('') | Out-Null
    }

    $lines.Add('## Supporting Files') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add(('- Support folder: `{0}`' -f $SupportFolderPath)) | Out-Null
    $lines.Add(('- Improvement plan JSON: `{0}`' -f $JsonOutPath)) | Out-Null
    $lines.Add(('- Remediation snippets: `{0}`' -f $SnippetOutPath)) | Out-Null
    $lines.Add('') | Out-Null
    return ($lines -join [Environment]::NewLine)
}

$originalAssessmentInputPath = $AssessmentJsonPath
$inputSources = Resolve-ImprovementPlanInputSources -Path $AssessmentJsonPath
if ($inputSources.Count -gt 1) {
    $resolvedInputRoot = (Resolve-Path -Path $AssessmentJsonPath).Path
    if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
        $OutputFolder = $resolvedInputRoot
    }
    if (-not (Test-Path -Path $OutputFolder)) {
        $null = New-Item -ItemType Directory -Path $OutputFolder -Force
    }
    $OutputFolder = (Resolve-Path -Path $OutputFolder).Path
    if ([string]::IsNullOrWhiteSpace($OutputPrefix)) {
        $OutputPrefix = Split-Path -Path $resolvedInputRoot -Leaf
    }
    $OutputPrefix = Get-ImprovementOutputPrefix -OutputPrefix $OutputPrefix

    $generatedAt = Get-Date
    $supportFolder = Join-Path -Path $OutputFolder -ChildPath 'Support'
    if (-not (Test-Path -Path $supportFolder)) {
        $null = New-Item -ItemType Directory -Path $supportFolder -Force
    }
    $customerAssessmentReportMarkdownPaths = New-Object System.Collections.Generic.List[string]
    $roadmapRemediationPlanPaths = New-Object System.Collections.Generic.List[string]

    $sourcePayloads = New-Object System.Collections.Generic.List[object]
    $tenantNames = New-Object System.Collections.Generic.List[string]
    $customerAssessmentReportPaths = New-Object System.Collections.Generic.List[string]
    $snippetBlocks = New-Object System.Collections.Generic.List[string]
    $snippetBlocks.Add('<#') | Out-Null
    $snippetBlocks.Add('Remediation snippets grouped by manifest source. Review and test each command before use.') | Out-Null
    $snippetBlocks.Add('#>') | Out-Null
    $snippetBlocks.Add('') | Out-Null

    foreach ($source in @($inputSources)) {
        $safeSourceLabel = Convert-ToArrayaSafeFileComponent -Value $source.SourceLabel
        $perSourceOutputFolder = Join-Path -Path $supportFolder -ChildPath ("PerSource\" + $safeSourceLabel)
        if (-not (Test-Path -Path $perSourceOutputFolder)) {
            $null = New-Item -ItemType Directory -Path $perSourceOutputFolder -Force
        }

        $sourceResult = & $PSCommandPath `
            -AssessmentJsonPath $source.InputPath `
            -OutputFolder $perSourceOutputFolder `
            -OutputPrefix $safeSourceLabel `
            -IncludeLegacyArtifacts:$IncludeLegacyArtifacts `
            -PassThru `
            -Quiet

        if (-not $sourceResult) {
            continue
        }

        $sourcePayload = Get-Content -Path $sourceResult.JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 25
        $sourcePayload | Add-Member -NotePropertyName SourceInputPath -NotePropertyValue $source.InputPath -Force
        $sourcePayload | Add-Member -NotePropertyName SourceType -NotePropertyValue $source.SourceType -Force
        $sourcePayload | Add-Member -NotePropertyName SourceLabel -NotePropertyValue $source.SourceLabel -Force
        $sourcePayloads.Add($sourcePayload) | Out-Null

        $sourceTenantName = Get-TenantDisplayName -LegacyData $null -OutputPrefix $source.SourceLabel
        if (-not [string]::IsNullOrWhiteSpace($sourceTenantName)) {
            $tenantNames.Add($sourceTenantName) | Out-Null
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$sourceResult.CustomerAssessmentReportPath)) {
            $customerAssessmentReportPaths.Add([string]$sourceResult.CustomerAssessmentReportPath) | Out-Null
        }
        if ($sourceResult.PSObject.Properties.Name -contains 'CustomerAssessmentReportMarkdownPath' -and -not [string]::IsNullOrWhiteSpace([string]$sourceResult.CustomerAssessmentReportMarkdownPath)) {
            $customerAssessmentReportMarkdownPaths.Add([string]$sourceResult.CustomerAssessmentReportMarkdownPath) | Out-Null
        }
        if ($sourceResult.PSObject.Properties.Name -contains 'RoadmapRemediationPlanPath' -and -not [string]::IsNullOrWhiteSpace([string]$sourceResult.RoadmapRemediationPlanPath)) {
            $roadmapRemediationPlanPaths.Add([string]$sourceResult.RoadmapRemediationPlanPath) | Out-Null
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$sourceResult.RemediationPs1Path) -and (Test-Path -Path $sourceResult.RemediationPs1Path)) {
            $snippetBlocks.Add("##############################") | Out-Null
            $snippetBlocks.Add("# Source: $($source.InputPath)") | Out-Null
            $snippetBlocks.Add("##############################") | Out-Null
            foreach ($line in (Get-Content -Path $sourceResult.RemediationPs1Path -Encoding UTF8)) {
                $snippetBlocks.Add($line) | Out-Null
            }
            $snippetBlocks.Add('') | Out-Null
        }
    }

    $tenantName = @($tenantNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique | Select-Object -First 1)
    if ($tenantName.Count -eq 0) {
        $tenantName = @((Split-Path -Path $resolvedInputRoot -Leaf))
    }

    $artifactPrefix = Get-TenantArtifactFilePrefix -TenantName ([string]$tenantName[0]) -Fallback $OutputPrefix
    $jsonOutPath = Join-Path -Path $supportFolder -ChildPath ("{0}-Plan.json" -f $artifactPrefix)
    $engineerMdOutPath = Join-Path -Path $supportFolder -ChildPath ("{0}-EngPack.md" -f $artifactPrefix)
    $snippetOutPath = Join-Path -Path $supportFolder -ChildPath ("{0}-Snips.ps1" -f $artifactPrefix)

    $engineerActionPackMarkdown = New-MultiSourceEngineerActionPack -SourcePayloads $sourcePayloads.ToArray() -TenantName ([string]$tenantName[0]) -GeneratedAt $generatedAt -SupportFolderPath $supportFolder -JsonOutPath $jsonOutPath -SnippetOutPath $snippetOutPath

    [System.IO.File]::WriteAllText($engineerMdOutPath, $engineerActionPackMarkdown, [System.Text.UTF8Encoding]::new($false))
    Set-Content -Path $snippetOutPath -Value ($snippetBlocks -join [Environment]::NewLine) -Encoding UTF8

    $multiSourcePayload = [PSCustomObject]@{
        GeneratedAt  = $generatedAt.ToString('o')
        SourceFolder = $resolvedInputRoot
        SourceCount  = $sourcePayloads.Count
        Sources      = $sourcePayloads.ToArray()
    }
    $multiSourcePayload | ConvertTo-Json -Depth 30 | Set-Content -Path $jsonOutPath -Encoding UTF8

    $outputSummary = [PSCustomObject]@{
        FindingsCount                        = (@($sourcePayloads | ForEach-Object { @(Convert-ArrayaObjectToArray $_.Findings).Count } | Measure-Object -Sum).Sum)
        SupportFolderPath                    = $supportFolder
        JsonPath                             = $jsonOutPath
        CsvPath                              = $null
        MarkdownPath                         = $null
        CustomerAssessmentReportPath         = $null
        CustomerAssessmentReportPaths        = $customerAssessmentReportPaths.ToArray()
        CustomerAssessmentReportMarkdownPath = $null
        CustomerAssessmentReportMarkdownPaths = $customerAssessmentReportMarkdownPaths.ToArray()
        RoadmapRemediationPlanPath          = $null
        RoadmapRemediationPlanPaths         = $roadmapRemediationPlanPaths.ToArray()
        EngineerActionPackPath               = $engineerMdOutPath
        RemediationPs1Path                   = $snippetOutPath
    }

    if (-not $Quiet) {
        Write-Host 'Improvement plan generated.'
        foreach ($customerAssessmentReportPath in $customerAssessmentReportPaths.ToArray()) {
            Write-Host "  Customer report : $customerAssessmentReportPath" -ForegroundColor Green
        }
        foreach ($customerAssessmentMarkdownPath in $customerAssessmentReportMarkdownPaths.ToArray()) {
            Write-Host "  Customer markdown: $customerAssessmentMarkdownPath" -ForegroundColor DarkGreen
        }
        foreach ($roadmapRemediationPlanPath in $roadmapRemediationPlanPaths.ToArray()) {
            Write-Host "  Roadmap report   : $roadmapRemediationPlanPath" -ForegroundColor Yellow
        }
        Write-Host "  Engineer pack    : $engineerMdOutPath" -ForegroundColor Cyan
        Write-Host "  Support folder   : $supportFolder" -ForegroundColor DarkGray
    }

    if ($PassThru) {
        $outputSummary
    }
    return
}

$sourceInputMetadata = $inputSources[0]
$AssessmentJsonPath = $sourceInputMetadata.InputPath

$snapshotContext = Import-ArrayaTenantSnapshotContext -Path $AssessmentJsonPath -Purpose ImprovementPlan
$AssessmentJsonPath = $snapshotContext.Path
$tenantData = $snapshotContext.LegacyData
$snapshotDerived = $snapshotContext.Derived
$snapshotDiagnostics = $snapshotContext.Diagnostics
$assessmentVersionLabel = Resolve-ArrayaAssessmentVersionLabel -SnapshotContext $snapshotContext

$outputContext = Resolve-ArrayaSnapshotOutputContext -PrimaryInputPath $AssessmentJsonPath -OutputFolder $OutputFolder -OutputPrefix $OutputPrefix
$OutputFolder = $outputContext.OutputFolder
$OutputPrefix = Get-ImprovementOutputPrefix -OutputPrefix $outputContext.OutputPrefix

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
$recipientRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('AllRecipients', 'Recipients'))
$mailboxRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('AllMailboxes', 'MailboxFullDetails'))
$primaryMailboxStatsRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('PrimaryMailboxStats'))
$archiveMailboxRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('ArchiveMailboxes'))
$archiveMailboxStatsRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('ArchiveMailboxStats'))
$inactiveMailboxRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('InactiveMailboxes', 'InactiveMailboxDetails'))
$nonUserMailboxRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('NonUserMailboxes'))
$emailActivitySummary = Get-ArrayaObjectValue -Object $tenantData -Names @('EmailActivitySummary')
$emailActivityTopSenders = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('EmailActivityTopSenders'))
$emailActivityTopReceivers = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('EmailActivityTopReceivers'))
$connectorRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('MailFlowConnectors'))
$remoteDomainRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('RemoteDomains'))
$publicFolderRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('PublicFolderDetails'))
$smtpRelayServiceAccounts = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('SMTPRelayServiceAccounts'))
$sharePointRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('SharePoint'))
$oneDriveRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('OneDrive'))
$teamRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('AllTeams'))
$unifiedGroupRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('UnifiedGroups', 'EntraIDGroups'))
$smtpRelaySummary = Get-ArrayaObjectValue -Object $tenantData -Names @('SMTPRelaySummary')
$smtpRelayConfig = Get-ArrayaObjectValue -Object $tenantData -Names @('SMTPRelayConfig')
$authConfig = Get-ArrayaObjectValue -Object $tenantData -Names @('AuthenticationConfig')
$mfaSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('MfaRegistrationSummary', 'MFARegistrationSummary', 'MfaRegistration', 'MFARegistration')
$mfaRegistrationDetails = @(
    Convert-ToImprovementCollectionRows `
        -Value (Get-ArrayaObjectValue -Object $tenantData -Names @('MfaRegistrationDetails', 'MFARegistrationDetails')) `
        -MarkerNames @('UserPrincipalName', 'Id', 'DisplayName')
)
$mfaEnrollmentSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('MfaEnrollmentSummary')
$mfaEnforcementSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('MfaEnforcementSummary')
$mfaEnforcementGapUsers = @(
    Convert-ToImprovementCollectionRows `
        -Value (Get-ArrayaObjectValue -Object $tenantData -Names @('MfaEnforcementGapUsers')) `
        -MarkerNames @('DisplayName', 'UserPrincipalName', 'GapCategory')
)
$mfaEnforcementScopeReview = @(
    Convert-ToImprovementCollectionRows `
        -Value (Get-ArrayaObjectValue -Object $tenantData -Names @('MfaEnforcementScopeReview')) `
        -MarkerNames @('PolicyName', 'ScopeType', 'DisplayName')
)
$adminMfaSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('AdminMfaSummary')
$adminMfaRegistrationGaps = @(
    Convert-ToImprovementCollectionRows `
        -Value (Get-ArrayaObjectValue -Object $tenantData -Names @('AdminMfaRegistrationGaps')) `
        -MarkerNames @('DisplayName', 'UserPrincipalName', 'MfaRegistrationState')
)
$adminMfaEnforcementGaps = @(
    Convert-ToImprovementCollectionRows `
        -Value (Get-ArrayaObjectValue -Object $tenantData -Names @('AdminMfaEnforcementGaps')) `
        -MarkerNames @('DisplayName', 'UserPrincipalName', 'MfaEnforcementState')
)
$ownershipSummary = Get-ArrayaObjectValue -Object $snapshotDerived -Names @('OwnershipGovernanceSummary')
$unmanagedObjects = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $snapshotDerived -Names @('UnmanagedObjects'))
$oneDriveOwnerMismatches = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $snapshotDerived -Names @('OneDriveOwnerMismatches'))
$tenantInfoSummary = Get-ArrayaObjectValue -Object $snapshotDerived -Names @('TenantInfoSummary')
$authSummary = Get-ArrayaObjectValue -Object $snapshotDerived -Names @('AuthenticationConfigSummary')
$passwordLifecycleSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('PasswordLifecycleSummary')
if (-not $passwordLifecycleSummary) {
    $passwordLifecycleSummary = Get-ArrayaObjectValue -Object $snapshotDerived -Names @('PasswordLifecycleSummary')
}
$mfaDerivedSummary = Get-ArrayaObjectValue -Object $snapshotDerived -Names @('MfaRegistrationSummary')
$mfaEnrollmentDerivedSummary = Get-ArrayaObjectValue -Object $snapshotDerived -Names @('MfaEnrollmentSummary')
$mfaEnforcementDerivedSummary = Get-ArrayaObjectValue -Object $snapshotDerived -Names @('MfaEnforcementSummary')
$conditionalAccessSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('ConditionalAccessPolicySummary')
$securityDefaultsPolicy = Get-ArrayaObjectValue -Object $tenantData -Names @('SecurityDefaultsPolicy')
$rawEnterpriseApplications = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('EnterpriseApplications'))
$authenticationSsoApplications = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('AuthenticationSSOApplications'))
$enterpriseApplicationByIdentity = @{}
$enterpriseApplications = New-Object System.Collections.Generic.List[object]

function Test-AssessmentEnterpriseApplicationRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow
    )

    if ($null -eq $ApplicationRow) {
        return $false
    }

    if (-not $ApplicationRow.PSObject -or $ApplicationRow.PSObject.Properties.Count -eq 0) {
        return $false
    }

    $identitySignals = @(
        (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('ServicePrincipalId', 'Id')),
        (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('AppId')),
        (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('DisplayName')),
        (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('SSOMode', 'PreferredSingleSignOnMode'))
    )

    foreach ($signal in @($identitySignals)) {
        $signalText = Convert-ToArrayaDisplayText -Value $signal -Default ''
        if (-not [string]::IsNullOrWhiteSpace($signalText)) {
            return $true
        }
    }

    return (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('SsoEnabled'))) -eq $true
}

function Get-AssessmentEnterpriseApplicationSignalState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow,
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $false)]
        [string]$Default = 'Unavailable'
    )

    $state = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $ApplicationRow -Names @($Name)) -Default $Default
    if ([string]::IsNullOrWhiteSpace($state)) {
        return $Default
    }

    return $state
}

function Test-AssessmentEnterpriseApplicationSsoEnabled {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $ApplicationRow
    )

    $explicitSsoEnabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('SsoEnabled'))
    $preferredSingleSignOnMode = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $ApplicationRow -Names @('SSOMode', 'PreferredSingleSignOnMode')) -Default ''
    return (
        ($explicitSsoEnabled -eq $true) -or
        (-not [string]::IsNullOrWhiteSpace($preferredSingleSignOnMode) -and $preferredSingleSignOnMode -ne 'notSupported')
    )
}

function Get-AssessmentEnterpriseApplicationCoverageState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$TotalCount,
        [Parameter(Mandatory = $false)]
        [int]$CollectedCount = 0,
        [Parameter(Mandatory = $false)]
        [int]$PartialCount = 0,
        [Parameter(Mandatory = $false)]
        [int]$UnavailableCount = 0,
        [Parameter(Mandatory = $false)]
        [int]$NotApplicableCount = 0,
        [Parameter(Mandatory = $false)]
        [switch]$SupportsNotApplicable
    )

    if ($TotalCount -eq 0) {
        return 'Unavailable'
    }

    if ($SupportsNotApplicable -and (($CollectedCount + $PartialCount + $UnavailableCount) -eq 0) -and $NotApplicableCount -gt 0) {
        return 'NotApplicable'
    }

    if (($PartialCount + $UnavailableCount) -eq 0) {
        return 'Collected'
    }

    if (($CollectedCount + $PartialCount) -gt 0) {
        return 'Partial'
    }

    return 'Unavailable'
}

function New-AssessmentEnterpriseApplicationSummaryFromRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object[]]$ApplicationRows = @()
    )

    $rows = @($ApplicationRows)
    $applicationSourceCollectedCount = @($rows | Where-Object { (Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ApplicationSourceState') -eq 'Collected' }).Count
    $applicationSourceUnavailableCount = @($rows | Where-Object { (Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ApplicationSourceState') -eq 'Unavailable' }).Count
    $ownerSignalCollectedCount = @($rows | Where-Object { (Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'OwnerSignalState') -eq 'Collected' }).Count
    $ownerSignalUnavailableCount = @($rows | Where-Object { (Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'OwnerSignalState') -eq 'Unavailable' }).Count
    $ownerSignalNotApplicableCount = @($rows | Where-Object { (Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'OwnerSignalState') -eq 'NotApplicable' }).Count
    $redirectUriSignalCollectedCount = @($rows | Where-Object { (Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'RedirectUriSignalState') -eq 'Collected' }).Count
    $redirectUriSignalPartialCount = @($rows | Where-Object { (Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'RedirectUriSignalState') -eq 'Partial' }).Count
    $redirectUriSignalUnavailableCount = @($rows | Where-Object { (Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'RedirectUriSignalState') -eq 'Unavailable' }).Count
    $activitySignalCollectedCount = @($rows | Where-Object { (Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState') -eq 'Collected' }).Count
    $activitySignalPartialCount = @($rows | Where-Object { (Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState') -eq 'Partial' }).Count
    $activitySignalUnavailableCount = @($rows | Where-Object { (Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState') -eq 'Unavailable' }).Count

    return [pscustomobject]@{
        TotalEnterpriseApplications      = $rows.Count
        FirstPartyApplications           = @($rows | Where-Object { [string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'First Party' }).Count
        ThirdPartyApplications           = @($rows | Where-Object { [string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'Third Party' }).Count
        UnknownSourceApplications        = @($rows | Where-Object { [string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'Unknown' }).Count
        ApplicationsWithHighPrivilege    = @($rows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('HighPrivilegePermissionCount'))) -gt 0 }).Count
        ApplicationsWithDelegatedGrants  = @($rows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('DelegatedPermissionGrantCount'))) -gt 0 }).Count
        ApplicationsWithApplicationPerms = @($rows | Where-Object { (Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('ApplicationPermissionCount'))) -gt 0 }).Count
        FirstPartyAppsWithoutOwners      = @($rows | Where-Object {
            ([string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'First Party') -and
            ((Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'OwnerSignalState') -eq 'Collected') -and
            ((Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OwnerCount'))) -le 0)
        }).Count
        ThirdPartyAppsWithApplicationPerms = @($rows | Where-Object {
            ([string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'Third Party') -and
            ((Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('ApplicationPermissionCount'))) -gt 0)
        }).Count
        ApplicationsWithNoRecentActivity = @($rows | Where-Object {
            ((Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState') -eq 'Collected') -and
            ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasRecentActivity'))) -eq $false)
        }).Count
        ApplicationsWithInsecureRedirectUris = @($rows | Where-Object {
            $redirectState = Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'RedirectUriSignalState'
            (($redirectState -eq 'Collected') -or ($redirectState -eq 'Partial')) -and
            ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasInsecureRedirectUris'))) -eq $true)
        }).Count
        SsoEnabledApplications           = @($rows | Where-Object { Test-AssessmentEnterpriseApplicationSsoEnabled -ApplicationRow $_ }).Count
        ApplicationSourceCoverageState   = Get-AssessmentEnterpriseApplicationCoverageState -TotalCount $rows.Count -CollectedCount $applicationSourceCollectedCount -UnavailableCount $applicationSourceUnavailableCount
        ApplicationSourceCollectedCount  = $applicationSourceCollectedCount
        ApplicationSourceUnavailableCount = $applicationSourceUnavailableCount
        OwnerSignalCoverageState         = Get-AssessmentEnterpriseApplicationCoverageState -TotalCount $rows.Count -CollectedCount $ownerSignalCollectedCount -UnavailableCount $ownerSignalUnavailableCount -NotApplicableCount $ownerSignalNotApplicableCount -SupportsNotApplicable
        OwnerSignalCollectedCount        = $ownerSignalCollectedCount
        OwnerSignalUnavailableCount      = $ownerSignalUnavailableCount
        OwnerSignalNotApplicableCount    = $ownerSignalNotApplicableCount
        RedirectUriSignalCoverageState   = Get-AssessmentEnterpriseApplicationCoverageState -TotalCount $rows.Count -CollectedCount $redirectUriSignalCollectedCount -PartialCount $redirectUriSignalPartialCount -UnavailableCount $redirectUriSignalUnavailableCount
        RedirectUriSignalCollectedCount  = $redirectUriSignalCollectedCount
        RedirectUriSignalPartialCount    = $redirectUriSignalPartialCount
        RedirectUriSignalUnavailableCount = $redirectUriSignalUnavailableCount
        ActivitySignalCoverageState      = Get-AssessmentEnterpriseApplicationCoverageState -TotalCount $rows.Count -CollectedCount $activitySignalCollectedCount -PartialCount $activitySignalPartialCount -UnavailableCount $activitySignalUnavailableCount
        ActivitySignalCollectedCount     = $activitySignalCollectedCount
        ActivitySignalPartialCount       = $activitySignalPartialCount
        ActivitySignalUnavailableCount   = $activitySignalUnavailableCount
    }
}

foreach ($applicationRow in @(@($rawEnterpriseApplications) + @($authenticationSsoApplications))) {
    if (-not (Test-AssessmentEnterpriseApplicationRow -ApplicationRow $applicationRow)) {
        continue
    }

    $identityParts = @(
        Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $applicationRow -Names @('ServicePrincipalId', 'Id')) -Default ''
        Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $applicationRow -Names @('AppId')) -Default ''
        Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $applicationRow -Names @('DisplayName')) -Default ''
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    if ($identityParts.Count -eq 0) {
        $enterpriseApplications.Add($applicationRow) | Out-Null
        continue
    }

    $applicationIdentity = ($identityParts -join '|').ToLowerInvariant()
    if ($enterpriseApplicationByIdentity.ContainsKey($applicationIdentity)) {
        $existing = $enterpriseApplicationByIdentity[$applicationIdentity]
        $mergedProperties = [ordered]@{}
        foreach ($property in @($existing.PSObject.Properties)) {
            $mergedProperties[$property.Name] = $property.Value
        }
        foreach ($property in @($applicationRow.PSObject.Properties)) {
            $currentValue = $null
            $hasCurrentValue = $mergedProperties.Contains($property.Name)
            if ($hasCurrentValue) {
                $currentValue = $mergedProperties[$property.Name]
            }

            if (-not $hasCurrentValue -or $null -eq $currentValue -or [string]::IsNullOrWhiteSpace([string]$currentValue)) {
                $mergedProperties[$property.Name] = $property.Value
            }
        }
        $enterpriseApplicationByIdentity[$applicationIdentity] = [pscustomobject]$mergedProperties
        continue
    }

    $enterpriseApplicationByIdentity[$applicationIdentity] = $applicationRow
}

foreach ($applicationIdentity in @($enterpriseApplicationByIdentity.Keys | Sort-Object)) {
    $enterpriseApplications.Add($enterpriseApplicationByIdentity[$applicationIdentity]) | Out-Null
}
$enterpriseApplications = @($enterpriseApplications.ToArray())
$enterpriseApplicationSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('EnterpriseApplicationSummary')
$enterpriseApplicationRowSummaryRecord = if ($enterpriseApplications.Count -gt 0) { New-AssessmentEnterpriseApplicationSummaryFromRows -ApplicationRows $enterpriseApplications } else { $null }
$existingEnterpriseApplicationSummaryRecord = if ($enterpriseApplicationSummary) { Get-ArrayaObjectValue -Object $enterpriseApplicationSummary -Names @('Summary') } else { $null }
$existingTotalEnterpriseApplications = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $existingEnterpriseApplicationSummaryRecord -Names @('TotalEnterpriseApplications', 'EnterpriseApplicationCount'))
$preferRowBasedEnterpriseApplicationSummary = ($rawEnterpriseApplications.Count -gt 0) -and ($null -ne $enterpriseApplicationRowSummaryRecord)

if ($preferRowBasedEnterpriseApplicationSummary) {
    $enterpriseApplicationSummary = [pscustomobject]@{
        Summary = $enterpriseApplicationRowSummaryRecord
    }
}
elseif ($null -eq $enterpriseApplicationSummary -and $null -ne $enterpriseApplicationRowSummaryRecord) {
    $enterpriseApplicationSummary = [pscustomobject]@{
        Summary = $enterpriseApplicationRowSummaryRecord
    }
}
elseif ($null -ne $enterpriseApplicationSummary -and $null -ne $enterpriseApplicationRowSummaryRecord) {
    $mergedEnterpriseApplicationSummary = [ordered]@{}
    if ($null -ne $existingEnterpriseApplicationSummaryRecord) {
        foreach ($property in @($existingEnterpriseApplicationSummaryRecord.PSObject.Properties)) {
            $mergedEnterpriseApplicationSummary[$property.Name] = $property.Value
        }
    }

    foreach ($property in @($enterpriseApplicationRowSummaryRecord.PSObject.Properties)) {
        $hasCurrentValue = $mergedEnterpriseApplicationSummary.Contains($property.Name)
        $currentValue = if ($hasCurrentValue) { $mergedEnterpriseApplicationSummary[$property.Name] } else { $null }
        if (-not $hasCurrentValue -or $null -eq $currentValue -or [string]::IsNullOrWhiteSpace([string]$currentValue)) {
            $mergedEnterpriseApplicationSummary[$property.Name] = $property.Value
        }
    }

    if ($null -eq $existingTotalEnterpriseApplications -or $existingTotalEnterpriseApplications -eq 0) {
        foreach ($property in @($enterpriseApplicationRowSummaryRecord.PSObject.Properties)) {
            $mergedEnterpriseApplicationSummary[$property.Name] = $property.Value
        }
    }

    $enterpriseApplicationSummary = [pscustomobject]@{
        Summary = [pscustomobject]$mergedEnterpriseApplicationSummary
    }
}
$guestSignInSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('GuestSignInSummary')
$privilegedAccessSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('PrivilegedAccessSummary')
$inboxRulesExternalForwarding = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('InboxRulesExternalForwarding'))
$inboxRuleForwardingSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('InboxRuleForwardingSummary')
$forwardingPolicySummary = Get-ArrayaObjectValue -Object $tenantData -Names @('ForwardingPolicySummary')
$sharedMailboxGovernanceSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('SharedMailboxGovernanceSummary')
$sharePointSharingSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('SharePointSharingSummary')
$externalSharingSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('ExternalSharingSummary')
$externalSharingSiteOverrides = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('ExternalSharingSiteOverrides'))
$externalExposureFindings = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('ExternalExposureFindings'))
$collaborationActivitySummary = Get-ArrayaObjectValue -Object $tenantData -Names @('CollaborationActivitySummary')
$teamsVoiceSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('TeamsVoiceSummary')
$deviceManagementSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('DeviceManagementSummary')
$spamFilteringSummary = Get-ArrayaObjectValue -Object $tenantData -Names @('SpamFilteringSummary')
$adConnectConfiguration = Get-ArrayaObjectValue -Object $tenantData -Names @('AdConnectConfiguration')
$externalIdentityRestrictions = Get-ArrayaObjectValue -Object $tenantData -Names @('ExternalIdentityRestrictions')
$guestAccessConfiguration = Get-ArrayaObjectValue -Object $tenantData -Names @('GuestAccessConfiguration')
$retentionPolicyRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('RetentionPolicies', 'CompliancePolicies'))
$dlpPolicyRows = Convert-ArrayaObjectToArray (Get-ArrayaObjectValue -Object $tenantData -Names @('DlpPolicies'))

if (-not $mfaEnrollmentSummary) {
    $mfaEnrollmentSummary = $mfaEnrollmentDerivedSummary
}
if (-not $mfaEnforcementSummary) {
    $mfaEnforcementSummary = $mfaEnforcementDerivedSummary
}
if (-not $mfaSummary) {
    $mfaSummary = $mfaDerivedSummary
}

$mfaEnrollmentSummary = Get-CustomerMfaEnrollmentSummary -ExistingSummary $mfaEnrollmentSummary -MfaSummary $mfaSummary -MfaRegistrationDetails $mfaRegistrationDetails
$mfaEnforcementSummary = Get-CustomerMfaEnforcementSummary -ExistingSummary $mfaEnforcementSummary -ConditionalAccessPolicies $caPolicies -ConditionalAccessSummary $conditionalAccessSummary -SecurityDefaultsPolicy $securityDefaultsPolicy

if ($retentionPolicyRows.Count -eq 0 -and $mailboxRows.Count -gt 0) {
    $retentionPolicyRows = @(
        $mailboxRows |
            Group-Object {
                $policyName = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('RetentionPolicy')) -Default ''
                $hasHoldSignals =
                    (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('LitigationHoldEnabled'))) -eq $true -or
                    (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('RetentionHoldEnabled'))) -eq $true -or
                    (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DelayHoldApplied'))) -eq $true

                if (-not [string]::IsNullOrWhiteSpace($policyName)) {
                    return $policyName
                }
                if ($hasHoldSignals) {
                    return 'HoldSignalsWithoutNamedPolicy'
                }

                return $null
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Name) } |
            ForEach-Object {
                [pscustomobject]@{
                    PolicyName                 = $_.Name
                    MailboxCount               = $_.Count
                    LitigationHoldMailboxCount = @($_.Group | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('LitigationHoldEnabled'))) -eq $true }).Count
                    RetentionHoldMailboxCount  = @($_.Group | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('RetentionHoldEnabled'))) -eq $true }).Count
                    DelayHoldMailboxCount      = @($_.Group | Where-Object { (Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('DelayHoldApplied'))) -eq $true }).Count
                }
            }
    )
}

if (-not $passwordLifecycleSummary -and $adConnectConfiguration) {
    $adConnectSummary = Get-ArrayaObjectValue -Object $adConnectConfiguration -Names @('Summary')
    if ($adConnectSummary) {
        $derivedPasswordLifecycleSummary = [ordered]@{}
        foreach ($field in @(
            'PasswordWriteback',
            'PasswordWritebackEnabled',
            'PassThroughAuthentication',
            'PassThroughAuthenticationEnabled',
            'SelfServicePasswordReset',
            'SelfServicePasswordResetEnabled',
            'OnPremisesSyncEnabled',
            'OnPremisesLastSyncDateTime'
        )) {
            $fieldValue = Get-ArrayaObjectValue -Object $adConnectSummary -Names @($field)
            if ($null -ne $fieldValue) {
                $derivedPasswordLifecycleSummary[$field] = $fieldValue
            }
        }

        if ($derivedPasswordLifecycleSummary.Count -gt 0) {
            $passwordLifecycleSummary = [pscustomobject]$derivedPasswordLifecycleSummary
        }
    }
}

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
    $registeredUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummary -Names @('RegisteredUsers', 'RegisteredUserCount', 'MfaRegisteredUsers'))
    $totalUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummary -Names @('TotalUsers', 'UserCount', 'TotalUserCount'))
    $registrationPct = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummary -Names @('RegistrationPercent', 'RegisteredPercent', 'MfaRegistrationPercent'))
    if ($null -eq $registrationPct -and $null -ne $registeredUsers -and $null -ne $totalUsers -and $totalUsers -gt 0) {
        $registrationPct = [math]::Round(($registeredUsers / $totalUsers) * 100, 2)
    }
    if ($null -ne $registrationPct) {
        if ($registrationPct -lt 70) {
            Add-HeuristicFinding -Store $findingStore -RuleId 'MFA-001' -Area 'Identity Governance' -Category 'Identity Governance' -Severity 'High' -Finding 'MFA enrollment appears low.' -Recommendation 'Drive registration completion, reduce weak-method reliance, and then confirm enforcement through Conditional Access.' -CurrentValue "$registrationPct%" -TargetValue '>= 90%' -RelatedWorksheet 'MfaEnrollmentSummary' -RelatedSection 'MFA Enrollment'
        } elseif ($registrationPct -lt 90) {
            Add-HeuristicFinding -Store $findingStore -RuleId 'MFA-001' -Area 'Identity Governance' -Category 'Identity Governance' -Severity 'Medium' -Finding 'MFA enrollment is below the target adoption threshold.' -Recommendation 'Drive remaining users through registration completion and verify which methods they are using before broad enforcement.' -CurrentValue "$registrationPct%" -TargetValue '>= 90%' -RelatedWorksheet 'MfaEnrollmentSummary' -RelatedSection 'MFA Enrollment'
        }
    }
}

if (-not (Test-DerivedCoverage -Tags @('mfa method', 'weak method') -DerivedFindings $derivedFindings)) {
    $weakMethodOnlyUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummary -Names @('UsersWithWeakMethodsOnly'))
    $weakDefaultMethodUsers = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnrollmentSummary -Names @('UsersWithWeakDefaultMethod'))
    $weakMethodBreakdown = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $mfaEnrollmentSummary -Names @('WeakMethodBreakdown')) -Default ''
    if (($null -ne $weakMethodOnlyUsers -and $weakMethodOnlyUsers -gt 0) -or ($null -ne $weakDefaultMethodUsers -and $weakDefaultMethodUsers -gt 0)) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'MFA-002' -Area 'Identity Governance' -Category 'Identity Governance' -Severity 'Medium' -Finding 'MFA enrollment still relies on weaker methods for part of the tenant.' -Recommendation 'Move users from SMS, voice, and email-based MFA toward stronger app-based or phishing-resistant methods, with exceptions documented explicitly.' -CurrentValue ("{0} user(s) have only weak methods; {1} user(s) default to weak methods{2}" -f $(if ($null -eq $weakMethodOnlyUsers) { 0 } else { $weakMethodOnlyUsers }), $(if ($null -eq $weakDefaultMethodUsers) { 0 } else { $weakDefaultMethodUsers }), $(if ([string]::IsNullOrWhiteSpace($weakMethodBreakdown)) { '' } else { "; Weak methods observed: $weakMethodBreakdown" })) -TargetValue 'Registered users rely on stronger MFA methods, with weak methods limited to approved exceptions' -RelatedWorksheet 'MfaEnrollmentSummary' -RelatedSection 'MFA Enrollment'
    }
}

if (-not (Test-DerivedCoverage -Tags @('mfa enforcement') -DerivedFindings $derivedFindings)) {
    $enabledMfaPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummary -Names @('EnabledPoliciesRequiringMfa'))
    $reportOnlyMfaPolicies = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $mfaEnforcementSummary -Names @('ReportOnlyPoliciesRequiringMfa'))
    $securityDefaultsEnabled = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $mfaEnforcementSummary -Names @('SecurityDefaultsEnabled'))
    if ((($null -eq $enabledMfaPolicies) -or ($enabledMfaPolicies -eq 0)) -and ($securityDefaultsEnabled -ne $true)) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'MFA-003' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'High' -Finding 'Active MFA enforcement was not clearly detected.' -Recommendation 'Move validated MFA coverage into enabled Conditional Access policies or confirm that Security Defaults is the intended enforcement model.' -CurrentValue ("Enabled MFA enforcement policies={0}; report-only MFA policies={1}; Security Defaults={2}" -f $(if ($null -eq $enabledMfaPolicies) { 0 } else { $enabledMfaPolicies }), $(if ($null -eq $reportOnlyMfaPolicies) { 0 } else { $reportOnlyMfaPolicies }), $(if ($securityDefaultsEnabled -eq $true) { 'Enabled' } elseif ($securityDefaultsEnabled -eq $false) { 'Disabled' } else { 'Not validated' })) -TargetValue 'MFA enforcement active for the intended population through one documented baseline model' -RelatedWorksheet 'MfaEnforcementSummary' -RelatedSection 'MFA Enforcement'
    }
}

if (-not (Test-DerivedCoverage -Tags @('global administrator', 'privileged') -DerivedFindings $derivedFindings) -and $snapshotMetrics.GlobalAdminCount -gt $MaxGlobalAdmins) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'ADMIN-001' -Area 'Identity Governance' -Category 'Privileged Access' -Severity 'High' -Finding 'Global administrator count exceeds the recommended threshold.' -Recommendation 'Reduce active Global Administrator assignments, remove stale admins from the role entirely, move infrequent administrators to lower-privilege roles such as Global Reader where possible, and use eligible activation with approval for full tenant-wide access where that operating model is supported.' -CurrentValue "$($snapshotMetrics.GlobalAdminCount) accounts" -TargetValue "<= $MaxGlobalAdmins accounts" -RelatedWorksheet 'Admins' -RelatedSection 'Privileged Access'
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
    Add-HeuristicFinding -Store $findingStore -RuleId 'ID-005' -Area 'Identity Governance' -Category 'Identity Governance' -Severity $(if ($inactiveGuests90Days -ge 10) { 'High' } else { 'Medium' }) -Finding 'Inactive guest accounts were identified by the guest sign-in governance summary.' -Recommendation 'Review stale guest identities, validate sponsor ownership, and remove or disable guests that are no longer required.' -CurrentValue "$inactiveGuests90Days inactive guest account(s) over 90 days" -TargetValue 'Inactive guest accounts reviewed and dispositioned' -Source 'Summary/GuestSignIn' -RelatedWorksheet 'GuestSignInSummary' -RelatedSection 'Guest Access'
}

$privilegedSummaryRecord = if ($privilegedAccessSummary) { Get-ArrayaObjectValue -Object $privilegedAccessSummary -Names @('Summary') } else { $null }
$stalePrivilegedSummary = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $privilegedSummaryRecord -Names @('StalePrivilegedAccounts90Days'))
if ($null -ne $stalePrivilegedSummary -and $stalePrivilegedSummary -gt 0) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'ID-006' -Area 'Identity Governance' -Category 'Identity Governance' -Severity 'Medium' -Finding 'Privileged identities with stale sign-in activity were identified by the privileged-access governance summary.' -Recommendation 'Review stale privileged accounts, remove unused role assignments, and validate emergency access documentation.' -CurrentValue "$stalePrivilegedSummary stale privileged account(s) over 90 days" -TargetValue '0 stale privileged accounts' -Source 'Summary/PrivilegedAccess' -RelatedWorksheet 'PrivilegedAccessSummary' -RelatedSection 'Privileged Access'
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

$thirdPartyAppsWithApplicationPermissions = @(
    $enterpriseApplications | Where-Object {
        ([string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'Third Party') -and
        ((Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ApplicationSourceState') -eq 'Collected') -and
        ((Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('ApplicationPermissionCount'))) -gt 0)
    }
)
if ($thirdPartyAppsWithApplicationPermissions.Count -gt 0) {
    $thirdPartyApplicationExamples = @(
        $thirdPartyAppsWithApplicationPermissions |
            ForEach-Object { Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default '' } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First 5
    )
    $currentValue = "{0} third-party enterprise application(s) have application permissions" -f $thirdPartyAppsWithApplicationPermissions.Count
    if ($thirdPartyApplicationExamples.Count -gt 0) {
        $currentValue = "{0}; examples: {1}" -f $currentValue, ($thirdPartyApplicationExamples -join ', ')
    }

    Add-HeuristicFinding -Store $findingStore -RuleId 'ID-009' -Area 'Identity Governance' -Category 'Identity Governance' -Severity $(if ($thirdPartyAppsWithApplicationPermissions.Count -ge 5) { 'High' } else { 'Medium' }) -Finding 'Third-party enterprise applications with application permissions were detected.' -Recommendation 'Review third-party apps that have application permissions, validate business need and vendor trust, and reduce consent scope where possible.' -CurrentValue $currentValue -TargetValue 'Third-party application permissions reviewed and reduced to approved least-privilege scope' -Source 'Summary/EnterpriseApplications' -RelatedWorksheet 'EnterpriseApplications' -RelatedSection 'Enterprise Apps'
}

$firstPartyAppsWithoutOwners = @(
    $enterpriseApplications | Where-Object {
        ([string](Get-ArrayaObjectValue -Object $_ -Names @('ApplicationSource')) -eq 'First Party') -and
        ((Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'OwnerSignalState') -eq 'Collected') -and
        ((Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('OwnerCount'))) -le 0)
    }
)
if ($firstPartyAppsWithoutOwners.Count -gt 0) {
    $ownerlessExamples = @(
        $firstPartyAppsWithoutOwners |
            ForEach-Object { Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default '' } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First 5
    )
    $currentValue = "{0} first-party enterprise application(s) do not show an owner signal" -f $firstPartyAppsWithoutOwners.Count
    if ($ownerlessExamples.Count -gt 0) {
        $currentValue = "{0}; examples: {1}" -f $currentValue, ($ownerlessExamples -join ', ')
    }

    Add-HeuristicFinding -Store $findingStore -RuleId 'ID-010' -Area 'Identity Governance' -Category 'Identity Governance' -Severity $(if ($firstPartyAppsWithoutOwners.Count -ge 10) { 'High' } else { 'Medium' }) -Finding 'First-party enterprise applications without an owner signal were detected.' -Recommendation 'Assign accountable owners to first-party app registrations and retire applications that no longer have an approved sponsor.' -CurrentValue $currentValue -TargetValue 'Each first-party app registration has a documented owner and approved business purpose' -Source 'Summary/EnterpriseApplications' -RelatedWorksheet 'EnterpriseApplications' -RelatedSection 'Enterprise Apps'
}

$appsWithRedirectUriRisk = @(
    $enterpriseApplications | Where-Object {
        $redirectSignalState = Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'RedirectUriSignalState'
        (($redirectSignalState -eq 'Collected') -or ($redirectSignalState -eq 'Partial')) -and
        ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasInsecureRedirectUris'))) -eq $true)
    }
)
if ($appsWithRedirectUriRisk.Count -gt 0) {
    $redirectUriRiskExamples = @(
        $appsWithRedirectUriRisk |
            ForEach-Object {
                $displayName = Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default ''
                $riskCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $_ -Names @('InsecureRedirectUriCount'))
                if (-not [string]::IsNullOrWhiteSpace($displayName)) {
                    if ($null -ne $riskCount -and $riskCount -gt 0) {
                        '{0} ({1} flagged URI(s))' -f $displayName, $riskCount
                    }
                    else {
                        $displayName
                    }
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First 5
    )
    $currentValue = "{0} enterprise application(s) include redirect URI patterns that warrant security review" -f $appsWithRedirectUriRisk.Count
    if ($redirectUriRiskExamples.Count -gt 0) {
        $currentValue = "{0}; examples: {1}" -f $currentValue, ($redirectUriRiskExamples -join ', ')
    }

    Add-HeuristicFinding -Store $findingStore -RuleId 'ID-011' -Area 'Identity Governance' -Category 'Identity Governance' -Severity 'Medium' -Finding 'Enterprise applications with redirect URI patterns that warrant review were detected.' -Recommendation 'Review flagged redirect URIs, remove legacy or broad redirect patterns, and keep only approved production endpoints.' -CurrentValue $currentValue -TargetValue 'Redirect URIs reduced to approved, documented endpoints' -Source 'Summary/EnterpriseApplications' -RelatedWorksheet 'EnterpriseApplications' -RelatedSection 'Enterprise Apps'
}

$inactiveEnterpriseApps = @(
    $enterpriseApplications | Where-Object {
        ((Get-AssessmentEnterpriseApplicationSignalState -ApplicationRow $_ -Name 'ActivitySignalState') -eq 'Collected') -and
        ((Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $_ -Names @('HasRecentActivity'))) -eq $false)
    }
)
if ($inactiveEnterpriseApps.Count -gt 0) {
    $inactiveExamples = @(
        $inactiveEnterpriseApps |
            ForEach-Object { Convert-ToArrayaDisplayText -Value (Get-ArrayaObjectValue -Object $_ -Names @('DisplayName')) -Default '' } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First 5
    )
    $currentValue = "{0} enterprise application(s) do not show a recent activity signal" -f $inactiveEnterpriseApps.Count
    if ($inactiveExamples.Count -gt 0) {
        $currentValue = "{0}; examples: {1}" -f $currentValue, ($inactiveExamples -join ', ')
    }

    Add-HeuristicFinding -Store $findingStore -RuleId 'ID-012' -Area 'Identity Governance' -Category 'Identity Governance' -Severity $(if ($inactiveEnterpriseApps.Count -ge 10) { 'Medium' } else { 'Low' }) -Finding 'Enterprise applications without a recent activity signal were detected.' -Recommendation 'Review inactive enterprise apps, validate whether they are still required, and disable or remove stale integrations where appropriate.' -CurrentValue $currentValue -TargetValue 'Inactive enterprise apps reviewed and dispositioned' -Source 'Summary/EnterpriseApplications' -RelatedWorksheet 'EnterpriseApplications' -RelatedSection 'Enterprise Apps'
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
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-008' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Medium' -Finding 'Conditional Access summary does not show guest or external-user coverage.' -Recommendation 'Add or validate guest and external-user Conditional Access coverage.' -CurrentValue 'Guest/external-user coverage not detected in summary' -TargetValue 'Guest/external-user coverage documented and enabled' -Source 'Summary/ConditionalAccess' -RelatedWorksheet 'ConditionalAccessPolicySummary' -RelatedSection 'Conditional Access'
    }
    if ($hasPrivilegedCoverage -eq $false) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-009' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Medium' -Finding 'Conditional Access summary does not show privileged-role coverage.' -Recommendation 'Add or validate privileged-role Conditional Access protection for administrative identities.' -CurrentValue 'Privileged-role coverage not detected in summary' -TargetValue 'Privileged-role coverage documented and enabled' -Source 'Summary/ConditionalAccess' -RelatedWorksheet 'ConditionalAccessPolicySummary' -RelatedSection 'Conditional Access'
    }
    if ($hasCompliantDeviceRequirement -eq $false) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-010' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Low' -Finding 'Conditional Access summary does not show a compliant-device requirement.' -Recommendation 'Decide whether managed-access scenarios should require compliant devices.' -CurrentValue 'Compliant-device requirement not detected in summary' -TargetValue 'Compliant-device requirement reviewed and documented' -Source 'Summary/ConditionalAccess' -RelatedWorksheet 'ConditionalAccessPolicySummary' -RelatedSection 'Conditional Access'
    }
    if ($hasRiskCoverage -eq $false) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-011' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Low' -Finding 'Conditional Access summary does not show risk-based controls.' -Recommendation 'Review whether sign-in and user-risk controls belong in the baseline access model.' -CurrentValue 'Risk-based Conditional Access not detected in summary' -TargetValue 'Risk-based Conditional Access reviewed and documented' -Source 'Summary/ConditionalAccess' -RelatedWorksheet 'ConditionalAccessPolicySummary' -RelatedSection 'Conditional Access'
    }
    if ($null -ne $policiesWithExclusions -and $policiesWithExclusions -ge 5) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'CA-012' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'Medium' -Finding 'Conditional Access summary shows a higher number of policies with exclusions.' -Recommendation 'Review exclusion sprawl and remove broad bypass patterns that are no longer required.' -CurrentValue "$policiesWithExclusions policy/policies with exclusions" -TargetValue 'Exclusions minimized and documented' -Source 'Summary/ConditionalAccess' -RelatedWorksheet 'ConditionalAccessPolicySummary' -RelatedSection 'Conditional Access'
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
    Add-HeuristicFinding -Store $findingStore -RuleId 'EX-006' -Area 'Exchange Hygiene' -Category 'Exchange Hygiene' -Severity 'High' -Finding 'Inbox rules with external forwarding targets were detected.' -Recommendation 'Review external inbox-rule forwarding, confirm approved use cases, and remove or document exceptions that must remain.' -CurrentValue $ex006CurrentValue -TargetValue 'All external inbox-rule forwarding paths reviewed and either approved or removed, with tenant forwarding policy aligned to the approved baseline' -Source 'Summary/InboxRules' -RelatedWorksheet 'InboxRulesExternalForwarding' -RelatedSection 'Inbox Rules' -WhyFlagged 'External forwarding rules were detected, and the tenant forwarding settings show that at least some forwarding paths may still be permitted.' -ExampleAction 'Example: export the external inbox-rule list, review the target domains against approved forwarding use cases, and compare the result to the tenant forwarding baseline.'
}

$sharedMailboxSummaryRecord = if ($sharedMailboxGovernanceSummary) { Get-ArrayaObjectValue -Object $sharedMailboxGovernanceSummary -Names @('Summary') } else { $null }
$oversizedSharedMailboxCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $sharedMailboxSummaryRecord -Names @('OversizedSharedMailboxes'))
$ownerlessSharedMailboxCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $sharedMailboxSummaryRecord -Names @('SharedMailboxesWithoutOwnerSignal'))
if (($null -ne $oversizedSharedMailboxCount -and $oversizedSharedMailboxCount -gt 0) -or ($null -ne $ownerlessSharedMailboxCount -and $ownerlessSharedMailboxCount -gt 0)) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'EX-007' -Area 'Exchange Hygiene' -Category 'Exchange Hygiene' -Severity 'Low' -Finding 'Shared mailbox governance summary shows ownership or growth gaps.' -Recommendation 'Review oversized shared mailboxes and assign accountable ownership signals for long-lived shared mailbox workloads.' -CurrentValue "$oversizedSharedMailboxCount oversized; $ownerlessSharedMailboxCount lacking owner signal" -TargetValue 'Shared mailbox ownership and growth governance documented' -Source 'Summary/SharedMailboxGovernance' -RelatedWorksheet 'SharedMailboxGovernanceSummary' -RelatedSection 'Shared Mailboxes'
}

$sharePointSharingRecord = if ($sharePointSharingSummary) { Get-ArrayaObjectValue -Object $sharePointSharingSummary -Names @('Summary') } else { $null }
$externalSharingSummaryRecord = if ($externalSharingSummary) { Get-ArrayaObjectValue -Object $externalSharingSummary -Names @('Summary') } else { $null }
$tenantSharingCapability = [string](Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('TenantSharingCapability'))
if ([string]::IsNullOrWhiteSpace($tenantSharingCapability)) {
    $tenantSharingCapability = [string](Get-ArrayaObjectValue -Object $sharePointSharingRecord -Names @('TenantSharingCapability'))
}
$defaultSharingLinkType = [string](Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('DefaultSharingLinkType'))
if ([string]::IsNullOrWhiteSpace($defaultSharingLinkType)) {
    $defaultSharingLinkType = [string](Get-ArrayaObjectValue -Object $sharePointSharingRecord -Names @('DefaultSharingLinkType'))
}
$sharingDomainRestrictionMode = [string](Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('SharingDomainRestrictionMode'))
$anonymousLinkExpirationDays = [string](Get-ArrayaObjectValue -Object $externalSharingSummaryRecord -Names @('AnonymousLinkExpirationInDays'))
$requireInvitedUserMatch = [string](Get-ArrayaObjectValue -Object $sharePointSharingRecord -Names @('RequireInvitedUserMatch'))
$preventExternalUsersFromResharing = [string](Get-ArrayaObjectValue -Object $sharePointSharingRecord -Names @('PreventExternalUsersFromResharing'))
$oneDriveSharingCapability = [string](Get-ArrayaObjectValue -Object $sharePointSharingRecord -Names @('OneDriveSharingCapability'))
$externalIdentityRestrictionsRecord = if ($externalIdentityRestrictions) { Get-ArrayaObjectValue -Object $externalIdentityRestrictions -Names @('Summary') } else { $null }
$guestAccessConfigurationRecord = if ($guestAccessConfiguration) { Get-ArrayaObjectValue -Object $guestAccessConfiguration -Names @('Summary') } else { $null }
$allowInvitesFrom = [string](Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('AllowInvitesFrom'))
if ([string]::IsNullOrWhiteSpace($allowInvitesFrom)) {
    $allowInvitesFrom = [string](Get-ArrayaObjectValue -Object $guestAccessConfigurationRecord -Names @('GuestInvitationControl'))
}
$crossTenantPartnerCount = Convert-ArrayaToNumber (Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('CrossTenantPartnerCount'))
$defaultInboundMfaTrust = [string](Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('DefaultInboundMfaTrust'))
$defaultOutboundMfaTrust = [string](Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('DefaultOutboundMfaTrust'))
$hasCrossTenantAccessPolicy = [string](Get-ArrayaObjectValue -Object $externalIdentityRestrictionsRecord -Names @('HasCrossTenantAccessPolicy'))
$guestCoverageSummary = Convert-ToArrayaBoolean (Get-ArrayaObjectValue -Object $guestAccessConfigurationRecord -Names @('ConditionalAccessGuestCoverage'))
if ($tenantSharingCapability -match 'ExternalUserAndGuestSharing|ExternalUserSharingOnly') {
    Add-HeuristicFinding -Store $findingStore -RuleId 'COL-005' -Area 'SharePoint / OneDrive Governance' -Category 'SharePoint / OneDrive Governance' -Severity 'Medium' -Finding 'SharePoint sharing is configured to allow external sharing at the tenant level.' -Recommendation 'Confirm that tenant-wide sharing settings still match the intended external collaboration model.' -CurrentValue "SharingCapability=$tenantSharingCapability; OneDriveSharingCapability=$oneDriveSharingCapability; DefaultSharingLinkType=$defaultSharingLinkType; SharingDomainRestrictionMode=$sharingDomainRestrictionMode; RequireInvitedUserMatch=$requireInvitedUserMatch; PreventExternalUsersFromResharing=$preventExternalUsersFromResharing" -TargetValue 'Tenant sharing posture documented and aligned to policy' -Source 'Summary/SharePointSharing' -RelatedWorksheet 'SharePointSharingSummary' -RelatedSection 'Sharing'
}
if ($defaultSharingLinkType -match 'AnonymousAccess') {
    Add-HeuristicFinding -Store $findingStore -RuleId 'COL-006' -Area 'SharePoint / OneDrive Governance' -Category 'SharePoint / OneDrive Governance' -Severity 'Medium' -Finding 'Anonymous links appear to remain the default sharing link type.' -Recommendation 'Review default link behavior and decide whether named-user links should be the baseline.' -CurrentValue "DefaultSharingLinkType=$defaultSharingLinkType; AnonymousLinkExpirationInDays=$anonymousLinkExpirationDays; RequireInvitedUserMatch=$requireInvitedUserMatch" -TargetValue 'Default sharing link type aligned to policy' -Source 'Summary/SharePointSharing' -RelatedWorksheet 'SharePointSharingSummary' -RelatedSection 'Sharing'
}
if ($tenantSharingCapability -match 'ExternalUserAndGuestSharing|ExternalUserSharingOnly' -and $guestCoverageSummary -eq $false) {
    Add-HeuristicFinding -Store $findingStore -RuleId 'CA-013' -Area 'Conditional Access' -Category 'Conditional Access Quality' -Severity 'High' -Finding 'External sharing is enabled, but the review does not show guest Conditional Access coverage.' -Recommendation 'Add or validate guest and external-user Conditional Access coverage for externally shared workloads.' -CurrentValue "SharingCapability=$tenantSharingCapability; ConditionalAccessGuestCoverage=$guestCoverageSummary" -TargetValue 'Guest and external-user access covered by documented Conditional Access policy' -Source 'Summary/ExternalExposure' -RelatedWorksheet 'ExternalExposureFindings' -RelatedSection 'External Exposure Review'
}
if ($externalExposureFindings.Count -gt 0) {
    $staleExternallyExposedAssets = @($externalExposureFindings | Where-Object { [string]$_.ExposureCategory -eq 'Stale externally shared content' })
    $mismatchedExternallyExposedAssets = @($externalExposureFindings | Where-Object { [string]$_.ExposureCategory -eq 'Externally sharable OneDrive with ownership mismatch' })
    $trustReviewRows = @($externalExposureFindings | Where-Object { [string]$_.ExposureCategory -eq 'Cross-tenant trust posture review' -or [string]$_.ExposureCategory -eq 'Guest invitation posture review' })

    if ($staleExternallyExposedAssets.Count -gt 0 -or $mismatchedExternallyExposedAssets.Count -gt 0) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'COL-007' -Area 'SharePoint / OneDrive Governance' -Category 'SharePoint / OneDrive Governance' -Severity 'High' -Finding 'Externally exposed SharePoint or OneDrive locations now show stale activity or ownership drift.' -Recommendation 'Review stale externally exposed sites and OneDrives, then close ownership or lifecycle gaps before access remains open by default.' -CurrentValue "$($staleExternallyExposedAssets.Count) stale externally exposed location(s); $($mismatchedExternallyExposedAssets.Count) externally exposed ownership mismatch(es)" -TargetValue 'Externally exposed stale or mismatched collaboration locations reviewed and governed' -Source 'Summary/ExternalExposure' -RelatedWorksheet 'ExternalExposureFindings' -RelatedSection 'External Exposure Review'
    }

    $guestHeavyOrDormantRows = @(
        $externalExposureFindings | Where-Object {
            [string]$_.ExposureCategory -in @('Guest-heavy dormant Team', 'Guest-heavy Team', 'Dormant guest-enabled Team', 'Guest-enabled ownerless Team', 'Guest-enabled ownerless Microsoft 365 Group')
        }
    )
    if ($guestHeavyOrDormantRows.Count -gt 0) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'TM-008' -Area 'Teams / M365 Groups Governance' -Category 'Teams / M365 Groups Governance' -Severity 'Medium' -Finding 'Externally relevant Teams or groups show guest-heavy, dormant, or ownerless patterns.' -Recommendation 'Review guest-enabled Teams and Microsoft 365 groups that are dormant, guest-heavy, or ownerless.' -CurrentValue "$($guestHeavyOrDormantRows.Count) externally relevant Team/group exposure row(s)" -TargetValue 'Guest-enabled Teams and groups reviewed, owned, and governed' -Source 'Summary/ExternalExposure' -RelatedWorksheet 'ExternalExposureFindings' -RelatedSection 'External Exposure Review'
    }

    if ($trustReviewRows.Count -gt 0) {
        Add-HeuristicFinding -Store $findingStore -RuleId 'ID-008' -Area 'Identity Governance' -Category 'Identity Governance' -Severity 'Medium' -Finding 'External invitation or cross-tenant trust posture should be reviewed against the intended external-access baseline.' -Recommendation 'Review guest invitation controls and cross-tenant trust settings against the approved external-access baseline.' -CurrentValue "$($trustReviewRows.Count) external identity / trust review row(s); AllowInvitesFrom=$allowInvitesFrom; DefaultInboundMfaTrust=$defaultInboundMfaTrust; DefaultOutboundMfaTrust=$defaultOutboundMfaTrust" -TargetValue 'External invitation and cross-tenant trust posture documented and aligned to policy' -Source 'Summary/ExternalExposure' -RelatedWorksheet 'ExternalExposureFindings' -RelatedSection 'External Exposure Review'
    }
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
$resolvedWorkstreamSummaryRows = [object[]]@()
if ($null -ne $workstreamSummaries) {
    if ($workstreamSummaries.PSObject.Methods.Name -contains 'ToArray') {
        $resolvedWorkstreamSummaryRows = @($workstreamSummaries.ToArray())
    }
    else {
        $resolvedWorkstreamSummaryRows = @($workstreamSummaries)
    }
}

$resolvedFindingRows = [object[]]@()
if ($null -ne $sortedFindings) {
    if ($sortedFindings.PSObject.Methods.Name -contains 'ToArray') {
        $resolvedFindingRows = @($sortedFindings.ToArray())
    }
    else {
        $resolvedFindingRows = @($sortedFindings)
    }
}

$visibleWorkstreamSummaries = [object[]]@(Resolve-VisibleWorkstreamSummaries -Summaries $resolvedWorkstreamSummaryRows -Findings $resolvedFindingRows)
$sortedWorkstreamSummaries = Get-SortedWorkstreamSummaries -Summaries @($visibleWorkstreamSummaries)
$tenantName = Get-TenantDisplayName -TenantInfoSummary $(if ($tenantInfoSummary) { Get-ArrayaObjectValue -Object $tenantInfoSummary -Names @('Summary') } else { $null }) -LegacyData $tenantData -OutputPrefix $OutputPrefix
$customerAssessmentSignals = [pscustomobject]@{
    Admins                     = $adminRows
    ConditionalAccessPolicies  = $caPolicies
    ConditionalAccessSummary   = $conditionalAccessSummary
    AuthenticationConfig       = $authConfig
    MfaRegistrationSummary     = $mfaSummary
    MfaRegistrationDetails     = $mfaRegistrationDetails
    MfaEnrollmentSummary       = $mfaEnrollmentSummary
    MfaEnforcementSummary      = $mfaEnforcementSummary
    MfaEnforcementGapUsers     = $mfaEnforcementGapUsers
    MfaEnforcementScopeReview  = $mfaEnforcementScopeReview
    AdminMfaSummary            = $adminMfaSummary
    AdminMfaRegistrationGaps   = $adminMfaRegistrationGaps
    AdminMfaEnforcementGaps    = $adminMfaEnforcementGaps
    EnterpriseApplications     = $enterpriseApplications
    AuthenticationSSOApplications = $authenticationSsoApplications
    EnterpriseApplicationSummary = $enterpriseApplicationSummary
    GuestSignInSummary         = $guestSignInSummary
    PrivilegedAccessSummary    = $privilegedAccessSummary
    DeviceDetails              = $deviceRows
    DeviceManagementSummary    = $deviceManagementSummary
    AllRecipients              = $recipientRows
    AllMailboxes               = $mailboxRows
    PrimaryMailboxStats        = $primaryMailboxStatsRows
    ArchiveMailboxes           = $archiveMailboxRows
    ArchiveMailboxStats        = $archiveMailboxStatsRows
    InactiveMailboxes          = $inactiveMailboxRows
    NonUserMailboxes           = $nonUserMailboxRows
    EmailActivitySummary       = $emailActivitySummary
    EmailActivityTopSenders    = $emailActivityTopSenders
    EmailActivityTopReceivers  = $emailActivityTopReceivers
    InboxRulesExternalForwarding = $inboxRulesExternalForwarding
    InboxRuleForwardingSummary = $inboxRuleForwardingSummary
    ForwardingPolicySummary    = $forwardingPolicySummary
    MailFlowConnectors         = $connectorRows
    RemoteDomains              = $remoteDomainRows
    SharedMailboxGovernanceSummary = $sharedMailboxGovernanceSummary
    PublicFolderDetails        = $publicFolderRows
    SharePoint                 = $sharePointRows
    OneDrive                   = $oneDriveRows
    AllTeams                   = $teamRows
    TeamsVoiceSummary          = $teamsVoiceSummary
    UnifiedGroups              = $unifiedGroupRows
    SharePointSharingSummary   = $sharePointSharingSummary
    ExternalSharingSummary     = $externalSharingSummary
    ExternalSharingSiteOverrides = $externalSharingSiteOverrides
    ExternalExposureFindings   = $externalExposureFindings
    CollaborationActivitySummary = $collaborationActivitySummary
    OneDriveOwnerMismatches    = $oneDriveOwnerMismatches
    Domains                    = $domainRows
    LicenseSKUs                = $licenseRows
    TenantInfoSummary          = $tenantInfoSummary
    SecuritySecureScore        = $secureScoreRows
    SMTPRelaySummary           = $smtpRelaySummary
    SMTPRelayConfig            = $smtpRelayConfig
    SMTPRelayServiceAccounts   = $smtpRelayServiceAccounts
    SpamFilteringSummary       = $spamFilteringSummary
    AuthenticationConfigSummary = $authConfigSummary
    ExternalIdentityRestrictions = $externalIdentityRestrictions
    GuestAccessConfiguration   = $guestAccessConfiguration
    RetentionPolicies          = $retentionPolicyRows
    DlpPolicies                = $dlpPolicyRows
    PasswordLifecycleSummary   = $passwordLifecycleSummary
    AdConnectConfiguration     = $adConnectConfiguration
    SecurityDefaultsPolicy     = $securityDefaultsPolicy
    Users                      = $userRows
}
$technicalObservations = Get-CustomerTechnicalObservations -Signals $customerAssessmentSignals -Findings $sortedFindings -WorkstreamSummaries $sortedWorkstreamSummaries
$consultativeSummaries = Get-CustomerConsultativeSummaries -Signals $customerAssessmentSignals -Findings $sortedFindings -ExecutiveThemes (Get-CustomerExecutiveThemes -Findings $sortedFindings -Count 5) -RoadmapActions (Get-CustomerRoadmapActions -Findings $sortedFindings -MaxPerPhase 5) -OwnerGroups @($sortedFindings | Group-Object OwnerTeam | Sort-Object Count -Descending)

$supportFolder = Join-Path -Path $OutputFolder -ChildPath 'Support'
if (-not (Test-Path -Path $supportFolder)) {
    $null = New-Item -ItemType Directory -Path $supportFolder -Force
}

$artifactPrefix = Get-TenantArtifactFilePrefix -TenantName $tenantName -Fallback $OutputPrefix
$customerReportDeliveryDate = $generatedAt.ToString('yyyy-MM-dd')
$jsonOutPath = Join-Path -Path $supportFolder -ChildPath ("{0}-Plan.json" -f $artifactPrefix)
$csvOutPath = if ($IncludeLegacyArtifacts) { Join-Path -Path $supportFolder -ChildPath ("{0}-Plan.csv" -f $artifactPrefix) } else { $null }
$mdOutPath = if ($IncludeLegacyArtifacts) { Join-Path -Path $supportFolder -ChildPath ("{0}-Plan.md" -f $artifactPrefix) } else { $null }
$customerAssessmentReportOutPath = Join-Path -Path $supportFolder -ChildPath ("{0}-Microsoft 365 Tenant Best Practices Assessment-{1}.docx" -f $artifactPrefix, $customerReportDeliveryDate)
$customerAssessmentReportMarkdownOutPath = Join-Path -Path $supportFolder -ChildPath ("{0}-CustRpt.md" -f $artifactPrefix)
$roadmapRemediationPlanOutPath = Join-Path -Path $supportFolder -ChildPath ("{0}-Microsoft 365 Remediation Roadmap-{1}.docx" -f $artifactPrefix, $customerReportDeliveryDate)
$engineerMdOutPath = Join-Path -Path $supportFolder -ChildPath ("{0}-EngPack.md" -f $artifactPrefix)
$snippetOutPath = Join-Path -Path $supportFolder -ChildPath ("{0}-Snips.ps1" -f $artifactPrefix)

$deliverables = [ordered]@{
    ImprovementPlanJson       = $jsonOutPath
    CustomerAssessmentReport  = $customerAssessmentReportOutPath
    CustomerAssessmentReportMarkdown = $customerAssessmentReportMarkdownOutPath
    RoadmapRemediationPlan    = $roadmapRemediationPlanOutPath
    EngineerActionPack        = $engineerMdOutPath
    RemediationSnippets       = $snippetOutPath
    SupportFolder             = $supportFolder
}
if ($IncludeLegacyArtifacts) {
    $deliverables['ImprovementPlanCsv'] = $csvOutPath
    $deliverables['ImprovementPlanMarkdown'] = $mdOutPath
}

$payload = [PSCustomObject]@{
    GeneratedAt = $generatedAt.ToString('o')
    AssessmentVersion = $assessmentVersionLabel
    SourceFile  = (Resolve-Path -Path $AssessmentJsonPath).Path
    Thresholds  = [PSCustomObject]@{
        StaleDeviceDays = $StaleDeviceDays
        MaxGlobalAdmins = $MaxGlobalAdmins
    }
    Deliverables = [PSCustomObject]$deliverables
    WorkstreamSummaries = $sortedWorkstreamSummaries
    Findings           = $sortedFindings
    TechnicalObservations = $technicalObservations
    ConsultativeSummaries = $consultativeSummaries
    ExternalExposureFindings = $externalExposureFindings
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

$customerSourceModel = New-CustomerReportSourceModel -TenantName $tenantName -AssessmentJsonPath $AssessmentJsonPath -SourceInputPath $AssessmentJsonPath -SourceType 'Snapshot' -SourceLabel ([System.IO.Path]::GetFileNameWithoutExtension($AssessmentJsonPath)) -Findings $sortedFindings -WorkstreamSummaries $sortedWorkstreamSummaries -TechnicalObservations $technicalObservations -ConsultativeSummaries $consultativeSummaries -AssessmentVersion $assessmentVersionLabel
$customerAssessmentTemplatePath = Get-CustomerAssessmentTemplatePath
$customerAssessmentBlocks = New-CustomerAssessmentDocumentBlocks -SourceModel $customerSourceModel -Signals $customerAssessmentSignals -GeneratedAt $generatedAt
Write-CustomerAssessmentDocxFromModel -TemplatePath $customerAssessmentTemplatePath -OutputPath $customerAssessmentReportOutPath -TenantName $tenantName -GeneratedAt $generatedAt -Blocks $customerAssessmentBlocks
Write-CustomerAssessmentMarkdownFromBlocks -Blocks $customerAssessmentBlocks -OutputPath $customerAssessmentReportMarkdownOutPath
$roadmapRemediationTemplatePath = Get-RoadmapRemediationTemplatePath
$roadmapRemediationBlocks = New-RoadmapRemediationDocumentBlocks -SourceModel $customerSourceModel -Signals $customerAssessmentSignals -GeneratedAt $generatedAt
Write-CustomerAssessmentDocxFromModel -TemplatePath $roadmapRemediationTemplatePath -OutputPath $roadmapRemediationPlanOutPath -TenantName $tenantName -GeneratedAt $generatedAt -Blocks $roadmapRemediationBlocks -DocumentTitle "$tenantName Microsoft 365 Remediation Roadmap"

$engineerActionPackMarkdown = New-EngineerActionPack -Findings $sortedFindings -WorkstreamSummaries $sortedWorkstreamSummaries -TenantName $tenantName -AssessmentJsonPath $AssessmentJsonPath -GeneratedAt $generatedAt -JsonOutPath $jsonOutPath -CsvOutPath $csvOutPath -SnippetOutPath $snippetOutPath -SupportFolderPath $supportFolder
Set-Content -Path $engineerMdOutPath -Value $engineerActionPackMarkdown -Encoding UTF8

$snippetLibrary = @{
    'SEC-001' = @'
# Secure Score review
Connect-MgGraph -Scopes 'SecurityEvents.Read.All'
Get-MgSecuritySecureScore -Top 1 | Format-List *
'@
    'CA-001' = @'
# Conditional Access baseline validation
Connect-MgGraph -Scopes 'Policy.Read.All'
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
    'MFA-002' = @'
# MFA method quality review
Connect-MgGraph -Scopes 'Reports.Read.All'
Get-MgReportAuthenticationMethodUserRegistrationDetail -All |
    Select-Object UserPrincipalName, DefaultMfaMethod, MethodsRegistered
'@
    'MFA-003' = @'
# MFA enforcement review
Connect-MgGraph -Scopes 'Policy.Read.All'
Get-MgIdentityConditionalAccessPolicy |
    Select-Object DisplayName, State, @{N='GrantControls';E={$_.GrantControls.BuiltInControls -join ','}}
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
    CustomerAssessmentReportPath  = $customerAssessmentReportOutPath
    CustomerAssessmentReportMarkdownPath = $customerAssessmentReportMarkdownOutPath
    RoadmapRemediationPlanPath    = $roadmapRemediationPlanOutPath
    CustomerRemediationReportPath = $null
    CustomerRemediationReportMarkdownPath = $null
    EngineerActionPackPath        = $engineerMdOutPath
    RemediationPs1Path            = $snippetOutPath
}

if (-not $Quiet) {
    Write-Host 'Improvement plan generated.'
    Write-Host "  Customer report  : $customerAssessmentReportOutPath" -ForegroundColor Green
    Write-Host "  Customer markdown: $customerAssessmentReportMarkdownOutPath" -ForegroundColor DarkGreen
    Write-Host "  Roadmap report   : $roadmapRemediationPlanOutPath" -ForegroundColor Yellow
    Write-Host "  Engineer pack    : $engineerMdOutPath" -ForegroundColor Cyan
    Write-Host "  Support folder   : $supportFolder" -ForegroundColor DarkGray
}

if ($PassThru) {
    $outputSummary
}
