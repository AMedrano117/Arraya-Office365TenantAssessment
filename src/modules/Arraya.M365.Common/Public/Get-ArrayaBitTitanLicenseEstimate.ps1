function Get-ArrayaBitTitanLicenseEstimate {
    <#
    .SYNOPSIS
        Estimates BitTitan mailbox migration licensing for a tenant.

    .DESCRIPTION
        BitTitan is used for mailbox data only. SharePoint, OneDrive, and Teams content are
        migrated with ShareGate and are deliberately outside this function - it accepts no
        site or drive rows, so the mailbox-only boundary is enforced by the signature.

        The function is pure: no live service calls, no script-scoped state. Everything it
        needs arrives as parameters, which is what makes the sizing math unit-testable.

        Two vendor-valid unit-count mixes are always returned so a Solutions Engineer can choose after
        applying the commercial rates available to them:
          - Workload fit: Mailbox licenses in BlockSizeGB increments for mailbox-only objects;
                          one User Migration Bundle for each object with archive content.
          - UMB-led:      one User Migration Bundle per licensable user or archive-bearing
                          object; mailbox-only non-user objects remain mailbox-license units.

        Microsoft 365 Group mailbox conversations remain visible but are report-only by
        default, matching the delivery practice of not automatically licensing them. The
        caller can explicitly include their mailbox-license units for an exception scenario.

        Tenant Migration Bundle is surfaced as an optional commercial scenario only. It is
        not counted automatically because this mailbox-only function does not receive the
        Teams and SharePoint library scope needed to size its Flex Collaboration component,
        and the normal delivery practice uses ShareGate for those workloads.

        SKU names, entitlements, and thresholds live in the model file rather than in code
        so they can be updated when BitTitan changes its price list.

    .PARAMETER MailboxRows
        Mailbox planning rows: active plus de-duplicated inactive mailboxes.

    .PARAMETER GroupMailboxRows
        Microsoft 365 Group mailbox rows. These are not returned by Get-EXOMailbox without
        -GroupMailbox, so they arrive from their own collection pass.

    .PARAMETER UserRows
        Directory user rows, used to establish which mailboxes have a licensable user behind
        them for the bundle path.

    .PARAMETER Model
        Parsed bittitan-license-model.json. Defaults to the copy shipped in
        src/config/baseline.

    .PARAMETER IncludeGroupMailboxLicenseUnits
        Overrides the normal practice policy and includes Microsoft 365 Group mailboxes with
        data in the mailbox-license unit paths. Without this switch, group mailboxes remain in
        ObjectLines with their discovered sizes but contribute no automatic license units.

    .OUTPUTS
        PSCustomObject with ObjectLines, SkuTotals, PathComparison, and Drivers collections.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$MailboxRows,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$GroupMailboxRows,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$UserRows,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $Model,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeGroupMailboxLicenseUnits
    )

    if ($null -eq $Model) {
        $Model = Get-ArrayaBitTitanLicenseModel
    }

    $blockSizeGB = [double](Get-ArrayaObjectValue -Object $Model.Skus.MigrationWizMailbox -Names @('BlockSizeGB'))
    if ($blockSizeGB -le 0) { $blockSizeGB = 50 }

    $mailboxSkuName = [string](Get-ArrayaObjectValue -Object $Model.Skus.MigrationWizMailbox -Names @('Name'))
    $bundleSkuName = [string](Get-ArrayaObjectValue -Object $Model.Skus.UserMigrationBundle -Names @('Name'))
    $emptyGroupFloorGB = [double](Get-ArrayaObjectValue -Object $Model.Thresholds -Names @('EmptyGroupMailboxFloorGB'))
    $largeMailboxGB = [double](Get-ArrayaObjectValue -Object $Model.Thresholds -Names @('LargeMailboxGB'))
    if ($largeMailboxGB -le 0) { $largeMailboxGB = 50 }

    $bundleEligibleTypes = @(Get-ArrayaObjectValue -Object $Model.BundleEligibility -Names @('RecipientTypeDetails'))
    if ($bundleEligibleTypes.Count -eq 0) { $bundleEligibleTypes = @('UserMailbox') }

    # Users with a mailbox are the bundle denominator. Index every identifier a mailbox row
    # might join on, because UPN and primary SMTP diverge often enough to matter.
    $licensableUserKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($userRow in @($UserRows)) {
        if ($null -eq $userRow) { continue }
        foreach ($identifier in @(
                (Get-ArrayaObjectValue -Object $userRow -Names @('UserPrincipalName')),
                (Get-ArrayaObjectValue -Object $userRow -Names @('Mail')),
                (Get-ArrayaObjectValue -Object $userRow -Names @('MailboxPrimarySmtpAddress'))
            )) {
            if (-not [string]::IsNullOrWhiteSpace([string]$identifier)) {
                [void]$licensableUserKeys.Add(([string]$identifier).Trim())
            }
        }
    }

    $objectLines = New-Object System.Collections.Generic.List[object]

    foreach ($mailboxRow in @($MailboxRows)) {
        if ($null -eq $mailboxRow) { continue }
        $objectLines.Add((New-ArrayaBitTitanObjectLine `
            -Record $mailboxRow `
            -ObjectType 'Mailbox' `
            -BlockSizeGB $blockSizeGB `
            -SkuName $mailboxSkuName `
            -BundleSkuName $bundleSkuName `
            -EmptyFloorGB $null `
            -BundleEligibleTypes $bundleEligibleTypes `
            -LicensableUserKeys $licensableUserKeys)) | Out-Null
    }

    foreach ($groupMailboxRow in @($GroupMailboxRows)) {
        if ($null -eq $groupMailboxRow) { continue }
        $objectLines.Add((New-ArrayaBitTitanObjectLine `
            -Record $groupMailboxRow `
            -ObjectType 'GroupMailbox' `
            -BlockSizeGB $blockSizeGB `
            -SkuName $mailboxSkuName `
            -BundleSkuName $bundleSkuName `
            -EmptyFloorGB $emptyGroupFloorGB `
            -BundleEligibleTypes $bundleEligibleTypes `
            -ReportOnly:(-not $IncludeGroupMailboxLicenseUnits) `
            -LicensableUserKeys $licensableUserKeys)) | Out-Null
    }

    $allLines = @($objectLines.ToArray())
    $billableLines = @($allLines | Where-Object { $null -ne $_.LicenseUnits })
    $needsDataLines = @($allLines | Where-Object { $_.RecommendedSku -eq 'Needs data' })
    $emptyShellLines = @($allLines | Where-Object { $_.RecommendedSku -eq 'No migration' })
    $groupMailboxDataLines = @($allLines | Where-Object {
            $_.ObjectType -eq 'GroupMailbox' -and
            $null -ne $_.SizeGB -and
            [double]$_.SizeGB -gt $emptyGroupFloorGB
        })
    $reportOnlyGroupLines = @($groupMailboxDataLines | Where-Object { $_.EstimateStatus -eq 'Report only - manual decision' })

    # Canonical practice mix used by proposal-facing output. This is deliberately a
    # recipient-category view rather than a competing commercial-path comparison: mailbox-
    # only objects consume Mailbox Migration license blocks, archive-bearing objects consume
    # one UMB planning unit, and Microsoft 365 Group conversations remain report-only unless
    # the customer explicitly elects to migrate them.
    $licenseMixCategoryOrder = @(
        'Office 365 Groups (report-only / optional)'
        'Resources'
        'Shared Mailboxes'
        'Inactive'
        'Inactive + archive'
        'User Mailboxes'
        'User Mailboxes + archive'
    )
    $licenseMixRowsByCategory = @{}
    foreach ($category in $licenseMixCategoryOrder) {
        $licenseMixRowsByCategory[$category] = New-Object System.Collections.Generic.List[object]
    }

    foreach ($line in @($allLines)) {
        $recipientType = [string]$line.RecipientType
        $category = if ([string]$line.ObjectType -eq 'GroupMailbox' -or $recipientType -eq 'GroupMailbox') {
            'Office 365 Groups (report-only / optional)'
        } elseif ($line.IsInactiveMailbox -eq $true -and $line.HasArchive -eq $true) {
            'Inactive + archive'
        } elseif ($line.IsInactiveMailbox -eq $true) {
            'Inactive'
        } elseif ($recipientType -match '(?i)^(room|equipment|scheduling|resource)') {
            'Resources'
        } elseif ($recipientType -match '(?i)^shared') {
            'Shared Mailboxes'
        } elseif ($recipientType -match '(?i)^user' -and $line.HasArchive -eq $true) {
            'User Mailboxes + archive'
        } elseif ($recipientType -match '(?i)^user') {
            'User Mailboxes'
        } else {
            # Keep uncommon service/non-user mailbox types visible in the fixed category
            # contract instead of dropping them from the canonical totals.
            'Resources'
        }
        $licenseMixRowsByCategory[$category].Add($line) | Out-Null
    }

    $licenseMixBreakdown = @(
        foreach ($category in $licenseMixCategoryOrder) {
            $categoryLines = @($licenseMixRowsByCategory[$category].ToArray())
            $mailboxLicenseLines = @($categoryLines | Where-Object { $null -ne $_.LicenseUnits -and $_.HasArchive -ne $true })
            $umbLines = @($categoryLines | Where-Object { $null -ne $_.LicenseUnits -and $_.HasArchive -eq $true })
            $needsDataCategoryLines = @($categoryLines | Where-Object {
                    [string]$_.RecommendedSku -match '(?i)needs data|needs sizing'
                })
            $optionalCategoryLines = @($categoryLines | Where-Object {
                    [string]$_.ObjectType -eq 'GroupMailbox' -and [string]$_.RecommendedSku -eq 'Report only'
                })
            $categoryDataGB = [math]::Round((@($categoryLines | ForEach-Object {
                            if ($null -ne $_.SizeGB) { [double]$_.SizeGB } else { 0 }
                        } | Measure-Object -Sum).Sum), 3)

            $treatment = switch ($category) {
                'Office 365 Groups (report-only / optional)' { 'Report only by default; migrate conversations only by customer decision.' }
                'Resources' { 'Mailbox Migration when mailbox-only; archive content requires a UMB planning unit and eligibility validation.' }
                'Shared Mailboxes' { 'Mailbox Migration when mailbox-only; archive content requires a UMB planning unit and eligibility validation.' }
                'Inactive' { 'Restore or activate, then migrate with Mailbox Migration license blocks.' }
                'Inactive + archive' { 'Restore or activate, then migrate with one UMB planning unit; validate eligibility.' }
                'User Mailboxes' { 'Mailbox Migration license blocks for mailbox-only scope.' }
                'User Mailboxes + archive' { 'One User Migration Bundle planning unit per archive-bearing user.' }
            }

            $notes = if ($category -eq 'Office 365 Groups (report-only / optional)') {
                ("{0} optional conversation migration object(s), {1} measured empty object(s), and {2} object(s) needing mailbox evidence. No automatic group-mailbox license units are included." -f
                    $optionalCategoryLines.Count,
                    @($categoryLines | Where-Object { [string]$_.RecommendedSku -eq 'No migration' }).Count,
                    $needsDataCategoryLines.Count)
            } else {
                'Mailbox Migration and UMB values are license-unit counts. Needs Data objects are excluded until sizing and archive evidence are complete.'
            }

            [pscustomobject]@{
                Category                         = $category
                ObjectCount                      = $categoryLines.Count
                MailboxMigrationLicenseUnits     = [int](@($mailboxLicenseLines | ForEach-Object { [int]$_.LicenseUnits } | Measure-Object -Sum).Sum)
                UserMigrationBundleUnits         = [int](@($umbLines | ForEach-Object { [int]$_.LicenseUnits } | Measure-Object -Sum).Sum)
                ReportOnlyObjectCount             = $(if ($category -eq 'Office 365 Groups (report-only / optional)') { $categoryLines.Count } else { 0 })
                OptionalMigrationObjectCount      = $optionalCategoryLines.Count
                NeedsDataObjectCount              = $needsDataCategoryLines.Count
                DataGB                            = $categoryDataGB
                UmbEligibilityValidationCount     = @($categoryLines | Where-Object { $_.UmbEligibilityValidationRequired -eq $true }).Count
                Treatment                         = $treatment
                Notes                             = $notes
            }
        }
    )

    # Workload-fit mix: mailbox licenses are valid only when no archive is migrated. An
    # archive-bearing object needs UMB; it cannot be represented by stacked mailbox units.
    $workloadMailboxLines = @($billableLines | Where-Object { $_.HasArchive -ne $true })
    $workloadUmbLines = @($billableLines | Where-Object { $_.HasArchive -eq $true })
    $workloadMailboxUnits = [int](@($workloadMailboxLines | ForEach-Object { [int]$_.LicenseUnits } | Measure-Object -Sum).Sum)
    $workloadUmbCount = $workloadUmbLines.Count

    # UMB-led mix: UMB covers every archive-bearing object plus mailbox-only user mailboxes
    # with a matching directory identity. Remaining non-user mailbox-only objects use mailbox
    # license blocks. Distinct object rows prevent archive mailboxes from being counted twice.
    $bundleLines = @($billableLines | Where-Object { $_.HasArchive -eq $true -or $_.BundleEligible -eq $true })
    $bundleResidualLines = @($billableLines | Where-Object { $_.HasArchive -ne $true -and $_.BundleEligible -ne $true })
    $bundleCount = $bundleLines.Count
    $tmbEligibleUserCount = @($bundleLines | Where-Object { $_.BundleEligible -eq $true }).Count
    $bundleResidualUnits = [int](@($bundleResidualLines | ForEach-Object { [int]$_.LicenseUnits } | Measure-Object -Sum).Sum)

    $skuTotals = @(
        [pscustomobject]@{
            Path         = 'Workload fit'
            Sku          = $mailboxSkuName
            Units        = $workloadMailboxUnits
            ObjectCount  = $workloadMailboxLines.Count
            Notes        = $(if ($IncludeGroupMailboxLicenseUnits) {
                    ("Mailbox licenses in {0} GB blocks across mailboxes, including Microsoft 365 Group mailbox exceptions." -f $blockSizeGB)
                } else {
                    ("Mailbox licenses in {0} GB blocks. Microsoft 365 Group mailboxes are report-only and excluded from these units by practice policy." -f $blockSizeGB)
                })
        },
        [pscustomobject]@{
            Path         = 'Workload fit'
            Sku          = $bundleSkuName
            Units        = $workloadUmbCount
            ObjectCount  = $workloadUmbLines.Count
            Notes        = 'One UMB planning unit for each archive-bearing object. Archive content is not eligible for a Mailbox Migration license; validate non-user and restored inactive UMB eligibility.'
        },
        [pscustomobject]@{
            Path         = 'UMB-led'
            Sku          = $bundleSkuName
            Units        = $bundleCount
            ObjectCount  = $bundleCount
            Notes        = 'One UMB planning unit per licensable user or archive-bearing object. Validate non-user and restored inactive eligibility. OneDrive and document entitlements go unused when ShareGate handles file workloads.'
        },
        [pscustomobject]@{
            Path         = 'UMB-led'
            Sku          = $mailboxSkuName
            Units        = $bundleResidualUnits
            ObjectCount  = $bundleResidualLines.Count
            Notes        = $(if ($IncludeGroupMailboxLicenseUnits) {
                    'Shared, room, equipment, and explicitly included Microsoft 365 Group mailboxes stay a-la-carte on the User Migration Bundle path.'
                } else {
                    'Shared, room, and equipment mailboxes stay a-la-carte on the User Migration Bundle path. Microsoft 365 Group mailboxes are report-only.'
                })
        }
    )

    $totalDataGB = [math]::Round((@($allLines | ForEach-Object { [double](Convert-ArrayaToNumber -Value $_.SizeGB) } | Measure-Object -Sum).Sum), 3)
    $archiveEnabledCount = @($allLines | Where-Object { $_.HasActiveArchive -eq $true }).Count
    $overThresholdCount = @($allLines | Where-Object { $null -ne $_.SizeGB -and [double]$_.SizeGB -gt $largeMailboxGB }).Count
    $sizedLines = @($allLines | Where-Object { $null -ne $_.SizeGB })
    $averageSizeGB = if ($sizedLines.Count -gt 0) {
        [math]::Round((@($sizedLines | ForEach-Object { [double]$_.SizeGB } | Measure-Object -Average).Average), 3)
    } else { $null }
    $largestLine = @($sizedLines | Sort-Object { [double]$_.SizeGB } -Descending | Select-Object -First 1) | Select-Object -First 1
    $archivePercent = if ($allLines.Count -gt 0) { [math]::Round((($archiveEnabledCount / $allLines.Count) * 100), 1) } else { 0 }

    # This assessment deliberately has no price inputs. Unit counts are evidence for the SE,
    # not a cost comparison, so selecting a "cheaper" path here would be false precision.
    $recommendedPath = 'SE selection required'
    $recommendationReason = 'No prices or customer-specific commercial terms are emitted. Compare the unit baskets using the rates available to the Solutions Engineer.'

    $pathComparison = @(
        [pscustomobject]@{
            Path             = 'Workload fit'
            PrimarySku       = $mailboxSkuName
            PrimaryUnits     = $workloadMailboxUnits
            SecondarySku     = $bundleSkuName
            SecondaryUnits   = $workloadUmbCount
            TotalUnits       = ($workloadMailboxUnits + $workloadUmbCount)
            IsRecommended    = $false
            SelectionStatus  = 'SE commercial comparison required'
            Notes            = ("Mailbox-only objects use {0} GB Mailbox license blocks; archive-bearing objects use UMB. Group mailbox exceptions are {1}." -f $blockSizeGB, $(if ($IncludeGroupMailboxLicenseUnits) { 'included' } else { 'report-only' }))
        },
        [pscustomobject]@{
            Path             = 'UMB-led'
            PrimarySku       = $bundleSkuName
            PrimaryUnits     = $bundleCount
            SecondarySku     = $mailboxSkuName
            SecondaryUnits   = $bundleResidualUnits
            TotalUnits       = ($bundleCount + $bundleResidualUnits)
            IsRecommended    = $false
            SelectionStatus  = 'SE commercial comparison required'
            Notes            = 'UMB covers licensable users and every archive-bearing object; mailbox-only non-user objects use Mailbox license blocks. OneDrive and document entitlements go unused because ShareGate handles file workloads.'
        }
    )

    $tenantBundleModel = Get-ArrayaObjectValue -Object (Get-ArrayaObjectValue -Object $Model -Names @('OptionalCommercialOptions')) -Names @('TenantMigrationBundle')
    $tenantBundleName = [string](Get-ArrayaObjectValue -Object $tenantBundleModel -Names @('Name'))
    if ([string]::IsNullOrWhiteSpace($tenantBundleName)) { $tenantBundleName = 'Tenant Migration Bundle' }
    $tenantBundleCoverage = @((Get-ArrayaObjectValue -Object $tenantBundleModel -Names @('Includes')) | ForEach-Object { [string]$_ }) -join '; '
    if ([string]::IsNullOrWhiteSpace($tenantBundleCoverage)) {
        $tenantBundleCoverage = 'One User Migration Bundle plus one Flex Collaboration License'
    }
    $tenantBundleExclusions = @((Get-ArrayaObjectValue -Object $tenantBundleModel -Names @('Excludes')) | ForEach-Object { [string]$_ }) -join '; '
    if ([string]::IsNullOrWhiteSpace($tenantBundleExclusions)) { $tenantBundleExclusions = 'Public folders' }
    $tenantBundlePrivateChatGuidance = [string](Get-ArrayaObjectValue -Object $tenantBundleModel -Names @('PrivateChatGuidance'))
    if ([string]::IsNullOrWhiteSpace($tenantBundlePrivateChatGuidance)) {
        $tenantBundlePrivateChatGuidance = 'Teams Private Chat (PCH) uses a separate Collaboration (Private Chats) project. Validate current capability, limitations, and licensing entitlement; do not assume TMB/UMB inclusion or exclusion.'
    }
    $tenantBundlePrivateChatGuideUrl = [string](Get-ArrayaObjectValue -Object $tenantBundleModel -Names @('PrivateChatGuideUrl'))
    if ([string]::IsNullOrWhiteSpace($tenantBundlePrivateChatGuideUrl)) {
        $tenantBundlePrivateChatGuideUrl = 'https://help.bittitan.com/hc/en-us/articles/25603590557979-Teams-Private-Chat-Migration-Guide'
    }
    $tenantBundleGettingStartedUrl = [string](Get-ArrayaObjectValue -Object $tenantBundleModel -Names @('GettingStartedUrl'))
    if ([string]::IsNullOrWhiteSpace($tenantBundleGettingStartedUrl)) {
        $tenantBundleGettingStartedUrl = 'https://help.bittitan.com/hc/en-us/articles/1260806785629-Getting-Started-with-Migrations'
    }
    $tenantBundleGuidanceParagraph = ("Tenant Migration Bundle is optional and is not Arraya's default. Do not add its {0} candidate unit(s) to the practical mix above ({1} Mailbox Migration license unit(s) plus {2} User Migration Bundle unit(s)); TMB is an alternate packaging option to review only when BitTitan, instead of ShareGate, will also migrate qualifying collaboration content. Each TMB pairs one User Migration Bundle with one Flex Collaboration License. One Flex entitlement covers one Team or one SharePoint library up to 100 GB, so exact collaboration demand requires Team-to-library mapping and per-object sizing; site count alone is not enough. ShareGate normally carries SharePoint, OneDrive, and Teams, which normally leaves the included Flex entitlement unused. Public folders are excluded and require separate planning. Teams Private Chat (PCH) uses a separate Collaboration (Private Chats) project, and BitTitan's official guidance is changing and currently inconsistent about PCH entitlement; validate current capability, limitations, and licensing in the PCH migration guide and Getting Started documentation when scoping and again before ordering. This assessment includes no prices or automatic cost-based recommendation." -f $tmbEligibleUserCount, $workloadMailboxUnits, $workloadUmbCount)

    $planningOptions = @(
        [pscustomobject]@{
            Option              = $tenantBundleName
            DefaultPractice     = 'Not used by default'
            PlanningUnits       = $tmbEligibleUserCount
            Eligibility         = 'Microsoft 365 tenant-to-tenant migration where BitTitan will also migrate Teams or SharePoint libraries.'
            WorkloadCoverage    = $tenantBundleCoverage
            UnitBasis           = ("{0} TMB unit(s) track the UMB-eligible users surfaced by this assessment and include the same number of Flex Collaboration entitlements. Collaboration demand still needs Team/library mapping and per-object 100 GB sizing." -f $tmbEligibleUserCount)
            WhenToReview        = 'Review only when BitTitan, rather than ShareGate, will carry collaboration workloads and the included Flex units can be used.'
            WhyNormallyExcluded = 'ShareGate owns SharePoint, OneDrive, and Teams in the normal delivery model, so the included Flex Collaboration entitlement is normally unused.'
            Exclusions          = $tenantBundleExclusions
            PrivateChatValidation = $tenantBundlePrivateChatGuidance
            PrivateChatGuideUrl = $tenantBundlePrivateChatGuideUrl
            GettingStartedUrl   = $tenantBundleGettingStartedUrl
            GuidanceParagraph   = $tenantBundleGuidanceParagraph
            Decision            = 'Optional SE commercial and tooling decision. Extra Flex demand is Needs Data because collected site counts are not document-library counts. Archive-bearing non-user objects and restored inactive identities need separate UMB/TMB eligibility validation. Teams PCH capability and licensing must be revalidated against current BitTitan terms. No automatic recommendation or dollar cost is produced.'
        }
    )

    $inactiveLines = @($allLines | Where-Object { $_.IsInactiveMailbox -eq $true })
    $inactiveArchiveLines = @($inactiveLines | Where-Object { $_.HasArchive -eq $true })
    $inactiveArchiveDataGB = [math]::Round((@($inactiveArchiveLines | ForEach-Object {
                    [double](Convert-ArrayaToNumber -Value $_.ArchiveSizeGB) +
                    [double](Convert-ArrayaToNumber -Value $_.ArchiveDeletedItemsGB)
                } | Measure-Object -Sum).Sum), 3)

    $drivers = @(
        [pscustomobject]@{ Driver = 'Total mailbox data (GB)'; Value = $totalDataGB; Notes = 'Primary plus deleted plus archive plus archive deleted.' }
        [pscustomobject]@{ Driver = 'Archive-enabled mailboxes'; Value = $archiveEnabledCount; Notes = ("{0}% of sized objects." -f $archivePercent) }
        [pscustomobject]@{ Driver = ("Mailboxes over {0} GB" -f $largeMailboxGB); Value = $overThresholdCount; Notes = 'Mailbox-only objects can require additional mailbox-license capacity; archive-bearing objects use UMB.' }
        [pscustomobject]@{ Driver = 'Average mailbox size (GB)'; Value = $averageSizeGB; Notes = 'Across objects with a known size.' }
        [pscustomobject]@{ Driver = 'Largest mailbox (GB)'; Value = $(if ($largestLine) { $largestLine.SizeGB } else { $null }); Notes = $(if ($largestLine) { [string]$largestLine.DisplayName } else { 'No sized objects.' }) }
        [pscustomobject]@{ Driver = 'Objects needing license evidence'; Value = $needsDataLines.Count; Notes = 'Missing mailbox statistics or archive state. Excluded from license totals; treat unit baskets as a floor.' }
        [pscustomobject]@{ Driver = 'Empty group mailboxes'; Value = $emptyShellLines.Count; Notes = ("Microsoft 365 Group mailboxes at or below {0} GB. Excluded from licensing - these groups exist for their SharePoint site." -f $emptyGroupFloorGB) }
        [pscustomobject]@{ Driver = 'Group mailboxes with data'; Value = $groupMailboxDataLines.Count; Notes = $(if ($IncludeGroupMailboxLicenseUnits) { 'Included by explicit override.' } else { 'Visible for disposition, but excluded from automatic BitTitan license units by practice policy.' }) }
        [pscustomobject]@{ Driver = 'Inactive mailboxes'; Value = $inactiveLines.Count; Notes = 'Assumes each in-scope inactive mailbox is restored or activated before a regular BitTitan migration job.' }
        [pscustomobject]@{ Driver = 'Inactive mailboxes with archive'; Value = $inactiveArchiveLines.Count; Notes = ("Archive plus archive deleted data: {0} GB." -f $inactiveArchiveDataGB) }
        [pscustomobject]@{ Driver = 'License path selection'; Value = $recommendedPath; Notes = $recommendationReason }
        [pscustomobject]@{ Driver = 'Tenant Migration Bundle'; Value = 'Optional / non-default'; Notes = 'Review only if BitTitan will migrate Teams or SharePoint. Normal ShareGate delivery leaves the included Flex Collaboration entitlement unused.' }
    )

    return [pscustomobject]@{
        ObjectLines     = $allLines
        LicenseMixBreakdown = $licenseMixBreakdown
        SkuTotals       = $skuTotals
        PathComparison  = $pathComparison
        PlanningOptions = $planningOptions
        Drivers         = $drivers
        RecommendedPath = $recommendedPath
        BlockSizeGB     = $blockSizeGB
        TotalDataGB     = $totalDataGB
        NeedsDataCount  = $needsDataLines.Count
        EmptyShellCount = $emptyShellLines.Count
        GroupMailboxDataCount = $groupMailboxDataLines.Count
        ExcludedGroupMailboxCount = $reportOnlyGroupLines.Count
        GroupMailboxLicensingMode = $(if ($IncludeGroupMailboxLicenseUnits) { 'Included by explicit override' } else { 'Report only (practice default)' })
        InactiveMailboxCount = $inactiveLines.Count
        InactiveMailboxWithArchiveCount = $inactiveArchiveLines.Count
        InactiveArchiveDataGB = $inactiveArchiveDataGB
    }
}

function Get-ArrayaBitTitanLicenseModel {
    <#
    .SYNOPSIS
        Loads the BitTitan license model from src/config/baseline, with a code fallback.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        # PSScriptRoot is <repo>/src/modules/Arraya.M365.Common/Public
        $repoRoot = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\..'))
        $Path = Join-Path -Path $repoRoot -ChildPath 'src\config\baseline\bittitan-license-model.json'
    }

    if (Test-Path -Path $Path) {
        try {
            return (Get-Content -Path $Path -Raw | ConvertFrom-Json)
        }
        catch {
            Write-Warning ("Unable to parse BitTitan license model at '{0}'. Falling back to built-in defaults. {1}" -f $Path, $_.Exception.Message)
        }
    }

    # Built-in fallback keeps the engine usable when the config is missing, at the cost of
    # thresholds that nobody has reviewed against the current price list.
    return [pscustomobject]@{
        ModelVersion = '1.1-fallback'
        Skus = [pscustomobject]@{
            MigrationWizMailbox = [pscustomobject]@{ Name = 'MigrationWiz-Mailbox'; BlockSizeGB = 50 }
            UserMigrationBundle = [pscustomobject]@{ Name = 'User Migration Bundle'; PerUser = $true }
        }
        Thresholds = [pscustomobject]@{
            LargeMailboxGB = 50
            VeryLargeMailboxGB = 100
            LargeArchiveGB = 100
            EmptyGroupMailboxFloorGB = 0
        }
        BundleEligibility = [pscustomobject]@{ RecipientTypeDetails = @('UserMailbox'); RequiresBackingUser = $true }
        PracticePolicy = [pscustomobject]@{
            GroupMailboxLicensing = 'ReportOnly'
            InactiveMailboxTreatment = 'RestoreThenMigrate'
            EmitDollarCosts = $false
            AutomaticLicenseRecommendation = $false
        }
        OptionalCommercialOptions = [pscustomobject]@{
            TenantMigrationBundle = [pscustomobject]@{
                Name = 'Tenant Migration Bundle'
                DefaultIncluded = $false
                Includes = @('One User Migration Bundle', 'One Flex Collaboration License')
                Excludes = @('Public folders')
                RequiresCurrentValidation = @('Teams private chat (PCH) capability and licensing')
                PrivateChatProjectType = 'Collaboration (Private Chats)'
                PrivateChatGuidance = 'Teams Private Chat (PCH) uses a separate Collaboration (Private Chats) project. Validate current capability, limitations, and licensing entitlement; do not assume TMB/UMB inclusion or exclusion.'
                PrivateChatGuideUrl = 'https://help.bittitan.com/hc/en-us/articles/25603590557979-Teams-Private-Chat-Migration-Guide'
                GettingStartedUrl = 'https://help.bittitan.com/hc/en-us/articles/1260806785629-Getting-Started-with-Migrations'
            }
        }
    }
}

function New-ArrayaBitTitanObjectLine {
    <#
    .SYNOPSIS
        Classifies one mailbox into a BitTitan SKU line item.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        $Record,

        [Parameter(Mandatory = $true)]
        [string]$ObjectType,

        [Parameter(Mandatory = $true)]
        [double]$BlockSizeGB,

        [Parameter(Mandatory = $true)]
        [string]$SkuName,

        [Parameter(Mandatory = $true)]
        [string]$BundleSkuName,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $EmptyFloorGB,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [string[]]$BundleEligibleTypes,

        [Parameter(Mandatory = $true)]
        $LicensableUserKeys,

        [Parameter(Mandatory = $false)]
        [switch]$ReportOnly
    )

    $displayName = [string](Get-ArrayaObjectValue -Object $Record -Names @('DisplayName'))
    $identity = [string](Get-ArrayaObjectValue -Object $Record -Names @('PrimarySmtpAddress', 'UserPrincipalName', 'Identity', 'ExchangeGuid'))
    $recipientType = [string](Get-ArrayaObjectValue -Object $Record -Names @('RecipientTypeDetails'))
    $archiveStatusValue = Get-ArrayaObjectValue -Object $Record -Names @('ArchiveStatus')
    $archiveStatus = [string]$archiveStatusValue
    $inactiveValue = Get-ArrayaObjectValue -Object $Record -Names @('IsInactiveMailbox', 'IsInactive')
    $isInactiveMailbox = ($inactiveValue -eq $true -or [string]$inactiveValue -match '^(?i:true)$')

    # Every size read goes through the same normalizer. The legacy inline math mixed direct
    # [double] casts with a size parser, so a value like "120 GB" behaved differently
    # depending on which threshold was being evaluated.
    $primaryGB = Convert-ArrayaBitTitanSizeToGb -Value (Get-ArrayaObjectValue -Object $Record -Names @('MailboxSizeGB', 'MBXSizeGB'))
    $deletedGB = Convert-ArrayaBitTitanSizeToGb -Value (Get-ArrayaObjectValue -Object $Record -Names @('DeletedItemsGB'))
    $archiveGB = Convert-ArrayaBitTitanSizeToGb -Value (Get-ArrayaObjectValue -Object $Record -Names @('ArchiveSizeGB'))
    $archiveDeletedGB = Convert-ArrayaBitTitanSizeToGb -Value (Get-ArrayaObjectValue -Object $Record -Names @('ArchiveDeletedItemsGB'))
    $archiveGuid = [string](Get-ArrayaObjectValue -Object $Record -Names @('ArchiveGuid'))
    $hasActiveArchive = ($archiveStatus -match '^(?i:active|enabled)$')
    $hasArchive = (
        $hasActiveArchive -or
        ($null -ne $archiveGB -and [double]$archiveGB -gt 0) -or
        ($null -ne $archiveDeletedGB -and [double]$archiveDeletedGB -gt 0) -or
        (-not [string]::IsNullOrWhiteSpace($archiveGuid) -and $archiveGuid -notmatch '^0{8}-0{4}-0{4}-0{4}-0{12}$')
    )
    $archiveStateKnown = (
        $null -ne $archiveStatusValue -and
        -not [string]::IsNullOrWhiteSpace($archiveStatus) -and
        $archiveStatus -notmatch '^(?i:unknown|unavailable|n/a)$'
    )
    if (-not $archiveStateKnown -and (
            ($null -ne $archiveGB) -or
            ($null -ne $archiveDeletedGB) -or
            (-not [string]::IsNullOrWhiteSpace($archiveGuid)))) {
        $archiveStateKnown = $true
    }
    $migrationPrerequisite = if ($isInactiveMailbox) {
        'Restore or activate the inactive mailbox before running a regular BitTitan mailbox/archive migration job.'
    } else { $null }

    $sizeParts = @($primaryGB, $deletedGB, $archiveGB, $archiveDeletedGB) | Where-Object { $null -ne $_ }
    $totalGB = if ($sizeParts.Count -gt 0) {
        [math]::Round((($sizeParts | Measure-Object -Sum).Sum), 3)
    } else {
        $null
    }

    $bundleEligible = (
        ($BundleEligibleTypes -contains $recipientType) -and
        (-not [string]::IsNullOrWhiteSpace($identity)) -and
        $LicensableUserKeys.Contains($identity.Trim())
    )
    $umbEligibilityValidationRequired = ($hasArchive -and -not $bundleEligible)

    # No size data means no defensible license count. Bucket it rather than dropping it,
    # which is what the previous model did by returning a null license type.
    if ($null -eq $totalGB) {
        return [pscustomobject]@{
            ObjectType       = $ObjectType
            Identity         = $identity
            DisplayName      = $displayName
            RecipientType    = $recipientType
            SizeGB           = $null
            IsInactiveMailbox = $isInactiveMailbox
            ArchiveStatus    = $archiveStatus
            HasArchive       = $hasArchive
            HasActiveArchive = $hasActiveArchive
            ArchiveSizeGB    = $archiveGB
            ArchiveDeletedItemsGB = $archiveDeletedGB
            MigrationPrerequisite = $migrationPrerequisite
            RecommendedSku   = $(if ($ReportOnly) { 'Report only - needs sizing data' } else { 'Needs data' })
            LicenseUnits     = $null
            BundleEligible   = $bundleEligible
            UmbEligibilityValidationRequired = $umbEligibilityValidationRequired
            EstimateStatus   = $(if ($ReportOnly) { 'Report only - manual decision' } else { 'Needs sizing data' })
            SizingBasis      = 'No mailbox statistics available'
            Notes            = $(if ($ReportOnly) {
                    'Microsoft 365 Group mailbox is inventory-only by practice policy. Collect statistics only if the customer elects to migrate its conversations.'
                } elseif ($isInactiveMailbox) {
                    'Excluded from license totals until statistics are available. Restore or activate the inactive mailbox before migration.'
                } else {
                    'Excluded from license totals. Collect mailbox statistics before quoting.'
                })
        }
    }

    # Without archive-state evidence, mailbox sizing alone cannot determine whether a
    # Mailbox license is valid or UMB is required. Treat it as a quote condition rather than
    # silently assuming no archive.
    if (-not $archiveStateKnown -and $ObjectType -ne 'GroupMailbox') {
        return [pscustomobject]@{
            ObjectType       = $ObjectType
            Identity         = $identity
            DisplayName      = $displayName
            RecipientType    = $recipientType
            SizeGB           = $totalGB
            IsInactiveMailbox = $isInactiveMailbox
            ArchiveStatus    = $archiveStatus
            HasArchive       = $null
            HasActiveArchive = $false
            ArchiveSizeGB    = $archiveGB
            ArchiveDeletedItemsGB = $archiveDeletedGB
            MigrationPrerequisite = $migrationPrerequisite
            RecommendedSku   = 'Needs data'
            LicenseUnits     = $null
            BundleEligible   = $bundleEligible
            UmbEligibilityValidationRequired = $false
            EstimateStatus   = 'Needs archive status'
            SizingBasis      = ("{0} GB discovered; archive state unknown" -f $totalGB)
            Notes            = 'Confirm whether an archive exists. Mailbox licenses are only valid for mailbox-only migrations; archive content requires UMB.'
        }
    }

    # Group mailboxes below the floor exist only for their SharePoint site. Counting them as
    # mailbox licenses is the single largest source of quote inflation on group-heavy tenants.
    if ($ObjectType -eq 'GroupMailbox' -and $null -ne $EmptyFloorGB -and $totalGB -le [double]$EmptyFloorGB) {
        return [pscustomobject]@{
            ObjectType       = $ObjectType
            Identity         = $identity
            DisplayName      = $displayName
            RecipientType    = $recipientType
            SizeGB           = $totalGB
            IsInactiveMailbox = $isInactiveMailbox
            ArchiveStatus    = $archiveStatus
            HasArchive       = $hasArchive
            HasActiveArchive = $hasActiveArchive
            ArchiveSizeGB    = $archiveGB
            ArchiveDeletedItemsGB = $archiveDeletedGB
            MigrationPrerequisite = $migrationPrerequisite
            RecommendedSku   = 'No migration'
            LicenseUnits     = $null
            BundleEligible   = $false
            UmbEligibilityValidationRequired = $false
            EstimateStatus   = 'Excluded - no mail data'
            SizingBasis      = ("At or below the {0} GB empty-group floor" -f $EmptyFloorGB)
            Notes            = 'Group has no mail data to migrate. Its SharePoint site is ShareGate scope.'
        }
    }

    if ($ObjectType -eq 'GroupMailbox' -and $ReportOnly) {
        return [pscustomobject]@{
            ObjectType       = $ObjectType
            Identity         = $identity
            DisplayName      = $displayName
            RecipientType    = $recipientType
            SizeGB           = $totalGB
            IsInactiveMailbox = $isInactiveMailbox
            ArchiveStatus    = $archiveStatus
            HasArchive       = $hasArchive
            HasActiveArchive = $hasActiveArchive
            ArchiveSizeGB    = $archiveGB
            ArchiveDeletedItemsGB = $archiveDeletedGB
            MigrationPrerequisite = $migrationPrerequisite
            RecommendedSku   = 'Report only'
            LicenseUnits     = $null
            BundleEligible   = $false
            UmbEligibilityValidationRequired = $false
            EstimateStatus   = 'Report only - manual decision'
            SizingBasis      = ("{0} GB discovered; no automatic license unit" -f $totalGB)
            Notes            = 'Microsoft 365 Group mailbox conversations remain visible for customer disposition, but practice policy excludes them from automatic BitTitan license counts.'
        }
    }

    $units = if ($hasArchive) { 1 } else { [int][math]::Max(1, [math]::Ceiling($totalGB / $BlockSizeGB)) }
    $selectedSku = if ($hasArchive) { $BundleSkuName } else { $SkuName }

    return [pscustomobject]@{
        ObjectType       = $ObjectType
        Identity         = $identity
        DisplayName      = $displayName
        RecipientType    = $recipientType
        SizeGB           = $totalGB
        IsInactiveMailbox = $isInactiveMailbox
        ArchiveStatus    = $archiveStatus
        HasArchive       = $hasArchive
        HasActiveArchive = $hasActiveArchive
        ArchiveSizeGB    = $archiveGB
        ArchiveDeletedItemsGB = $archiveDeletedGB
        MigrationPrerequisite = $migrationPrerequisite
        RecommendedSku   = $selectedSku
        LicenseUnits     = $units
        BundleEligible   = $bundleEligible
        UmbEligibilityValidationRequired = $umbEligibilityValidationRequired
        EstimateStatus   = $(if ($umbEligibilityValidationRequired) { 'Included as UMB planning unit - validate eligibility' } else { 'Included in unit estimate' })
        SizingBasis      = $(if ($hasArchive) {
                ("Archive present; one {0} for this migration object ({1} GB discovered total)" -f $BundleSkuName, $totalGB)
            } else {
                ("{0} GB over {1} GB mailbox-license blocks" -f $totalGB, $BlockSizeGB)
            })
        Notes            = $(if ($isInactiveMailbox -and $hasArchive -and $umbEligibilityValidationRequired) {
                'Assumes restore or activation before migration. Archive requires one UMB planning unit; validate UMB eligibility after restoring the identity.'
            } elseif ($isInactiveMailbox -and $hasArchive) {
                'Assumes restore or activation before migration. Archive is included in the sizing total.'
            } elseif ($isInactiveMailbox) {
                'Assumes restore or activation before migration.'
            } elseif ($hasArchive -and $umbEligibilityValidationRequired) {
                'Archive requires one UMB planning unit. Validate UMB eligibility for this non-user or unmatched identity before quoting.'
            } elseif ($hasArchive) {
                'Archive included in the sizing total.'
            } else { $null })
    }
}

function Convert-ArrayaBitTitanSizeToGb {
    <#
    .SYNOPSIS
        Normalizes a size value to GB, accepting numbers and Exchange-style size strings.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) { return $null }

    if ($Value -is [double] -or $Value -is [float] -or $Value -is [decimal] -or
        $Value -is [int] -or $Value -is [long] -or $Value -is [int64]) {
        return [math]::Round(([double]$Value), 3)
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text) -or $text -in @('N/A', 'Unknown', 'Unavailable')) {
        return $null
    }

    # Exchange renders sizes as "1.234 GB (1,325,400,064 bytes)". Prefer the leading
    # human-readable value and its unit; the byte count in parentheses is redundant.
    if ($text -match '^\s*(?<number>-?[\d,]+(?:\.\d+)?)\s*(?<unit>B|KB|MB|GB|TB)?\b') {
        $numberText = $Matches['number'] -replace ',', ''
        $numeric = Convert-ArrayaToNumber -Value $numberText
        if ($null -eq $numeric) { return $null }

        $unit = if ($Matches['unit']) { $Matches['unit'].ToUpperInvariant() } else { 'GB' }
        $gb = switch ($unit) {
            'B'  { [double]$numeric / 1GB }
            'KB' { [double]$numeric / 1MB }
            'MB' { [double]$numeric / 1KB }
            'TB' { [double]$numeric * 1024 }
            default { [double]$numeric }
        }

        return [math]::Round($gb, 3)
    }

    return $null
}
