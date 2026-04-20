function Get-ArrayaAssessmentPipelineNamedPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Object,
        [Parameter(Mandatory = $true)]
        [string[]]$Names
    )

    if ($null -eq $Object) {
        return $null
    }

    foreach ($name in $Names) {
        if ($Object -is [System.Collections.IDictionary] -and $Object.Contains($name)) {
            return $Object[$name]
        }

        $matchedProperty = @($Object.PSObject.Properties.Match($name) | Select-Object -First 1)
        if ($matchedProperty.Count -gt 0) {
            return $matchedProperty[0].Value
        }

        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $camelName = if ($name.Length -gt 1) {
                '{0}{1}' -f $name.Substring(0, 1).ToLowerInvariant(), $name.Substring(1)
            }
            else {
                $name.ToLowerInvariant()
            }

            $camelProperty = @($Object.PSObject.Properties.Match($camelName) | Select-Object -First 1)
            if ($camelProperty.Count -gt 0) {
                return $camelProperty[0].Value
            }
        }
    }

    return $null
}

function Get-ArrayaAssessmentPipelineFirstPopulatedValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object[]]$Values
    )

    foreach ($value in $Values) {
        if ($null -eq $value) {
            continue
        }
        if ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) {
            continue
        }
        return $value
    }

    return $null
}

function Get-ArrayaAssessmentPipelineLicenseReferenceMap {
    [CmdletBinding()]
    param()

    $referenceMapInitialized = Get-Variable -Name PipelineLicenseReferenceMapInitialized -Scope Script -ErrorAction SilentlyContinue
    $referenceMapVariable = Get-Variable -Name PipelineLicenseReferenceMap -Scope Script -ErrorAction SilentlyContinue
    if ($referenceMapInitialized -and [bool]$referenceMapInitialized.Value -and $referenceMapVariable -and $referenceMapVariable.Value) {
        return $script:PipelineLicenseReferenceMap
    }

    $script:PipelineLicenseReferenceMap = @{}
    $script:PipelineLicenseReferenceMapInitialized = $true

    $knownCsvUri = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
    $docsUri = 'https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference'
    $cacheFile = Join-Path -Path $env:TEMP -ChildPath 'arraya-license-service-plan-reference.csv'

    try {
        $useCachedCsv = $false
        if (Test-Path -Path $cacheFile) {
            $cacheAge = (Get-Date) - (Get-Item -Path $cacheFile).LastWriteTime
            if ($cacheAge.TotalDays -lt 7) {
                $useCachedCsv = $true
            }
        }

        if (-not $useCachedCsv) {
            $csvUri = $knownCsvUri
            try {
                $docsResponse = Invoke-WebRequest -Uri $docsUri -UseBasicParsing -ErrorAction Stop
                $pattern = "https://download\.microsoft\.com/download/[^\s""'<>]+Product%20names%20and%20service%20plan%20identifiers%20for%20licensing\.csv"
                $match = [regex]::Match($docsResponse.Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($match.Success -and -not [string]::IsNullOrWhiteSpace($match.Value)) {
                    $csvUri = $match.Value
                }
            }
            catch {}

            Invoke-WebRequest -Uri $csvUri -OutFile $cacheFile -UseBasicParsing -ErrorAction Stop
        }

        $csvRows = Import-Csv -Path $cacheFile -ErrorAction Stop
        foreach ($row in $csvRows) {
            $stringId = [string]$row.String_Id
            $displayName = [string]$row.Product_Display_Name
            if ([string]::IsNullOrWhiteSpace($stringId) -or [string]::IsNullOrWhiteSpace($displayName)) {
                continue
            }

            if (-not $script:PipelineLicenseReferenceMap.ContainsKey($stringId)) {
                $script:PipelineLicenseReferenceMap[$stringId] = $displayName
            }

            $normalizedStringId = ($stringId -replace '\s*_\s*', '_').Trim()
            if (-not [string]::IsNullOrWhiteSpace($normalizedStringId) -and -not $script:PipelineLicenseReferenceMap.ContainsKey($normalizedStringId)) {
                $script:PipelineLicenseReferenceMap[$normalizedStringId] = $displayName
            }
        }
    }
    catch {
        $script:PipelineLicenseReferenceMap = @{}
    }

    return $script:PipelineLicenseReferenceMap
}

function Get-ArrayaAssessmentPipelineFriendlyProductName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$SkuPartNumber
    )

    if ([string]::IsNullOrWhiteSpace($SkuPartNumber)) {
        return $SkuPartNumber
    }

    $referenceMap = Get-ArrayaAssessmentPipelineLicenseReferenceMap
    if ($referenceMap.ContainsKey($SkuPartNumber)) {
        return $referenceMap[$SkuPartNumber]
    }

    $normalizedSkuPartNumber = ($SkuPartNumber -replace '\s*_\s*', '_').Trim()
    if ($referenceMap.ContainsKey($normalizedSkuPartNumber)) {
        return $referenceMap[$normalizedSkuPartNumber]
    }

    switch -Regex ($normalizedSkuPartNumber.ToUpperInvariant()) {
        '^SPE_E([35])$' { return "Microsoft 365 E$($Matches[1])" }
        '^SPE_F([13])$' { return "Microsoft 365 F$($Matches[1])" }
        '^MCOPSTN1$' { return 'Microsoft Teams Phone Standard' }
        '^MCOCAP$' { return 'Microsoft Teams Shared Space' }
        '^MICROSOFT_TEAMS_ENTERPRISE_NEW$' { return 'Microsoft Teams Enterprise' }
        '^POWERAPPS_PER_USER$' { return 'Power Apps Premium' }
        '^EXCHANGEENTERPRISE$' { return 'Exchange Online (Plan 2)' }
        '^CPC_E_(\d+)C_(\d+)GB_(\d+)GB$' { return "Windows 365 Enterprise $($Matches[1]) vCPU $($Matches[2]) GB $($Matches[3]) GB" }
        '^WINDOWS_365_S_(\d+)VCPU_(\d+)GB_(\d+)GB$' { return "Windows 365 Shared Use $($Matches[1]) vCPU $($Matches[2]) GB $($Matches[3]) GB" }
    }

    $friendly = $normalizedSkuPartNumber -replace '[_\-]+', ' '
    $friendly = $friendly -replace '\bM365\b', 'Microsoft 365'
    $friendly = $friendly -replace '\bO365\b', 'Office 365'
    $friendly = $friendly -replace '\bEXCHANGE\b', 'Exchange'
    $friendly = $friendly -replace '\bPOWERAPPS\b', 'Power Apps'

    $tokens = $friendly -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
        if ($_ -match '^[A-Z0-9]{2,}$' -or $_ -match '^[EF]\d$') {
            $_
        }
        else {
            (Get-Culture).TextInfo.ToTitleCase($_.ToLowerInvariant())
        }
    }

    $friendly = ($tokens -join ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($friendly)) {
        return $normalizedSkuPartNumber
    }

    return $friendly
}

function Get-ArrayaAssessmentPipelineLicenseClassification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$License,
        [Parameter(Mandatory = $false)]
        [int]$UserCount = 0
    )

    $skuPartNumber = [string]$License.SkuPartNumber
    $appliesTo = [string]$License.AppliesTo
    $isTrialFlag = [bool]($License.PSObject.Properties['IsTrial'] -and $License.IsTrial -eq $true)
    $isFreeOrTrialFlag = [bool]($License.PSObject.Properties['IsFreeOrTrial'] -and $License.IsFreeOrTrial -eq $true)
    $ignoreLifecycle = [bool]($License.PSObject.Properties['IgnoreLifecycle'] -and $License.IgnoreLifecycle -eq $true)
    $purchased = 0
    $consumed = 0

    if ($License.PSObject.Properties['PurchasedUnits'] -and $null -ne $License.PurchasedUnits -and $License.PurchasedUnits -ne 'N/A' -and $License.PurchasedUnits -ne '') {
        try { $purchased = [int64]$License.PurchasedUnits } catch { $purchased = 0 }
    }
    if ($License.PSObject.Properties['ConsumedUnits'] -and $null -ne $License.ConsumedUnits -and $License.ConsumedUnits -ne 'N/A' -and $License.ConsumedUnits -ne '') {
        try { $consumed = [int64]$License.ConsumedUnits } catch { $consumed = 0 }
    }

    if ($isTrialFlag) {
        return [pscustomobject]@{ IsPaid = $false; LicenseClass = 'FreeTrialBenefit'; Reason = 'Marked as trial by subscription metadata' }
    }
    if ($isFreeOrTrialFlag) {
        return [pscustomobject]@{ IsPaid = $false; LicenseClass = 'FreeTrialBenefit'; Reason = 'Marked as free or trial by lifecycle metadata' }
    }
    if ($ignoreLifecycle) {
        return [pscustomobject]@{ IsPaid = $false; LicenseClass = 'FreeTrialBenefit'; Reason = 'Ignored due to lifecycle metadata anomaly' }
    }

    $alwaysExcludeSkus = @(
        'MCOPSTNC', 'STREAM', 'FORMS_PRO', 'POWER_BI_STANDARD', 'RIGHTSMANAGEMENT_ADHOC',
        'PROJECT_MADEIRA_PREVIEW_IW_SKU', 'DYN365_ENTERPRISE_P1_IW', 'POWERAPPS_INDIVIDUAL_USER',
        'POWERAPPS_DEV', 'POWERAPPS_VIRAL', 'FLOW_FREE', 'CCIBOTS_PRIVPREV_VIRAL',
        'Power_Pages_vTrial_for_Makers', 'WINDOWS_STORE'
    )

    if ($alwaysExcludeSkus -contains $skuPartNumber) {
        return [pscustomobject]@{ IsPaid = $false; LicenseClass = 'FreeTrialBenefit'; Reason = 'Excluded by known freemium or benefit SKU list' }
    }

    if ($skuPartNumber -match 'TRIAL|EXPLORATORY|FREE|VIRAL|_FACULTY|_STUDENT|PREVIEW|_IW($|_)|ADHOC|INDIVIDUAL|_DEV($|_)|_TRIAL($|_)|FOR_MAKERS') {
        return [pscustomobject]@{ IsPaid = $false; LicenseClass = 'FreeTrialBenefit'; Reason = 'Excluded by SKU naming pattern' }
    }

    if ($UserCount -gt 0 -and $appliesTo -eq 'User') {
        $seatThreshold = [math]::Max(($UserCount * 10), 5000)
        $consumedThreshold = [math]::Max(($UserCount * 2), 250)
        if ($purchased -ge $seatThreshold -and $consumed -le $consumedThreshold) {
            return [pscustomobject]@{ IsPaid = $false; LicenseClass = 'FreeTrialBenefit'; Reason = "Excluded by tenant user-count heuristic ($purchased seats for $UserCount users)" }
        }
    }

    return [pscustomobject]@{ IsPaid = $true; LicenseClass = 'Paid'; Reason = 'Included as paid license inventory' }
}

function Get-ArrayaAssessmentPipelineTenantOverviewInfo {
    [CmdletBinding()]
    param()

    $org = Get-ArrayaAssessmentPipelineTenantOrganization
    $initialDomain = $null
    $defaultDomain = $null
    $verifiedDomains = @()
    if ($org -and $org.PSObject.Properties['VerifiedDomains']) {
        $verifiedDomains = @($org.VerifiedDomains)
        $initial = @($verifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -First 1)
        $default = @($verifiedDomains | Where-Object { $_.IsDefault -eq $true } | Select-Object -First 1)
        if ($initial.Count -gt 0) { $initialDomain = [string](Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $initial[0] -Names @('Name', 'Id')) }
        if ($default.Count -gt 0) { $defaultDomain = [string](Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $default[0] -Names @('Name', 'Id')) }
    }

    $multiGeoEnabled = $null
    $multiGeoAllowed = @()
    $multiGeoCentral = $null
    if ($org.PSObject.Properties['IsMultipleGeolocationEnabled']) {
        $multiGeoEnabled = [bool]$org.IsMultipleGeolocationEnabled
    }
    if ($org.PSObject.Properties['MultiGeoConfiguration'] -and $org.MultiGeoConfiguration) {
        $cfg = $org.MultiGeoConfiguration
        if ($cfg.PSObject.Properties['AllowedDataLocations'] -and $cfg.AllowedDataLocations) {
            $multiGeoAllowed = @($cfg.AllowedDataLocations)
            if ($multiGeoAllowed.Count -gt 1) { $multiGeoEnabled = $true }
        }
        if ($cfg.PSObject.Properties['PreferredDataLocation'] -and $cfg.PreferredDataLocation) {
            $multiGeoCentral = $cfg.PreferredDataLocation
        }
    }

    $selfServiceAnswer = 'Not collected'
    $selfServiceNotes = 'MSCommerce module not installed'
    $selfServiceDetails = $null
    try {
        if (-not (Get-Module -ListAvailable -Name MSCommerce)) {
            Install-Module -Name MSCommerce -Scope CurrentUser -Force -ErrorAction Stop
        }

        if ($PSVersionTable.PSVersion.Major -ge 6) {
            Import-Module MSCommerce -UseWindowsPowerShell -ErrorAction Stop -WarningAction SilentlyContinue
        }
        else {
            Import-Module MSCommerce -ErrorAction Stop -WarningAction SilentlyContinue
        }

        try {
            Connect-MSCommerce -ErrorAction Stop | Out-Null
            $sspPolicies = Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase -ErrorAction Stop
            if ($sspPolicies -and $sspPolicies.Count -gt 0) {
                $enabled = ($sspPolicies | Where-Object { $_.Value -eq 'Enabled' }).Count
                $disabled = ($sspPolicies | Where-Object { $_.Value -eq 'Disabled' }).Count
                $trialOnly = ($sspPolicies | Where-Object { $_.Value -eq 'OnlyTrialsWithoutPaymentMethod' }).Count
                $total = $sspPolicies.Count
                if ($enabled -gt 0) {
                    $selfServiceAnswer = "Yes (Enabled for $enabled/$total products)"
                }
                else {
                    $selfServiceAnswer = "No (Disabled for all $total products)"
                }
                $selfServiceNotes = "Disabled: $disabled; Trial-only: $trialOnly"
                $selfServiceDetails = [pscustomobject]@{
                    EnabledCount  = $enabled
                    DisabledCount = $disabled
                    TrialOnlyCount = $trialOnly
                    TotalCount = $total
                }
            }
            else {
                $selfServiceAnswer = 'Unknown'
                $selfServiceNotes = 'No policy data returned'
            }
        }
        catch {
            $selfServiceAnswer = 'Unknown'
            $selfServiceNotes = "MSCommerce connection or policy retrieval failed: $($_.Exception.Message)"
        }
    }
    catch {
        $selfServiceAnswer = 'Unknown'
        $selfServiceNotes = "MSCommerce module install/load failed: $($_.Exception.Message)"
    }

    return [pscustomobject]@{
        DisplayName           = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $org -Names @('DisplayName')
        TenantId              = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $org -Names @('Id', 'TenantId')
        InitialDomain         = $initialDomain
        DefaultDomain         = $defaultDomain
        PreferredDataLocation = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $org -Names @('PreferredDataLocation')
        Country               = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $org -Names @('Country')
        CountryLetterCode     = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $org -Names @('CountryLetterCode')
        MultiGeoEnabled       = $multiGeoEnabled
        MultiGeoAllowed       = $multiGeoAllowed
        MultiGeoCentral       = $multiGeoCentral
        SelfServicePurchase   = [pscustomobject]@{
            Answer  = $selfServiceAnswer
            Notes   = $selfServiceNotes
            Details = $selfServiceDetails
        }
        AzureResourceUsage    = [pscustomobject]@{
            Answer  = 'Not collected'
            Notes   = 'Azure module not used in this report'
            Details = $null
        }
    }
}

function Get-ArrayaAssessmentPipelineLicenseSkus {
    [CmdletBinding()]
    param()

    $licenseSkus = @{}
    $skus = @(Get-MgSubscribedSku -ProgressAction SilentlyContinue -ErrorAction Stop | Where-Object { $_.AppliesTo })
    $subscriptions = @()
    try {
        $subscriptions = @(Get-MgDirectorySubscription -All -ProgressAction SilentlyContinue -ErrorAction SilentlyContinue)
    }
    catch {}

    $subscriptionLookup = @{}
    foreach ($sub in $subscriptions) {
        if ($sub.SkuId) {
            $subscriptionLookup[$sub.SkuId.ToString()] = $sub
        }
    }

    foreach ($sku in $skus) {
        $skuId = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $sku -Names @('SkuId')
        if ($null -eq $skuId) {
            continue
        }
        $skuIdText = [string]$skuId.ToString()
        $subMatch = if ($subscriptionLookup.ContainsKey($skuIdText)) { $subscriptionLookup[$skuIdText] } else { $null }
        $skuPartNumber = [string](Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $sku -Names @('SkuPartNumber'))
        $friendlySkuName = Get-ArrayaAssessmentPipelineFriendlyProductName -SkuPartNumber $skuPartNumber
        $prepaidUnits = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object (Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $sku -Names @('PrepaidUnits')) -Names @('Enabled')
        $consumedUnits = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $sku -Names @('ConsumedUnits')
        $servicePlans = @(Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $sku -Names @('ServicePlans'))
        $servicePlanNames = @($servicePlans | ForEach-Object { Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $_ -Names @('ServicePlanName') } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })

        $skuDetails = [pscustomobject]@{
            AccountSkuId    = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $sku -Names @('AccountSkuId')
            AccountName     = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $sku -Names @('AccountName')
            AppliesTo       = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $sku -Names @('AppliesTo')
            CapabilityStatus = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $sku -Names @('CapabilityStatus')
            SkuId           = $skuId
            SkuPartNumber   = $skuPartNumber
            SkuFriendlyName = $friendlySkuName
            ConsumedUnits   = $consumedUnits
            PurchasedUnits  = $prepaidUnits
            RemainingUnits  = (($prepaidUnits | ForEach-Object { [int64]$_ }) - ($consumedUnits | ForEach-Object { [int64]$_ }))
            ServicePlansCount = $servicePlanNames.Count
            ServicePlans    = ($servicePlanNames -join ',')
        }

        if ($subMatch) {
            $skuDetails | Add-Member -MemberType NoteProperty -Name IsTrial -Value $subMatch.IsTrial -Force
            $skuDetails | Add-Member -MemberType NoteProperty -Name SubscriptionStatus -Value $subMatch.Status -Force
            $skuDetails | Add-Member -MemberType NoteProperty -Name NextLifecycleDateTime -Value $subMatch.NextLifecycleDateTime -Force
            $skuDetails | Add-Member -MemberType NoteProperty -Name TotalLicenses -Value $subMatch.TotalLicenses -Force

            $nextLifecycle = $subMatch.NextLifecycleDateTime
            $isFreeOrTrial = ($null -eq $nextLifecycle -or $nextLifecycle -eq '')
            $isIgnoredLifecycle = $false
            if ($nextLifecycle) {
                try {
                    $dt = [datetime]$nextLifecycle
                    if ($dt.Year -eq 9999) { $isIgnoredLifecycle = $true }
                }
                catch {}
            }
            if ($skuPartNumber -eq 'MCOPSTNC') {
                $isFreeOrTrial = $true
            }
            $skuDetails | Add-Member -MemberType NoteProperty -Name IsFreeOrTrial -Value $isFreeOrTrial -Force
            $skuDetails | Add-Member -MemberType NoteProperty -Name IgnoreLifecycle -Value $isIgnoredLifecycle -Force
        }

        $skuClassification = Get-ArrayaAssessmentPipelineLicenseClassification -License $skuDetails
        $skuDetails | Add-Member -MemberType NoteProperty -Name IsPaid -Value $skuClassification.IsPaid -Force
        $skuDetails | Add-Member -MemberType NoteProperty -Name LicenseClass -Value $skuClassification.LicenseClass -Force
        $skuDetails | Add-Member -MemberType NoteProperty -Name LicenseClassificationReason -Value $skuClassification.Reason -Force

        $licenseSkus[$skuIdText] = $skuDetails
    }

    return $licenseSkus
}

function Get-ArrayaAssessmentPipelineAdConnectSyncDetails {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$TenantInfo,
        [Parameter(Mandatory = $false)]
        [object]$ExistingPasswordLifecycleSummary
    )

    $org = $null
    try { $org = Get-ArrayaAssessmentPipelineTenantOrganization } catch {}

    $summary = [pscustomobject]@{
            OnPremisesSyncEnabled                           = if ($org) { Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $org -Names @('OnPremisesSyncEnabled') } else { $null }
            OnPremisesLastSyncDateTime                     = if ($org) { Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $org -Names @('OnPremisesLastSyncDateTime') } else { $null }
        PasswordSyncEnabled                            = $null
        PasswordWritebackEnabled                       = $null
        CloudPasswordPolicyForPasswordSyncedUsersEnabled = $null
        UserForcePasswordChangeOnLogonEnabled          = $null
        DeviceWritebackEnabled                         = $null
        UnifiedGroupWritebackEnabled                   = $null
        UserWritebackEnabled                           = $null
        PassThroughAuthenticationEnabled               = $null
    }

    $serviceDetails = @()
    $syncErrors = @()
    $syncFeatureCollectionNote = $null
    try {
        $syncServiceResponse = $null
        if (Get-Command Get-MgDirectoryOnPremiseSynchronization -ErrorAction SilentlyContinue) {
            $syncServiceResponse = Get-MgDirectoryOnPremiseSynchronization -ErrorAction Stop
        }
        elseif (Get-Variable -Name GraphHeaders -Scope Global -ErrorAction SilentlyContinue) {
            $syncServiceResponse = Office365Custom\Get-GraphData -Uri 'https://graph.microsoft.com/v1.0/directory/onPremisesSynchronization' -Activity 'Fetching directory synchronization service features'
        }

        $syncServiceObject = @($syncServiceResponse) | Select-Object -First 1
        if ($syncServiceObject) {
            $features = if ($syncServiceObject.PSObject.Properties['Features']) { $syncServiceObject.Features } else { $syncServiceObject.features }
            if ($features) {
                foreach ($mapping in @(
                    @{ Summary = 'PasswordSyncEnabled'; PropertyNames = @('PasswordSyncEnabled', 'PasswordHashSyncEnabled') },
                    @{ Summary = 'PasswordWritebackEnabled'; PropertyNames = @('PasswordWritebackEnabled') },
                    @{ Summary = 'CloudPasswordPolicyForPasswordSyncedUsersEnabled'; PropertyNames = @('CloudPasswordPolicyForPasswordSyncedUsersEnabled') },
                    @{ Summary = 'UserForcePasswordChangeOnLogonEnabled'; PropertyNames = @('UserForcePasswordChangeOnLogonEnabled') },
                    @{ Summary = 'DeviceWritebackEnabled'; PropertyNames = @('DeviceWritebackEnabled') },
                    @{ Summary = 'UnifiedGroupWritebackEnabled'; PropertyNames = @('UnifiedGroupWritebackEnabled') },
                    @{ Summary = 'UserWritebackEnabled'; PropertyNames = @('UserWritebackEnabled') },
                    @{ Summary = 'PassThroughAuthenticationEnabled'; PropertyNames = @('PassThroughAuthenticationEnabled', 'PassthroughAuthenticationEnabled', 'PassThroughAuthentication') }
                )) {
                    $featureValue = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $features -Names $mapping.PropertyNames
                    if ($null -ne $featureValue) {
                        try { $summary.$($mapping.Summary) = [bool]$featureValue } catch { $summary.$($mapping.Summary) = $featureValue }
                    }
                }
            }
        }
    }
    catch {
        $syncFeatureCollectionNote = if ($_.Exception.Message -match 'Authorization_RequestDenied|Insufficient privileges|403 Forbidden') {
            'Not available in current auth mode'
        }
        else {
            'Not surfaced in current source'
        }
    }

    $passwordFeatureFallback = $null
    if ($summary.OnPremisesSyncEnabled -eq $true -and -not [string]::IsNullOrWhiteSpace($syncFeatureCollectionNote)) {
        $passwordFeatureFallback = $syncFeatureCollectionNote
    }

    $passwordLifecycleSummary = [pscustomobject]@{
        PasswordWriteback                            = Get-ArrayaAssessmentPipelineFirstPopulatedValue @($summary.PasswordWritebackEnabled, (Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('PasswordWriteback', 'PasswordWritebackEnabled')), $passwordFeatureFallback)
        PasswordWritebackEnabled                     = Get-ArrayaAssessmentPipelineFirstPopulatedValue @($summary.PasswordWritebackEnabled, (Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('PasswordWritebackEnabled', 'PasswordWriteback')), $passwordFeatureFallback)
        PasswordSyncEnabled                          = Get-ArrayaAssessmentPipelineFirstPopulatedValue @($summary.PasswordSyncEnabled, (Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('PasswordSyncEnabled', 'PasswordHashSyncEnabled')), $passwordFeatureFallback)
        CloudPasswordPolicyForPasswordSyncedUsersEnabled = Get-ArrayaAssessmentPipelineFirstPopulatedValue @($summary.CloudPasswordPolicyForPasswordSyncedUsersEnabled, (Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('CloudPasswordPolicyForPasswordSyncedUsersEnabled')), $passwordFeatureFallback)
        UserForcePasswordChangeOnLogonEnabled        = Get-ArrayaAssessmentPipelineFirstPopulatedValue @($summary.UserForcePasswordChangeOnLogonEnabled, (Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('UserForcePasswordChangeOnLogonEnabled')), $passwordFeatureFallback)
        PassThroughAuthentication                    = Get-ArrayaAssessmentPipelineFirstPopulatedValue @($summary.PassThroughAuthenticationEnabled, (Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('PassThroughAuthentication', 'PassThroughAuthenticationEnabled')), $passwordFeatureFallback)
        PassThroughAuthenticationEnabled             = Get-ArrayaAssessmentPipelineFirstPopulatedValue @($summary.PassThroughAuthenticationEnabled, (Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('PassThroughAuthenticationEnabled', 'PassThroughAuthentication')), $passwordFeatureFallback)
        SelfServicePasswordReset                     = Get-ArrayaAssessmentPipelineFirstPopulatedValue @((Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('SelfServicePasswordReset', 'SelfServicePasswordResetEnabled')))
        SelfServicePasswordResetEnabled              = Get-ArrayaAssessmentPipelineFirstPopulatedValue @((Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('SelfServicePasswordResetEnabled', 'SelfServicePasswordReset')))
        DeviceWritebackEnabled                       = Get-ArrayaAssessmentPipelineFirstPopulatedValue @($summary.DeviceWritebackEnabled, (Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('DeviceWritebackEnabled')), $passwordFeatureFallback)
        UnifiedGroupWritebackEnabled                 = Get-ArrayaAssessmentPipelineFirstPopulatedValue @($summary.UnifiedGroupWritebackEnabled, (Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('UnifiedGroupWritebackEnabled')), $passwordFeatureFallback)
        UserWritebackEnabled                         = Get-ArrayaAssessmentPipelineFirstPopulatedValue @($summary.UserWritebackEnabled, (Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('UserWritebackEnabled')), $passwordFeatureFallback)
        OnPremisesSyncEnabled                        = Get-ArrayaAssessmentPipelineFirstPopulatedValue @($summary.OnPremisesSyncEnabled, (Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('OnPremisesSyncEnabled')))
        OnPremisesLastSyncDateTime                   = Get-ArrayaAssessmentPipelineFirstPopulatedValue @($summary.OnPremisesLastSyncDateTime, (Get-ArrayaObjectValue -Object $ExistingPasswordLifecycleSummary -Names @('OnPremisesLastSyncDateTime')))
        FeatureCollectionNote                        = $syncFeatureCollectionNote
    }

    if (Get-Command Get-AzureADConnectHealthSyncServices -ErrorAction Ignore) {
        try {
            $services = Get-AzureADConnectHealthSyncServices -ErrorAction SilentlyContinue
            foreach ($svc in $services) {
                $serviceDetails += [pscustomobject]@{
                    ServiceName  = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $svc -Names @('ServiceName', 'Name')
                    ServerName   = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $svc -Names @('ServerName', 'Server')
                    LastSyncTime = Get-ArrayaAssessmentPipelineNamedPropertyValue -Object $svc -Names @('LastSyncTime', 'LastSyncDateTime')
                }
            }
        }
        catch {}
    }

    if (Get-Command Get-AzureADConnectHealthSyncErrors -ErrorAction Ignore) {
        try {
            $syncErrors = @(Get-AzureADConnectHealthSyncErrors -ErrorAction SilentlyContinue)
        }
        catch {}
    }
    elseif (Get-Command Get-AzureADConnectHealthSyncAlert -ErrorAction Ignore) {
        try {
            $syncErrors = @(Get-AzureADConnectHealthSyncAlert -ErrorAction SilentlyContinue)
        }
        catch {}
    }

    return [pscustomobject]@{
        AdConnectConfiguration = [ordered]@{
            Summary      = $summary
            SyncServices = $serviceDetails
            RecentErrors = $syncErrors | Select-Object -First 10
            ErrorCount   = ($syncErrors | Measure-Object).Count
        }
        PasswordLifecycleSummary = $passwordLifecycleSummary
    }
}

function Invoke-ArrayaAssessmentPipelineStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$Index,
        [Parameter(Mandatory = $true)]
        [int]$Total,
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $percent = [math]::Round(($Index / [math]::Max($Total, 1)) * 100, 2)
    $prefix = "  [{0}/{1} | {2}%] {3}" -f $Index, $Total, $percent, $Name
    Write-Host $prefix -ForegroundColor White
    $stepStart = Get-Date
    $result = & $ScriptBlock
    $duration = (Get-Date) - $stepStart
    Write-Host ("{0} - Completed in {1}" -f $prefix, $duration.ToString('hh\:mm\:ss')) -ForegroundColor DarkGray
    return $result
}

function Invoke-ArrayaAssessmentPipelineProfileAwareStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$Index,
        [Parameter(Mandatory = $true)]
        [int]$Total,
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [bool]$Enabled,
        [Parameter(Mandatory = $false)]
        [string]$SkipReason = 'Not required for this run.',
        [Parameter(Mandatory = $false)]
        [scriptblock]$ScriptBlock
    )

    $percent = [math]::Round(($Index / [math]::Max($Total, 1)) * 100, 2)
    $prefix = "  [{0}/{1} | {2}%] {3}" -f $Index, $Total, $percent, $Name
    Write-Host $prefix -ForegroundColor White
    $stepStart = Get-Date

    if (-not $Enabled) {
        $duration = (Get-Date) - $stepStart
        $message = if ([string]::IsNullOrWhiteSpace($SkipReason)) { 'Not required for this run.' } else { $SkipReason }
        Write-Host ("{0} - Skipped in {1} - {2}" -f $prefix, $duration.ToString('hh\:mm\:ss'), $message) -ForegroundColor DarkGray
        return [pscustomobject]@{
            Status  = 'Skipped'
            Message = $message
        }
    }

    if ($ScriptBlock) {
        $result = & $ScriptBlock
    }
    else {
        $result = $null
    }

    $duration = (Get-Date) - $stepStart
    Write-Host ("{0} - Completed in {1}" -f $prefix, $duration.ToString('hh\:mm\:ss')) -ForegroundColor DarkGray
    return $result
}

function Invoke-ArrayaAssessmentTenantOverviewPhase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,
        [Parameter(Mandatory = $false)]
        [string]$InputSnapshotPath
    )

    $inputSnapshot = if (-not [string]::IsNullOrWhiteSpace($InputSnapshotPath)) {
        Import-ArrayaTenantSnapshot -Path $InputSnapshotPath -SkipValidation
    }
    else {
        $null
    }

    $tenantStatsHash = if ($inputSnapshot) {
        Convert-ArrayaSnapshotToLegacyTenantStatsHash -Snapshot $inputSnapshot
    }
    else {
        @{}
    }

    Write-Host ("Output profile: {0} (scope: {1})" -f $Context.OutputProfileLabel, ([string]$Context.ReportingMode).ToLowerInvariant()) -ForegroundColor DarkGray
    Write-Host 'Microsoft 365 Tenant Assessment' -ForegroundColor White
    Write-Host ''
    Write-Host '[1/6] Tenant Overview' -ForegroundColor White
    Write-Host ('-' * 72) -ForegroundColor DarkGray

    $tenantInfo = Invoke-ArrayaAssessmentPipelineStep -Index 1 -Total 3 -Name 'Tenant overview' -ScriptBlock {
        Get-ArrayaAssessmentPipelineTenantOverviewInfo
    }
    $tenantStatsHash['TenantInfo'] = $tenantInfo

    $licenseSkus = Invoke-ArrayaAssessmentPipelineStep -Index 2 -Total 3 -Name 'License SKUs' -ScriptBlock {
        Get-ArrayaAssessmentPipelineLicenseSkus
    }
    $tenantStatsHash['LicenseSKUs'] = $licenseSkus

    $adConnectResult = Invoke-ArrayaAssessmentPipelineStep -Index 3 -Total 3 -Name 'AD Connect sync details' -ScriptBlock {
        Get-ArrayaAssessmentPipelineAdConnectSyncDetails -TenantInfo $tenantInfo -ExistingPasswordLifecycleSummary (Get-ArrayaObjectValue -Object $tenantStatsHash -Names @('PasswordLifecycleSummary'))
    }
    $tenantStatsHash['AdConnectConfiguration'] = $adConnectResult.AdConnectConfiguration
    $tenantStatsHash['PasswordLifecycleSummary'] = $adConnectResult.PasswordLifecycleSummary

    $metadata = [ordered]@{}
    $collectionPlan = [ordered]@{}
    $diagnostics = [ordered]@{}
    if ($inputSnapshot) {
        if ($inputSnapshot.Contains('Metadata') -and ($inputSnapshot['Metadata'] -is [System.Collections.IDictionary])) {
            foreach ($entry in $inputSnapshot['Metadata'].GetEnumerator()) {
                $metadata[[string]$entry.Key] = $entry.Value
            }
        }
        if ($inputSnapshot.Contains('CollectionPlan') -and ($inputSnapshot['CollectionPlan'] -is [System.Collections.IDictionary])) {
            foreach ($entry in $inputSnapshot['CollectionPlan'].GetEnumerator()) {
                $collectionPlan[[string]$entry.Key] = $entry.Value
            }
        }
        if ($inputSnapshot.Contains('Diagnostics') -and ($inputSnapshot['Diagnostics'] -is [System.Collections.IDictionary])) {
            foreach ($entry in $inputSnapshot['Diagnostics'].GetEnumerator()) {
                $diagnostics[[string]$entry.Key] = $entry.Value
            }
        }
    }
    $metadata['GeneratedAt'] = (Get-Date).ToString('o')
    $metadata['OutputProfile'] = $Context.OutputProfile
    $metadata['OutputProfileLabel'] = $Context.OutputProfileLabel
    $metadata['ReportingMode'] = $Context.ReportingMode

    $snapshot = Convert-ArrayaLegacyTenantStatsToSnapshot -TenantStatsHash $tenantStatsHash -Metadata $metadata -CollectionPlan $collectionPlan -Diagnostics $diagnostics
    $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'TenantOverview'
    Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $checkpointPath

    return [pscustomobject]@{
        Phase          = 'TenantOverview'
        CheckpointPath = $checkpointPath
    }
}
