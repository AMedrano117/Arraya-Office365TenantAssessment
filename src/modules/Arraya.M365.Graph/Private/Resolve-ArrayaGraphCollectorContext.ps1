function Get-ArrayaGraphCollectionDepthPolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('Minimum', 'Operator', 'Combined', 'Automation', 'All', 'Geek')]
        [string]$ReportingMode = 'Minimum'
    )

    $mode = $ReportingMode.ToLowerInvariant()
    if ($mode -eq 'combined') { $mode = 'operator' }
    $isMinimum = $mode -eq 'minimum'
    $isOperator = $mode -eq 'operator'
    $isAutomation = $mode -eq 'automation'
    $isAll = $mode -eq 'all'
    $isGeek = $mode -eq 'geek'

    return [PSCustomObject]@{
        ReportingMode                  = (Get-Culture).TextInfo.ToTitleCase($mode)
        CollectEntraGroupDeepDetails   = (-not $isMinimum)
        CollectEntraGroupMemberCounts  = (-not $isMinimum)
        CollectEntraGroupOwnerCounts   = (-not $isMinimum)
        CollectEntraGroupLicenseChecks = (-not $isMinimum)
        IsMinimum                      = $isMinimum
        IsOperator                     = $isOperator
        IsCombined                     = $isOperator
        IsAutomation                   = $isAutomation
        IsAll                          = $isAll
        IsGeek                         = $isGeek
    }
}

function Resolve-ArrayaGraphCollectorContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Context,
        [Parameter(Mandatory = $false)]
        [string]$DetailLevel = 'minimum'
    )

    if ($null -eq $Context) {
        $Context = New-ArrayaAssessmentContext -ReportingMode ((Get-Culture).TextInfo.ToTitleCase($DetailLevel.ToLowerInvariant()))
    }

    if (-not ($Context.TenantStats -is [System.Collections.IDictionary])) {
        $Context | Add-Member -NotePropertyName TenantStats -NotePropertyValue ([ordered]@{}) -Force
    }
    if (-not ($Context.Policies -is [System.Collections.IDictionary])) {
        $Context | Add-Member -NotePropertyName Policies -NotePropertyValue ([ordered]@{}) -Force
    }
    if (-not ($Context.Runtime -is [System.Collections.IDictionary])) {
        $Context | Add-Member -NotePropertyName Runtime -NotePropertyValue ([ordered]@{}) -Force
    }
    if (-not ($Context.Metadata -is [System.Collections.IDictionary])) {
        $Context | Add-Member -NotePropertyName Metadata -NotePropertyValue ([ordered]@{}) -Force
    }

    $resolvedMode = if ($Context.Policies.Contains('ReportingMode')) {
        [string]$Context.Policies['ReportingMode']
    }
    else {
        (Get-Culture).TextInfo.ToTitleCase($DetailLevel.ToLowerInvariant())
    }
    if ($resolvedMode -ieq 'Combined') {
        $resolvedMode = 'Operator'
    }
    $Context.Policies['ReportingMode'] = $resolvedMode

    if (-not $Context.Policies.Contains('CollectionDepth')) {
        $Context.Policies['CollectionDepth'] = Get-ArrayaGraphCollectionDepthPolicy -ReportingMode $resolvedMode
    }
    if (-not $Context.Metadata.Contains('StartedAt')) {
        $Context.Metadata['StartedAt'] = Get-Date
    }

    return $Context
}

function Get-ArrayaEntraGroupClassification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $GroupDetails
    )

    $membershipType = if ($GroupDetails.groupTypes -contains 'DynamicMembership') { 'Dynamic Group' } else { 'Assigned' }
    $membershipRule = if ($membershipType -eq 'Dynamic Group') { $GroupDetails.membershipRule } else { $null }

    if ($GroupDetails.groupTypes -contains 'Unified') {
        $groupType = 'Microsoft 365'
    }
    elseif ($GroupDetails.mailEnabled -eq $true -and $GroupDetails.securityEnabled -eq $true) {
        $groupType = 'Mail-enabled Security Group'
    }
    elseif ($GroupDetails.mailEnabled -eq $false -and $GroupDetails.securityEnabled -eq $true) {
        $groupType = 'Security Group'
    }
    elseif ($GroupDetails.mailEnabled -eq $true -and $GroupDetails.securityEnabled -eq $false) {
        $groupType = 'Distribution Group'
    }
    else {
        $groupType = 'Unknown'
    }

    return [PSCustomObject]@{
        MembershipType = $membershipType
        MembershipRule = $membershipRule
        GroupType      = $groupType
        Source         = if ($GroupDetails.onPremisesSyncEnabled) { 'On-Premises' } else { 'Cloud' }
    }
}

function Convert-ArrayaGraphBatchCountValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $Body
    )

    if ($null -eq $Body) {
        return $null
    }

    if ($Body -is [int] -or $Body -is [long] -or $Body -is [double] -or $Body -is [decimal]) {
        try {
            return [int]$Body
        }
        catch {
            return $null
        }
    }

    if ($Body -is [string]) {
        $parsedValue = 0
        if ([int]::TryParse($Body, [ref]$parsedValue)) {
            return $parsedValue
        }
    }

    foreach ($propertyName in @('value', '@odata.count')) {
        if ($Body.PSObject.Properties[$propertyName]) {
            try {
                return [int]$Body.PSObject.Properties[$propertyName].Value
            }
            catch {
            }
        }
    }

    return $null
}

function Invoke-ArrayaGraphBatchRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Requests,
        [Parameter(Mandatory = $false)]
        [string[]]$GraphAuthType = @('REST'),
        [Parameter(Mandatory = $false)]
        [string]$Activity = 'Graph batch request',
        [Parameter(Mandatory = $false)]
        [string]$ExportFileLocation
    )

    $result = @{}
    if (-not $Requests -or $Requests.Count -eq 0) {
        return $result
    }

    $batchEndpoint = 'https://graph.microsoft.com/v1.0/$batch'
    $chunkSize = 20
    for ($offset = 0; $offset -lt $Requests.Count; $offset += $chunkSize) {
        $chunk = @($Requests | Select-Object -Skip $offset -First $chunkSize)
        if ($chunk.Count -eq 0) {
            continue
        }

        $payload = @{ requests = $chunk } | ConvertTo-Json -Depth 10 -Compress
        $batchResponse = $null

        try {
            $canUseSdkBatch = (
                ($GraphAuthType -contains 'SDK') -and
                (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue) -and
                (Get-MgContext -ErrorAction SilentlyContinue)
            )
            if ($canUseSdkBatch) {
                $batchResponse = Invoke-QuietCommand -ScriptBlock {
                    Invoke-MgGraphRequest -Method POST -Uri $batchEndpoint -Body $payload -OutputType PSObject -ProgressAction SilentlyContinue -ErrorAction Stop
                }
            }
            else {
                $headers = @{}
                if ($global:GraphHeaders) {
                    $headers = $global:GraphHeaders.Clone()
                }
                elseif ($global:GraphToken) {
                    $headers = @{
                        'Content-Type'     = 'application/json'
                        'Authorization'    = "Bearer $global:GraphToken"
                        'ConsistencyLevel' = 'eventual'
                    }
                }
                if (-not $headers.ContainsKey('Content-Type')) {
                    $headers['Content-Type'] = 'application/json'
                }
                if (-not $headers.ContainsKey('ConsistencyLevel')) {
                    $headers['ConsistencyLevel'] = 'eventual'
                }

                $batchResponse = Invoke-RestMethod -Uri $batchEndpoint -Headers $headers -Method POST -Body $payload -ContentType 'application/json' -ErrorAction Stop
            }
        }
        catch {
            Write-Log -Type WARNING -Message "[Get-EntraIDGroups] $Activity failed for request chunk starting at index $offset. $($_.Exception.Message)" -ExportFileLocation $ExportFileLocation
            continue
        }

        foreach ($response in @($batchResponse.responses)) {
            if ($response -and $response.id) {
                $result[[string]$response.id] = $response
            }
        }
    }

    return $result
}
