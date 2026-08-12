<#
.SYNOPSIS
Merges completed Presales migration-scope decisions into a copy of an assessment snapshot.

.DESCRIPTION
Reads the MigrationScopeDecisions worksheet from a Presales workbook, or equivalent CSV,
validates its DecisionKey and Status values, and writes the normalized rows to
Derived.MigrationScopeDecisions in a new assessment snapshot. The source snapshot is never
overwritten. Use the output snapshot with Start-M365TenantAssessment.ps1 -Action Report so
the readiness, effort, and wave registers are rebuilt from the completed decisions.

.PARAMETER AssessmentJsonPath
Path to the source assessment snapshot JSON.

.PARAMETER DecisionInputPath
Path to the completed Presales XLSX workbook or a CSV containing the decision worksheet.

.PARAMETER OutputPath
Path for the merged snapshot. It must differ from AssessmentJsonPath.

.PARAMETER WorksheetName
Worksheet to import from XLSX. Defaults to MigrationScopeDecisions.

.PARAMETER Force
Allows an existing OutputPath to be replaced. The source snapshot still cannot be replaced.

.EXAMPLE
./src/scripts/operations/Merge-M365MigrationScopeDecisions.ps1 `
    -AssessmentJsonPath './Support/Contoso-Snap.json' `
    -DecisionInputPath './Contoso.xlsx' `
    -OutputPath './Support/Contoso-Snap-WithDecisions.json'

.EXAMPLE
./src/scripts/operations/Merge-M365MigrationScopeDecisions.ps1 `
    -AssessmentJsonPath './Support/Contoso-Snap.json' `
    -DecisionInputPath './MigrationScopeDecisions.csv' `
    -OutputPath './Support/Contoso-Snap-WithDecisions.json' `
    -WhatIf

.OUTPUTS
PSCustomObject describing the validated merge and whether the output snapshot was written.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [string]$AssessmentJsonPath,

    [Parameter(Mandatory = $true)]
    [Alias('CompletedWorkbookPath', 'CompletedDecisionsPath')]
    [string]$DecisionInputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$WorksheetName = 'MigrationScopeDecisions',

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-MigrationDecisionInputFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label was not found: $Path"
    }

    return (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
}

function Get-MigrationDecisionProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Row,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $Row) { return $null }
    if ($Row -is [System.Collections.IDictionary]) {
        foreach ($key in $Row.Keys) {
            if ([string]::Equals([string]$key, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $Row[$key]
            }
        }
        return $null
    }

    $property = $Row.PSObject.Properties |
        Where-Object { [string]::Equals($_.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase) } |
        Select-Object -First 1
    if ($property) { return $property.Value }
    return $null
}

function Test-MigrationDecisionProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Row,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $Row) { return $false }
    if ($Row -is [System.Collections.IDictionary]) {
        foreach ($key in $Row.Keys) {
            if ([string]::Equals([string]$key, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }
        }
        return $false
    }

    return [bool]($Row.PSObject.Properties |
        Where-Object { [string]::Equals($_.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase) } |
        Select-Object -First 1)
}

function ConvertTo-MigrationDecisionText {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)]$Value)

    if ($null -eq $Value) { return $null }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    return $text.Trim()
}

function ConvertTo-MigrationDecisionScopeValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Value,
        [Parameter(Mandatory = $true)]
        [int]$RowNumber
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    if ($Value -is [bool]) { return [bool]$Value }
    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
        if ([double]$Value -eq 1) { return $true }
        if ([double]$Value -eq 0) { return $false }
    }

    $normalized = ([string]$Value).Trim().ToLowerInvariant() -replace '\s+', ' '
    if ($normalized -in @('true', 'yes', 'y', '1', 'include', 'included', 'in scope', 'in-scope')) { return $true }
    if ($normalized -in @('false', 'no', 'n', '0', 'exclude', 'excluded', 'out of scope', 'out-of-scope')) { return $false }
    if ($normalized -in @('n/a', 'na', 'unknown', 'undecided', 'not decided')) { return $null }

    throw "Decision input row $RowNumber has invalid InScope value '$Value'. Use Yes/No, True/False, Include/Exclude, 1/0, or leave it blank."
}

function ConvertTo-MigrationDecisionStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        $Value,
        [Parameter(Mandatory = $true)]
        [int]$RowNumber
    )

    $status = ConvertTo-MigrationDecisionText -Value $Value
    if (-not $status) {
        throw "Decision input row $RowNumber has a blank Status."
    }

    $statusMap = @{
        'needs input'        = 'Needs Input'
        'needs confirmation' = 'Needs Confirmation'
        'needs data'         = 'Needs Data'
        'not started'        = 'Not Started'
        'open'               = 'Open'
        'pending'            = 'Pending'
        'draft'              = 'Draft'
        'in review'          = 'In Review'
        'blocked'            = 'Blocked'
        'deferred'           = 'Deferred'
        'rejected'           = 'Rejected'
        'confirmed'          = 'Confirmed'
        'approved'           = 'Approved'
        'complete'           = 'Complete'
        'completed'          = 'Completed'
        'resolved'           = 'Resolved'
    }
    $normalized = ($status.ToLowerInvariant() -replace '\s+', ' ').Trim()
    if (-not $statusMap.ContainsKey($normalized)) {
        $allowed = @($statusMap.Values | Sort-Object -Unique) -join ', '
        throw "Decision input row $RowNumber has invalid Status '$status'. Allowed values: $allowed."
    }

    return $statusMap[$normalized]
}

$resolvedAssessmentPath = Resolve-MigrationDecisionInputFile -Path $AssessmentJsonPath -Label 'Assessment snapshot'
$resolvedDecisionInputPath = Resolve-MigrationDecisionInputFile -Path $DecisionInputPath -Label 'Decision input'
$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)

if ([System.IO.Path]::GetExtension($resolvedOutputPath) -ne '.json') {
    throw "OutputPath must use the .json extension: $resolvedOutputPath"
}
if ([string]::Equals($resolvedAssessmentPath, $resolvedOutputPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'OutputPath must differ from AssessmentJsonPath. The source snapshot is never overwritten.'
}
if ([string]::Equals($resolvedDecisionInputPath, $resolvedOutputPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'OutputPath must differ from DecisionInputPath. The completed workbook or CSV is never overwritten.'
}
if ((Test-Path -LiteralPath $resolvedOutputPath -PathType Leaf) -and -not $Force) {
    throw "OutputPath already exists. Use -Force to replace it: $resolvedOutputPath"
}

$extension = [System.IO.Path]::GetExtension($resolvedDecisionInputPath).ToLowerInvariant()
$decisionRows = @()
switch ($extension) {
    '.csv' {
        $decisionRows = @(Import-Csv -LiteralPath $resolvedDecisionInputPath)
    }
    '.xlsx' {
        if (-not (Get-Command -Name Import-Excel -ErrorAction SilentlyContinue)) {
            if (Get-Module -ListAvailable -Name ImportExcel) {
                Import-Module ImportExcel -ErrorAction Stop
            }
            else {
                throw 'ImportExcel is required to read an XLSX decision workbook. Install ImportExcel or export the MigrationScopeDecisions worksheet as CSV.'
            }
        }
        try {
            $decisionRows = @(Import-Excel -Path $resolvedDecisionInputPath -WorksheetName $WorksheetName -ErrorAction Stop)
        }
        catch {
            throw "Unable to import worksheet '$WorksheetName' from '$resolvedDecisionInputPath': $($_.Exception.Message)"
        }
    }
    default {
        throw "Unsupported decision input extension '$extension'. Use .xlsx or .csv."
    }
}

if ($decisionRows.Count -eq 0) {
    throw "Decision input contains no rows: $resolvedDecisionInputPath"
}
if (-not (Test-MigrationDecisionProperty -Row $decisionRows[0] -Name 'DecisionKey')) {
    throw "Decision input is missing the required DecisionKey column: $resolvedDecisionInputPath"
}
if (-not (Test-MigrationDecisionProperty -Row $decisionRows[0] -Name 'Status')) {
    throw "Decision input is missing the required Status column: $resolvedDecisionInputPath"
}

$normalizedDecisionRows = New-Object System.Collections.Generic.List[object]
$seenDecisionKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$confirmedStatuses = @('Confirmed', 'Approved', 'Complete', 'Completed', 'Resolved')
$confirmedCount = 0
$inputRowNumber = 1
foreach ($decisionRow in $decisionRows) {
    $inputRowNumber++ # Header is row 1 in both the worksheet and CSV.
    $decisionKey = ConvertTo-MigrationDecisionText -Value (Get-MigrationDecisionProperty -Row $decisionRow -Name 'DecisionKey')
    if (-not $decisionKey) {
        throw "Decision input row $inputRowNumber has a blank DecisionKey."
    }
    if (-not $seenDecisionKeys.Add($decisionKey)) {
        throw "Decision input contains duplicate DecisionKey '$decisionKey' (comparison is case-insensitive)."
    }

    $status = ConvertTo-MigrationDecisionStatus -Value (Get-MigrationDecisionProperty -Row $decisionRow -Name 'Status') -RowNumber $inputRowNumber
    if ($status -in $confirmedStatuses) { $confirmedCount++ }

    $normalizedDecisionRows.Add([pscustomobject][ordered]@{
            DecisionKey            = $decisionKey
            DecisionPrompt         = ConvertTo-MigrationDecisionText -Value (Get-MigrationDecisionProperty -Row $decisionRow -Name 'DecisionPrompt')
            CustomerConfirmedValue = ConvertTo-MigrationDecisionText -Value (Get-MigrationDecisionProperty -Row $decisionRow -Name 'CustomerConfirmedValue')
            Status                 = $status
            InScope                = ConvertTo-MigrationDecisionScopeValue -Value (Get-MigrationDecisionProperty -Row $decisionRow -Name 'InScope') -RowNumber $inputRowNumber
            TargetMapping          = ConvertTo-MigrationDecisionText -Value (Get-MigrationDecisionProperty -Row $decisionRow -Name 'TargetMapping')
            MigrationTool          = ConvertTo-MigrationDecisionText -Value (Get-MigrationDecisionProperty -Row $decisionRow -Name 'MigrationTool')
            DecisionOwner          = ConvertTo-MigrationDecisionText -Value (Get-MigrationDecisionProperty -Row $decisionRow -Name 'DecisionOwner')
            DueDate                = ConvertTo-MigrationDecisionText -Value (Get-MigrationDecisionProperty -Row $decisionRow -Name 'DueDate')
            Notes                  = ConvertTo-MigrationDecisionText -Value (Get-MigrationDecisionProperty -Row $decisionRow -Name 'Notes')
        }) | Out-Null
}

$resolveRepoRootHelperPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\shared\Resolve-ArrayaRepoRoot.ps1'))
if (-not (Test-Path -LiteralPath $resolveRepoRootHelperPath -PathType Leaf)) {
    throw "Repo-root helper script not found: $resolveRepoRootHelperPath"
}
. $resolveRepoRootHelperPath

$repoRoot = Resolve-ArrayaRepoRoot -StartPath $PSScriptRoot
$commonManifestPath = Join-Path $repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
if (-not (Test-Path -LiteralPath $commonManifestPath -PathType Leaf)) {
    throw "Common module manifest not found: $commonManifestPath"
}
Import-Module -Name $commonManifestPath -Force -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop

$snapshot = Import-ArrayaTenantSnapshot -Path $resolvedAssessmentPath
if ($null -eq $snapshot) {
    throw "Assessment snapshot could not be loaded: $resolvedAssessmentPath"
}
if (-not $snapshot.Contains('Derived') -or -not ($snapshot['Derived'] -is [System.Collections.IDictionary])) {
    throw "Assessment snapshot does not contain a writable Derived section: $resolvedAssessmentPath"
}

$written = $false
if ($PSCmdlet.ShouldProcess($resolvedOutputPath, "write snapshot copy with $($normalizedDecisionRows.Count) migration-scope decision row(s)")) {
    $snapshot['Derived']['MigrationScopeDecisions'] = @($normalizedDecisionRows.ToArray())
    Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $resolvedOutputPath
    $written = $true
}

[pscustomobject]@{
    SourceSnapshotPath = $resolvedAssessmentPath
    DecisionInputPath  = $resolvedDecisionInputPath
    WorksheetName      = $(if ($extension -eq '.xlsx') { $WorksheetName } else { $null })
    OutputPath         = $resolvedOutputPath
    DecisionCount      = $normalizedDecisionRows.Count
    ConfirmedCount     = $confirmedCount
    Written            = $written
}
