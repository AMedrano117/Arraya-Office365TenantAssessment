function Get-ArrayaMigrationCriteriaCatalog {
    <#
    .SYNOPSIS
        Loads the migration evidence criteria catalog.

    .DESCRIPTION
        Reads src/config/baseline/migration-evidence-criteria.json, or the file given by -Path,
        and validates that every criterion carries the fields the register needs. A catalog
        that fails validation throws, because a silently incomplete catalog would hide
        evidence requirements.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        # PSScriptRoot is <repo>/src/modules/Arraya.M365.Common/Public
        $repoRoot = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\..'))
        $Path = Join-Path -Path $repoRoot -ChildPath 'src\config\baseline\migration-evidence-criteria.json'
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Migration evidence criteria catalog not found: $Path"
    }

    $catalog = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    $requiredFields = @(
        'CriterionId', 'Profile', 'SemanticKey', 'Criterion', 'Workloads', 'Requirement', 'Phase', 'Scope',
        'Population', 'Period', 'Meaning', 'Applicability', 'Coverage', 'Evidence', 'MissingEvidenceBehavior',
        'Impact', 'ResolutionEvidence', 'ResponsibleRole'
    )
    $allowed = @{
        Requirement             = @('Required', 'Optional')
        Phase                   = @('Scoping', 'Design', 'Execution')
        Applicability           = @('Always', 'WhenDirectorySynchronized')
        Coverage                = @('Supported', 'Unsupported')
        MissingEvidenceBehavior = @('PresalesBlocker', 'NeedsData', 'PlanningGap', 'Nonblocking')
        Impact                  = @('Readiness', 'Planning', 'Reference')
    }
    $allowedRules = @('Dataset', 'DirectorySync', 'ExchangeHybrid', 'MailboxHolds', 'MailboxDelegates', 'MailboxSizing', 'Unsupported')

    $problems = New-Object System.Collections.Generic.List[string]
    $seenIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($criterion in @($catalog.Criteria)) {
        $id = [string]$criterion.CriterionId
        if (-not $seenIds.Add($id)) { $problems.Add("Duplicate CriterionId '$id'.") | Out-Null }
        foreach ($field in $requiredFields) {
            if ($null -eq $criterion.PSObject.Properties[$field] -or $null -eq $criterion.$field -or ($criterion.$field -is [string] -and [string]::IsNullOrWhiteSpace($criterion.$field))) {
                $problems.Add("Criterion '$id' is missing '$field'.") | Out-Null
            }
        }
        foreach ($field in $allowed.Keys) {
            if ($null -ne $criterion.PSObject.Properties[$field] -and $allowed[$field] -notcontains [string]$criterion.$field) {
                $problems.Add("Criterion '$id' has an unsupported $field '$($criterion.$field)'.") | Out-Null
            }
        }
        if ($criterion.Evidence -and $allowedRules -notcontains [string]$criterion.Evidence.Rule) {
            $problems.Add("Criterion '$id' has an unsupported evidence rule '$($criterion.Evidence.Rule)'.") | Out-Null
        }
        if ([string]$criterion.MissingEvidenceBehavior -eq 'PresalesBlocker') {
            if ([string]$criterion.Profile -ne 'Presales') { $problems.Add("Criterion '$id' uses PresalesBlocker but is not owned by Presales.") | Out-Null }
            foreach ($field in @('RequiredDecision', 'BlockerImpact')) {
                if ([string]::IsNullOrWhiteSpace([string]$criterion.$field)) { $problems.Add("PresalesBlocker criterion '$id' must state '$field'.") | Out-Null }
            }
        }
        if ([string]$criterion.Coverage -eq 'Unsupported' -and [string]$criterion.Evidence.Rule -ne 'Unsupported') {
            $problems.Add("Criterion '$id' is Unsupported but names evidence rule '$($criterion.Evidence.Rule)'.") | Out-Null
        }
    }

    if ($problems.Count -gt 0) {
        throw ("Migration evidence criteria catalog '{0}' is invalid: {1}" -f $Path, ($problems.ToArray() -join ' '))
    }

    return $catalog
}
