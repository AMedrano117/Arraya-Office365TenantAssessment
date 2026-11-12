function Resolve-ArrayaMigrationCriteria {
    <#
    .SYNOPSIS
        Composes the evidence criteria for one or more output profiles.

    .DESCRIPTION
        Deterministic composition rules:

        - Profiles are processed in canonical order, and criteria in catalog order, so the
          same request in any order yields the same result.
        - Criteria are merged only when they are semantically identical: same SemanticKey,
          Scope, Population, Period, Meaning, evidence rule, decision, responsible role, and
          resolution evidence. Anything else stays a separate criterion, even with the same
          SemanticKey.
        - A merged criterion keeps its owning profile (the strictest requirement; ties go to
          the first profile in canonical order) and lists the contributing profiles.
        - The strictest rule wins: Required over Optional; PresalesBlocker over NeedsData over
          PlanningGap over Nonblocking; Readiness over Planning over Reference; and the union of
          evidence datasets and collector steps.
        - PresalesBlocker survives only when Presales is among the composed profiles; otherwise
          it becomes NeedsData, so another profile cannot inherit a Presales gate.
        - Evidence is never borrowed: each composed criterion is evaluated against its own
          datasets, so one profile cannot fill another profile's unrelated evidence gap.

        Profiles with no catalog entries contribute nothing.
    #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Returns the composed set of criteria for the requested profiles.')]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$OutputProfile,
        [Parameter(Mandatory = $false)]
        $Catalog
    )

    if ($null -eq $Catalog) { $Catalog = Get-ArrayaMigrationCriteriaCatalog }

    $canonicalOrder = @('Presales', 'TenantToTenantMigration', 'SolutionsEngineer', 'ExecutiveLevel', 'Geek', 'Machine')
    $profiles = @($canonicalOrder | Where-Object { $OutputProfile -contains $_ })
    $requirementRank = @{ Required = 2; Optional = 1 }
    $behaviorRank = @{ PresalesBlocker = 4; NeedsData = 3; PlanningGap = 2; Nonblocking = 1 }
    $impactRank = @{ Readiness = 3; Planning = 2; Reference = 1 }

    $toList = {
        param($Value)
        @($Value | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
    }
    $unionOrdered = {
        param([object[]]$First, [object[]]$Second)
        $result = New-Object System.Collections.Generic.List[string]
        foreach ($item in @($First) + @($Second)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$item) -and -not $result.Contains([string]$item)) { $result.Add([string]$item) | Out-Null }
        }
        @($result.ToArray())
    }

    $composed = New-Object System.Collections.Generic.List[object]
    $byIdentity = @{}
    foreach ($profileName in $profiles) {
        foreach ($criterion in @($Catalog.Criteria | Where-Object { [string]$_.Profile -eq $profileName })) {
            $identity = (@(
                    $criterion.SemanticKey, $criterion.Scope, $criterion.Population, $criterion.Period, $criterion.Meaning,
                    $criterion.Evidence.Rule, $criterion.DecisionKey, $criterion.ResponsibleRole, $criterion.ResolutionEvidence
                ) | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() }) -join '|'

            if (-not $byIdentity.ContainsKey($identity)) {
                $entry = [pscustomobject][ordered]@{
                    CriterionId             = [string]$criterion.CriterionId
                    SourceCriterionIds      = @([string]$criterion.CriterionId)
                    Profile                 = $profileName
                    ContributingProfiles    = @()
                    SemanticKey             = [string]$criterion.SemanticKey
                    Criterion               = [string]$criterion.Criterion
                    Workloads               = @(& $toList $criterion.Workloads)
                    Requirement             = [string]$criterion.Requirement
                    Phase                   = @([string]$criterion.Phase)
                    Scope                   = [string]$criterion.Scope
                    Population              = [string]$criterion.Population
                    Period                  = [string]$criterion.Period
                    Meaning                 = [string]$criterion.Meaning
                    Applicability           = [string]$criterion.Applicability
                    Coverage                = [string]$criterion.Coverage
                    EvidenceRule            = [string]$criterion.Evidence.Rule
                    Datasets                = @(& $toList $criterion.Evidence.Datasets)
                    CollectorSteps          = @(& $toList $criterion.Evidence.CollectorSteps)
                    EmptyIsUnexpected       = [bool]$criterion.Evidence.EmptyIsUnexpected
                    MissingEvidenceBehavior = [string]$criterion.MissingEvidenceBehavior
                    Impact                  = [string]$criterion.Impact
                    DecisionKey             = [string]$criterion.DecisionKey
                    RequiredDecision        = [string]$criterion.RequiredDecision
                    DecisionCategory        = [string]$criterion.DecisionCategory
                    BlockerImpact           = [string]$criterion.BlockerImpact
                    ResolutionEvidence      = [string]$criterion.ResolutionEvidence
                    ResponsibleRole         = [string]$criterion.ResponsibleRole
                    ComplexityFlagItem      = [string]$criterion.ComplexityFlagItem
                }
                $byIdentity[$identity] = $entry
                $composed.Add($entry) | Out-Null
                continue
            }

            $entry = $byIdentity[$identity]
            $incomingIsStricter = $requirementRank[[string]$criterion.Requirement] -gt $requirementRank[$entry.Requirement]
            if ($incomingIsStricter) {
                $entry.ContributingProfiles = @(& $unionOrdered @($entry.Profile) $entry.ContributingProfiles)
                $entry.Profile = $profileName
                $entry.CriterionId = [string]$criterion.CriterionId
                $entry.Criterion = [string]$criterion.Criterion
                $entry.Requirement = [string]$criterion.Requirement
            }
            elseif ($entry.Profile -ne $profileName -and $entry.ContributingProfiles -notcontains $profileName) {
                $entry.ContributingProfiles = @(& $unionOrdered $entry.ContributingProfiles @($profileName))
            }
            $entry.ContributingProfiles = @($canonicalOrder | Where-Object { $entry.ContributingProfiles -contains $_ -and $_ -ne $entry.Profile })
            $entry.SourceCriterionIds = @(& $unionOrdered $entry.SourceCriterionIds @([string]$criterion.CriterionId))
            $entry.Workloads = @(& $unionOrdered $entry.Workloads @(& $toList $criterion.Workloads))
            $entry.Phase = @(& $unionOrdered $entry.Phase @([string]$criterion.Phase))
            $entry.Datasets = @(& $unionOrdered $entry.Datasets @(& $toList $criterion.Evidence.Datasets))
            $entry.CollectorSteps = @(& $unionOrdered $entry.CollectorSteps @(& $toList $criterion.Evidence.CollectorSteps))
            $entry.EmptyIsUnexpected = $entry.EmptyIsUnexpected -or [bool]$criterion.Evidence.EmptyIsUnexpected
            if ($behaviorRank[[string]$criterion.MissingEvidenceBehavior] -gt $behaviorRank[$entry.MissingEvidenceBehavior]) {
                $entry.MissingEvidenceBehavior = [string]$criterion.MissingEvidenceBehavior
                $entry.BlockerImpact = [string]$criterion.BlockerImpact
                if (-not [string]::IsNullOrWhiteSpace([string]$criterion.ComplexityFlagItem)) { $entry.ComplexityFlagItem = [string]$criterion.ComplexityFlagItem }
            }
            if ($impactRank[[string]$criterion.Impact] -gt $impactRank[$entry.Impact]) { $entry.Impact = [string]$criterion.Impact }
        }
    }

    foreach ($entry in $composed) {
        $involvedProfiles = @($entry.Profile) + @($entry.ContributingProfiles)
        if ($entry.MissingEvidenceBehavior -eq 'PresalesBlocker' -and $involvedProfiles -notcontains 'Presales') {
            $entry.MissingEvidenceBehavior = 'NeedsData'
        }
    }

    return @($composed.ToArray())
}
