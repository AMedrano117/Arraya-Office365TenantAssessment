# Improvement Plan Rule Taxonomy

Use this guide to understand what the rule names in the `Improve` output mean, where they come from, and how to explain them to an operator or customer.

## What A Rule Is

Each row in the improvement plan is a normalized finding with:

- `RuleId`: the internal identifier used by this repo
- `Source`: where the finding came from
- `Area`: the workstream or workload area
- `Category`: the subtopic inside the workstream
- `RelatedWorksheet`: the supporting dataset in the assessment snapshot/workbook

The rule names are mostly internal to this assessment solution. They are not usually Microsoft-native control IDs.

## Where Rules Come From

The improvement plan combines two main sources:

### 1. Derived assessment findings

These come from the assessment's own derived model:

- `Derived.BestPracticeFindings`
- `Derived.Findings`
- `Derived.BestPractices`

You will usually see these in the output with sources such as:

- `Derived/BestPracticeFindings`
- `Derived/Findings`
- `Summary/BestPractices`

These findings are built from Microsoft 365 data collected by the assessment, then normalized into customer/operator-friendly findings.

### 2. Heuristic improvement rules

These are explicit built-in rules in the `Improve` workflow. They use collected snapshot data and summary objects to create operator-focused findings that may not already exist in the derived assessment layer.

Examples:

- `ID-007`: high-privilege enterprise applications
- `CA-012`: Conditional Access exclusion sprawl
- `DEV-006`: unsupported operating systems
- `EX-007`: shared mailbox ownership/growth governance

## Rule Prefixes

These prefixes are internal taxonomy labels used by this repo.

- `SEC-*`: security configuration and Secure Score posture
- `CA-*`: Conditional Access quality and coverage
- `MFA-*`: MFA registration or adoption
- `ID-*`: identity governance
- `ADMIN-*`: privileged admin posture
- `DOMAIN-*`: domain hygiene
- `LIC-*`: licensing optimization
- `DEV-*`: device / endpoint posture
- `EX-*`: Exchange hygiene
- `COL-*`: SharePoint / OneDrive governance
- `TM-*`: Teams / M365 Groups governance
- `DIAG-*`: collection diagnostics

Hashed IDs such as `AREA-...`, `SECURESCORE-...`, `DOMAINS-...`, or `TEAMSCOLLABORATION-...` are normalized IDs generated for derived findings. They are not hand-authored rule IDs.

## How To Tell Whether A Rule Is Microsoft-Backed Or Repo-Built

Use the `Source` column.

- `Derived/...`
  - assessment-derived finding based on collected Microsoft data and repo logic
- `Summary/...`
  - repo-built rule using computed summary tables
- `Hybrid/...`
  - repo-built rule using collected tenant workload data directly
- `Heuristic`
  - generic repo-built fallback rule

Practical interpretation:

- Microsoft provides the raw service data
- this repo evaluates that data against guidance and thresholds
- the rule ID, wording, severity, and recommendation are produced by this repo

## Examples

### Derived-oriented findings

- `SECURESCORE-...`
  - normalized finding built from Secure Score-derived assessment output
- `AREA-...`
  - area summary generated from `Derived.BestPractices`

### Built-in improvement rules

- `ID-005`
  - inactive guest accounts found from `GuestSignInSummary`
- `ID-007`
  - enterprise apps with broad/high-privilege permissions from `EnterpriseApplications`
- `CA-002`
  - report-only Conditional Access policies
- `CA-010`
  - compliant-device requirement not detected in Conditional Access summary
- `CA-012`
  - many Conditional Access exclusions detected
- `EX-007`
  - shared mailboxes lacking owner signals or showing growth risk
- `DEV-005`
  - meaningful unmanaged-device population
- `DEV-006`
  - unsupported operating-system versions detected
- `COL-005`
  - external sharing detected in SharePoint posture summary
- `COL-006`
  - anonymous link defaults detected in SharePoint posture summary

## How Operators Should Read The Plan

For each finding, use this order:

1. `Severity` and `PriorityBand`
2. `Area` and `OwnerTeam`
3. `Finding`
4. `CurrentValue`
5. `TargetValue`
6. `RelatedWorksheet`
7. `Source`

This lets the operator answer:

- how urgent is this?
- which team should own it?
- what evidence supports it?
- is this derived from broader assessment logic or from a specific built-in rule?

## How To Explain It To A Customer

Use this wording:

- the assessment collects Microsoft 365 tenant data
- the tool then applies built-in analysis and best-practice logic
- some findings come from the assessment's best-practice model
- some findings come from targeted remediation rules added for operator planning

Avoid saying that every rule is a Microsoft-native rule. That would be misleading.

## Related Operator Docs

- [RUN.md](../../RUN.md)
- [Operator Guide](../../RUN.md)
- [Workflow Overview](../architecture/workflow-overview.md)
