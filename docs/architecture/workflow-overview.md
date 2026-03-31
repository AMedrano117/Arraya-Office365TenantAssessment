# Workflow Overview

This document explains the supported launcher actions in `src/scripts/operations/Start-M365TenantAssessment.ps1`, how they differ, and where they overlap.

## At A Glance

| Action | Primary purpose | Input source | Primary outputs | Uses output profiles? | Notes |
|---|---|---|---|---|---|
| `M365` | Full Microsoft 365 assessment run | Live tenant connections | Customer remediation HTML, engineer action pack, support artifacts, compatibility assessment artifacts when requested | Yes | Operator remediation-first collect + export path |
| `M365Collect` | Collect tenant data only | Live tenant connections | JSON snapshot, manifest | Yes | Same core engine as `M365`, but collection-only |
| `M365Export` | Rebuild artifacts from a saved snapshot | Existing tenant JSON snapshot | Customer remediation HTML, engineer action pack, support artifacts, compatibility assessment artifacts when requested | Yes | Same core engine as `M365`, but export-only |
| `AD` | Active Directory assessment | Live on-prem AD | AD workbook/export set | No | Separate legacy workflow for AD objects and infra |
| `Improve` | Build an improvement plan from one snapshot | Existing tenant JSON snapshot | Customer remediation HTML, engineer action pack, support JSON, support snippets | No | Snapshot post-processing workflow |
| `Compare` | Compare two snapshots over time | Two tenant JSON snapshots | Comparison JSON, CSV, Markdown | No | Snapshot post-processing workflow |

## Workflow Details

### `M365`

`M365` is the primary end-to-end workflow. It performs live tenant collection, builds derived reporting tables, and then runs the export pipeline that creates the human-readable and machine-readable artifacts selected by the output profile.

This action is profile-driven. The profile determines both collection depth and which artifacts should be generated. The merged-profile behavior is resolved once in the runner before the core assessment script is invoked.

### `M365Collect`

`M365Collect` is not a separate assessment engine. It routes to the same core tenant assessment script as `M365`, but enables collection-only mode and forces JSON output on while disabling workbook, HTML, questionnaire, and PDF generation.

Use this when the goal is to capture tenant state once and defer export/report generation until later.

### `M365Export`

`M365Export` is the export-only counterpart to `M365Collect`. It loads a previously saved tenant snapshot, validates it, converts it back into the in-memory structure expected by the legacy engine, and then runs the same export pipeline used by the full `M365` path.

Use this when collection has already been completed or when the team wants to regenerate artifacts under a different output profile without reconnecting to the tenant.

### `AD`

`AD` is a separate Active Directory assessment path. It inventories AD forests, domain controllers, FSMO roles, GPOs, users, groups, computers, and OUs, then exports those results.

This workflow is not part of the Microsoft 365 snapshot/export model and does not share the M365 output-profile system.

### `Improve`

`Improve` is a snapshot post-processing workflow. It reads a single tenant snapshot and produces an improvement plan in JSON, CSV, and Markdown, plus a PowerShell snippet file with remediation starting points.

It can also use Graph fallback calls when explicitly requested and when certain datasets are missing from the snapshot.

### `Compare`

`Compare` is another snapshot post-processing workflow. It reads a baseline snapshot and a current snapshot, recalculates a fixed set of posture metrics, and writes JSON, CSV, and Markdown comparison outputs.

This is intended for trend or before/after review rather than full artifact regeneration.

## Key Differences

### Input model

- `M365`, `M365Collect`, and `AD` use live connections.
- `M365Export`, `Improve`, and `Compare` use saved JSON snapshots.

### Output model

- `M365` and `M365Export` are artifact-generation workflows.
- `M365Collect` is a capture workflow.
- `Improve` and `Compare` are analytic workflows on top of snapshots.
- `AD` is its own assessment/export path.

### Output profile support

- Only the `M365` family uses `OutputProfile`.
- `AD`, `Improve`, and `Compare` have fixed behavior and fixed output types.

## Overlap Review

### 1. The `M365` family is one workflow with three execution modes

`M365`, `M365Collect`, and `M365Export` all route through the same runner logic and the same core legacy script. The real differences are mode flags and export behavior:

- `M365`: live collect + export
- `M365Collect`: live collect only
- `M365Export`: snapshot export only

This is the clearest overlap in the repo. From a maintainer perspective, these actions are better understood as stages or modes of one pipeline than as fully separate workflows.

### 2. `Improve` and `Compare` duplicate snapshot metric logic

Both snapshot workflows recalculate overlapping posture signals from the tenant snapshot, including:

- Secure Score
- Conditional Access counts
- Global administrator counts
- Unverified domain counts
- License utilization pressure
- Stale device posture

`Improve` uses those values to generate findings and recommendations. `Compare` uses them to classify change over time. The metric extraction logic is similar enough that it should be a shared module rather than duplicated script logic.

### 3. Snapshot-based actions share similar validation and path handling needs

`M365Export`, `Improve`, and `Compare` all depend on saved tenant snapshots and need:

- path validation
- snapshot validation
- output-folder defaults
- predictable naming

That shared surface area is a good target for consolidation even if the business logic stays separate.

## Refactor Opportunities

### Highest-value changes

1. Create a shared snapshot metrics module.
   - Centralize metric extraction for Secure Score, Conditional Access, admins, domains, licenses, devices, and related helpers.
   - Have `Improve` consume metrics plus rules.
   - Have `Compare` consume the same metrics plus comparison logic.

2. Reframe the `M365` family as pipeline stages.
   - Keep the user-facing actions for compatibility.
   - Internally model them as one workflow with explicit modes such as `Full`, `CollectOnly`, and `ExportOnly`.
   - This would make the control flow easier to reason about and reduce duplicate launcher/runner handling.

3. Centralize snapshot input and output conventions.
   - Shared helper for resolving snapshot paths, default output folders, and snapshot validation.
   - Reuse it across `M365Export`, `Improve`, and `Compare`.

### Lower-risk follow-up work

4. Introduce a common post-processing base for snapshot workflows.
   - Shared bootstrap for loading the common module, reading snapshots, and emitting JSON/CSV/Markdown outputs.

## Practical Recommendation

If the team wants to improve maintainability without changing user-facing behavior, the safest sequence is:

1. Extract shared snapshot metric helpers from `Improve` and `Compare`.
2. Extract shared snapshot bootstrap and output helpers for snapshot-driven actions.
3. Normalize the `M365` family into a single internal pipeline with three modes while keeping the existing launcher actions as compatibility aliases.

That order reduces duplication first, then clarifies workflow intent, without forcing a large user-facing redesign up front.
