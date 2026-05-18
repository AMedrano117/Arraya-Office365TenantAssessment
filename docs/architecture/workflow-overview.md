# Workflow Overview

This document explains the supported launcher actions in `src/scripts/operations/Start-M365TenantAssessment.ps1`, how they differ, and where they overlap.

## At A Glance

| Action | Primary purpose | Input source | Primary outputs | Uses output profiles? | Notes |
|---|---|---|---|---|---|
| `Full` | Full Microsoft 365 assessment run | Live tenant connections | `Deliverables` human reports/workbook plus `Support` machine artifacts | Yes | Standard end-to-end operator path |
| `Preflight` | Validate connections and permissions only | Live tenant connections | Console status output only | No | Run before a long collection to catch auth gaps early |
| `Collect` | Collect tenant data only | Live tenant connections | `Support` JSON snapshot and manifest | Yes | Same core engine as `Full`, collection-only mode |
| `Report` | Rebuild artifacts from a saved snapshot | Existing tenant JSON snapshot | `Deliverables` human reports/workbook plus `Support` machine artifacts | Yes | Same core engine as `Full`, export-only mode |
| `AD` | Active Directory assessment | Live on-prem AD | AD workbook/export set | No | Separate legacy workflow for AD objects and infra |
| `Improve` | Build an improvement plan from one snapshot | Existing tenant JSON snapshot | `Deliverables` customer/roadmap/engineer reports plus `Support` JSON/snippets | No | Snapshot post-processing workflow |
| `Compare` | Compare two snapshots over time | Two tenant JSON snapshots | Comparison JSON, CSV, Markdown | No | Snapshot post-processing workflow |

## Workflow Details

### `Full`

`Full` is the primary end-to-end workflow. It performs live tenant collection, builds derived reporting tables, and then runs the export pipeline that creates the human-readable and machine-readable artifacts selected by the output profile.

This action is profile-driven. The profile determines both collection depth and which artifacts should be generated. The merged-profile behavior is resolved once in the runner before the core assessment script is invoked.

### `Preflight`

`Preflight` validates tenant connections and checks that the operator (or app registration) has the permissions required for the active output profile. It does not collect data or generate reports.

Use this on a new machine or against a new app registration before committing to a full collection run.

### `Collect`

`Collect` is not a separate assessment engine. It routes to the same core tenant assessment script as `Full`, but enables collection-only mode and forces JSON output on while disabling workbook, HTML, questionnaire, and PDF generation.

Use this when the goal is to capture tenant state once and defer report generation until later.

### `Report`

`Report` is the export-only counterpart to `Collect`. It loads a previously saved tenant snapshot, validates it, converts it back into the in-memory structure expected by the legacy engine, and then runs the same export pipeline used by the full `Full` path.

Use this when collection has already been completed or when the team wants to regenerate artifacts under a different output profile without reconnecting to the tenant.

### `AD`

`AD` is a separate Active Directory assessment path. It inventories AD forests, domain controllers, FSMO roles, GPOs, users, groups, computers, and OUs, then exports those results.

This workflow is not part of the Microsoft 365 snapshot/export model and does not share the M365 output-profile system.

### `Improve`

`Improve` is a snapshot post-processing workflow. It reads a single tenant snapshot and produces customer, roadmap, and engineer deliverables under `Deliverables`, plus JSON/snippet support artifacts under `Support`. Solutions Engineer evidence coverage is kept in the engineer pack and `Support` artifacts rather than shown as process detail in the customer-facing DOCX files. CSV and Markdown planning artifacts are only produced when legacy artifact output is requested.

It can also use Graph fallback calls when explicitly requested and when certain datasets are missing from the snapshot.

### `Compare`

`Compare` is another snapshot post-processing workflow. It reads a baseline snapshot and a current snapshot, recalculates a fixed set of posture metrics, and writes JSON, CSV, and Markdown comparison outputs.

This is intended for trend or before/after review rather than full artifact regeneration.

## Key Differences

### Input model

- `Full`, `Collect`, `Preflight`, and `AD` use live connections.
- `Report`, `Improve`, and `Compare` use saved JSON snapshots.

### Output model

- `Full` and `Report` are artifact-generation workflows.
- `Preflight` is a validation-only workflow with no artifact output.
- `Collect` is a capture workflow.
- `Improve` and `Compare` are analytic workflows on top of snapshots.
- `AD` is its own assessment/export path.

### Output profile support

- Only the `Full`, `Collect`, and `Report` family uses `OutputProfile`.
- `Preflight`, `AD`, `Improve`, and `Compare` have fixed behavior and fixed output types.

## Overlap Review

### 1. The `Full` family is one workflow with three execution modes

`Full`, `Collect`, and `Report` all route through the same runner logic and the same core legacy script. The real differences are mode flags and export behavior:

- `Full`: live collect + export
- `Collect`: live collect only
- `Report`: snapshot export only

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

`Report`, `Improve`, and `Compare` all depend on saved tenant snapshots and need:

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

2. Reframe the `Full` family as pipeline stages.
   - Keep the user-facing actions for compatibility.
   - Internally model them as one workflow with explicit modes: `Full`, `CollectOnly`, and `ExportOnly`.
   - This would make the control flow easier to reason about and reduce duplicate launcher/runner handling.

3. Centralize snapshot input and output conventions.
   - Shared helper for resolving snapshot paths, default output folders, and snapshot validation.
   - Reuse it across `Report`, `Improve`, and `Compare`.

### Lower-risk follow-up work

4. Introduce a common post-processing base for snapshot workflows.
   - Shared bootstrap for loading the common module, reading snapshots, and emitting JSON/CSV/Markdown outputs.

## Practical Recommendation

If the team wants to improve maintainability without changing user-facing behavior, the safest sequence is:

1. Extract shared snapshot metric helpers from `Improve` and `Compare`.
2. Extract shared snapshot bootstrap and output helpers for snapshot-driven actions.
3. Normalize the `Full` family into a single internal pipeline with three modes while keeping the existing launcher actions as compatibility aliases.

That order reduces duplication first, then clarifies workflow intent, without forcing a large user-facing redesign up front.
