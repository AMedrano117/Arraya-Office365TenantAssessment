# Assessment Pipeline Redesign

This document describes the parallel redesign line introduced on the `redesign/phase-pipeline` branch.

## Branch Strategy

- `main` remains the maintenance line for bug fixes, tenant-accuracy corrections, and small UX/reporting improvements.
- `redesign/phase-pipeline` is the long-lived architecture branch for the new phase-based pipeline runtime.
- Backports from `main` to the redesign branch should be selective and limited to correctness or shared business-rule changes.

## Runtime Model

The redesign introduces a checkpoint-backed pipeline with explicit phase contracts:

1. `Connection`
2. `TenantOverview`
3. `Identity`
4. `Exchange`
5. `Collaboration`
6. `Endpoint`
7. `Governance`
8. `Export`

Each phase now has:

- a stable name and execution order
- declared dependencies
- a checkpoint file when the phase emits data
- a resumable state entry in `pipeline.state.json`

## Current Bridge Design

The redesign does not rewrite all collectors in one pass. Instead:

- `Arraya.M365.AssessmentPipeline` owns phase order, checkpoint paths, resume behavior, and phase state.
- `Arraya.M365.AssessmentRunner` routes Microsoft 365 assessment commands through the new pipeline runtime.
- the legacy `Get-FullTenantReportDetails.ps1` script remains a temporary execution bridge and can now run a single phase with checkpoint output.

That lets the team:

- rerun one phase without replaying the full assessment
- resume from the last completed checkpoint
- export from a saved checkpoint
- add focused tests around orchestration before moving deeper collectors out of the monolith

## Key Entry Points

- `Invoke-M365TenantPipeline`
- `Resume-M365TenantPipeline`
- `src/scripts/assessments/tenant-wide/Invoke-M365TenantPipeline.ps1`
- `src/scripts/assessments/tenant-wide/Resume-M365TenantPipeline.ps1`

## Checkpoint Layout

By default, the pipeline writes state under:

```text
<ExportPath>\Support\Pipeline
```

Artifacts include:

- `pipeline.state.json`
- `01-Connection.snapshot.json`
- `02-TenantOverview.snapshot.json`
- `03-Identity.snapshot.json`
- and so on for each checkpointed phase

## Next Extraction Order

1. keep the current bridge stable
2. move connection/preflight logic fully under the pipeline module
3. peel Tenant Overview and Identity orchestration out of the legacy script
4. continue phase-by-phase until the legacy script is no longer the execution bridge
