# T2T Profile Composition and Evidence States

Date: 2026-11-11
Branch: `feature/t2t-p0-correctness`
Related documents: [P0 status](T2T-P0-Implementation-Status.md), [backlog](T2T-Migration-Assessment-Backlog.md)

This pass makes profile scope explicit, composes multiple profiles deterministically, completes collector evidence states, separates hold evidence from migration treatment, narrows Presales blockers, and adds the `MigrationCriteriaRegister` projection. No collector was added.

## Decisions resolved

| Open decision after P0 | Resolution |
| --- | --- |
| How to commit the review documents | Committed first, as originally written, then the P0 batch, then this pass. |
| Whether unknown synchronization and hybrid state block the Presales quote | Narrowed. Unknown synchronization blocks (identity design cannot be scoped). Unknown Exchange hybrid is `NeedsData` and a non-blocking quote condition (coexistence is a planning decision). The rule is read from the criteria catalog. |
| Whether the hold row may never be Ready | Replaced. Complete hold evidence with no holds is `Ready` and `Clear`. Incomplete evidence is `Partial` or `NeedsData` and `NotEstablished`. |
| The Presales hold blocker required Purview, which the Presales profile does not collect | Split (catalog review, 2026-11-11). `PS-CMP-001` and `T2T-CMP-001` now cover mailbox-level hold fields only; `PS-CMP-001` stays a `PresalesBlocker`. New `PS-CMP-002` and `T2T-CMP-002` cover organization-wide retention policies as `NeedsData`. Matching rows merge in combined runs. |
| `ID-02` cited for two different decisions | The directory synchronization criteria no longer cite `ID-02`, which remains the identity mapping sign-off. The catalog loader now rejects one key with two different decisions. |
| `new-migrationbatch` matched the native cross-tenant method | Removed. The cmdlet also runs hybrid, IMAP, staged, and cutover migrations, so a confirmed method must name cross-tenant migration explicitly. |

## What changed

| Area | Change |
| --- | --- |
| Profile scope | `src/config/baseline/migration-evidence-criteria.json` defines 13 Presales and 15 TenantToTenantMigration criteria. TenantToTenantMigration lists enterprise applications, SSO, Entra groups, Microsoft 365 Groups, devices, Conditional Access, and Purview as explicit criteria; SSO identity configuration and the synchronization server are `Unsupported`. |
| Composition | `Resolve-ArrayaMigrationCriteria` merges only semantically identical criteria, keeps owning and contributing profiles, applies the strictest rule, never lets another profile inherit a Presales blocker, and is order-independent. Combined runs now build the register too. |
| Collector states | `Get-ArrayaDatasetEvidenceState` returns a `CollectorState` from the ten-value vocabulary. Public folders, mail flow rules and connectors, domains, remote domains, and Purview policies now record their own state through `Set-ArrayaCollectorEvidenceState`, kept in snapshot diagnostics. An empty dataset with no collector record or step result is now `Unknown`, not complete. |
| Holds | `Get-ArrayaMailboxHoldTreatment` returns evidence status and migration treatment separately. `-EvidenceScope` evaluates mailbox hold fields, organization-wide Exchange retention policies, or both (`All`, the default, used by the cutover summary). `BlocksMigration` uses `HoldMethodRules`, sourced from Microsoft documentation that native cross-tenant mailbox migration blocks mailboxes on any hold. |
| Presales | `PresalesBlocker` requires missing required evidence for a gating decision, and each blocker row names the missing evidence, blocked decision, impact, resolution evidence, role, and profile. |
| Output | `MigrationCriteriaRegister` worksheet in the Presales and tenant-to-tenant workbooks, with a stable 25-column schema, and the same rows under `Derived` in the snapshot. |
| Fixed in passing | The domain collector dropped every domain row when the tenant had no recipients (an empty array failed parameter binding). |

## Tests

`tests/unit/T2TProfileComposition.Tests.ps1`: 36 fictional scenarios. Against the previous commit, 34 fail; the 2 that pass are compatibility checks of the cutover pack CSV columns and existing worksheet columns. One P0 assertion was tightened (an empty dataset with no collector record is no longer complete).

The catalog review fixes added 9 scenarios and extended 1 (45 in total). Before the fixes, 9 fail; the one that passes is a guard that missing mailbox hold fields still block Presales. Full Pester suite: 516 total, 515 passed, 0 failed, 1 skipped. Python: 54 and 19 tests pass.

## Compatibility

No existing worksheet, column, profile name, parameter, decision key, or cutover CSV column changed. Additive: one worksheet, one snapshot derived table, one snapshot diagnostics table, eleven exported functions in total across both passes. Replayed snapshots that carry no collector results now report empty datasets as `Unknown`.

## Not done

- Collection-level composition (P1-18): a combined run still turns off delegate collection and the tenant-to-tenant shaping. The register now shows that as `NotRequested` instead of hiding it.
- Mailbox enumeration still suppresses its own errors (P1-19, remaining part).
- P1-12, the source-versus-destination comparison, and new collectors.
