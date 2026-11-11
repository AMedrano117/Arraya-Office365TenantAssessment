# T2T Review Summary

Review date: 2026-11-10
Branch: `feature/t2t-migration-assessment-review`
Documents: [coverage](T2T-Migration-Assessment-Coverage.md), [backlog](T2T-Migration-Assessment-Backlog.md), [discovery questionnaire](T2T-Migration-Discovery-Questionnaire.md), [synthetic scenario spec](T2T-Synthetic-Scenario-Spec.md), [application identity mapping template](../templates/T2T-Application-Identity-Mapping-Template.csv)

## Final assessment

The assessment does not yet capture enough evidence to plan a complex tenant-to-tenant migration with a short customer-defined deadline, a hybrid destination, new sign-in and mail domains, guest replacement, a SharePoint rename, and SAML application continuity. It is a sound base for the Exchange and collaboration parts of that plan and has the right decision-register design to grow into the rest.

Three things stand between the current state and an adequate one:

1. Ten outputs are wrong or unsupported today and were reproduced offline. They should be corrected before the tool is used to plan a migration.
2. The profile intended for migration work collects no application, group, device, Conditional Access, or compliance evidence and omits the planning registers.
3. There is no two-tenant model, so nothing about the destination or about collisions between the tenants can be evaluated.

No code was changed in this review. The P0 corrections were implemented afterwards on branch `feature/t2t-p0-correctness`; see [T2T-P0-Implementation-Status.md](T2T-P0-Implementation-Status.md). The sections below describe the state at the time of the review.

Design correction: the reusable assessment is milestone-neutral. It defines no "Day 1", no fixed phases, and no required cutover. Deadlines, waves, and cutover dates are engagement inputs.

## What is already adequate

- Mailbox inventory, sizing, archives, inactive mailboxes, and third-party tool license sizing.
- Full Access, Send As, Send on Behalf, and calendar-permission inventories with collection states.
- SMTP AUTH and relay evidence with explicit evidence states.
- SharePoint, OneDrive, Microsoft 365 Group, and Teams inventories, storage reconciliation, and activity.
- The `DecisionKey` register: discovered evidence, customer confirmation, owner, and due date are separate; section approvals cannot confirm object rows; nothing is pre-approved.
- Enterprise application inventory, activity, reply URLs, and credential metadata (in the profiles that collect it).
- Read-only collection, tenant-mismatch guard, collector status classification, and a passing test baseline (404 Pester tests, 73 Python tests).

## What is missing

- Source and destination roles and any comparison between two tenants.
- SAML NameID, claims, assignments, service principal signing certificates, SCIM matching, and an application identity mapping worksheet.
- Hybrid identity detail: sync client, server, staging mode, health, filtering, rules.
- Calendar and meeting evidence and the report language that separates copied content from operational continuity.
- A group disposition register and cross-references from groups to their consumers.
- A way to attach engagement-defined milestones, waves, and sequencing dependencies to items, and governance decisions for licensing, access, tooling, DNS authority, application owners, endpoints, communications, and rollback.
- Licensing readiness, tenant rename readiness, domain removal dependencies.
- Endpoint and file-server discovery, which the tenant cannot provide.

## What must be corrected first (P0)

| ID | Defect | Reproduced |
| --- | --- | --- |
| P0-01 | AD Connect and Exchange hybrid are flagged "Configured" for every tenant | R1, R2 |
| P0-02 | Unknown sync and sign-in state is reported as cloud-only | R4, R8 |
| P0-03 | Items with no collected data are scored "Ready" and dropped from the cutover checklist | R3, R5, R6 |
| P0-04 | Multi-Geo is never detected; the questionnaire says none was found | R7 |
| P0-05 | Failed DNS lookups and `mx.microsoft` records are classified as third-party mail routing | R9 |
| P0-06 | The tenant-to-tenant profile is described as deep readiness, collects the least identity evidence, and does not say what it skipped | Code reading |
| P0-07 | "Mailboxes on hold" counts litigation hold only | R10 |

All seven are narrow, need no new permission, add no runtime, and do not change snapshot or worksheet schemas. Each has a fail-first test defined in the scenario spec.

## What should remain a questionnaire or a separate collector

| Topic | Approach |
| --- | --- |
| Entra Connect server, staging mode, version, filtering, custom rules, destination OU, writable DC, AD administration server | Questionnaire, plus an optional on-server export with an import contract |
| Application account key, JIT behavior, rename and bulk-map support, dual identity provider trust, test environment, cutover and rollback | Vendor and application-owner questionnaire feeding the mapping worksheet |
| Destination SKU, agreement terms, Teams bundling, seat limits | Customer and commercial confirmation; the tool reports evidence only |
| DNS authority and registrar access; production mail flow validation | Customer confirmation and test evidence; never derived from tenant verification |
| Gateway, archiving, journaling, and phishing-report vendors | Customer and vendor questionnaire, with connector evidence as context |
| File servers | Separate optional discovery module or import of the migration tool's source analysis |
| Endpoints, GPO, VPN, Wi-Fi, certificates, BitLocker, profiles | Separate optional endpoint and AD collector plus questionnaire |
| SharePoint item-level issues and permissions | Import from the migration tool pre-check |
| Calendar items | Report semantics by default; opt-in counts-only collection |
| Teams chat method and fidelity | Vendor questionnaire; activity counts as context |
| Legal and compliance determinations | Customer legal and compliance owners; the tool lists evidence and open decisions |

## Recommended implementation sequence

1. P0-01 to P0-07 as one change set with fail-first tests.
2. P1-01 and P1-12: profile scope and the readiness output model.
3. P1-02, P1-04, P1-03: application SSO evidence, signing certificate expiry, mapping worksheet.
4. P1-05 and P1-11: two-tenant comparison and guest replacement.
5. P1-06, P1-15, P1-16, P1-07: hybrid identity, domains, rename, licensing.
6. P1-08, P1-13, P1-14, P1-09 part A, P1-10, P1-17.
7. P2 by project demand.

## Decisions needed

- Authorization to implement the P0 change set.
- Whether the tenant-to-tenant profile should grow to full collection (P1-01) or the Presales and tenant-to-tenant profiles should be merged into one migration profile.
- Whether a second, optional Graph permission for provisioning evidence (`Synchronization.Read.All`) is acceptable for customer app registrations.
- Whether opt-in calendar metadata collection (P1-09 part B) is acceptable at all, given its sensitivity.
- A written, approved procedure for read-only validation in the internal test tenant, if any behavior in section 9 of the scenario spec needs it.

## Limitations of this review

- Static review plus offline reproduction. No live collection was run, so collector behavior against real API responses was not observed.
- The 1.2 MB legacy collector and the improvement-plan generator were read selectively, guided by the migration code paths. Best-practice findings were not reviewed.
- Third-party tool behavior was not verified. Microsoft limits were checked on the review date and will change.
- Line numbers refer to commit `66f48a8`.
