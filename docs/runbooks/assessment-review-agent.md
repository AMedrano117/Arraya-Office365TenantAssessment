# Assessment review agent (MVP)

The assessment review agent reads one completed M365 Copilot Readiness
Assessment package and produces an independent engineering review of it. The
package stays read-only: the agent never writes to it, the collected evidence,
the snapshots or the existing deliverables.

The agent is deterministic. It does not call a language model. Every
conclusion comes from retained package data, a versioned Microsoft benchmark
registry, and rules in `tools/assessment_review/`.

## Run

From the repository root:

    .venv\Scripts\python.exe -B -m tools.assessment_review --package "<package folder>" --review-date 2026-10-28

Outputs are written to `generated/assessment-review/<AssessmentId>/<RunId>/`,
which git ignores:

| File | Audience |
| --- | --- |
| Executive-Readiness-Summary.docx | Leadership and business sponsors |
| Technical-Assessment-Review.docx | Senior engineers |
| Customer-Action-Plan.xlsx | Customer responsible roles (confidential: names users, devices, sites and apps) |
| Review-Meeting-Brief.docx | Review meeting participants |
| review-results.json | Arraya engineering record (not sent to the customer) |
| delivery-manifest.json | SHA-256 of every output and input, plus publication status |

### How the documents are laid out

Every document opens with the readiness answer, then the recommendations, then
the detail. Each action has one number (#1, #2, ...) that is the same in all four
documents, and each action has the same plain wording everywhere:

- a short instruction ("Turn on the existing MFA policy for all users")
- what was found, with the collected counts
- why it matters
- how to fix it
- what shows it is done

Actions are grouped by timing: Fix before rollout, Check before rollout,
Improve security, Information needed, and Plan for later. Within each group
they are ordered by priority, then by dependency step (foundation before
AI-specific controls). Items that need no action, and duplicates, are left out
of the customer documents but listed in the technical review and in the Review
trace sheet.

- **Executive summary** (about 2 pages): the answer, the group counts, the top
  recommendations (every Fix before rollout item plus the leading
  pre-rollout check in each area), a status line per area, the decisions
  needed before rollout, and next steps.
- **Technical review**: the summary, all recommendations in order, what changed
  from the original assessment, then every finding in detail. The detail labels
  each evidence note by source (collected data, Microsoft guidance, reviewer
  view, suggested practice, your decision, missing evidence). Package checks,
  the controls, conflicts and the Microsoft guidance list are in appendices.
- **Action plan workbook**: a Start Here sheet; an Action Plan sheet with
  eleven columns (#, When, Priority, Action, Owner, What we found, How to fix,
  Affected items, Done when, Status, Ref) where Status is a drop-down left
  blank for the customer; a Decisions sheet; the named lists (below); a Review
  trace sheet with the engineering trace for every item; and the Microsoft
  guidance sheet.
- **Meeting brief**: goals, agenda, decisions needed, the before-rollout
  checklist, what changed, and the information to send after the meeting.

### Named lists in the action plan

The Affected items column points each action to a sheet that names the
specific objects to fix. Every list row has its own Action needed (or Options
to fix) text, the related action numbers, and a blank Status drop-down. The
lists are built by `tools/assessment_review/details.py` from the retained
collection and from collected inputs whose hashes verify:

| Sheet | Lists | Source |
| --- | --- | --- |
| Copilot users | People holding a Copilot license, sign-in and MFA state | M365 users, licenses, registration report |
| Admin accounts | Every admin role holder, roles, permanent or eligible, MFA methods and strength, last sign-in | Role assignments and schedules, registration report, sign-in activity |
| Users without MFA | Accounts with no MFA method, licensed or not, last sign-in | Registration report, M365 users, sign-in activity |
| Sign-in policies | Conditional Access policies, state, scope and named exclusions | Conditional Access policies |
| Risky accounts | Accounts at risk, level, recent detections, where to investigate | Identity Protection |
| Devices | Entra devices, join type, managed, compliant, last sign-in, Defender state | Directory devices, Defender for Endpoint |
| Overshared sites | Sites and OneDrives with Everyone, EEEU or broad links, the exact grants, options to fix | Data Access Governance CSV exports |
| Sites | Every SharePoint site, owner, last activity, storage, sharing setting | SharePoint site inventory and usage report |
| Guests | Guest accounts and last sign-in | Guest users, sign-in activity |
| Apps with high access | Apps with high-impact delegated or application permissions, or unverified publishers; consent, last use, owners | Service principals, grants, app role assignments, sign-in activity |
| Security incidents | Open incidents, alerts, alert recommendations and the Defender portal link | Defender XDR incidents and alerts |

Two things the package does not contain, so the lists say so instead of
guessing: app registration secrets and certificates (only service principal
credentials are collected, and only their presence and expiry are shown), and
Microsoft 365 group owners (group-connected sites show the group as owner).

The wording comes from `tools/assessment_review/plain.py`. When the engine adds a
finding type, add a template there; until then the item falls back to the
original report's text.

Exit codes: 0 success, 1 review stopped, 2 output validation failed,
3 customer-ready publication stopped. When code 3 is returned, rerun with
`--allow-engineering-only` to write outputs marked
"ENGINEERING REVIEW ONLY - NOT CUSTOMER-READY".

Pass `--review-date` explicitly. The same package and review date always
produce byte-identical outputs.

## What the agent checks

1. **Package.** It verifies identity across the run record, collection,
   snapshot, seeds, history and evidence records. It checks the snapshot and
   collected-input hashes, and classifies each source as Complete,
   CompleteEmpty, Partial, Failed, Unavailable, NotRequested, NotApplicable or
   Unknown. It also checks dates, sentinels, truncation, broken evidence
   references, broken detail links, and disagreement between the workbooks,
   the HTML and the JSON.
2. **Each finding (workflow steps 1 to 7).**
   - It reconstructs the evidence and the script conclusion.
   - It runs consistency checks: counts, records versus entities, period,
     configuration versus operation, registration versus enforcement,
     inventory versus entitlement, gap versus failure, export versus
     telemetry, confidence, and output agreement.
   - It compares the finding with the benchmark registry.
   - It assigns exactly one review decision and one readiness treatment, and
     defines the closure evidence.
3. **Added findings.** It adds missing findings only from retained evidence,
   for example emergency access, Global Administrator count or endpoint
   coverage. Any dependency step that has no coverage becomes an evidence
   request.
4. **Readiness decision.** Readiness is judged for the tenant; no pilot group
   is needed. A confirmed tenant-wide gap that the original report rated as a
   blocker stays a blocker. Script blockers the evidence does not support are
   re-rated with the reason stated. Adoption planning (use cases and success
   measures) is outside the scope of this review.

Every statement is typed as one of the following:

- collector fact
- script-derived conclusion
- Microsoft-documented guidance
- reviewer interpretation
- suggested hardening practice
- customer policy decision
- evidence gap

## Microsoft benchmark registry

`tools/assessment_review/benchmarks.json` holds the versioned registry
(BenchmarkVersion). Each entry has a Microsoft Learn URL, a verification
date, a requirement type, its licensing dependency and a limitation. The
loader rejects entries without a Microsoft Learn URL. When guidance changes,
re-verify the page, update the entry, and increment BenchmarkVersion.

## Privacy

Script text is redacted before it reaches any output. The redaction removes
the following:

- object names and person names taken from the collection
- email addresses and UPNs
- GUIDs, IP addresses, tenant host names and phone numbers
- tokens, secrets and Graph payloads

Validation then scans the Word documents, the workbook summary sheets and
review-results.json against the same deny-list and patterns. Evidence records
that carry affected objects or identity data are classified Restricted. Their
IDs appear only in the technical review and in review-results.json, never in
the executive summary, the brief or the action plan.

The named list sheets are the one place that names tenant objects, by design,
so the customer can act on them. They are written only to the action plan
workbook (never to the Word documents or review-results.json), the Start Here
sheet marks the workbook confidential, and validation checks them for tokens,
secrets and raw API payloads. Secret and certificate values are never read.

## Tests

    .venv\Scripts\python.exe -B -m unittest discover -s tests/assessment_review -t .

The tests build fictional packages (Northwind Traders, Tailspin Toys) in
temporary folders. They cover the package, review-rule, privacy and
deliverable cases in the MVP test plan, including the cross-customer and
determinism checks.

Expert validation against a real package is opt-in and stores no customer
data:

    $env:ASSESSMENT_REVIEW_VALIDATION_PACKAGE = "<package folder>"
    .venv\Scripts\python.exe -B -m unittest tests.assessment_review.test_expert_validation -v

## Limits

- The agent reviews; it does not close findings, approve exceptions, change
  the tenant, publish to SharePoint or write back to the engine.
- Customer action status stays "Not recorded" until the customer supplies a
  status. Owners are recorded as roles only.
- Operator review is required before any customer distribution.
