# Customer deliverables MVP

This downstream renderer accepts the atomic projection **1.2.0** package in
inputs/assessment-package/ and generates exactly:

- Executive Readiness Summary.docx
- Customer Action Plan.xlsx
- Review Meeting Brief.docx

The destination is generated/AssessmentId/RunId/. The qualification remains
**ENGINEERING VALIDATION ONLY — NOT CUSTOMER-READY**. Operator review is not
permission to publish or distribute to customers.

## Run

From the repository root, using the existing virtual environment:

    .venv\Scripts\python.exe -B tools\customer_deliverables.py

Dependencies are pinned in [customer-deliverables-requirements.txt](../../tools/customer-deliverables-requirements.txt).
They were already installed in the repository virtual environment when this
workflow was built. The generator prints its validation result and file hashes
to stdout; it does not add a receipt or other deliverable to the output folder.

## Input contract

Keep these four files from one export together:

- agent-demo-manifest.json
- agent-customer-review.json
- agent-knowledge.md
- agent-evidence-index.json

The generator verifies the manifest self-hash, every listed exact-file SHA-256,
the review's canonical self-hash contract, version 1.2.0, source-hash metadata,
AssessmentId, RunId, internal-use qualification, and recorded ID/count
reconciliation. It rejects this MVP's package unless it contains exactly five
recorded blockers and twenty customer actions. Validation does not recompute
assessment findings, priorities, or metrics. No schema or assessment engine is
modified or invoked.

The evidence index is read only for package hashes, kind/version, source-hash
metadata, and assessment/run identity. None of its evidence entries or technical
content enters report generation. No restricted snapshot is opened.
Markdown integrity is verified, but this MVP renders explanations from the
approved JSON fields and fixed neutral workflow prose; Markdown cannot override
JSON.

## Content and privacy

The action sheet has the ten requested columns and a filterable Excel table.
The Source States sheet preserves each action's recorded identities, technical
status, evidence sufficiency, lifecycle, governance treatment, and customer
action status. Assessment Context preserves recorded counts, qualification,
lifecycle/governance records using explicit safe fields, and assessment identity.
Null or empty source values display as Not recorded. Roles are copied from the
source; no individuals or dates are assigned.

The two Word documents retain source blocker IDs, actions, closure evidence,
readiness decision, and evidence qualifications. The executive summary includes
source-confirmation actions and source-supported configuration observations.
The meeting brief includes all recorded evidence-gap IDs and all twenty actions'
closure-evidence requests.

Only explicit scalar fields enter the documents. Private containers, restricted
references, approved-document links, and raw evidence are excluded. Pattern and
known-private-value checks reject detected identity leakage before writing.
These checks do not independently establish whether every natural-language
name was correctly redacted upstream; the approved customer projection remains
the privacy boundary. Source text is stored as literal Excel strings, never
formulas.

## Validation and repeat runs

Generation occurs twice in memory and must produce identical bytes. OOXML/ZIP
timestamps are fixed solely for reproducibility; they are not assessment dates.
All Office ZIP/XML parts are parsed, Word documents are reopened with python-docx,
and the workbook is reopened with openpyxl. The workbook's twenty rows and state
records must match the JSON field by field. Every source blocker is checked in
each document's blocker section. Privacy checks run against exported XML text.

Inputs are hashed again before output publication and after generation. Files
are staged and validated before exposing the final three-file directory.
Repeating a run with identical inputs is idempotent. An existing different
output, operator-edited file, or unexpected fourth file causes a stop rather
than an overwrite. Byte determinism is qualified for the pinned runtime and
dependencies; it is not promised across arbitrary library changes.

Native Word/Excel open checks may be run by the operator. Successful structural
or native-open checks do not claim a human visual review. The engineering
qualification and operator review requirement remain.

## Automated tests

    .venv\Scripts\python.exe -B -m unittest discover -s tests -p test_customer_deliverables.py -v

Tests construct fictional packages in temporary directories. They cover
integrity/version/identity failures, mismatched source counts and priorities,
privacy leakage, Markdown non-authority, literal spreadsheet text, source-field
preservation, deterministic reruns, input preservation, and overwrite refusal.
They never read the customer's input package.

