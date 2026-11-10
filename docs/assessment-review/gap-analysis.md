# Assessment review agent: gap analysis (2026-10-28)

## Starting point

**Assessment engine (read-only).** A completed package already holds
everything an independent review needs:

- run identity
- a hashed snapshot of 45 customer findings, 21 actions, 19 controls and
  1,074 evidence records
- source states for each collector
- per-file hashes for the collected inputs
- the customer workbook, the Technical Evidence workbook and the HTML
  deliverables

The engine evaluates and renders its own conclusions. It does not challenge
them.

**This repository.** `tools/customer_deliverables.py` renders an approved
projection into three documents. It never re-evaluates conclusions, never
compares them with evidence or Microsoft guidance, and requires the approved
`agent-*` inputs.

## Gaps closed by the MVP

| Need | Before | MVP |
|---|---|---|
| Package validation | Engine-side only | `package.py`: identity, hashes, source classes, dates, links, workbook and HTML versus JSON |
| Independent review | None | `engine.py`: workflow steps 1 to 7 per finding, typed statements, one decision and one treatment |
| Microsoft benchmarks | Unverified claims flagged by the engine | `benchmarks.json`: 25 Microsoft Learn entries verified 2026-10-28 |
| Missing findings | None | Rules for pilot scope, Global Administrator count, emergency access, endpoint coverage, incident process, agents, shadow AI and AI investigation |
| Deliverables | Three re-rendered documents | Four review deliverables plus review-results.json and delivery-manifest.json |
| Privacy | Pattern checks | Redaction with a collection-derived deny-list, restricted evidence classes and output scanning |
| Tests | Renderer tests | Fictional package, rule, privacy and deliverable tests, plus opt-in expert validation |

## Not in scope

The MVP does not include the following:

- Copilot Studio deployment
- SharePoint publishing
- write-back to the engine
- finding closure
- governance approval
- tenant changes
- new collectors
- contract migration
- maturity scoring
- automatic delivery to the customer
