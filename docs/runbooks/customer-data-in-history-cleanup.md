# Customer Data In Repository History - Outstanding Cleanup

**Status:** `main` is clean on both remotes. Seven other branches are not.
**Raised:** 2026-08-17

## What happened

On 2026-03-04, commit `a4f3a41` ("Refactor code structure for improved readability and
maintainability") added 96 assessment artifact files under `artifacts/test-runs/2026-03-04/`.
Commit `6a25934` ("cleaned up artifacts") deleted them the same day. Deleting a file does not
remove it from history, so the artifacts remained reachable from every branch descended from
those commits.

The artifacts are real customer output, not sample data. They identify the client by name in
plain text, for example:

```
**Date:** 2026-03-04
**Prepared By:** amedrano (automated assessment)
**Client:** Cavehouse Brewery
```

The set includes `.xlsx` workbooks, HTML assessment and best-practices reports, full report
logs, and tenant-to-tenant questionnaires.

## What has been done

At some point before 2026-08-17 the history was rewritten to exclude these artifacts and the
clean result was pushed to the GitHub remote's `main`. The Azure DevOps remote was never
reconciled, which is why the two remotes carried different histories that diverged at
`460d00f` (2026-02-25).

On 2026-08-17, `origin/main` (Azure DevOps) was force-updated from `28861d9` to `457e6ae`,
replacing the polluted history with the clean one.

This was verified safe before the push. Of the 127 commits then unique to `origin/main`,
124 had identical subjects to commits already present on the clean history. Only three were
unique: the two artifact commits above, and `fac684a` (Entra tenant-ID parameterization),
whose content was already present via `66ded29`. No code was lost.

A local safety ref preserving the old tip is at `backup/origin-main-pre-rewrite` (`28861d9`).
Delete it once the rewrite is confirmed good, since it still holds the customer data.

## What remains

### 1. Branches still carrying the artifacts

Each of these still reaches the `a4f3a41` / `6a25934` commits and must be rebuilt onto the
clean history or deleted if abandoned:

**Azure DevOps (`origin`)**
- `agent/presales-migration-scoping`
- `dev/collector-performance-refactor`
- `dev/python-migration`
- `fix/claude-code-review-remediation`
- `redesign/phase-pipeline`

**GitHub (`github`)**
- `agent/presales-migration-scoping`

To re-check the current list:

```powershell
git fetch --all
foreach ($ref in (git for-each-ref --format='%(refname)' refs/remotes/)) {
    if (git log $ref --oneline -1 -- 'artifacts/test-runs/*') { "CONTAINS: $ref" }
}
```

### 2. Force-pushing does not erase the old commits

Rewriting a branch only makes the old commits unreachable. On Azure Repos they stay fetchable
by SHA until garbage collection, which is not exposed on demand. If the requirement is that
this data be genuinely gone rather than merely off `main`, a force-push is not sufficient.
Options, in increasing order of certainty:

1. Delete or rebuild every ref that reaches the commits, then open an Azure DevOps support
   request to force garbage collection.
2. Recreate the repository from the clean history and archive the original.

The same consideration applies to the GitHub remote, where unreachable commits remain
accessible until GitHub GC, and any fork or existing clone retains them regardless.

### 3. Decide whether this needs escalating

Customer assessment output committed to a shared repository may carry notification or
contractual obligations. That determination is a business call, not a technical one, and has
not been made.

## Preventing recurrence

`.gitignore` now has anchored rules for `runs/`, `artifacts/**`, `reports/**`, `output/**`,
and root-level test output. These protect future commits only.

Codex's review also recommended a pre-commit or CI customer-data scan. That is not yet
implemented and would be the durable fix, since ignore rules do not help when a path is
added deliberately or with `git add -f`.
