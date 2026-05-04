# Modern Workplace Tenant Assessment

PowerShell automation for Microsoft 365 tenant assessments, reporting, and improvement planning.

For operator-focused setup and execution guidance, use [RUN.md](RUN.md).

## Operator References
- [RUN.md](RUN.md)
- [Tenant Assessment Quick Start](docs/runbooks/tenant-assessment-quick-start.md)
- [App Registration Setup](docs/runbooks/app-registration-setup.md)
- [Certificate Auth Setup](docs/runbooks/certificate-auth-setup.md)
- [Improvement Plan Rule Taxonomy](docs/runbooks/improvement-plan-rule-taxonomy.md)

## User Quick Start
Run the launcher and pick an action from the menu:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1
```

Users do not need to manually import modules. The launcher imports `Arraya.M365.AssessmentRunner`, and the tenant assessment now owns its own workload-aware login flow. `Office365Custom` remains available for shared helpers and legacy scripts, but the main assessment no longer relies on `Connect-Office365` as its startup bootstrap.
The assessment also now validates required access as each workload connects, instead of deferring one large permission preflight block until after all authentication work is finished.

## Auth Modes
- `Delegated`: interactive sign-in for full module compatibility.
- `Certificate`: noninteractive app auth for Graph and Exchange, with Graph-based SharePoint fallback in PowerShell 7.
- `Client secret`: compatibility mode for Graph app auth. In this workflow it is not equivalent to certificate auth: Exchange Online falls back to delegated sign-in, Purview compliance app auth is unsupported, and certificate auth remains the recommended production path.

If you do not pass `-AuthMode`, the assessment defaults to delegated interactive sign-in. For backward compatibility, supplying `-CertificateThumbprint` still switches the run to certificate auth, and supplying `-ClientSecret` still switches the run to client-secret auth.
For safer secret handling, the launcher and runner also accept `-ClientSecretSecure` or `-ClientSecretCredential` so the secret does not need to stay as a plain string in shell history.
The assessment now stages authentication by workload instead of trying to connect every possible Microsoft 365 surface up front. Graph and Exchange remain the baseline live-collection dependencies. Purview is only connected when the active run needs retention/DLP collection. SharePoint admin PowerShell and Teams PowerShell are treated as optional workload-specific connections with explicit fallback behavior. In interactive mode, the assessment now intentionally connects Exchange Online and Purview before Microsoft Graph so Exchange-family auth gets a clean session before Graph interactive auth runs.
If you already connected the required workloads for the current run in the current session, you can run with `-SkipAuth` to reuse those sessions and bypass the assessment's authentication bootstrap. `-SkipAuth` now validates only the workloads the active profile actually needs.
When the active profile includes Purview retention or DLP collection, `-SkipAuth` expects an already-usable compliance PowerShell session, not just imported cmdlet names.
Use `-SkipPermissionPreflight` if you want to bypass the staged workload access checks during startup and let the collection continue until a later collector hits missing access.
The launcher now also reuses the repo modules already loaded from this repository in the current PowerShell session instead of force-reimporting them on every run.
If you want the assessment to continue without the startup permission gate, you can add `-SkipPermissionPreflight`. This skips the required-access validation at the beginning of the run and allows collection to continue on a best-effort basis, which means missing permissions may still surface later as individual workload failures.
The permission preflight now shows the specific check in progress, then prints a short summary of how many checks succeeded and how many access gaps remain. It fast-passes core Graph permissions from the current token claims and reserves live probes for the more ambiguous endpoints, so startup validation stays more accurate without making every Graph scope wait on a network call. Non-blocking checks such as `OnPremDirectorySynchronization.Read.All` are reported as structured warnings in the preflight summary instead of only surfacing as raw Graph 403 noise.
Progress output also now uses plain-language governance step names instead of older internal `Tier B` terminology.
The operator docs now distinguish between the core Graph permission set used by the standard assessment and optional extended-enrichment permissions, so the app registration ask is easier to defend and keep least-privileged.
App-based Teams collection now keeps team and channel inventory as part of the standard run, but member-count and guest-count enrichment are treated as optional. If `TeamMember.Read.All` or `TeamMember.ReadWrite.All` is not granted, the assessment logs a one-time warning and continues without that deeper Teams membership expansion.

Retention and DLP policy collection uses Purview compliance PowerShell through `Connect-IPPSSession`, not the main Graph collector path. In this workflow:
- delegated auth is supported
- interactive delegated auth first uses a normal `Connect-IPPSSession` attempt, then retries with `-DisableWAM` when supported
- device-code fallback for Purview is only available when the installed `ExchangeOnlineManagement` module exposes `Connect-IPPSSession -Device`
- certificate auth is supported using `AppId + Organization + CertificateThumbprint`
- client-secret auth is not supported for Purview compliance collection
- `ExchangeOnlineManagement` must be available because it provides `Connect-IPPSSession`
- when Purview compliance connection fails, the assessment now reports the auth path, tenant organization value, and next-step guidance in the preflight output so the operator can see what still needs to be corrected
- if Microsoft Graph interactive auth has already broken later Exchange or Purview auth in the current shell, start a fresh PowerShell session before retrying the assessment
- a certificate/app combination can succeed in one tenant and fail in another if the target tenant does not expose the Purview retention/DLP cmdlets to that app session; the preflight now calls that out explicitly when the compliance endpoint accepts auth but does not assign those cmdlets

For SharePoint and OneDrive collection, certificate-based and client-secret runs now skip importing the `Microsoft.Online.SharePoint.PowerShell` module entirely. Those app-based runs rely on Microsoft Graph collection instead of `Connect-SPOService`. Interactive runs only attempt SharePoint admin PowerShell when the active assessment run can benefit from it.

Examples:

```powershell
.\src\scripts\operations\Start-M365TenantAssessment.ps1 -Action M365

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -SkipImprove

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -AuthMode Interactive

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -SkipAuth

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -SkipPermissionPreflight

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -AuthMode Certificate `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -AuthMode ClientSecret `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -ClientSecret $env:ARRAYA_M365_CLIENT_SECRET

$clientSecret = Read-Host 'Client secret' -AsSecureString
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -AuthMode ClientSecret `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -ClientSecretSecure $clientSecret

.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365 `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>' `
  -ExportPath 'C:\Assessment-Outputs' `
  -OutputProfile SolutionsEngineer,ExecutiveLevel
```

## Collection Vs Export
The M365 workflow now supports explicit separation between data collection and artifact export:

- `M365Preflight`: runs only workload connection plus permission preflight, prints readiness, and exits without collection, export, or `Improve`.
- `M365Collect`: runs discovery/collection only and writes a JSON snapshot.
- `M365Export`: loads a previously collected JSON snapshot and generates artifacts without re-collecting tenant data.
- `M365`: the standard full assessment path. It now runs the assessment deliverables and the `Improve` post-processing step in one flow unless you pass `-SkipImprove`.
- each run also writes a `*.manifest.json` artifact index next to the primary export filename.
- JSON snapshots now use a versioned V2 contract with explicit sections: `Metadata`, `CollectionPlan`, `Data`, `Derived`, and `Diagnostics`.
- Import is backward compatible with older V1 snapshots through an in-memory adapter.

Examples:

```powershell
# Connection and permission preflight only
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365Preflight `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>'

# Collect data only (JSON snapshot)
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365Collect `
  -TenantId '<tenant-guid>' `
  -ClientId '<app-id>' `
  -CertificateThumbprint '<cert-thumbprint>' `
  -OutputProfile SolutionsEngineer `
  -SkipPermissionPreflight

# Export artifacts from an existing JSON snapshot
.\src\scripts\operations\Start-M365TenantAssessment.ps1 `
  -Action M365Export `
  -OutputProfile ExecutiveLevel
```

## Assessment Outputs
For a default `-Action M365` run, the top-level operator deliverables are now:

- `*-CustomerAssessmentReport.docx`: the primary customer-facing deliverable generated from `Improve`
- `Support\*-Microsoft 365 Remediation Roadmap-<date>.docx`: the companion executive remediation roadmap generated alongside the customer report
- `*-EngineerActionPack.md`: the primary engineer-facing remediation deliverable

Support and machine-readable artifacts are written under `Support\...`:

- `Support\*-ImprovementPlan.json`: the machine-readable remediation payload
- `Support\*-AssessmentSnapshot.json`: the assessment snapshot used by `Improve`, replay, and export workflows
- `Support\*.manifest.json`: artifact index for the run
- `Support\*-RemediationSnippets.ps1`: support helper commands

Logs and troubleshooting output are written under `Debugging\...`.

The workflow still preserves a JSON assessment snapshot so `Improve`, `M365Export`, and later comparison/report replay can work reliably.

Legacy assessment artifacts are still available, but they are now compatibility outputs rather than default deliverables:

- `*-TenantSnapshot.html`
- `*-BestPracticesSnapshot.html`
- `*-TenantToTenantQuestionnaire.md`
- `*.pdf`

Workbook output is still a default deliverable for `SolutionsEngineer` and `TenantToTenantMigration`. Use `-IncludeLegacyAssessmentArtifacts` when you explicitly want the older HTML / questionnaire / PDF artifact family in the same run. Use `-IncludeLegacyArtifacts` when you also want the older `Improve` CSV/Markdown planning artifacts.

For operator guidance on how to interpret `Improve` findings and rule IDs, see [Improvement Plan Rule Taxonomy](docs/runbooks/improvement-plan-rule-taxonomy.md).

## Output Profiles
The assessment run supports six operator-facing output profiles:

- `Presales`: scope `Minimum`
- `ExecutiveLevel`: scope `Minimum`
- `SolutionsEngineer`: scope `Operator`
- `Machine`: scope `Automation`; intended for JSON-focused automation and replay paths
- `Geek`: scope `Geek`
- `TenantToTenantMigration`: scope `All`

`SolutionsEngineer` is the default profile.
You can run multiple profiles in one command by passing a comma-separated list (for example `SolutionsEngineer,ExecutiveLevel`).
When multiple profiles are supplied, the assessment runs once using the highest required reporting scope.
In the default `M365` flow, profiles now influence collection/reporting depth more than artifact sprawl. The consolidated remediation outputs stay the default regardless of profile, and the legacy workbook/HTML/questionnaire family is only added when you pass `-IncludeLegacyAssessmentArtifacts`.

Detail levels are intended to be read this way:

- `Minimum`: fastest leadership and presales story with lighter enrichment
- `Operator`: standard consultant/operator assessment depth without the heavy combined user/mailbox projection
- `Automation`: structured snapshot depth for `Improve`, export replay, and automation workflows
- `Geek`: deep engineer troubleshooting depth
- `All`: deepest migration-oriented collection for readiness, cutover analysis, and the full combined user/mailbox projection

`Combined` remains accepted as a compatibility alias for `Operator` when older wrappers or scripts still pass the legacy name.

`-RunImprove` remains available for `M365Collect` when you want to collect a snapshot and immediately post-process it, but it is no longer required for the main `M365` workflow.
When you run `Improve` separately, use `-LiveRefresh` if you want snapshot-plus-live-refresh behavior; it is a friendlier alias for `-UseGraphFallback`.

Reporting scope is not prompted interactively. Scope is automatically derived from the selected output profile.

## Runtime Notes
- In `Minimum` scope profiles, the collector depth policy trims high-cardinality enrichment to reduce runtime and memory pressure.
- Examples: Entra group deep membership/license expansion and detailed SSO app inventory are reduced in `Minimum`.
- The output contract is preserved: workbook tabs and report artifacts still generate with compatible values.
- Run logs now include collector duration and memory summaries (`[CollectorMetrics]`) plus inventory row counts (`[CollectorInventory]`) for hotspot review.

## HTML And PDF
- `CustomerAssessmentReport.docx` is now the primary customer-facing deliverable and is generated from the approved Word template.
- The best practices analysis HTML and full technical HTML are compatibility artifacts generated only when `-IncludeLegacyAssessmentArtifacts` is used.
- PDF is also a compatibility artifact in this model and remains skippable with `-SkipPdfReport`.
- PDF rendering uses a locally installed Chromium-based browser, preferring Google Chrome and falling back to Microsoft Edge.

If a supported browser is unavailable, the remediation HTML and markdown outputs still complete.

## Primary Entry Scripts
- `src/scripts/operations/Start-M365TenantAssessment.ps1`
- `src/scripts/assessments/tenant-wide/Invoke-M365FullTenantAssessment.ps1`
- `src/scripts/assessments/tenant-wide/Invoke-M365TenantDataCollection.ps1`
- `src/scripts/assessments/identity/Invoke-ActiveDirectoryTenantAssessment.ps1`
- `src/scripts/reporting/Invoke-M365TenantAssessmentExport.ps1`
- `src/scripts/reporting/Invoke-M365TenantImprovementPlan.ps1`
- `src/scripts/reporting/Invoke-M365TenantAssessmentComparison.ps1`

## Repo Layout
- `docs`: user-facing templates and questionnaires. The tenant-to-tenant questionnaire template lives under `docs/templates`.
- `src/modules/Arraya.M365.Common`: shared helpers and Office365Custom local import function.
- `src/modules/Arraya.M365.AssessmentRunner`: user-facing commands that execute assessment/report scripts.
- `src/scripts/migrated/legacy`: migrated legacy scripts kept for compatibility, including the core M365 assessment engine, HTML/PDF helper, and questionnaire exporter. The standalone `Invoke-EntraAppReport.ps1` file is retained as an unsupported reference script and now requires explicit opt-in to run.
- `src/vendor/Office365Custom/1.2.0`: vendored module used as shared function source.

## Development
1. Install PowerShell 7+.
2. Run `tools/bootstrap-dev.ps1`.
3. Validate with `tools/invoke-scriptanalyzer.ps1` and `tools/run-pester.ps1`.
