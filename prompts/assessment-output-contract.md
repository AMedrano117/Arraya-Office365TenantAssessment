# assessment-output-contract.md
## Required outputs
- `exports/findings.csv`
- `exports/execution-metadata.json`
- `exports/errors.json`
- optional `raw/raw-data.json`

## Findings columns
TenantId,Workload,Category,Control,CurrentState,RecommendedState,Severity,Status,Risk,Evidence,Remediation,AssessedAt

## Allowed Severity
Critical,High,Medium,Low,Info

## Allowed Status
Compliant,NonCompliant,Informational,NotAssessed,NotLicensed,PermissionDenied,Error
