# Repo Overview

Primary structure:
- `src/modules/Arraya.M365.Common`: shared helpers and local module bootstrapping.
- `src/modules/Arraya.M365.Graph`: Graph helper functions.
- `src/modules/Arraya.M365.Reporting`: reporting helper functions.
- `src/modules/Arraya.M365.AssessmentRunner`: user-facing commands that launch tenant/AD/graph/improvement workflows.
- `src/scripts/assessments`: entry scripts for tenant-wide and identity assessments.
- `src/scripts/reporting`: improvement and comparison reporting scripts.
- `src/scripts/operations`: menu launcher script for user-driven execution.
- `src/scripts/migrated/legacy`: migrated scripts preserved for compatibility.
- `src/vendor/Office365Custom/1.2.0`: vendored shared function module.
