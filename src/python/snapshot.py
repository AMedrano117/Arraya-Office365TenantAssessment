"""
Snapshot load/save — Python equivalent of Import/Export-ArrayaTenantSnapshot
and New/Update-ArrayaTenantSnapshot.

Snapshot schema v2:
{
  "SchemaVersion": 2,
  "Metadata": { "GeneratedAt": "<ISO8601>", ... },
  "CollectionPlan": { ... },
  "Data": {
    "Exchange": {}, "Identity": {}, "Collaboration": {},
    "Security": {}, "Governance": {}, "Tenant": {}, "Other": {}
  },
  "Derived": {},
  "Diagnostics": {
    "WarningCount": 0, "ErrorCount": 0,
    "WarningSummary": [], "ErrorSummary": [],
    "SourceCoverage": {}, "CollectorStats": {}
  }
}
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def load(path: str | Path, skip_validation: bool = False) -> dict | None:
    path = Path(path)
    if not path.exists():
        return None

    with path.open(encoding="utf-8") as f:
        root = json.load(f)

    if not isinstance(root, dict):
        raise ValueError(f"Snapshot file did not deserialize to an object: {path}")

    schema_version = 0
    try:
        schema_version = int(root.get("SchemaVersion", 0))
    except (TypeError, ValueError):
        schema_version = 0

    if schema_version >= 2:
        snapshot = root
    else:
        snapshot = _migrate_legacy(root)

    if not skip_validation:
        errors = _validate(snapshot)
        if errors:
            raise ValueError("Imported snapshot failed validation: " + "; ".join(errors))

    return snapshot


def save(snapshot: dict, path: str | Path) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    normalized = dict(snapshot)
    normalized.setdefault("SchemaVersion", 2)

    tmp = path.with_suffix(".tmp")
    try:
        with tmp.open("w", encoding="utf-8") as f:
            json.dump(normalized, f, indent=2, default=_json_default, ensure_ascii=False)
        tmp.replace(path)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise


def create(
    metadata: dict | None = None,
    collection_plan: dict | None = None,
    data: dict | None = None,
    derived: dict | None = None,
    diagnostics: dict | None = None,
) -> dict:
    snapshot: dict[str, Any] = {
        "SchemaVersion": 2,
        "Metadata": {"GeneratedAt": datetime.now(timezone.utc).isoformat()},
        "CollectionPlan": {},
        "Data": {
            "Exchange": {},
            "Identity": {},
            "Collaboration": {},
            "Security": {},
            "Governance": {},
            "Tenant": {},
            "Other": {},
        },
        "Derived": {},
        "Diagnostics": {
            "WarningCount": 0,
            "ErrorCount": 0,
            "WarningSummary": [],
            "ErrorSummary": [],
            "SourceCoverage": {},
            "CollectorStats": {},
        },
    }

    if metadata:
        snapshot["Metadata"].update(metadata)
    if collection_plan:
        snapshot["CollectionPlan"].update(collection_plan)
    if data:
        snapshot["Data"].update(data)
    if derived:
        snapshot["Derived"].update(derived)
    if diagnostics:
        snapshot["Diagnostics"].update(diagnostics)

    return snapshot


def get_data(snapshot: dict) -> dict:
    return snapshot.get("Data") or {}


def get_metadata(snapshot: dict) -> dict:
    return snapshot.get("Metadata") or {}


def get_section(snapshot: dict, section: str) -> dict:
    return get_data(snapshot).get(section) or {}


# ---------------------------------------------------------------------------
# Internals
# ---------------------------------------------------------------------------

def _migrate_legacy(root: dict) -> dict:
    if "Data" in root and isinstance(root["Data"], dict):
        legacy_data = dict(root["Data"])
    else:
        legacy_data = {k: v for k, v in root.items() if k != "SchemaVersion"}

    meta: dict[str, Any] = {}
    for field, dest in (("GeneratedAt", "GeneratedAt"), ("OutputProfile", "OutputProfileLabel"), ("ReportingMode", "ReportingMode")):
        if field in root:
            meta[dest] = root[field]

    snapshot = create(metadata=meta)
    snapshot["Data"]["Other"].update(legacy_data)
    return snapshot


def _validate(snapshot: dict) -> list[str]:
    errors: list[str] = []
    if not isinstance(snapshot, dict):
        errors.append("Snapshot is not a dict")
        return errors
    if snapshot.get("SchemaVersion", 0) < 2:
        errors.append(f"Unsupported SchemaVersion: {snapshot.get('SchemaVersion')}")
    if "Data" not in snapshot:
        errors.append("Missing 'Data' section")
    return errors


def _json_default(obj: Any) -> Any:
    if isinstance(obj, datetime):
        return obj.isoformat()
    if hasattr(obj, "__dict__"):
        return obj.__dict__
    return str(obj)
