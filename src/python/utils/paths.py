"""
Path resolution utilities — Python equivalent of:
  Resolve-ArrayaRepoRoot, Get-ExportPath, Resolve-ArrayaSnapshotOutputContext
"""

from __future__ import annotations

import re
from datetime import datetime, timezone
from pathlib import Path


def find_repo_root(start: str | Path | None = None) -> Path:
    start = Path(start or Path(__file__).resolve())
    if start.is_file():
        start = start.parent
    for candidate in [start, *start.parents]:
        if (candidate / ".git").exists() or (candidate / "pyproject.toml").exists():
            return candidate
    return start


def get_export_path(base_dir: str | Path, stem: str, ext: str) -> Path:
    base_dir = Path(base_dir)
    if not ext.startswith("."):
        ext = f".{ext}"
    return base_dir / f"{stem}{ext}"


def make_run_folder(base_dir: str | Path, label: str = "") -> Path:
    ts = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    folder_name = f"{ts}-{_sanitize(label)}" if label else ts
    run_dir = Path(base_dir) / folder_name
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def resolve_snapshot_output_context(snapshot: dict, export_path: str | Path) -> dict:
    meta = snapshot.get("Metadata") or {}
    generated_at = meta.get("GeneratedAt", "")
    profile = meta.get("OutputProfileLabel", meta.get("OutputProfile", "SolutionsEngineer"))

    export_path = Path(export_path)
    if export_path.suffix:
        output_dir = export_path.parent
        stem = export_path.stem
    else:
        output_dir = export_path
        tenant = _resolve_tenant_name(meta)
        stem = _sanitize(tenant) if tenant else "Assessment"

    return {
        "OutputDirectory": output_dir,
        "FileStem": stem,
        "Profile": profile,
        "GeneratedAt": generated_at,
    }


def _resolve_tenant_name(meta: dict) -> str:
    # Schema v2: Metadata.Tenant.DisplayName
    tenant_obj = meta.get("Tenant") or {}
    if isinstance(tenant_obj, dict):
        for key in ("DisplayName", "DefaultDomainName", "TenantId"):
            val = tenant_obj.get(key)
            if val and str(val).strip():
                return str(val).strip()
    # Flat fallback fields
    for key in ("TenantDisplayName", "TenantDomain", "OutputProfileLabel"):
        val = meta.get(key)
        if val and str(val).strip():
            return str(val).strip()
    return ""


def _sanitize(text: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(text)).strip("_. ")
