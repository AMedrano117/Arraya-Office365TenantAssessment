"""
Configuration loading — Python equivalent of:
  Get-ArrayaAssessmentOutputProfilePolicy
  Get-ArrayaAssessmentOutputRoot
"""

from __future__ import annotations

import json
import os
from pathlib import Path

VALID_PROFILES = frozenset(
    ["Presales", "SolutionsEngineer", "ExecutiveLevel", "TenantToTenantMigration", "Geek", "Machine"]
)

_DEFAULT_POLICY = {
    "ExportCsv": True,
    "ExportJson": True,
    "ExportHtml": False,
    "IncludeRaw": False,
    "GenerateJson": True,
}


def get_repo_root(start: str | Path | None = None) -> Path:
    start = Path(start or Path(__file__).resolve())
    if start.is_file():
        start = start.parent
    for candidate in [start, *start.parents]:
        if (candidate / ".git").exists() or (candidate / "pyproject.toml").exists():
            return candidate
    return start


def get_output_root(fallback: str | Path | None = None) -> Path:
    env_root = os.environ.get("ARRAYA_ASSESSMENT_OUTPUT_ROOT")
    if env_root:
        return Path(env_root)
    repo_root = get_repo_root()
    output_dir = repo_root / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def get_output_profile_policy(profile: str) -> dict:
    profile = profile.strip()
    if profile not in VALID_PROFILES:
        raise ValueError(
            f"Invalid output profile '{profile}'. Valid: {sorted(VALID_PROFILES)}"
        )

    baseline_path = get_repo_root() / "src" / "config" / "baseline" / "default-report-settings.json"
    if baseline_path.exists():
        try:
            with baseline_path.open(encoding="utf-8") as f:
                base = json.load(f)
            policy = dict(_DEFAULT_POLICY)
            policy.update(base)
            return policy
        except Exception:
            pass

    return dict(_DEFAULT_POLICY)


def load_baseline_json(filename: str) -> dict:
    path = get_repo_root() / "src" / "config" / "baseline" / filename
    if not path.exists():
        return {}
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def parse_profiles(raw: str | list[str]) -> list[str]:
    if isinstance(raw, list):
        tokens = raw
    else:
        tokens = [raw]
    result: list[str] = []
    seen: set[str] = set()
    for token in tokens:
        for part in str(token).split(","):
            p = part.strip()
            if p and p not in seen:
                result.append(p)
                seen.add(p)
    return result or ["SolutionsEngineer"]
