"""
Snapshot comparison — Python equivalent of:
  Compare-M365TenantAssessmentSnapshots / Invoke-M365TenantAssessmentComparison
"""

from __future__ import annotations

from typing import Any

from ..snapshot import get_data, get_metadata


def compare(baseline: dict, current: dict) -> dict:
    delta: dict[str, Any] = {
        "Metadata": {
            "BaselineGeneratedAt": get_metadata(baseline).get("GeneratedAt"),
            "CurrentGeneratedAt": get_metadata(current).get("GeneratedAt"),
            "BaselineProfile": get_metadata(baseline).get("OutputProfileLabel"),
            "CurrentProfile": get_metadata(current).get("OutputProfileLabel"),
        },
        "Changes": {},
        "Summary": {"Added": 0, "Removed": 0, "Changed": 0, "Unchanged": 0},
    }

    base_data = get_data(baseline)
    curr_data = get_data(current)

    all_sections = set(base_data) | set(curr_data)
    for section in sorted(all_sections):
        base_section = base_data.get(section) or {}
        curr_section = curr_data.get(section) or {}
        section_delta = _diff_dict(base_section, curr_section)
        if section_delta["changes"]:
            delta["Changes"][section] = section_delta
            delta["Summary"]["Added"] += section_delta["added"]
            delta["Summary"]["Removed"] += section_delta["removed"]
            delta["Summary"]["Changed"] += section_delta["changed"]
            delta["Summary"]["Unchanged"] += section_delta["unchanged"]

    return delta


def _diff_dict(base: dict, curr: dict) -> dict:
    changes: list[dict] = []
    added = removed = changed = unchanged = 0

    all_keys = set(base) | set(curr)
    for key in sorted(all_keys):
        in_base = key in base
        in_curr = key in curr

        if in_base and not in_curr:
            changes.append({"Key": key, "Change": "Removed", "Before": _summarize(base[key]), "After": None})
            removed += 1
        elif in_curr and not in_base:
            changes.append({"Key": key, "Change": "Added", "Before": None, "After": _summarize(curr[key])})
            added += 1
        else:
            before = _summarize(base[key])
            after = _summarize(curr[key])
            if before != after:
                changes.append({"Key": key, "Change": "Modified", "Before": before, "After": after})
                changed += 1
            else:
                unchanged += 1

    return {"changes": changes, "added": added, "removed": removed, "changed": changed, "unchanged": unchanged}


def _summarize(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (str, bool, int, float)):
        return value
    if isinstance(value, dict):
        return f"<dict keys={sorted(value.keys())}>"
    if isinstance(value, list):
        return f"<list len={len(value)}>"
    return str(value)
