"""
Improvement plan generation — Python equivalent of New-M365TenantImprovementPlan.
Reads findings from snapshot and writes structured remediation documents.
"""

from __future__ import annotations

import json
import logging
from pathlib import Path

from ..snapshot import get_data, get_metadata
from ..utils.converters import ensure_list

log = logging.getLogger(__name__)

_SEVERITY_ORDER = {"Risk": 0, "Warning": 1, "Info": 2}


def generate(snapshot: dict, output_dir: str | Path) -> None:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    findings = _extract_findings(snapshot)
    if not findings:
        log.warning("No findings in snapshot — improvement plan will be empty.")

    findings.sort(key=lambda f: _SEVERITY_ORDER.get(str(f.get("Severity", "Info")), 99))

    _write_json_plan(findings, output_dir / "improvement-plan.json")
    _write_markdown_plan(findings, snapshot, output_dir / "improvement-plan.md")
    log.info("Improvement plan written to: %s", output_dir)


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------

def _extract_findings(snapshot: dict) -> list[dict]:
    findings: list[dict] = []
    seen_keys: set[str] = set()

    def _add(collection) -> None:
        if isinstance(collection, dict):
            # Keyed findings dict — e.g. {"032-Hybrid": {Area:..., Severity:..., ...}}
            for key, val in collection.items():
                if isinstance(val, dict) and key not in seen_keys:
                    seen_keys.add(key)
                    findings.append(val)
        else:
            for item in ensure_list(collection):
                if isinstance(item, dict):
                    findings.append(item)

    # Individual findings live in Derived.BestPracticeFindings (preferred) or Derived.Findings.
    # Both are identical dicts in the current schema — only consume one to avoid duplicates.
    derived = snapshot.get("Derived") or {}
    primary = derived.get("BestPracticeFindings") or derived.get("Findings")
    if primary:
        _add(primary)

    # Also check per-section findings in Data (for future schema extensions).
    data = get_data(snapshot)
    for section in data.values():
        if not isinstance(section, dict):
            continue
        # Skip BestPractices — it is an area-level rollup, not individual findings.
        for key in ("BestPracticeFindings", "Findings"):
            if key in section:
                _add(section[key])

    return findings


# ---------------------------------------------------------------------------
# Writers
# ---------------------------------------------------------------------------

def _write_json_plan(findings: list[dict], path: Path) -> None:
    plan = {
        "Version": "1.0",
        "FindingCount": len(findings),
        "Findings": findings,
    }
    path.write_text(json.dumps(plan, indent=2, default=str), encoding="utf-8")
    log.debug("JSON plan: %s", path)


def _write_markdown_plan(findings: list[dict], snapshot: dict, path: Path) -> None:
    meta = get_metadata(snapshot)
    tenant_obj = meta.get("Tenant") or {}
    tenant = (
        (tenant_obj.get("DisplayName") if isinstance(tenant_obj, dict) else None)
        or meta.get("TenantDisplayName")
        or meta.get("TenantDomain")
        or "Tenant"
    )
    lines: list[str] = [
        f"# Improvement Plan - {tenant}",
        "",
        f"> Generated: {meta.get('GeneratedAt', '')}  |  Findings: {len(findings)}",
        "",
    ]

    by_severity: dict[str, list[dict]] = {}
    for f in findings:
        sev = str(f.get("Severity", "Info"))
        by_severity.setdefault(sev, []).append(f)

    for sev in sorted(_SEVERITY_ORDER, key=lambda s: _SEVERITY_ORDER[s]):
        group = by_severity.get(sev)
        if not group:
            continue
        lines += [f"## {sev} ({len(group)})", ""]
        for finding in group:
            area = finding.get("Area", "")
            category = finding.get("Category", "")
            title = f"{area} - {category}" if area and category else (area or category or "Untitled Finding")
            message = finding.get("Message") or finding.get("Description") or finding.get("Details") or ""
            remediation = finding.get("RecommendedAction") or finding.get("Remediation") or ""
            lines += [
                f"### {title}",
                "",
                message,
                "",
            ]
            if remediation:
                lines += ["**Recommended Action:**", "", remediation, ""]
            lines.append("---")
            lines.append("")

    path.write_text("\n".join(lines), encoding="utf-8")
    log.debug("Markdown plan: %s", path)
