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

_SEVERITY_ORDER = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3, "Info": 4}


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
    data = get_data(snapshot)
    findings: list[dict] = []

    for section in data.values():
        if not isinstance(section, dict):
            continue
        for key in ("BestPracticeFindings", "BestPractices", "Findings"):
            if key in section:
                findings.extend(ensure_list(section[key]))

    derived = snapshot.get("Derived") or {}
    for key in ("BestPracticeFindings", "Findings"):
        if key in derived:
            findings.extend(ensure_list(derived[key]))

    return [f for f in findings if isinstance(f, dict)]


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
    tenant = meta.get("TenantDisplayName") or meta.get("TenantDomain") or "Tenant"
    lines: list[str] = [
        f"# Improvement Plan — {tenant}",
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
            title = finding.get("Title") or finding.get("Name") or "Untitled Finding"
            description = finding.get("Description") or finding.get("Details") or ""
            remediation = finding.get("Remediation") or finding.get("RecommendedAction") or ""
            lines += [
                f"### {title}",
                "",
                description,
                "",
            ]
            if remediation:
                lines += ["**Remediation:**", "", remediation, ""]
            lines.append("---")
            lines.append("")

    path.write_text("\n".join(lines), encoding="utf-8")
    log.debug("Markdown plan: %s", path)
