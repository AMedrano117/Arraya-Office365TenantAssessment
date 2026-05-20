"""
Engineer Action Pack generation — Python equivalent of New-EngineerActionPack.
Generates EngPack.md from plan data (produced by analysis.plan.generate).
"""

from __future__ import annotations

import logging
from pathlib import Path

from ..snapshot import get_metadata

log = logging.getLogger(__name__)

_PHASE_ORDER: dict[str, int] = {"Immediate": 0, "Near Term": 1, "Planned": 2, "Monitor": 3}
_SEV_ORDER: dict[str, int] = {"High": 0, "Medium": 1, "Low": 2}


def generate(
    plan: dict,
    snapshot: dict,
    output_path: Path,
    snapshot_path: str | None = None,
    support_dir: Path | None = None,
    coverage: dict | None = None,
) -> None:
    """Write EngPack.md to output_path."""
    output_path = Path(output_path)
    if output_path.suffix.lower() != ".md":
        output_path = output_path.with_suffix(".md")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    meta = get_metadata(snapshot)
    tenant_obj = meta.get("Tenant") or {}
    tenant_name = (
        (tenant_obj.get("DisplayName") if isinstance(tenant_obj, dict) else None)
        or meta.get("TenantDisplayName")
        or plan.get("TenantName")
        or "Tenant"
    )
    generated_at = plan.get("GeneratedAt", meta.get("GeneratedAt", ""))

    findings: list[dict] = plan.get("Findings", [])
    findings_sorted = sorted(
        findings,
        key=lambda f: (
            _SEV_ORDER.get(f.get("Severity", "Low"), 99),
            _PHASE_ORDER.get(f.get("RoadmapPhase", "Monitor"), 99),
        ),
    )
    ws_summaries: list[dict] = plan.get("WorkstreamSummaries", [])

    header_lines: list[str] = [
        f"# {tenant_name} Engineer Action Pack",
        "",
        f"Generated: {generated_at}",
    ]
    if snapshot_path:
        header_lines.append(f"Source snapshot: {snapshot_path}")
    header_lines += [
        "",
        "Use this pack to review the current tenant state, confirm each finding, "
        "and plan the remediation work in a way that fits your change process.",
        "",
    ]
    lines: list[str] = header_lines

    # Engineering Summary
    phase_counts: dict[str, int] = {p: 0 for p in ("Immediate", "Near Term", "Planned", "Monitor")}
    for f in findings:
        phase = f.get("RoadmapPhase", "Monitor")
        if phase in phase_counts:
            phase_counts[phase] += 1

    lines += [
        "## Engineering Summary",
        "",
        f"- Total findings: {len(findings)}",
        f"- Immediate: {phase_counts['Immediate']}",
        f"- Near Term: {phase_counts['Near Term']}",
        f"- Planned: {phase_counts['Planned']}",
        f"- Monitor: {phase_counts['Monitor']}",
        "",
    ]

    # Workstream Summary table
    if ws_summaries:
        lines += [
            "## Workstream Summary",
            "",
            "| Severity | Workstream | Area | Open Findings | Top Signals |",
            "|---|---|---|---|---|",
        ]
        for ws in ws_summaries:
            lines.append(
                f"| {ws.get('Severity','')} "
                f"| {ws.get('Workstream','')} "
                f"| {ws.get('Area','')} "
                f"| {ws.get('OpenFindings',0)} "
                f"| {_md(ws.get('TopSignals',''))} |"
            )
        lines.append("")

    # Before You Start
    lines += [
        "## Before You Start",
        "",
        "- Validate tenant admin roles, Graph scopes, Exchange connectivity, "
        "and any pilot exclusions before enforcement changes.",
        "- Use the support artifacts for backlog import, validation context, "
        "and change-record attachment as needed.",
        "",
    ]

    # Findings To Work
    lines += [
        "## Findings To Work",
        "",
        "| Severity | Phase | Workstream | Reference | What Needs Attention | "
        "Why It Matters | Technical Remediation | Current Evidence | Where To Verify | Done When |",
        "|---|---|---|---|---|---|---|---|---|---|",
    ]
    for f in findings_sorted:
        lines.append(
            f"| {_md(f.get('Severity'))} "
            f"| {_md(f.get('RoadmapPhase'))} "
            f"| {_md(f.get('Workstream'))} "
            f"| {_md(f.get('RuleId'))} "
            f"| {_md(f.get('Finding'))} "
            f"| {_md(f.get('WhyFlagged'))} "
            f"| {_md(f.get('TechnicalRemediation'))} "
            f"| {_md(f.get('CurrentEvidence'))} "
            f"| {_md(f.get('EvidenceLocation'))} "
            "| |"
        )

    lines += [
        "",
        "## Verification Checklist",
        "",
        "- Confirm each recommendation against the latest tenant state before making changes.",
        "- Pilot security and access-policy changes with a limited scope before broad enforcement.",
        "- Re-run `M365Collect` and `Improve` after remediation milestones to measure delta and retire closed findings.",
        "",
    ]

    # Assessment Evidence Coverage (from coverage report if available)
    if coverage:
        covered = coverage.get("CoveredObjectiveCount", 0)
        review = coverage.get("ReviewObjectiveCount", 0)
        total = coverage.get("ObjectiveCount", 0)
        lines += [
            "## Assessment Evidence Coverage",
            "",
            f"**{covered} of {total} objectives fully covered** | {review} objective(s) require review",
            "",
            "| Objective | Area | Status | Confidence | Datasets |",
            "|---|---|---|---|---|",
        ]
        for obj in coverage.get("Objectives", []):
            present = obj.get("PresentDatasetCount", 0)
            expected = obj.get("ExpectedDatasetCount", 0)
            lines.append(
                f"| {_md(obj.get('ObjectiveId', ''))} "
                f"| {_md(obj.get('Area', ''))} "
                f"| {_md(obj.get('Status', ''))} "
                f"| {_md(obj.get('Confidence', ''))} "
                f"| {present} of {expected} |"
            )

        gaps = coverage.get("CoverageGaps", [])
        if gaps:
            lines += [
                "",
                "### Coverage Gaps",
                "",
                "| Gap | Area | Limitation |",
                "|---|---|---|",
            ]
            for gap in gaps:
                lines.append(
                    f"| {_md(gap.get('GapId', ''))} "
                    f"| {_md(gap.get('Area', ''))} "
                    f"| {_md(gap.get('CurrentLimitation', ''))} |"
                )
        lines.append("")

    if support_dir:
        stem = output_path.stem.replace("-EngPack", "")
        lines += [
            "## Supporting Files",
            "",
            f"- Support folder: `{support_dir}`",
            f"- Improvement plan JSON: `{support_dir / (stem + '-Plan.json')}`",
            "",
        ]

    output_path.write_text("\n".join(lines), encoding="utf-8")
    log.info("EngPack written: %s (%d findings)", output_path, len(findings))


def _md(value: object) -> str:
    """Sanitize a value for use in a markdown table cell."""
    return str(value or "").replace("|", "/").replace("\n", " ").replace("\r", "")
