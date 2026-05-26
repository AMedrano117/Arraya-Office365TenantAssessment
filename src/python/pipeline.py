"""
Export pipeline — drives Python report generation from a snapshot.
Equivalent of Invoke-M365TenantAssessmentExportPipeline.
"""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path

from rich.console import Console

from . import snapshot as snap
from .analysis import plan as plan_module, evidence_coverage as coverage_module
from .reporting import excel, html, docx, markdown, improvement_plan, engpack, roadmap
from .utils.paths import resolve_snapshot_output_context

log = logging.getLogger(__name__)
_console = Console()


def run(
    snapshot_path: str | Path,
    output_dir: str | Path | None = None,
    profiles: list[str] | None = None,
    skip_excel: bool = False,
    skip_html: bool = False,
    skip_docx: bool = False,
    skip_improve: bool = False,
    docx_template: str | Path | None = None,
    verbose: bool = False,
) -> dict[str, Path]:
    snapshot_path = Path(snapshot_path)
    snapshot_data = snap.load(snapshot_path)
    if snapshot_data is None:
        raise FileNotFoundError(f"Snapshot not found: {snapshot_path}")

    # If no explicit output_dir and the snapshot lives in a Support folder,
    # step up to the assessment root so outputs land beside Support, not inside it.
    if output_dir:
        base_dir = Path(output_dir)
    else:
        snap_parent = snapshot_path.parent
        base_dir = snap_parent.parent if snap_parent.name.lower() == "support" else snap_parent

    ctx = resolve_snapshot_output_context(snapshot_data, base_dir)
    stem: str = ctx["FileStem"]

    # Date-stamp each run so multiple runs are easy to distinguish
    run_stamp = datetime.now().strftime("%Y-%m-%d-%H%M")
    out_dir   = ctx["OutputDirectory"] / run_stamp

    deliverables = out_dir / "Deliverables"
    support      = out_dir / "Support"
    deliverables.mkdir(parents=True, exist_ok=True)
    support.mkdir(parents=True, exist_ok=True)

    artifacts: dict[str, Path] = {}

    # -----------------------------------------------------------------------
    # Analysis: generate enriched Plan.json (always — other steps depend on it)
    # -----------------------------------------------------------------------
    _console.print("  [dim]Analyzing snapshot...[/dim]")
    log.info("Generating plan: %s", support)
    generated_plan = plan_module.generate(snapshot_data, support, tenant_name=stem)
    artifacts["plan"] = support / f"{stem}-Plan.json"
    if verbose:
        _console.print(f"  [dim]Plan[/dim]      {artifacts['plan'].name}")

    # -----------------------------------------------------------------------
    # Reports
    # -----------------------------------------------------------------------
    _console.print("  [dim]Generating reports...[/dim]")
    if not skip_excel:
        xlsx_path = deliverables / f"{stem}-{run_stamp}-Tenant Details.xlsx"
        log.info("Generating Excel workbook: %s", xlsx_path)
        _flatten_and_export_excel(snapshot_data, xlsx_path)
        artifacts["excel"] = xlsx_path
        if verbose:
            _console.print(f"  [dim]Excel[/dim]     {xlsx_path.name}")

    if not skip_html:
        html_path = deliverables / f"{stem}-{run_stamp}-Report.html"
        log.info("Generating HTML report: %s", html_path)
        html.generate(snapshot_data, html_path, plan=generated_plan)
        artifacts["html"] = html_path
        if verbose:
            _console.print(f"  [dim]HTML[/dim]      {html_path.name}")

    if not skip_docx:
        docx_path = deliverables / f"{stem}-{run_stamp}-Best Practices Assessment.docx"
        log.info("Generating Word document: %s", docx_path)
        docx.generate(snapshot_data, docx_path, template_path=docx_template, plan=generated_plan)
        artifacts["docx"] = docx_path
        if verbose:
            _console.print(f"  [dim]Word[/dim]      {docx_path.name}")

    if not skip_improve:
        # Evidence coverage JSON (support artifact for SE review)
        coverage_data = coverage_module.generate(
            snapshot_data,
            support,
            tenant_name=stem,
            snapshot_path=str(snapshot_path),
        )
        coverage_path = support / f"{stem}-SolutionsEngineerEvidenceCoverage.json"
        artifacts["coverage"] = coverage_path
        if verbose:
            _console.print(f"  [dim]Coverage[/dim]  {coverage_path.name}")

        # EngPack.docx (engineer-facing deliverable)
        engpack_path = deliverables / f"{stem}-{run_stamp}-EngPack.docx"
        log.info("Generating EngPack: %s", engpack_path)
        engpack.generate(
            generated_plan,
            snapshot_data,
            engpack_path,
            snapshot_path=str(snapshot_path),
            support_dir=support,
            coverage=coverage_data,
        )
        artifacts["engpack"] = engpack_path
        if verbose:
            _console.print(f"  [dim]EngPack[/dim]   {engpack_path.name}")

        # Remediation Roadmap Word doc
        roadmap_path = deliverables / f"{stem}-{run_stamp}-Microsoft 365 Remediation Roadmap.docx"
        log.info("Generating roadmap: %s", roadmap_path)
        roadmap.generate(generated_plan, snapshot_data, roadmap_path)
        artifacts["roadmap"] = roadmap_path
        if verbose:
            _console.print(f"  [dim]Roadmap[/dim]   {roadmap_path.name}")

        # Basic improvement plan markdown (lightweight summary)
        improve_dir = out_dir / "ImprovementPlan"
        log.info("Generating improvement plan: %s", improve_dir)
        improvement_plan.generate(snapshot_data, improve_dir)
        artifacts["improve_dir"] = improve_dir
        if verbose:
            _console.print(f"  [dim]Improve[/dim]   {improve_dir.name}/")

    log.info("Pipeline complete. Artifacts: %s", list(artifacts.values()))
    return artifacts


def _flatten_and_export_excel(snapshot_data: dict, xlsx_path: Path) -> None:
    flat: dict = {}
    data = snapshot_data.get("Data") or {}
    for section in data.values():
        if isinstance(section, dict):
            flat.update(section)
    derived = snapshot_data.get("Derived") or {}
    flat.update(derived)

    _enrich_teams(flat)
    _build_mailbox_summary(flat)

    excel.export(flat, xlsx_path)


def _enrich_teams(flat: dict) -> None:
    """Merge TeamOwners, TeamMemberCounts, PrivateChannels into each AllTeams entry."""
    all_teams = flat.get("AllTeams")
    if not isinstance(all_teams, dict) or not all_teams:
        return
    owners_by_team   = flat.get("TeamOwners") or {}
    counts_by_team   = flat.get("TeamMemberCounts") or {}
    channels_by_team = flat.get("PrivateChannels") or {}

    enriched: dict = {}
    for tid, team in all_teams.items():
        owners = owners_by_team.get(tid) or []
        entry  = dict(team)
        entry["MemberCount"]         = counts_by_team.get(tid)
        entry["OwnerCount"]          = len(owners)
        entry["OwnerNames"]          = "; ".join(
            o.get("displayName", "") for o in owners if isinstance(o, dict)
        )
        entry["PrivateChannelCount"] = len(channels_by_team.get(tid) or [])
        enriched[tid] = entry
    flat["AllTeams"] = enriched


def _build_mailbox_summary(flat: dict) -> None:
    """
    1. Enrich SharedMailboxes (bare UPN list) with display names and storage data.
    2. Build SharedMailboxGovernanceSummary — mailbox type counts.
    3. Build NonUserMailboxes — combined enriched list (shared + equipment + room).
    """
    shared_upns   = flat.get("SharedMailboxes") or []
    equipment_upns = flat.get("EquipmentMailboxes") or []
    room_mbx      = flat.get("RoomMailboxes") or []
    mbx_details   = flat.get("MailboxUsageDetails") or []
    users         = flat.get("Users") or {}

    # Already enriched (re-entrant guard)
    if shared_upns and isinstance(shared_upns[0], dict):
        return

    # Build UPN -> mailbox usage lookup (lowercase for case-insensitive match)
    mbx_by_upn: dict = {
        r.get("UserPrincipalName", "").lower(): r
        for r in mbx_details if r.get("UserPrincipalName")
    }

    def _enrich(upn: str, mailbox_type: str) -> dict:
        user_obj = users.get(upn) or {}
        mbx      = mbx_by_upn.get(upn.lower()) or {}
        return {
            "UserPrincipalName": upn,
            "DisplayName":       user_obj.get("displayName") or mbx.get("DisplayName", ""),
            "Department":        user_obj.get("department", ""),
            "AccountEnabled":    user_obj.get("accountEnabled"),
            "MailboxType":       mailbox_type,
            "StorageUsedGB":     mbx.get("StorageUsedGB"),
            "StorageQuotaGB":    mbx.get("StorageQuotaGB"),
            "ItemCount":         mbx.get("ItemCount"),
            "HasArchive":        mbx.get("HasArchive"),
            "LastActivityDate":  mbx.get("LastActivityDate"),
        }

    enriched_shared    = [_enrich(upn, "Shared")    for upn in shared_upns    if isinstance(upn, str)]
    enriched_equipment = [_enrich(upn, "Equipment") for upn in equipment_upns if isinstance(upn, str)]

    enriched_rooms = [
        {
            "UserPrincipalName": r.get("emailAddress", ""),
            "DisplayName":       r.get("displayName", ""),
            "Department":        r.get("building", ""),
            "AccountEnabled":    None,
            "MailboxType":       "Room",
            "StorageUsedGB":     None,
            "StorageQuotaGB":    None,
            "ItemCount":         None,
            "HasArchive":        None,
            "LastActivityDate":  None,
        }
        for r in room_mbx if isinstance(r, dict)
    ]

    flat["SharedMailboxes"]   = enriched_shared
    flat["EquipmentMailboxes"] = enriched_equipment

    # NonUserMailboxes — combined list for the existing Excel slot
    flat["NonUserMailboxes"] = enriched_shared + enriched_equipment + enriched_rooms

    # Mailbox type summary
    total_reported = flat.get("PrimaryMailboxStats", {}).get("MailboxCount") or 0
    n_shared    = len(enriched_shared)
    n_equipment = len(enriched_equipment)
    n_room      = len(enriched_rooms)
    flat["SharedMailboxGovernanceSummary"] = {
        "TotalMailboxesReported": total_reported,
        "UserMailboxes":          max(total_reported - n_shared - n_equipment, 0),
        # SharedMailboxCount mirrors the PS snapshot key name used by plan.py
        "SharedMailboxCount":     n_shared,
        "SharedMailboxes":        n_shared,
        "EquipmentMailboxes":     n_equipment,
        "RoomMailboxes":          n_room,
        "TotalNonUserMailboxes":  n_shared + n_equipment + n_room,
    }
