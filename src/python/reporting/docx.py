"""
Word document generation — Python equivalent of CustomerAssessmentDocx.ps1.
Uses python-docx to fill a template or generate from scratch.

Document structure:
  1. Cover
  2. Executive Summary  (WorkstreamSummaries + phase counts + ConsultativeSummaries)
  3. Workstream Reviews (narrative + top findings per workstream)
  4. Findings           (full findings list by severity)
  5. Appendix           (raw collected data tables)
"""

from __future__ import annotations

import logging
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor

from ..snapshot import get_data, get_metadata
from ..utils.converters import ensure_list

log = logging.getLogger(__name__)

_NAVY = RGBColor(0x1F, 0x38, 0x64)
_MAX_APPENDIX_ROWS = 500

_APPENDIX_EXCLUDED = {
    "AllMailboxes-MailIdentity",
    "AllMailboxes-UserPrincipalName",
    "AllMailboxes-PrimarySmtpAddress",
}

_PHASE_ORDER: dict[str, int] = {"Immediate": 0, "Near Term": 1, "Planned": 2, "Monitor": 3}
_SEV_ORDER: dict[str, int]   = {"High": 0, "Medium": 1, "Low": 2}


def generate(
    snapshot: dict,
    output_path: str | Path,
    template_path: str | Path | None = None,
    plan: dict | None = None,
) -> None:
    output_path = Path(output_path)
    if output_path.suffix.lower() != ".docx":
        output_path = output_path.with_suffix(".docx")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if template_path and Path(template_path).exists():
        doc = Document(str(template_path))
    else:
        doc = Document()
        _setup_styles(doc)

    _write_cover(doc, snapshot)
    _write_executive_summary(doc, snapshot, plan)
    _write_workstream_reviews(doc, plan)
    _write_findings(doc, plan)
    _write_appendix(doc, snapshot)

    doc.save(str(output_path))
    log.info("Word document saved: %s", output_path)


# ---------------------------------------------------------------------------
# Style setup
# ---------------------------------------------------------------------------

def _setup_styles(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)


# ---------------------------------------------------------------------------
# Cover
# ---------------------------------------------------------------------------

def _write_cover(doc: Document, snapshot: dict) -> None:
    meta = get_metadata(snapshot)
    tenant_obj = meta.get("Tenant") or {}
    tenant = (
        (tenant_obj.get("DisplayName") if isinstance(tenant_obj, dict) else None)
        or meta.get("TenantDisplayName")
        or meta.get("TenantDomain")
        or "Tenant"
    )
    generated_at = meta.get("GeneratedAt", "")

    p = doc.add_heading("Microsoft 365 Tenant Assessment", level=0)
    p.runs[0].font.color.rgb = _NAVY

    info = doc.add_paragraph()
    info.add_run(f"Tenant: {tenant}").bold = True
    doc.add_paragraph(f"Generated: {generated_at}")
    doc.add_page_break()


# ---------------------------------------------------------------------------
# Executive Summary
# ---------------------------------------------------------------------------

def _write_executive_summary(doc: Document, snapshot: dict, plan: dict | None) -> None:
    h = doc.add_heading("Executive Summary", level=1)
    h.runs[0].font.color.rgb = _NAVY

    findings: list[dict] = (plan or {}).get("Findings", [])
    ws_summaries: list[dict] = (plan or {}).get("WorkstreamSummaries", [])

    if not findings:
        # Fallback when no plan data
        doc.add_paragraph("No findings were identified in this assessment.")
        doc.add_page_break()
        return

    # Phase summary counts
    phase_counts: dict[str, int] = {p: 0 for p in ("Immediate", "Near Term", "Planned", "Monitor")}
    for f in findings:
        p = f.get("RoadmapPhase", "Monitor")
        if p in phase_counts:
            phase_counts[p] += 1

    doc.add_paragraph(
        f"This assessment identified {len(findings)} finding(s) across "
        f"{len({f.get('Workstream') for f in findings})} workstream(s)."
    )
    doc.add_paragraph(
        f"Immediate: {phase_counts['Immediate']}  |  "
        f"Near Term: {phase_counts['Near Term']}  |  "
        f"Planned: {phase_counts['Planned']}  |  "
        f"Monitor: {phase_counts['Monitor']}"
    )

    # Workstream summary table
    if ws_summaries:
        doc.add_paragraph()
        p = doc.add_paragraph()
        p.add_run("Assessment Workstream Overview").bold = True

        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        for idx, label in enumerate(("Severity", "Workstream", "Area", "Open Findings")):
            cell = tbl.rows[0].cells[idx]
            cell.text = label
            run = cell.paragraphs[0].runs[0]
            run.bold = True
            run.font.color.rgb = _NAVY

        for ws in ws_summaries:
            row = tbl.add_row().cells
            row[0].text = ws.get("Severity", "")
            row[1].text = ws.get("Workstream", "")
            row[2].text = ws.get("Area", "")
            row[3].text = str(ws.get("OpenFindings", 0))

    doc.add_page_break()


# ---------------------------------------------------------------------------
# Workstream Reviews (narrative per workstream)
# ---------------------------------------------------------------------------

def _write_workstream_reviews(doc: Document, plan: dict | None) -> None:
    if not plan:
        return

    consultative: dict = plan.get("ConsultativeSummaries", {})
    findings: list[dict] = plan.get("Findings", [])
    if not consultative and not findings:
        return

    h = doc.add_heading("Workstream Reviews", level=1)
    h.runs[0].font.color.rgb = _NAVY

    # Collect all workstreams present
    workstreams = sorted(
        {f.get("Workstream", "Governance") for f in findings},
        key=lambda ws: min(
            (_SEV_ORDER.get(f.get("Severity", "Low"), 99) for f in findings if f.get("Workstream") == ws),
            default=99,
        ),
    )

    for ws in workstreams:
        ws_key = f"{ws}ConsultativeSummary"
        summary = consultative.get(ws_key, {})
        narrative = summary.get("Narrative", "")
        ws_findings = [f for f in findings if f.get("Workstream") == ws]

        sh = doc.add_heading(ws, level=2)
        sh.runs[0].font.color.rgb = _NAVY

        if narrative:
            doc.add_paragraph(narrative)

        # Top High-severity findings for this workstream
        high_findings = [f for f in ws_findings if f.get("Severity") == "High"]
        if high_findings:
            p = doc.add_paragraph()
            p.add_run("Priority Items").bold = True
            for f in high_findings[:5]:
                area     = f.get("Area", "")
                category = f.get("Category", "")
                title    = f"{area} - {category}" if area and category else (area or category or "Finding")
                evidence = f.get("Finding", "")
                doc.add_paragraph(f"{title}: {evidence}", style="List Bullet")

        # Snapshot rows (signal / state pairs)
        snapshot_rows = summary.get("SnapshotRows", [])
        if snapshot_rows:
            doc.add_paragraph()
            tbl = doc.add_table(rows=1, cols=2)
            tbl.style = "Table Grid"
            for idx, label in enumerate(("Signal", "Current State")):
                cell = tbl.rows[0].cells[idx]
                cell.text = label
                run = cell.paragraphs[0].runs[0]
                run.bold = True
                run.font.color.rgb = _NAVY
            for row_data in snapshot_rows:
                row = tbl.add_row().cells
                row[0].text = str(row_data.get("Signal", ""))
                row[1].text = str(row_data.get("State", ""))

    doc.add_page_break()


# ---------------------------------------------------------------------------
# Findings Detail
# ---------------------------------------------------------------------------

def _write_findings(doc: Document, plan: dict | None) -> None:
    findings = (plan or {}).get("Findings", [])
    if not findings:
        return

    findings_sorted = sorted(
        findings,
        key=lambda f: (
            _SEV_ORDER.get(f.get("Severity", "Low"), 99),
            _PHASE_ORDER.get(f.get("RoadmapPhase", "Monitor"), 99),
        ),
    )

    h = doc.add_heading("Findings", level=1)
    h.runs[0].font.color.rgb = _NAVY

    by_phase: dict[str, list[dict]] = {}
    for f in findings_sorted:
        phase = f.get("RoadmapPhase", "Monitor")
        by_phase.setdefault(phase, []).append(f)

    for phase in ("Immediate", "Near Term", "Planned", "Monitor"):
        group = by_phase.get(phase)
        if not group:
            continue

        ph = doc.add_heading(f"{phase} ({len(group)})", level=2)
        ph.runs[0].font.color.rgb = _NAVY

        for finding in group:
            area     = finding.get("Area", "")
            category = finding.get("Category", "")
            title    = f"{area} - {category}" if area and category else (area or category or "Finding")
            message  = finding.get("Finding", "")
            rec      = finding.get("TechnicalRemediation", "")
            evidence = finding.get("EvidenceLocation", "")

            fh = doc.add_heading(title, level=3)
            fh.runs[0].font.color.rgb = _NAVY

            if message:
                doc.add_paragraph(message)

            if rec:
                p = doc.add_paragraph()
                p.add_run("Recommended Action: ").bold = True
                p.add_run(rec)

            if evidence:
                p = doc.add_paragraph()
                p.add_run("Where to Verify: ").bold = True
                p.add_run(evidence)

    doc.add_page_break()


# ---------------------------------------------------------------------------
# Data Appendix
# ---------------------------------------------------------------------------

def _write_appendix(doc: Document, snapshot: dict) -> None:
    h = doc.add_heading("Appendix: Collected Data", level=1)
    h.runs[0].font.color.rgb = _NAVY

    data = get_data(snapshot)
    for section_name, tables in data.items():
        if not isinstance(tables, dict) or not tables:
            continue

        sh = doc.add_heading(section_name, level=2)
        sh.runs[0].font.color.rgb = _NAVY

        for table_name, rows in tables.items():
            if table_name in _APPENDIX_EXCLUDED:
                continue
            rows_list = _coerce_to_row_list(rows)
            if not rows_list:
                continue

            doc.add_heading(table_name, level=3)

            if isinstance(rows_list[0], dict):
                headers = list(rows_list[0].keys())
                tbl = doc.add_table(rows=1, cols=len(headers))
                tbl.style = "Table Grid"

                hdr_row = tbl.rows[0]
                for idx, header in enumerate(headers):
                    cell = hdr_row.cells[idx]
                    cell.text = header
                    run = cell.paragraphs[0].runs[0]
                    run.bold = True
                    run.font.color.rgb = _NAVY

                for record in rows_list[:_MAX_APPENDIX_ROWS]:
                    if not isinstance(record, dict):
                        continue
                    row_cells = tbl.add_row().cells
                    for idx, header in enumerate(headers):
                        val = record.get(header)
                        row_cells[idx].text = str(val) if val is not None else ""
            else:
                for item in rows_list[:_MAX_APPENDIX_ROWS]:
                    doc.add_paragraph(str(item), style="List Bullet")


def _coerce_to_row_list(rows: object) -> list:
    if isinstance(rows, dict):
        vals = list(rows.values())
        if not vals:
            return []
        if all(not isinstance(v, (dict, list)) for v in vals):
            return [rows]
        if all(isinstance(v, dict) for v in vals):
            return vals
        if all(isinstance(v, list) for v in vals):
            return [r for sub in vals for r in sub if isinstance(r, dict)]
        return [rows]
    return ensure_list(rows)
