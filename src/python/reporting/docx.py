"""
Word document generation - customer-facing Best Practices Assessment.

Document structure:
  1. Cover
  2. Executive Summary
  3. Workstream Reviews (narrative + key observations per workstream)
  4. Findings (table view per phase)
  5. Appendix (raw collected data)
"""

from __future__ import annotations

import logging
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from ..snapshot import get_data, get_metadata
from ..utils.converters import ensure_list

log = logging.getLogger(__name__)

_NAVY   = RGBColor(0x1F, 0x38, 0x64)
_ORANGE = RGBColor(0xC5, 0x5A, 0x11)
_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
_LGRAY  = RGBColor(0xF2, 0xF2, 0xF2)

_SEV_BG: dict[str, str] = {
    "High":   "FFD9D9",
    "Medium": "FFE9CC",
    "Low":    "E2F0D9",
}
_SEV_FG: dict[str, RGBColor] = {
    "High":   RGBColor(0xC0, 0x00, 0x00),
    "Medium": RGBColor(0xC5, 0x5A, 0x11),
    "Low":    RGBColor(0x37, 0x63, 0x2F),
}

_MAX_APPENDIX_ROWS = 300

_APPENDIX_EXCLUDED = {
    "AllMailboxes-MailIdentity",
    "AllMailboxes-UserPrincipalName",
    "AllMailboxes-PrimarySmtpAddress",
}

_PHASE_ORDER: dict[str, int] = {"Immediate": 0, "Near Term": 1, "Planned": 2, "Monitor": 3}
_SEV_ORDER:   dict[str, int] = {"High": 0, "Medium": 1, "Low": 2}

_PHASE_DESCRIPTIONS: dict[str, str] = {
    "Immediate": (
        "These items present the highest risk and should be addressed within 30 days. "
        "Each represents a direct security or compliance exposure requiring prompt attention."
    ),
    "Near Term": (
        "These findings should be resolved within 60-90 days. They represent meaningful gaps "
        "that will improve the tenant's security posture and operational efficiency."
    ),
    "Planned": (
        "These items are recommended improvements to address within the next 6 months. "
        "While not urgent, resolving them will reduce long-term risk and administrative burden."
    ),
    "Monitor": (
        "These items reflect areas to watch over time. "
        "No immediate action is required, but they should be revisited at future assessment intervals."
    ),
}


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
# Style helpers
# ---------------------------------------------------------------------------

def _setup_styles(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)


def _set_cell_bg(cell, hex_color: str) -> None:
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def _h(doc: Document, text: str, level: int, color: RGBColor | None = None) -> None:
    p = doc.add_heading(text, level=level)
    if p.runs:
        p.runs[0].font.color.rgb = color or _NAVY
    p.paragraph_format.space_after = Pt(6)


def _accent_bar(doc: Document) -> None:
    """Thin orange separator line."""
    p = doc.add_paragraph()
    run = p.add_run("_" * 70)
    run.font.color.rgb = _ORANGE
    run.font.size = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.space_before = Pt(0)


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

    doc.add_paragraph()
    doc.add_paragraph()

    brand = doc.add_paragraph()
    run = brand.add_run("ARRAYA SOLUTIONS")
    run.font.color.rgb = _ORANGE
    run.font.size = Pt(12)
    run.bold = True

    doc.add_paragraph()

    title = doc.add_paragraph()
    run = title.add_run(tenant)
    run.bold = True
    run.font.size = Pt(26)
    run.font.color.rgb = _NAVY

    sub = doc.add_paragraph()
    run = sub.add_run("Microsoft 365 Best Practices Assessment")
    run.font.size = Pt(16)
    run.font.color.rgb = _NAVY

    doc.add_paragraph()
    _accent_bar(doc)
    doc.add_paragraph()

    info = doc.add_paragraph()
    info.add_run("Prepared by: ").bold = True
    info.add_run("Arraya Solutions")

    if generated_at:
        date_p = doc.add_paragraph()
        date_p.add_run("Assessment Date: ").bold = True
        date_p.add_run(generated_at[:10] if len(generated_at) > 10 else generated_at)

    doc.add_page_break()


# ---------------------------------------------------------------------------
# Executive Summary
# ---------------------------------------------------------------------------

def _write_executive_summary(doc: Document, snapshot: dict, plan: dict | None) -> None:
    _h(doc, "Executive Summary", 1)

    findings: list[dict] = (plan or {}).get("Findings", [])
    ws_summaries: list[dict] = (plan or {}).get("WorkstreamSummaries", [])

    if not findings:
        doc.add_paragraph(
            "This assessment did not identify any open findings. "
            "The tenant configuration aligns with Microsoft 365 best practices."
        )
        doc.add_page_break()
        return

    phase_counts: dict[str, int] = {p: 0 for p in ("Immediate", "Near Term", "Planned", "Monitor")}
    for f in findings:
        p = f.get("RoadmapPhase", "Monitor")
        if p in phase_counts:
            phase_counts[p] += 1

    workstream_count = len({f.get("Workstream") for f in findings if f.get("Workstream")})
    high_count = sum(1 for f in findings if f.get("Severity") == "High")

    doc.add_paragraph(
        f"Arraya's assessment of the {(get_metadata(snapshot).get('Tenant') or {}).get('DisplayName', 'tenant')} "
        f"Microsoft 365 environment identified {len(findings)} findings across {workstream_count} workstream(s). "
        f"Of these, {high_count} are rated High severity and require prompt attention. "
        f"The findings are organized into four execution phases to guide prioritization and remediation planning."
    )

    doc.add_paragraph()

    # Phase summary callout
    tbl = doc.add_table(rows=2, cols=4)
    tbl.style = "Table Grid"
    phases = [("Immediate", "C00000"), ("Near Term", "C55A11"), ("Planned", "BF8F00"), ("Monitor", "375623")]
    for i, (phase, hex_color) in enumerate(phases):
        hcell = tbl.rows[0].cells[i]
        _set_cell_bg(hcell, hex_color)
        p = hcell.paragraphs[0]
        p.clear()
        run = p.add_run(phase)
        run.bold = True
        run.font.color.rgb = _WHITE
        run.font.size = Pt(10)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        vcell = tbl.rows[1].cells[i]
        p2 = vcell.paragraphs[0]
        run2 = p2.add_run(str(phase_counts[phase]))
        run2.bold = True
        run2.font.size = Pt(20)
        run2.font.color.rgb = RGBColor(int(hex_color[:2], 16), int(hex_color[2:4], 16), int(hex_color[4:], 16))
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    if ws_summaries:
        doc.add_paragraph()
        _h(doc, "Workstream Overview", 2)

        cols = ["Severity", "Workstream", "Area", "Open Findings"]
        widths = [Inches(0.9), Inches(1.5), Inches(2.0), Inches(1.0)]
        tbl2 = doc.add_table(rows=1, cols=len(cols))
        tbl2.style = "Table Grid"
        hdr = tbl2.rows[0].cells
        for i, (label, w) in enumerate(zip(cols, widths)):
            _set_cell_bg(hdr[i], "1F3864")
            hdr[i].width = w
            p = hdr[i].paragraphs[0]
            p.clear()
            run = p.add_run(label)
            run.bold = True
            run.font.color.rgb = _WHITE
            run.font.size = Pt(9)

        for ws in ws_summaries:
            row = tbl2.add_row().cells
            sev = ws.get("Severity", "")
            row[0].text = sev
            row[1].text = ws.get("Workstream", "")
            row[2].text = ws.get("Area", "")
            row[3].text = str(ws.get("OpenFindings", 0))
            if sev in _SEV_BG:
                _set_cell_bg(row[0], _SEV_BG[sev])
                if row[0].paragraphs[0].runs:
                    row[0].paragraphs[0].runs[0].font.color.rgb = _SEV_FG[sev]
                    row[0].paragraphs[0].runs[0].bold = True
            for cell in row:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(9)

    doc.add_page_break()


# ---------------------------------------------------------------------------
# Workstream Reviews
# ---------------------------------------------------------------------------

def _write_workstream_reviews(doc: Document, plan: dict | None) -> None:
    if not plan:
        return

    consultative: dict = plan.get("ConsultativeSummaries", {})
    findings: list[dict] = plan.get("Findings", [])
    if not consultative and not findings:
        return

    _h(doc, "Workstream Reviews", 1)

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

        _h(doc, ws, 2, _NAVY)

        if narrative:
            doc.add_paragraph(narrative)

        # Key observations for high-severity findings only
        high_findings = [f for f in ws_findings if f.get("Severity") == "High"]
        if high_findings:
            doc.add_paragraph()
            p = doc.add_paragraph()
            run = p.add_run("Key Observations")
            run.bold = True
            run.font.color.rgb = _ORANGE

            for f in high_findings[:5]:
                area     = f.get("Area", "")
                category = f.get("Category", "")
                title    = f"{area} - {category}" if area and category else (area or category or "Finding")
                message  = f.get("Finding", "")
                bullet = doc.add_paragraph(style="List Bullet")
                run_title = bullet.add_run(f"{title}: ")
                run_title.bold = True
                bullet.add_run(message)

        snapshot_rows = summary.get("SnapshotRows", [])
        if snapshot_rows:
            doc.add_paragraph()
            tbl = doc.add_table(rows=1, cols=2)
            tbl.style = "Table Grid"
            _set_cell_bg(tbl.rows[0].cells[0], "1F3864")
            _set_cell_bg(tbl.rows[0].cells[1], "1F3864")
            for idx, label in enumerate(("Signal", "Current State")):
                p = tbl.rows[0].cells[idx].paragraphs[0]
                p.clear()
                run = p.add_run(label)
                run.bold = True
                run.font.color.rgb = _WHITE
                run.font.size = Pt(9)
            for row_data in snapshot_rows:
                row = tbl.add_row().cells
                row[0].text = str(row_data.get("Signal", ""))
                row[1].text = str(row_data.get("State", ""))
                for cell in row:
                    if cell.paragraphs[0].runs:
                        cell.paragraphs[0].runs[0].font.size = Pt(9)

        doc.add_paragraph()

    doc.add_page_break()


# ---------------------------------------------------------------------------
# Findings (table view per phase)
# ---------------------------------------------------------------------------

def _write_findings(doc: Document, plan: dict | None) -> None:
    findings = (plan or {}).get("Findings", [])
    if not findings:
        return

    _h(doc, "Findings", 1)

    by_phase: dict[str, list[dict]] = {}
    for f in sorted(
        findings,
        key=lambda f: (
            _SEV_ORDER.get(f.get("Severity", "Low"), 99),
            f.get("Workstream", ""),
        ),
    ):
        phase = f.get("RoadmapPhase", "Monitor")
        by_phase.setdefault(phase, []).append(f)

    cols   = ["Severity", "Workstream", "Area", "Observation", "Recommended Action"]
    widths = [Inches(0.75), Inches(1.1), Inches(1.1), Inches(2.25), Inches(2.25)]

    for phase in ("Immediate", "Near Term", "Planned", "Monitor"):
        group = by_phase.get(phase)
        if not group:
            continue

        _h(doc, f"{phase}  ({len(group)} finding{'s' if len(group) != 1 else ''})", 2)

        desc = _PHASE_DESCRIPTIONS.get(phase, "")
        if desc:
            p = doc.add_paragraph(desc)
            p.paragraph_format.space_after = Pt(6)

        tbl = doc.add_table(rows=1, cols=len(cols))
        tbl.style = "Table Grid"
        hdr = tbl.rows[0].cells
        for i, (label, w) in enumerate(zip(cols, widths)):
            _set_cell_bg(hdr[i], "1F3864")
            hdr[i].width = w
            p = hdr[i].paragraphs[0]
            p.clear()
            run = p.add_run(label)
            run.bold = True
            run.font.color.rgb = _WHITE
            run.font.size = Pt(9)

        for f in group:
            sev = f.get("Severity", "")
            row = tbl.add_row().cells
            values = [
                sev,
                f.get("Workstream", ""),
                f.get("Area", ""),
                f.get("Finding", ""),
                f.get("TechnicalRemediation", ""),
            ]
            for i, (val, w) in enumerate(zip(values, widths)):
                row[i].width = w
                p = row[i].paragraphs[0]
                run = p.add_run(str(val))
                run.font.size = Pt(9)
                if i == 0 and sev in _SEV_BG:
                    _set_cell_bg(row[i], _SEV_BG[sev])
                    run.font.color.rgb = _SEV_FG[sev]
                    run.bold = True

        doc.add_paragraph()

    doc.add_page_break()


# ---------------------------------------------------------------------------
# Data Appendix
# ---------------------------------------------------------------------------

def _write_appendix(doc: Document, snapshot: dict) -> None:
    _h(doc, "Appendix: Collected Data", 1)
    doc.add_paragraph(
        "The following tables contain the raw configuration data collected during the assessment. "
        "This information serves as the evidence base for the findings documented above."
    )
    doc.add_paragraph()

    data = get_data(snapshot)
    for section_name, tables in data.items():
        if not isinstance(tables, dict) or not tables:
            continue

        _h(doc, section_name, 2)

        for table_name, rows in tables.items():
            if table_name in _APPENDIX_EXCLUDED:
                continue
            rows_list = _coerce_to_row_list(rows)
            if not rows_list:
                continue

            _h(doc, table_name, 3)

            if isinstance(rows_list[0], dict):
                headers = list(rows_list[0].keys())
                tbl = doc.add_table(rows=1, cols=len(headers))
                tbl.style = "Table Grid"

                hdr_row = tbl.rows[0]
                for idx, header in enumerate(headers):
                    cell = hdr_row.cells[idx]
                    _set_cell_bg(cell, "1F3864")
                    p = cell.paragraphs[0]
                    p.clear()
                    run = p.add_run(header)
                    run.bold = True
                    run.font.color.rgb = _WHITE
                    run.font.size = Pt(8)

                for record in rows_list[:_MAX_APPENDIX_ROWS]:
                    if not isinstance(record, dict):
                        continue
                    row_cells = tbl.add_row().cells
                    for idx, header in enumerate(headers):
                        val = record.get(header)
                        p = row_cells[idx].paragraphs[0]
                        run = p.add_run(str(val) if val is not None else "")
                        run.font.size = Pt(8)
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
