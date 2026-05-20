"""
Word document generation - customer-facing Best Practices Assessment.
Follows the PS Best Practices Assessment structure:
  1.0 Introduction
  2.0 Project Scope
  3.0 Executive Summary
  4.0 Modern Workplace Recommendations
  5.0-10.0 Workstream Sections
  11.0 Workloads Outside Assessment Scope
  12.0 Appendix (documentation links, not raw data)
"""

from __future__ import annotations

import logging
import re
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

_SEV_BG: dict[str, str] = {
    "High":     "FFD9D9",
    "Medium":   "FFE9CC",
    "Low":      "E2F0D9",
    "Critical": "FFD9D9",
    "Info":     "EBF3FB",
}
_SEV_FG: dict[str, RGBColor] = {
    "High":     RGBColor(0xC0, 0x00, 0x00),
    "Medium":   RGBColor(0xC5, 0x5A, 0x11),
    "Low":      RGBColor(0x37, 0x63, 0x2F),
    "Critical": RGBColor(0xC0, 0x00, 0x00),
    "Info":     RGBColor(0x1F, 0x38, 0x64),
}

_PHASE_ORDER: dict[str, int] = {"Immediate": 0, "Near Term": 1, "Planned": 2, "Monitor": 3}
_SEV_ORDER:   dict[str, int] = {"Critical": 0, "High": 0, "Medium": 1, "Low": 2, "Info": 3}

_WS_CLUSTER_NAMES: dict[str, str] = {
    "Identity":      "Identity and privileged access governance",
    "Messaging":     "Messaging and mail flow security",
    "Collaboration": "Collaboration ownership and lifecycle governance",
    "Endpoint":      "Endpoint compliance and device control",
    "Security":      "Security baseline and access controls",
    "Governance":    "Licensing and tenant governance",
}

_WS_SECTION_TITLES: dict[str, str] = {
    "Identity":      "Identity and Access Review",
    "Messaging":     "Exchange Online and Messaging Review",
    "Collaboration": "Microsoft Teams and Collaboration",
    "Endpoint":      "Device and Endpoint Review",
    "Security":      "Security Baseline Review",
    "Governance":    "Governance and Compliance",
}

_WS_ORDER = ["Identity", "Messaging", "Collaboration", "Endpoint", "Security", "Governance"]

_DATA_COLLECTION_SCOPE = [
    ("Identity and Access", "Entra ID users, guests, device registrations, admin role assignments, MFA enrollment, enterprise applications, and privileged access signals"),
    ("Exchange Online", "Mailbox configuration, accepted domains, mail forwarding, shared mailboxes, mail flow rules, archive statistics, and recipient inventory"),
    ("Microsoft Teams", "Team lifecycle signals, guest access configuration, group ownership, and collaboration activity"),
    ("SharePoint and OneDrive", "External sharing posture, site storage, sharing overrides, and ownership governance signals"),
    ("Security", "Secure Score, spam filtering configuration, SMTP relay, and consent governance"),
    ("Governance", "DLP policies, retention policies, license assignment, lifecycle signals, and hybrid identity configuration"),
]

_NOT_IN_SCOPE = [
    "Microsoft Defender for Endpoint / Microsoft Defender XDR - device compliance posture, endpoint protection signals, and threat analytics",
    "Microsoft Power Platform (Power Apps, Power Automate, Power BI) - environment sprawl, connector usage, and platform-level governance",
    "Microsoft Copilot for Microsoft 365 - readiness posture, data oversharing risk, and prompt usage analytics",
    "Microsoft Purview Sensitivity Labels - label taxonomy, auto-labeling policy configuration, and classification coverage",
    "Microsoft Purview Audit Log - unified audit log review, alert policy coverage, and forensic event query scope",
]

_APPENDIX_LINKS: dict[str, list[tuple[str, str, str]]] = {
    "Identity and Access": [
        ("Conditional Access overview", "https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview", "Core reference for understanding and deploying Conditional Access policies"),
        ("Privileged Identity Management overview", "https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-configure", "Use PIM to reduce standing privileged access and require justification for elevation"),
        ("B2B collaboration fundamentals", "https://learn.microsoft.com/en-us/entra/external-id/b2b-fundamentals", "Foundation for managing guest identities and cross-tenant access"),
    ],
    "Authentication and MFA": [
        ("Manage authentication methods", "https://learn.microsoft.com/en-us/azure/active-directory/authentication/concept-authentication-methods-manage", "Centralize MFA method management and migrate from legacy per-user MFA"),
        ("Configure the admin consent workflow", "https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-admin-consent-workflow", "Reduce risk from unreviewed app consent grants"),
    ],
    "Exchange Online": [
        ("Control automatic external email forwarding", "https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/external-email-forwarding", "Block or limit external forwarding at the organization level"),
        ("Set up SPF, DKIM, and DMARC", "https://learn.microsoft.com/en-us/defender-office-365/email-authentication-dmarc-configure", "Harden domain authentication to reduce spoofing and phishing exposure"),
    ],
    "Teams and Collaboration": [
        ("Manage who can create Microsoft 365 Groups", "https://learn.microsoft.com/en-us/microsoft-365/solutions/manage-creation-of-groups", "Restrict group creation to reduce unmanaged collaboration sprawl"),
        ("Set expiration for Microsoft 365 groups", "https://learn.microsoft.com/en-us/entra/identity/users/groups-lifecycle", "Automate lifecycle governance for Teams and groups"),
        ("Overview of external sharing in SharePoint and OneDrive", "https://learn.microsoft.com/en-us/sharepoint/external-sharing-overview", "Understand and control tenant-level and site-level sharing settings"),
    ],
    "Governance and Compliance": [
        ("Learn about retention policies and retention labels", "https://learn.microsoft.com/en-us/purview/retention", "Baseline reference for configuring retention in Exchange and SharePoint"),
        ("Learn about data loss prevention", "https://learn.microsoft.com/en-us/purview/dlp-learn-about-dlp", "Overview of DLP policy configuration and scope"),
    ],
    "Endpoint": [
        ("Get started with device compliance policies in Microsoft Intune", "https://learn.microsoft.com/en-us/intune/intune-service/protect/device-compliance-get-started", "Foundation for defining and enforcing device compliance baselines"),
        ("Require compliant or hybrid Microsoft Entra joined device", "https://learn.microsoft.com/en-us/entra/identity/conditional-access/howto-conditional-access-policy-compliant-device", "Use Conditional Access to gate access on device compliance state"),
    ],
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

    meta          = get_metadata(snapshot)
    tenant_obj    = meta.get("Tenant") or {}
    tenant_name   = (
        (tenant_obj.get("DisplayName") if isinstance(tenant_obj, dict) else None)
        or meta.get("TenantDisplayName") or meta.get("TenantDomain") or "Tenant"
    )
    generated_at  = meta.get("GeneratedAt", "")
    date_label    = generated_at[:10] if len(generated_at) >= 10 else generated_at

    findings:    list[dict] = (plan or {}).get("Findings", [])
    ws_sums:     list[dict] = (plan or {}).get("WorkstreamSummaries", [])
    actions:     list[dict] = (plan or {}).get("RoadmapActions", [])
    consultative: dict       = (plan or {}).get("ConsultativeSummaries", {})

    _write_cover(doc, tenant_name, date_label)
    _write_version_history(doc, date_label)
    _write_introduction(doc, tenant_name, date_label, findings, actions)
    _write_project_scope(doc)
    _write_executive_summary(doc, tenant_name, findings, ws_sums, actions, consultative)
    _write_recommendations(doc, actions)
    _write_workstream_sections(doc, snapshot, findings, consultative, actions)
    _write_out_of_scope(doc)
    _write_appendix_links(doc)

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
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def _h(doc: Document, text: str, level: int, color: RGBColor | None = None) -> None:
    p = doc.add_heading(text, level=level)
    if p.runs:
        p.runs[0].font.color.rgb = color or _NAVY
    p.paragraph_format.space_after = Pt(6)


def _table_2col(doc: Document, rows: list[tuple[str, str]], hdr: tuple[str, str] | None = None,
                col_widths: tuple[float, float] = (2.0, 4.5)) -> None:
    """Build a two-column Signal | Current State table."""
    tbl = doc.add_table(rows=0, cols=2)
    tbl.style = "Table Grid"
    if hdr:
        r = tbl.add_row().cells
        _set_cell_bg(r[0], "1F3864")
        _set_cell_bg(r[1], "1F3864")
        for i, label in enumerate(hdr):
            r[i].width = Inches(col_widths[i])
            p = r[i].paragraphs[0]
            p.clear()
            run = p.add_run(label)
            run.bold = True
            run.font.color.rgb = _WHITE
            run.font.size = Pt(9)
    for left, right in rows:
        r = tbl.add_row().cells
        r[0].width = Inches(col_widths[0])
        r[1].width = Inches(col_widths[1])
        for i, val in enumerate((left, right)):
            p = r[i].paragraphs[0]
            run = p.add_run(str(val))
            run.font.size = Pt(9)
    doc.add_paragraph()


def _finding_row(tbl, f: dict) -> None:
    sev = f.get("Severity", "")
    row = tbl.add_row().cells
    vals = [
        sev,
        f.get("Area", ""),
        f.get("Finding", ""),
        f.get("TechnicalRemediation", ""),
    ]
    for i, val in enumerate(vals):
        p = row[i].paragraphs[0]
        run = p.add_run(str(val))
        run.font.size = Pt(9)
        if i == 0 and sev in _SEV_BG:
            _set_cell_bg(row[i], _SEV_BG[sev])
            run.font.color.rgb = _SEV_FG[sev]
            run.bold = True


def _findings_table(doc: Document, findings: list[dict]) -> None:
    """Render a 4-column findings table sorted by severity then phase."""
    if not findings:
        doc.add_paragraph("No findings identified for this workstream.", style="Normal")
        return
    sorted_f = sorted(
        findings,
        key=lambda f: (
            _SEV_ORDER.get(f.get("Severity", "Low"), 99),
            _PHASE_ORDER.get(f.get("RoadmapPhase", "Monitor"), 99),
        ),
    )
    cols   = ["Severity", "Area", "Observation", "Recommended Action"]
    widths = [0.75, 1.25, 2.5, 2.5]
    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    hdr = tbl.rows[0].cells
    for i, (label, w) in enumerate(zip(cols, widths)):
        _set_cell_bg(hdr[i], "1F3864")
        hdr[i].width = Inches(w)
        p = hdr[i].paragraphs[0]
        p.clear()
        run = p.add_run(label)
        run.bold = True
        run.font.color.rgb = _WHITE
        run.font.size = Pt(9)
    for f in sorted_f:
        _finding_row(tbl, f)
    doc.add_paragraph()


# ---------------------------------------------------------------------------
# Cover
# ---------------------------------------------------------------------------

def _write_cover(doc: Document, tenant_name: str, date_label: str) -> None:
    doc.add_paragraph()
    doc.add_paragraph()

    brand = doc.add_paragraph()
    run = brand.add_run("ARRAYA SOLUTIONS")
    run.font.color.rgb = _ORANGE
    run.font.size = Pt(12)
    run.bold = True

    doc.add_paragraph()

    try:
        title = doc.add_paragraph(style="Title")
    except Exception:
        title = doc.add_paragraph()
    run = title.add_run(f"{tenant_name} Microsoft 365 Tenant Best Practices Assessment")
    run.bold = True
    run.font.size = Pt(22)
    run.font.color.rgb = _NAVY

    try:
        sub = doc.add_paragraph(style="Subtitle")
    except Exception:
        sub = doc.add_paragraph()
    run = sub.add_run("Prepared by: Arraya Solutions")
    run.font.size = Pt(14)
    run.font.color.rgb = _ORANGE

    doc.add_paragraph()
    gen_p = doc.add_paragraph()
    gen_p.add_run("Generated: ").bold = True
    gen_p.add_run(date_label)
    rev_p = doc.add_paragraph()
    rev_p.add_run("Document revision: ").bold = True
    rev_p.add_run("1.0")

    doc.add_page_break()


# ---------------------------------------------------------------------------
# Version History
# ---------------------------------------------------------------------------

def _write_version_history(doc: Document, date_label: str) -> None:
    _h(doc, "Version History", 1)
    doc.add_paragraph("This section tracks the issued version of the customer report.")
    _table_2col(
        doc,
        [(date_label, "1.0  |  Arraya Solutions  |  Current report release")],
        hdr=("Date", "Revision  |  Author  |  Description"),
        col_widths=(1.2, 5.3),
    )
    doc.add_page_break()


# ---------------------------------------------------------------------------
# 1.0 Introduction
# ---------------------------------------------------------------------------

def _write_introduction(
    doc: Document,
    tenant_name: str,
    date_label: str,
    findings: list[dict],
    actions: list[dict],
) -> None:
    _h(doc, "1.0 Introduction", 1)
    doc.add_paragraph(
        f"This report summarizes the Microsoft 365 state observed in {tenant_name} and organizes "
        "the highest-value risks, decisions, and recommendations into a structured remediation path. "
        "Use the Executive Summary for the leadership-level risk brief, the Modern Workplace "
        "Recommendations section for the prioritized execution view, and the workstream sections "
        "for detailed signal analysis."
    )

    _h(doc, "Tenant Snapshot At A Glance", 2)

    high_count      = sum(1 for f in findings if f.get("Severity") in ("High", "Critical"))
    immediate_count = sum(1 for a in actions if a.get("RoadmapPhase") == "Immediate")
    ws_present      = sorted({f.get("Workstream", "") for f in findings if f.get("Workstream")})

    rows = [
        ("Scope reviewed",               ", ".join(ws_present) if ws_present else "Identity, Messaging, Collaboration, Endpoint, Security, Governance"),
        ("Total findings",               str(len(findings))),
        ("High-severity findings",       str(high_count)),
        ("Immediate actions required",   str(immediate_count)),
        ("Collection date",              date_label),
        ("Collection method",            "Read-only Microsoft Graph API calls and PowerShell"),
        ("Assessment companion documents", "Remediation Roadmap and Engineer Action Pack"),
    ]
    _table_2col(doc, rows, hdr=("Signal", "Current State"))
    doc.add_paragraph(
        "This report starts with an orientation snapshot of what was reviewed and where the "
        "detailed evidence appears later in the document. Each workstream section carries its "
        "own signal table and narrative so the evidence behind each recommendation is visible "
        "alongside the recommendation itself."
    )
    doc.add_page_break()


# ---------------------------------------------------------------------------
# 2.0 Project Scope
# ---------------------------------------------------------------------------

def _write_project_scope(doc: Document) -> None:
    _h(doc, "2.0 Project Scope", 1)
    doc.add_paragraph(
        "The scope of this report is limited to the Microsoft 365 signals surfaced in the tenant "
        "review across identity, devices, messaging, collaboration, security, and governance. "
        "Where a data point was not available, the report labels it clearly instead of inferring a value."
    )

    _h(doc, "2.1 Assessment Methodology", 2)
    doc.add_paragraph(
        "Tenant data was collected using read-only Microsoft Graph API calls and PowerShell. "
        "No configuration changes were made during the assessment. All findings are based on "
        "signals observable at the time of collection."
    )

    _h(doc, "Data Collection Scope", 3)
    _table_2col(doc, _DATA_COLLECTION_SCOPE, hdr=("Area", "Data Collected"), col_widths=(1.8, 4.7))

    _h(doc, "Workloads Not in Scope", 3)
    doc.add_paragraph(
        "The following capability areas were not collected or analyzed in this engagement. "
        "Section 11.0 documents each exclusion with context on why it falls outside the current scope."
    )
    for item in _NOT_IN_SCOPE:
        p = doc.add_paragraph(item, style="List Bullet")
        p.paragraph_format.space_after = Pt(3)

    doc.add_page_break()


# ---------------------------------------------------------------------------
# 3.0 Executive Summary
# ---------------------------------------------------------------------------

def _write_executive_summary(
    doc: Document,
    tenant_name: str,
    findings: list[dict],
    ws_sums: list[dict],
    actions: list[dict],
    consultative: dict,
) -> None:
    _h(doc, "3.0 Executive Summary", 1)

    # Identify worst workstreams
    from collections import defaultdict
    ws_counts: dict[str, int] = defaultdict(int)
    ws_areas:  dict[str, set] = defaultdict(set)
    for f in findings:
        ws = f.get("Workstream", "")
        ws_counts[ws] += 1
        ws_areas[ws].add(f.get("Area", ""))

    top_ws = sorted(ws_counts.items(), key=lambda x: -x[1])
    if top_ws:
        top_names = [_WS_CLUSTER_NAMES.get(w, w) for w, _ in top_ws[:3]]
        doc.add_paragraph(
            f"The report shows the clearest risk concentration in "
            f"{', '.join(top_names[:-1])}, and {top_names[-1]}. "
            "The sections below give leadership the decision-ready view of where risk is "
            "clustering and why those patterns matter now."
        )
    else:
        doc.add_paragraph(
            f"The assessment of {tenant_name} did not identify significant open findings. "
            "The tenant configuration is broadly aligned with Microsoft 365 best practices."
        )

    # Risk Clusters
    _h(doc, "Risk Clusters", 2)
    for ws, count in top_ws:
        cluster_name  = _WS_CLUSTER_NAMES.get(ws, ws)
        areas         = sorted(ws_areas[ws])
        area_text     = ", ".join(areas) if areas else ws
        cs_key        = f"{ws}ConsultativeSummary"
        narrative     = consultative.get(cs_key, {}).get("Narrative", "")
        top_signal    = narrative.split(".")[0] if narrative else ""
        bullet = doc.add_paragraph(style="List Bullet")
        run_title = bullet.add_run(f"{cluster_name}: ")
        run_title.bold = True
        detail = f"The tenant shows {count} related finding(s) concentrated across {area_text}."
        if top_signal:
            detail += f" {top_signal}."
        bullet.add_run(detail)
        bullet.paragraph_format.space_after = Pt(4)

    # Leadership Decision Brief
    _h(doc, "Leadership Decision Brief", 2)
    doc.add_paragraph(
        "The items below highlight the leadership approvals or owner decisions that would remove "
        "the biggest blockers to a cleaner operating posture."
    )
    immediate_actions = [a for a in actions if a.get("RoadmapPhase") in ("Immediate", "Near Term")]
    for action in immediate_actions[:6]:
        bullet = doc.add_paragraph(style="List Bullet")
        run_title = bullet.add_run(f"{action.get('ActionTitle', '')}: ")
        run_title.bold = True
        rec = action.get("RecommendedNextStep", "")
        bullet.add_run(rec[:200] + ("..." if len(rec) > 200 else ""))
        bullet.paragraph_format.space_after = Pt(4)

    # Overall Findings Summary
    _h(doc, "Overall Findings Summary", 2)
    if ws_sums:
        # Consolidate to one row per workstream (highest severity wins)
        ws_best: dict[str, dict] = {}
        for row in ws_sums:
            ws = row.get("Workstream", "")
            cur_sev = _SEV_ORDER.get(row.get("Severity", "Low"), 99)
            if ws not in ws_best or cur_sev < _SEV_ORDER.get(ws_best[ws].get("Severity", "Low"), 99):
                ws_best[ws] = row
        # Also sum OpenFindings per workstream
        ws_total_findings: dict[str, int] = defaultdict(int)
        for row in ws_sums:
            ws_total_findings[row.get("Workstream", "")] += int(row.get("OpenFindings", 0))

        cols   = ["Workstream", "Severity / Impact", "Open Findings", "What Stands Out"]
        widths = [1.2, 1.1, 1.0, 3.2]
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        hdr = tbl.rows[0].cells
        for i, (label, w) in enumerate(zip(cols, widths)):
            _set_cell_bg(hdr[i], "1F3864")
            hdr[i].width = Inches(w)
            p = hdr[i].paragraphs[0]
            p.clear()
            run = p.add_run(label)
            run.bold = True
            run.font.color.rgb = _WHITE
            run.font.size = Pt(9)

        for ws in _WS_ORDER:
            if ws not in ws_best:
                continue
            row_data = ws_best[ws]
            sev = row_data.get("Severity", "")
            row = tbl.add_row().cells
            row[0].text = ws
            row[1].text = sev
            row[2].text = str(ws_total_findings.get(ws, 0))
            top_sig = row_data.get("TopSignals", "")
            row[3].text = top_sig[:180] + ("..." if len(top_sig) > 180 else "")
            if sev in _SEV_BG:
                _set_cell_bg(row[1], _SEV_BG[sev])
                if row[1].paragraphs[0].runs:
                    row[1].paragraphs[0].runs[0].font.color.rgb = _SEV_FG[sev]
                    row[1].paragraphs[0].runs[0].bold = True
            for cell in row:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(9)

        doc.add_paragraph()
        doc.add_paragraph(
            "This table shows where findings are clustering before the report moves into the "
            "detailed workstream sections. Use Section 4.0 for the prioritized execution view."
        )

    doc.add_page_break()


# ---------------------------------------------------------------------------
# 4.0 Modern Workplace Recommendations
# ---------------------------------------------------------------------------

def _write_recommendations(doc: Document, actions: list[dict]) -> None:
    _h(doc, "4.0 Modern Workplace Recommendations", 1)
    doc.add_paragraph(
        "This section is the streamlined execution view for the report. Use it to prioritize "
        "the work, confirm ownership, and align on effort before moving into the detailed workstream sections."
    )

    if actions:
        cols   = ["Recommendation", "Criticality", "Level of Effort", "Rough PS Hours"]
        widths = [3.0, 1.0, 1.1, 1.1]
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        hdr = tbl.rows[0].cells
        for i, (label, w) in enumerate(zip(cols, widths)):
            _set_cell_bg(hdr[i], "1F3864")
            hdr[i].width = Inches(w)
            p = hdr[i].paragraphs[0]
            p.clear()
            run = p.add_run(label)
            run.bold = True
            run.font.color.rgb = _WHITE
            run.font.size = Pt(9)

        sorted_actions = sorted(
            actions,
            key=lambda a: (_PHASE_ORDER.get(a.get("RoadmapPhase", "Monitor"), 99), _SEV_ORDER.get(a.get("HighestSeverity", "Low"), 99)),
        )
        for action in sorted_actions:
            sev = action.get("HighestSeverity", "")
            row = tbl.add_row().cells
            row[0].text = action.get("ActionTitle", "")
            row[1].text = sev
            row[2].text = action.get("EffortTier", "")
            row[3].text = action.get("EstimatedPsHours", "")
            if sev in _SEV_BG:
                _set_cell_bg(row[1], _SEV_BG[sev])
                if row[1].paragraphs[0].runs:
                    row[1].paragraphs[0].runs[0].font.color.rgb = _SEV_FG[sev]
                    row[1].paragraphs[0].runs[0].bold = True
            for cell in row:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(9)

        doc.add_paragraph()
        doc.add_paragraph(
            "Level of Effort is an initial delivery-planning estimate to help sequence work at a glance. "
            "Rough PS Hours is a pre-sales scoping reference and should be validated with the delivery team."
        )

    # Completion Signals
    _h(doc, "Completion Signals", 2)
    doc.add_paragraph(
        "The following signals indicate that each area has reached a closed or stabilized state. "
        "Use these to confirm that a recommendation has been fully addressed, not just started."
    )
    for action in actions:
        success = action.get("SuccessCheck", "")
        ws      = action.get("Workstream", "")
        if success:
            bullet = doc.add_paragraph(style="List Bullet")
            run_title = bullet.add_run(f"{ws}: ")
            run_title.bold = True
            bullet.add_run(success)
            bullet.paragraph_format.space_after = Pt(3)

    doc.add_page_break()


# ---------------------------------------------------------------------------
# 5.0 - 10.0 Workstream Sections
# ---------------------------------------------------------------------------

def _write_workstream_sections(
    doc: Document,
    snapshot: dict,
    findings: list[dict],
    consultative: dict,
    actions: list[dict],
) -> None:
    section_num = 5
    for ws in _WS_ORDER:
        ws_findings = [f for f in findings if f.get("Workstream") == ws]
        cs_key      = f"{ws}ConsultativeSummary"
        cs          = consultative.get(cs_key, {})

        if not ws_findings and not cs:
            section_num += 1
            continue

        title = _WS_SECTION_TITLES.get(ws, ws)
        _h(doc, f"{section_num}.0 {title}", 1)

        # Key Signals table — metrics at a glance
        snapshot_rows = cs.get("SnapshotRows", [])
        if snapshot_rows:
            _table_2col(
                doc,
                [(r.get("Signal", ""), r.get("State", "")) for r in snapshot_rows],
                hdr=("Configuration Signal", "Current State"),
            )

        # Story narrative — specifics that the table doesn't show
        for para_text in _ws_story(ws, snapshot, ws_findings):
            p = doc.add_paragraph(para_text)
            p.paragraph_format.space_after = Pt(6)

        # Findings
        if ws_findings:
            _h(doc, "Findings", 2)
            _findings_table(doc, ws_findings)

        doc.add_page_break()
        section_num += 1


# ---------------------------------------------------------------------------
# Workstream story helpers — named specifics the Signal table doesn't cover
# ---------------------------------------------------------------------------

def _ws_story(ws: str, snapshot: dict, findings: list[dict]) -> list[str]:
    """Return 2-3 paragraphs of character-specific narrative for the workstream."""
    data    = snapshot.get("Data", {})
    derived = snapshot.get("Derived", {})
    fn = {
        "Identity":      _identity_story,
        "Messaging":     _messaging_story,
        "Collaboration": _collab_story,
        "Endpoint":      _endpoint_story,
        "Security":      _security_story,
        "Governance":    _governance_story,
    }.get(ws)
    return fn(data, derived, findings) if fn else []


def _f(findings: list[dict], *category_keywords: str) -> dict | None:
    for kw in category_keywords:
        for f in findings:
            if kw.lower() in str(f.get("Category", "")).lower() or kw.lower() in str(f.get("Area", "")).lower():
                return f
    return None


def _count(text: str) -> str:
    m = re.match(r"(\d+)", str(text))
    return m.group(1) if m else ""


def _identity_story(data: dict, derived: dict, findings: list[dict]) -> list[str]:
    users        = data.get("Identity", {}).get("Users", {})
    user_count   = len(users) if isinstance(users, dict) else 0
    priv         = data.get("Identity", {}).get("PrivilegedAccessRemediationSummary", [])
    ga_count = stale_ga = 0
    if isinstance(priv, list):
        for p in priv:
            sig = str(p.get("Signal", ""))
            if "Global Administrator" in sig:
                ga_count = p.get("Count", 0) or 0
            elif "Stale privileged" in sig:
                stale_ga = p.get("Count", 0) or 0

    inactive_f = _f(findings, "Inactive Users", "inactive user")
    inactive_pct = ""
    inactive_cnt = ""
    if inactive_f:
        txt = inactive_f.get("Finding", "")
        m = re.match(r"(\d+) enabled.*?\(([0-9.]+)%\)", txt)
        if m:
            inactive_cnt, inactive_pct = m.group(1), m.group(2)

    paras = []
    if user_count and ga_count:
        paras.append(
            f"The identity review covers {user_count} user accounts. "
            f"The most pressing concern is the volume of standing Global Administrators: "
            f"{ga_count} accounts hold the highest-privilege role in the tenant"
            + (f", and {stale_ga} of those appear stale with no recent sign-in activity" if stale_ga else "")
            + ". Global Administrator is the broadest credential in Microsoft 365 — "
            "reducing the count and requiring PIM-based elevation for those that remain "
            "closes the largest single privilege-escalation surface in the environment."
        )
    if inactive_cnt and inactive_pct:
        paras.append(
            f"{inactive_cnt} enabled member accounts ({inactive_pct}%) have not signed in within the last 180 days. "
            "Accounts that stay enabled past their useful life continue to carry active group memberships, "
            "license assignments, and in some cases delegated access to shared mailboxes and sites — "
            "none of which require the user to ever sign in again to retain that access."
        )
    return paras


def _messaging_story(data: dict, derived: dict, findings: list[dict]) -> list[str]:
    domain_f = _f(findings, "Domain Verification", "Unverified")
    dmarc_f  = _f(findings, "DMARC")
    fwd_f    = _f(findings, "Forwarding", "Mailbox Forwarding")

    paras = []
    domain_parts = []
    if domain_f:
        m = re.search(r"'([^']+)'", domain_f.get("Finding", ""))
        if m:
            domain_parts.append(f"'{m.group(1)}' is listed as an accepted domain but is not verified")
    if dmarc_f:
        m = re.search(r"'([^']+)'", dmarc_f.get("Finding", ""))
        if m:
            domain_parts.append(f"DMARC is not configured on '{m.group(1)}'")

    if domain_parts:
        combined = "; ".join(domain_parts)
        paras.append(
            f"Two domain authentication gaps compound each other: {combined}. "
            "An unverified accepted domain can create internal mail routing ambiguity and "
            "is harder to clean up once mail-flow rules or connectors reference it. "
            "Missing DMARC on the primary sending domain means there is no published policy "
            "governing what receiving servers should do with mail that fails authentication — "
            "spoofed messages can reach inboxes without a rejection or quarantine signal."
        )
    if fwd_f:
        cnt = _count(fwd_f.get("Finding", ""))
        label = f"{cnt} mailbox{'es' if cnt != '1' else ''} have" if cnt else "Mailboxes have"
        paras.append(
            f"{label} active forwarding configured to send copies outside the organization. "
            "External forwarding creates a persistent data-exit path that runs independently of "
            "DLP policies unless the policy is specifically scoped to cover outbound forwarding rules — "
            "which most default configurations do not do."
        )
    return paras


def _collab_story(data: dict, derived: dict, findings: list[dict]) -> list[str]:
    teams_raw = data.get("Collaboration", {}).get("TeamsGroupsCleanupCandidates", {})
    teams_list = list(teams_raw.values()) if isinstance(teams_raw, dict) else (teams_raw if isinstance(teams_raw, list) else [])

    unmanaged_raw  = derived.get("UnmanagedObjects", {})
    unmanaged_list = list(unmanaged_raw.values()) if isinstance(unmanaged_raw, dict) else (unmanaged_raw if isinstance(unmanaged_raw, list) else [])

    og = derived.get("OwnershipGovernanceSummary", {}).get("Summary", {})
    total_objects = og.get("TotalObjectsReviewed", len(unmanaged_list))

    # Named team examples — prefer ownerless
    ownerless_teams = [t for t in teams_list if "Ownerless" in str(t.get("RiskSignal", ""))]
    dormant_teams   = [t for t in teams_list if "Dormant" in str(t.get("RiskSignal", "")) and t not in ownerless_teams]
    sample          = (ownerless_teams + dormant_teams)[:3]
    team_examples   = []
    for t in sample:
        name  = t.get("Name", "")
        risk  = t.get("RiskSignal", "").split(";")[0].strip()
        last  = (t.get("LastActivityDate") or "")[:7] or "unknown"
        if name:
            team_examples.append(f"'{name}' ({risk.lower()}, last active {last})")

    # Named group examples from unmanaged objects
    group_examples = [
        u.get("DisplayName", "") for u in unmanaged_list
        if u.get("ObjectType") in ("Entra Group", "Microsoft 365 Group") and u.get("DisplayName")
    ][:2]

    paras = []
    if total_objects and team_examples:
        ex_text = "; ".join(team_examples[:2])
        paras.append(
            f"The collaboration review assessed {total_objects} objects across Teams, SharePoint, OneDrive, and Entra groups. "
            f"Specific examples from the cleanup candidate list include {ex_text}. "
            "These are not outliers — they represent the ownership and lifecycle pattern "
            "visible across the full unmanaged inventory."
        )
    if group_examples:
        names = " and ".join(f"'{n}'" for n in group_examples)
        paras.append(
            f"Groups like {names} appear in the unmanaged inventory with no assigned owner. "
            "Ownerless groups are difficult to decommission cleanly because there is no accountable "
            "party to confirm whether the group is still in active use, "
            "whether its membership is correct, or whether content associated with it can be archived."
        )
    return paras


def _endpoint_story(data: dict, derived: dict, findings: list[dict]) -> list[str]:
    devices       = data.get("Identity", {}).get("DeviceDetails", {})
    device_count  = len(devices) if isinstance(devices, dict) else 0
    compliance_f  = _f(findings, "Device Compliance", "Compliance")
    unmanaged_f   = _f(findings, "Unmanaged Device", "Unmanaged")
    stale_f       = _f(findings, "Stale Device", "Stale")

    paras = []
    if device_count:
        paras.append(
            f"The device review covers {device_count} registered device records, "
            "with Windows representing the dominant platform. "
            "The compliance and management gaps reinforce each other: "
            "unmanaged devices cannot satisfy compliance policy requirements, "
            "and where compliance is the signal driving a Conditional Access decision, "
            "an unmanaged device will always fail that check — "
            "meaning the CA policy is structurally unable to enforce the intended access boundary."
        )
    stats = []
    for f in [compliance_f, unmanaged_f, stale_f]:
        if f:
            s = f.get("Finding", "").split(".")[0].strip()
            if s:
                stats.append(s)
    if stats:
        paras.append(" ".join(s.rstrip(".") + "." for s in stats))
    return paras


def _security_story(data: dict, derived: dict, findings: list[dict]) -> list[str]:
    score_f   = _f(findings, "Secure Score")
    consent_f = _f(findings, "Consent Governance", "Consent")

    paras = []
    if score_f:
        score_text = score_f.get("Finding", "")
        paras.append(
            f"{score_text} "
            "Secure Score is a composite signal — the gaps pulling it down are distributed "
            "across identity governance, MFA enforcement, and sharing controls "
            "that are documented in detail across the other workstream sections. "
            "Closing the Immediate and Near Term findings will have the most direct impact on this number."
        )
    if consent_f:
        paras.append(
            "The consent governance posture is worth separate attention. "
            "The current permission grant policy allows users to authorize third-party applications "
            "to access tenant data without admin involvement. "
            "In practice this means a user clicking Accept on an unfamiliar OAuth prompt "
            "grants that application ongoing read access to their mail, calendar, or files "
            "without the event appearing in an admin consent queue."
        )
    return paras


def _governance_story(data: dict, derived: dict, findings: list[dict]) -> list[str]:
    inactive_f  = _f(findings, "Inactive Licensed", "Inactive License")
    duplicate_f = _f(findings, "Duplicate Assignment", "Duplicate")
    ownerless_f = _f(findings, "Ownerless License", "Ownerless Group")

    parts = []
    if inactive_f:
        cnt = _count(inactive_f.get("Finding", ""))
        if cnt:
            parts.append(f"{cnt} inactive or disabled user{'s' if cnt != '1' else ''} still hold paid licenses")
    if duplicate_f:
        cnt = _count(duplicate_f.get("Finding", ""))
        if cnt:
            parts.append(f"{cnt} user{'s' if cnt != '1' else ''} carry both direct and group-based assignment for the same SKU")
    if ownerless_f:
        parts.append("at least one license-managing group has no assigned owner")

    paras = []
    if parts:
        combined = "; ".join(parts[:-1]) + (f"; and {parts[-1]}" if len(parts) > 1 else parts[0])
        paras.append(
            f"The governance review surfaced three license hygiene issues: {combined}. "
            "Duplicate SKU assignments do not extend the user's access but inflate spend calculations "
            "and create noise in license audits. "
            "An ownerless license group creates a change-control gap — "
            "there is no accountable owner for what happens when membership changes "
            "propagate automatically to license assignment."
        )

    # At-capacity SKUs from snapshot
    lic_data    = data.get("Identity", {}).get("LicenseSKUs", {})
    at_capacity = []
    if isinstance(lic_data, dict):
        for v in lic_data.values():
            if isinstance(v, dict):
                consumed = v.get("ConsumedUnits") or 0
                prepaid  = (v.get("PrepaidUnits") or {}).get("Enabled") or 0
                name     = v.get("SkuPartNumber", "")
                if prepaid > 0 and consumed >= prepaid and name:
                    at_capacity.append(name)
    if at_capacity:
        names = ", ".join(f"'{n}'" for n in at_capacity[:3])
        paras.append(
            f"License SKU(s) at or above capacity: {names}. "
            "At-capacity SKUs block new assignments silently and can cause provisioning failures "
            "during onboarding if they are not tracked against upcoming headcount changes."
        )
    return paras


# ---------------------------------------------------------------------------
# 11.0 Workloads Outside Assessment Scope
# ---------------------------------------------------------------------------

def _write_out_of_scope(doc: Document) -> None:
    _h(doc, "11.0 Workloads Outside Assessment Scope", 1)
    doc.add_paragraph(
        "This report covered the workloads and signals described in Sections 1-10. "
        "The capability areas listed below were not collected or analyzed in this engagement."
    )
    rows = [
        ("Microsoft Defender for Endpoint / Microsoft Defender XDR",      "Device compliance posture, endpoint protection signals, and threat analytics"),
        ("Microsoft Power Platform (Power Apps, Power Automate, Power BI)", "Power Platform environment sprawl, connector usage, and platform-level governance"),
        ("Microsoft Copilot for Microsoft 365",                             "Readiness posture, data oversharing risk, and prompt usage analytics"),
        ("Microsoft Purview Sensitivity Labels",                            "Label taxonomy, auto-labeling policy configuration, and classification coverage"),
        ("Microsoft Purview Audit Log",                                     "Unified audit log review, alert policy coverage, and forensic event query scope"),
    ]
    _table_2col(doc, rows, hdr=("Workload", "Scope Note"), col_widths=(2.3, 4.2))
    doc.add_page_break()


# ---------------------------------------------------------------------------
# 12.0 Appendix (documentation links)
# ---------------------------------------------------------------------------

def _write_appendix_links(doc: Document) -> None:
    _h(doc, "12.0 Appendix", 1)
    doc.add_paragraph(
        "The appendix provides Microsoft documentation references organized by workstream. "
        "Use these as starting points for the recommended actions in Section 4.0."
    )

    section = 1
    for area, links in _APPENDIX_LINKS.items():
        _h(doc, f"12.{section} {area}", 2)
        for title, url, why in links:
            p = doc.add_paragraph(style="List Bullet")
            run = p.add_run(f"{title} - {url}")
            run.font.size = Pt(10)
            p.paragraph_format.space_after = Pt(2)
            note = doc.add_paragraph(f"Why it is relevant: {why}")
            note.paragraph_format.left_indent = Inches(0.25)
            note.paragraph_format.space_after = Pt(6)
            if note.runs:
                note.runs[0].font.size = Pt(9)
                note.runs[0].font.color.rgb = RGBColor(0x59, 0x59, 0x59)
        section += 1
