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
    _write_executive_summary(doc, tenant_name, findings, ws_sums, actions, consultative, snapshot=snapshot)
    _write_recommendations(doc, actions)
    _write_workstream_sections(doc, snapshot, findings, consultative, actions)
    _write_out_of_scope(doc)
    _write_appendix_links(doc)
    _write_offboarding(doc, tenant_name, snapshot=snapshot)

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
    immediate_count = sum(1 for f in findings if f.get("RoadmapPhase") == "Immediate")
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
    snapshot: dict | None = None,
) -> None:
    _h(doc, "3.0 Executive Summary", 1)

    from collections import defaultdict

    # Pull key numbers from snapshot for the lede
    snapshot = snapshot or {}
    snap_data    = snapshot.get("Data", {})
    snap_derived = snapshot.get("Derived", {})
    identity     = snap_data.get("Identity", {})

    admins_raw  = identity.get("Admins", []) or []
    admins_list = admins_raw if isinstance(admins_raw, list) else list(admins_raw.values())
    ga_count    = sum(1 for a in admins_list if a.get("RoleName") == "Global Administrator" and a.get("accountEnabled", True))

    mfa_gap_raw = identity.get("MfaEnforcementGapUsers", []) or []
    gap_count   = len(mfa_gap_raw) if isinstance(mfa_gap_raw, list) else 0

    mfa_sum  = snap_derived.get("MfaRegistrationSummary", {})
    reg_pct  = float(mfa_sum.get("RegistrationPercent", 0))

    sec          = snap_data.get("Security", {})
    score_data   = sec.get("SecuritySecureScore", {}) or {}
    current_s    = score_data.get("currentScore") or 0
    max_s        = score_data.get("maxScore") or 0
    score_pct    = round(current_s / max_s * 100) if max_s else 0
    comp_scores  = score_data.get("averageComparativeScores") or []
    avg_all      = next((s.get("averageScore") for s in comp_scores if s.get("basis") == "AllTenants"), None)

    collab       = snap_data.get("Collaboration", {})
    teams_raw    = collab.get("TeamsGroupsCleanupCandidates", {})
    teams_list   = list(teams_raw.values()) if isinstance(teams_raw, dict) else (teams_raw if isinstance(teams_raw, list) else [])
    cleanup_cnt  = len(teams_list)
    ownerless_cnt = sum(1 for t in teams_list if "Ownerless" in str(t.get("RiskSignal", "")))

    high_count      = sum(1 for f in findings if f.get("Severity") in ("High", "Critical"))
    immediate_count = sum(1 for f in findings if f.get("RoadmapPhase") == "Immediate")

    # Also fix GA count to only count enabled GAs (consistent with plan.py)
    ga_count = sum(1 for a in admins_list if a.get("RoleName") == "Global Administrator" and a.get("accountEnabled", True) is not False)

    # --- Opening lede: specific numbers, not workstream names ---
    lede_bullets: list[str] = []
    if ga_count:
        lede_bullets.append(
            f"{ga_count} standing Global Administrator account{'s' if ga_count != 1 else ''} "
            "with permanent, always-on tenant-level access"
        )
    if gap_count:
        lede_bullets.append(
            f"{gap_count} member account{'s' if gap_count != 1 else ''} that can authenticate "
            "with only a password -- no MFA enforcement, no Conditional Access requirement"
        )
    if cleanup_cnt:
        lede_bullets.append(
            f"{cleanup_cnt} Teams and Microsoft 365 groups flagged for lifecycle review"
            + (f", including {ownerless_cnt} with no assigned owner" if ownerless_cnt else "")
        )

    if lede_bullets:
        lede_intro = f"At the time of this assessment, {tenant_name} presented the following measurable risk conditions:"
        doc.add_paragraph(lede_intro)
        for bullet in lede_bullets:
            p = doc.add_paragraph(bullet, style="List Bullet")
            p.paragraph_format.space_after = Pt(3)

        score_str = ""
        if current_s and max_s:
            delta_str = f" -- {current_s - avg_all:+.0f} vs. cross-tenant average" if avg_all else ""
            score_str = (
                f" The Microsoft Secure Score stands at {current_s:.0f} of {max_s:.0f} points "
                f"({score_pct}%{delta_str})."
            )
        doc.add_paragraph(
            "These are not aspirational improvements." + score_str + " "
            "They are measurable, remediable gaps in the tenant's security baseline that affect "
            "every user and resource in the environment."
        )
    elif not findings:
        doc.add_paragraph(
            f"The assessment of {tenant_name} did not identify significant open findings. "
            "The tenant configuration is broadly aligned with Microsoft 365 best practices."
        )

    # --- Business impact / leadership framing ---
    impact_parts: list[str] = []
    if ga_count:
        threshold = "above the recommended maximum of 4 for most organizations" if ga_count > 4 else "in a standing state"
        impact_parts.append(
            f"The {ga_count} standing Global Administrator{'s are' if ga_count != 1 else ' is'} {threshold}. "
            "Global Administrator is the broadest credential in the tenant -- a single compromised account "
            "at that level can disable MFA enforcement, export all user data, create backdoor accounts, "
            "and persist access even after an incident response has begun. "
            "The risk does not require a sophisticated attack: it requires only that one of those accounts "
            "clicks a phishing link or is included in a credential-stuffing list."
        )
    if gap_count:
        impact_parts.append(
            f"The {gap_count} account{'s' if gap_count != 1 else ''} without MFA enforcement "
            "represent the highest-probability entry point into the environment. "
            "Password spray and credential-phishing campaigns succeed specifically against accounts where "
            "no enforcing Conditional Access policy exists. "
            + (f"Of the {reg_pct:.0f}% of users who have registered an MFA method, "
               "registration alone provides zero protection without a policy that requires it on sign-in. "
               if reg_pct else "")
            + "Closing this gap requires Conditional Access policy changes, not user action."
        )
    if impact_parts:
        _h(doc, "What This Means for Leadership", 2)
        for part in impact_parts:
            p = doc.add_paragraph(part)
            p.paragraph_format.space_after = Pt(6)

    # --- Transition to tables ---
    if high_count or immediate_count:
        counts: list[str] = []
        if high_count:
            counts.append(f"{high_count} High or Critical finding{'s' if high_count != 1 else ''}")
        if immediate_count:
            counts.append(f"{immediate_count} Immediate action{'s' if immediate_count != 1 else ''} requiring prioritization")
        doc.add_paragraph(
            "The assessment identified " + " and ".join(counts) + ". "
            "The tables below summarize findings by workstream and surface the actions that "
            "require a leadership decision, owner assignment, or budget commitment. "
            "Evidence and technical remediation steps for each finding appear in the workstream sections."
        )

    # Findings by Workstream table
    _h(doc, "Findings by Workstream", 2)
    if ws_sums:
        ws_best: dict[str, dict] = {}
        for row in ws_sums:
            ws = row.get("Workstream", "")
            cur_sev = _SEV_ORDER.get(row.get("Severity", "Low"), 99)
            if ws not in ws_best or cur_sev < _SEV_ORDER.get(ws_best[ws].get("Severity", "Low"), 99):
                ws_best[ws] = row
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

    # Priority Actions for Leadership - structured table, not a bullet list
    _h(doc, "Priority Actions for Leadership", 2)
    doc.add_paragraph(
        "The table below shows the actions that require leadership prioritization or owner decision. "
        "Immediate items address active risk; Near Term items prevent risk from compounding."
    )
    priority_actions = [a for a in actions if a.get("RoadmapPhase") in ("Immediate", "Near Term")]
    if priority_actions:
        cols2   = ["Phase", "Action", "Workstream", "Why It Matters"]
        widths2 = [0.9, 2.2, 1.0, 2.4]
        tbl2 = doc.add_table(rows=1, cols=4)
        tbl2.style = "Table Grid"
        hdr2 = tbl2.rows[0].cells
        for i, (label, w) in enumerate(zip(cols2, widths2)):
            _set_cell_bg(hdr2[i], "1F3864")
            hdr2[i].width = Inches(w)
            p = hdr2[i].paragraphs[0]
            p.clear()
            run = p.add_run(label)
            run.bold = True
            run.font.color.rgb = _WHITE
            run.font.size = Pt(9)

        for action in priority_actions[:10]:
            phase = action.get("RoadmapPhase", "")
            why   = action.get("WhyItMatters", action.get("BusinessValue", ""))
            row = tbl2.add_row().cells
            row[0].text = phase
            row[1].text = action.get("ActionTitle", "")
            row[2].text = action.get("Workstream", "")
            row[3].text = why[:180] + ("..." if len(why) > 180 else "")
            phase_bg = "FFD9D9" if phase == "Immediate" else "FFE9CC"
            phase_fg = RGBColor(0xC0, 0x00, 0x00) if phase == "Immediate" else _ORANGE
            _set_cell_bg(row[0], phase_bg)
            if row[0].paragraphs[0].runs:
                row[0].paragraphs[0].runs[0].font.color.rgb = phase_fg
                row[0].paragraphs[0].runs[0].bold = True
            for cell in row:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(9)

        doc.add_paragraph()

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

        # Workstream-specific technical data tables
        _h(doc, "Supporting Data", 2)
        _ws_data_tables(ws, doc, snapshot)

        # Story narrative — specifics that the table doesn't show
        for para_text in _ws_story(ws, snapshot, ws_findings):
            p = doc.add_paragraph(para_text)
            p.paragraph_format.space_after = Pt(6)

        # Per-workstream recommended actions (sourced from roadmap actions)
        ws_actions = [a for a in actions if a.get("Workstream") == ws]
        if ws_actions:
            _h(doc, "Recommended Actions", 2)
            doc.add_paragraph(
                "The following actions are recommended for this workstream, in priority order. "
                "Effort and hour estimates are pre-sales planning references; confirm with the delivery team."
            )
            sorted_ws_actions = sorted(
                ws_actions,
                key=lambda a: (
                    _PHASE_ORDER.get(a.get("RoadmapPhase", "Monitor"), 99),
                    _SEV_ORDER.get(a.get("HighestSeverity", "Low"), 99),
                ),
            )
            cols   = ["Phase", "Action", "Effort", "Why It Matters"]
            widths = [0.8, 2.0, 0.7, 3.0]
            tbl = doc.add_table(rows=1, cols=4)
            tbl.style = "Table Grid"
            hdr = tbl.rows[0].cells
            for i, (label, w) in enumerate(zip(cols, widths)):
                _set_cell_bg(hdr[i], "1F3864")
                hdr[i].width = Inches(w)
                p = hdr[i].paragraphs[0]; p.clear()
                run = p.add_run(label)
                run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
            for action in sorted_ws_actions:
                phase = action.get("RoadmapPhase", "")
                why   = action.get("WhyItMatters", action.get("BusinessValue", ""))
                row = tbl.add_row().cells
                row[0].text = phase
                row[1].text = action.get("ActionTitle", "")
                row[2].text = action.get("EffortTier", "")
                row[3].text = (why[:200] + "..." if len(why) > 200 else why)
                phase_bg = "FFD9D9" if phase == "Immediate" else ("FFE9CC" if phase == "Near Term" else "EBF3FB")
                phase_fg = RGBColor(0xC0, 0x00, 0x00) if phase == "Immediate" else (_ORANGE if phase == "Near Term" else _NAVY)
                _set_cell_bg(row[0], phase_bg)
                if row[0].paragraphs[0].runs:
                    row[0].paragraphs[0].runs[0].font.color.rgb = phase_fg
                    row[0].paragraphs[0].runs[0].bold = True
                for cell in row:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.font.size = Pt(9)
            doc.add_paragraph()

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
    identity     = data.get("Identity", {})
    users        = identity.get("Users", {})
    user_list    = list(users.values()) if isinstance(users, dict) else (users if isinstance(users, list) else [])
    user_count   = len(user_list)
    members      = [u for u in user_list if str(u.get("userType") or u.get("UserType") or "").lower() == "member"]
    guests       = [u for u in user_list if str(u.get("userType") or u.get("UserType") or "").lower() == "guest"]

    # Count enabled GAs directly from Admins list for narrative consistency
    admins_story = identity.get("Admins", []) or []
    admins_story = admins_story if isinstance(admins_story, list) else list(admins_story.values())
    ga_count = sum(1 for a in admins_story if a.get("RoleName") == "Global Administrator" and a.get("accountEnabled", True) is not False)
    # Stale = enabled GAs with no sign-in activity (or no sign-in data)
    from datetime import datetime as _dt2, timezone as _tz2
    _now2 = _dt2.now(_tz2.utc)
    def _stale_ga(a: dict) -> bool:
        sia = a.get("signInActivity") or {}
        last = sia.get("lastSignInDateTime") or sia.get("lastNonInteractiveSignInDateTime")
        if not last:
            return True
        try:
            return (_now2 - _dt2.fromisoformat(last.replace("Z", "+00:00"))).days >= 180
        except (ValueError, AttributeError):
            return True
    enabled_gas_story = [a for a in admins_story if a.get("RoleName") == "Global Administrator" and a.get("accountEnabled", True) is not False]
    stale_ga = sum(1 for a in enabled_gas_story if _stale_ga(a))

    mfa_gap  = identity.get("MfaEnforcementGapUsers", [])
    gap_count = len(mfa_gap) if isinstance(mfa_gap, list) else 0

    mfa_sum  = derived.get("MfaRegistrationSummary", {})
    reg_pct  = float(mfa_sum.get("RegistrationPercent", 0))
    weak_only = int(mfa_sum.get("UsersWithWeakMethodsOnly", 0))

    ca_opt_raw = identity.get("ConditionalAccessOptimization") or []
    ca_opt_list = ca_opt_raw if isinstance(ca_opt_raw, list) else list(ca_opt_raw.values()) if isinstance(ca_opt_raw, dict) else []
    report_only_policies = [x for x in ca_opt_list if "report" in str(x.get("Signal", "")).lower()]
    excl_policies        = [x for x in ca_opt_list if "exclusion" in str(x.get("Signal", "")).lower()]

    paras = []

    # Para 1: user landscape and privilege posture
    enabled_members_story = [u for u in members if u.get("accountEnabled") or u.get("AccountEnabled")]
    enabled_guests_story  = [u for u in guests  if u.get("accountEnabled") or u.get("AccountEnabled")]
    user_str = (
        f"{len(members)} internal member{'s' if len(members) != 1 else ''} "
        f"({len(enabled_members_story)} active) and "
        f"{len(guests)} guest{'s' if len(guests) != 1 else ''} ({len(enabled_guests_story)} active)"
        if (members or guests) else f"{user_count} accounts"
    )
    if user_count:
        ga_str = (
            f" The most urgent privilege concern is the {ga_count} standing Global Administrator "
            f"account{'s' if ga_count != 1 else ''}"
            + (f" — none show recent interactive sign-in activity" if stale_ga == ga_count and stale_ga
               else (f", {stale_ga} of which show no recent interactive sign-in activity" if stale_ga else ""))
            + ". Global Administrator is the broadest credential in the tenant — a compromised account "
            "at that level bypasses every role-scoped control in the environment."
            if ga_count else ""
        )
        paras.append(
            f"The identity review covers {user_str}.{ga_str}"
        )

    # Para 2: MFA enforcement gap — specific and high-impact
    if gap_count:
        paras.append(
            f"{gap_count} member account{'s' if gap_count != 1 else ''} are not covered by any "
            "Conditional Access policy that enforces MFA or device compliance. "
            "These accounts can authenticate to Microsoft 365 with only a password — "
            "no second factor required regardless of location, device, or sign-in risk level. "
            + (f"Of the {int(reg_pct)}% of users who have registered an MFA method, "
               f"registration alone provides no protection without an enforcing policy. " if reg_pct else "")
            + (f"{weak_only} registered user{'s' if weak_only != 1 else ''} rely exclusively on SMS or phone call — "
               "methods vulnerable to SIM-swap and real-time phishing. " if weak_only else "")
            + "Enforcement through Conditional Access is the only control that converts "
            "method registration into an actual access requirement."
        )

    # Para 3: CA policy quality gaps
    ca_parts = []
    if report_only_policies:
        n = len(report_only_policies)
        ca_parts.append(f"{n} polic{'ies' if n != 1 else 'y'} remain in report-only mode and enforce nothing")
    if excl_policies:
        n = len(excl_policies)
        ca_parts.append(f"{n} polic{'ies' if n != 1 else 'y'} contain exclusions that may leave specific accounts or groups unprotected")
    if ca_parts:
        paras.append(
            "The Conditional Access policy set has structural quality gaps beyond unenforced MFA: "
            + "; ".join(ca_parts) + ". "
            "Report-only policies are a useful testing tool but not a security control — "
            "they log what would happen without preventing anything."
        )

    return paras


def _messaging_story(data: dict, derived: dict, findings: list[dict]) -> list[str]:
    domain_f = _f(findings, "Domain Verification", "Unverified")
    dmarc_f  = _f(findings, "DMARC")
    fwd_f    = _f(findings, "Forwarding", "Mailbox Forwarding")

    ex = data.get("Exchange", {})
    activity = ex.get("EmailActivitySummary", {})
    if isinstance(activity, list) and activity:
        activity = activity[0]
    top_senders = ex.get("EmailActivityTopSenders", []) or []

    paras = []

    # Para 1: email activity context from Graph reporting data
    if isinstance(activity, dict) and activity:
        send_count    = activity.get("Send Count") or activity.get("SendCount") or activity.get("send_count")
        receive_count = activity.get("Receive Count") or activity.get("ReceiveCount") or activity.get("receive_count")
        period        = activity.get("Report Period") or activity.get("ReportPeriod") or "30"
        active_senders = len([s for s in top_senders if isinstance(s, dict) and int(s.get("Send Count", s.get("SendCount", 0)) or 0) > 0]) if top_senders else None

        activity_parts = []
        if send_count:
            activity_parts.append(f"{int(send_count):,} messages sent")
        if receive_count:
            activity_parts.append(f"{int(receive_count):,} received")
        if activity_parts:
            sender_str = f", with {active_senders} distinct active senders" if active_senders else ""
            paras.append(
                f"Email activity over the last {period} days: {' and '.join(activity_parts)}{sender_str}. "
                "This provides a baseline for normal mail volume — significant deviations from this pattern "
                "can indicate a compromised account sending at scale, bulk forwarding rules, or a misconfigured connector."
            )

    # Para 2: domain and anti-spoofing gaps
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

    # Para 3: forwarding
    if fwd_f:
        cnt = _count(fwd_f.get("Finding", ""))
        label = f"{cnt} mailbox{'es' if cnt != '1' else ''} have" if cnt else "Mailboxes have"
        paras.append(
            f"{label} active forwarding configured to send copies outside the organization. "
            "External forwarding creates a persistent data-exit path that runs independently of "
            "DLP policies unless the policy is specifically scoped to cover outbound forwarding rules."
        )

    return paras


def _collab_story(data: dict, derived: dict, findings: list[dict]) -> list[str]:
    collab     = data.get("Collaboration", {})
    teams_raw  = collab.get("TeamsGroupsCleanupCandidates", {})
    teams_list = list(teams_raw.values()) if isinstance(teams_raw, dict) else (teams_raw if isinstance(teams_raw, list) else [])

    activity     = collab.get("CollaborationActivitySummary", {})
    total_grp    = activity.get("TotalUnifiedGroups") or activity.get("TotalGroups") or 0
    active_grp   = activity.get("GroupsWithActivityInPeriod") or activity.get("ActiveGroups") or 0
    inactive_grp = max(0, total_grp - active_grp)

    og = derived.get("OwnershipGovernanceSummary", {}).get("Summary", {})
    total_objects  = og.get("TotalObjectsReviewed", 0)
    ownerless_cnt  = og.get("OwnerlessObjectCount", 0)
    cleanup_cnt    = og.get("CleanupCandidateCount", len(teams_list))

    # Named team examples — prefer ownerless, then dormant
    ownerless_teams = [t for t in teams_list if "Ownerless" in str(t.get("RiskSignal", ""))]
    dormant_teams   = [t for t in teams_list if "Dormant" in str(t.get("RiskSignal", "")) and t not in ownerless_teams]
    sample          = (ownerless_teams + dormant_teams)[:3]
    team_examples   = []
    for t in sample:
        name = t.get("Name", "")
        risk = t.get("RiskSignal", "").split(";")[0].strip()
        last = (t.get("LastActivityDate") or "")[:7] or "no recorded activity"
        if name:
            team_examples.append(f"'{name}' ({risk.lower()}, last active {last})")

    paras = []

    # Para 1: activity baseline from Graph reporting data
    if total_grp:
        inactive_pct = round(inactive_grp / total_grp * 100) if total_grp else 0
        paras.append(
            f"The collaboration review assessed {total_grp} Microsoft 365 groups and Teams. "
            f"{active_grp} ({100 - inactive_pct}%) showed activity in the last 30 days; "
            f"{inactive_grp} ({inactive_pct}%) showed none. "
            "Inactive groups continue to hold permissions, memberships, and associated SharePoint sites "
            "regardless of their activity state — they remain a live access surface even when no one is using them."
        )

    # Para 2: named cleanup candidates with specifics
    if team_examples:
        ex_text = "; ".join(team_examples[:2])
        ownership_str = f" {ownerless_cnt} have no assigned owner." if ownerless_cnt else ""
        paras.append(
            f"{cleanup_cnt} group{'s' if cleanup_cnt != 1 else ''} are identified as cleanup candidates.{ownership_str} "
            f"Examples from the candidate list: {ex_text}. "
            "These are not outliers — they represent the ownership and lifecycle pattern across the inventory. "
            "Ownerless groups are the hardest to decommission: there is no accountable party to confirm "
            "whether the group is still in use, whether its membership is correct, or whether its content can be archived."
        )
    elif cleanup_cnt:
        paras.append(
            f"{cleanup_cnt} group{'s' if cleanup_cnt != 1 else ''} are flagged as cleanup candidates "
            + (f"including {ownerless_cnt} with no assigned owner. " if ownerless_cnt else ". ")
            + "Without an owner, there is no accountable party for lifecycle decisions on these objects."
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

    sec          = data.get("Security", {})
    score_data   = sec.get("SecuritySecureScore", {}) or {}
    controls     = sec.get("SecureScoreActions", []) or []

    current  = score_data.get("currentScore") or 0
    max_s    = score_data.get("maxScore") or 0
    comp_scores = score_data.get("averageComparativeScores") or []
    avg_all  = next((s.get("averageScore") for s in comp_scores if s.get("basis") == "AllTenants"), None)

    not_impl = [c for c in controls if isinstance(c, dict) and c.get("implementationStatus") == "notImplemented"]
    top_controls = sorted(not_impl, key=lambda c: -(c.get("maxScore") or 0))[:3]

    paras = []

    # Para 1: actual score numbers
    if current and max_s:
        pct = round(current / max_s * 100)
        avg_str = (f" The cross-tenant average is {avg_all:.0f} points." if avg_all else "")
        paras.append(
            f"The Microsoft Secure Score for this tenant is {current:.0f} out of {max_s:.0f} points ({pct}%).{avg_str} "
            "Each point represents a specific control Microsoft rates as reducing risk for this tenant profile. "
            "Secure Score is a composite signal — the gaps pulling it down are distributed across identity governance, "
            "MFA enforcement, device compliance, and sharing controls documented in the other workstream sections."
        )
    elif score_f:
        paras.append(
            score_f.get("Finding", "") + " "
            "Closing the Immediate and Near Term findings will have the most direct impact on this number."
        )

    # Para 2: top unimplemented controls by point value
    if top_controls:
        control_lines = "; ".join(
            f"'{c.get('title', '')}' ({c.get('maxScore', 0):.0f} pts)"
            for c in top_controls
        )
        paras.append(
            f"The three highest-point unimplemented controls are: {control_lines}. "
            "Addressing these will produce the largest score improvement and close the gaps "
            "Microsoft rates as most significant for this tenant."
        )

    # Para 3: consent governance
    if consent_f:
        paras.append(
            "The consent governance posture is worth separate attention. "
            "The current permission grant policy allows users to authorize third-party applications "
            "to access tenant data without admin involvement. "
            "In practice this means a user clicking Accept on an OAuth prompt "
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
            f"The governance review surfaced license hygiene issues: {combined}. "
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

    # Password lifecycle from Graph data
    pwd = data.get("Governance", {}).get("PasswordLifecycleSummary", {}) or {}
    if pwd:
        validity      = pwd.get("PasswordValidityPeriodInDays")
        never_expires = pwd.get("PasswordNeverExpires") or validity == 2147483647
        sspr          = pwd.get("SelfServicePasswordResetEnabled")
        pwd_parts = []
        if never_expires:
            pwd_parts.append("passwords on the default domain are set to never expire")
        elif validity:
            pwd_parts.append(f"passwords expire every {validity} days")
        if sspr is True:
            pwd_parts.append("self-service password reset is enabled")
        elif sspr is False:
            pwd_parts.append("self-service password reset is not enabled — users must contact the helpdesk to reset credentials")
        if pwd_parts:
            paras.append(
                "Password lifecycle: " + "; ".join(pwd_parts) + ". "
                + ("A never-expire policy shifts full reliance onto MFA and CA enforcement — "
                   "if those controls have gaps, a compromised credential remains valid indefinitely. "
                   if never_expires else "")
            )

    return paras


# ---------------------------------------------------------------------------
# Workstream data tables — technical detail tables per workstream section
# ---------------------------------------------------------------------------

def _tbl_label(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = _NAVY
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)


def _ws_data_tables(ws: str, doc: Document, snapshot: dict) -> None:
    if ws == "Identity":
        _identity_data_tables(doc, snapshot)
    elif ws == "Messaging":
        _messaging_data_tables(doc, snapshot)
    elif ws == "Collaboration":
        _collab_data_tables(doc, snapshot)
    elif ws == "Endpoint":
        _endpoint_data_tables(doc, snapshot)
    elif ws == "Security":
        _security_data_tables(doc, snapshot)
    elif ws == "Governance":
        _governance_data_tables(doc, snapshot)


def _identity_data_tables(doc: Document, snapshot: dict) -> None:
    data    = snapshot.get("Data", {})
    derived = snapshot.get("Derived", {})

    # --- User Account Breakdown ---
    users_raw = data.get("Identity", {}).get("Users", {})
    users_list = list(users_raw.values()) if isinstance(users_raw, dict) else (users_raw if isinstance(users_raw, list) else [])
    if users_list:
        members          = [u for u in users_list if str(u.get("userType") or u.get("UserType") or "").lower() == "member"]
        guests           = [u for u in users_list if str(u.get("userType") or u.get("UserType") or "").lower() == "guest"]
        enabled_members  = [u for u in members if u.get("accountEnabled") or u.get("AccountEnabled")]
        enabled_guests   = [u for u in guests  if u.get("accountEnabled") or u.get("AccountEnabled")]

        _tbl_label(doc, "User Account Breakdown")
        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        hdr = tbl.rows[0].cells
        for i, (label, w) in enumerate(zip(["Account Type", "Total", "Enabled"], [2.5, 1.0, 1.0])):
            _set_cell_bg(hdr[i], "1F3864")
            hdr[i].width = Inches(w)
            p = hdr[i].paragraphs[0]; p.clear()
            run = p.add_run(label)
            run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)

        for acct_type, total, enabled in [
            ("Internal Members", len(members),  len(enabled_members)),
            ("Guests",           len(guests),   len(enabled_guests)),
            ("Total",            len(users_list), len(enabled_members) + len(enabled_guests)),
        ]:
            row = tbl.add_row().cells
            for i, val in enumerate([acct_type, str(total), str(enabled)]):
                p = row[i].paragraphs[0]
                run = p.add_run(val)
                run.font.size = Pt(9)
                if acct_type == "Total":
                    run.bold = True
        doc.add_paragraph()

    # --- MFA Registration Coverage ---
    mfa = derived.get("MfaRegistrationSummary", {})
    if mfa and mfa.get("TotalUsers"):
        total_mfa  = mfa.get("TotalUsers", 0)
        registered = mfa.get("RegisteredUsers", 0)
        not_reg    = mfa.get("NotRegisteredUsers", 0)
        reg_pct    = float(mfa.get("RegistrationPercent", 0.0))

        rows: list[tuple[str, str]] = [
            ("Users reviewed for MFA",                     str(total_mfa)),
            ("Registered for MFA",                         f"{registered} ({reg_pct:.1f}%)"),
            ("Not registered",                             f"{not_reg} ({100.0 - reg_pct:.1f}%)"),
            ("With phishing-resistant method (WHfB/FIDO2)", str(mfa.get("UsersWithPhishingResistantMethods", 0))),
            ("With weak method only (SMS / phone)",        str(mfa.get("UsersWithWeakMethodsOnly", 0))),
        ]
        method_counts: dict = mfa.get("MethodCounts", {}) or {}
        if method_counts:
            rows.append(("-- Authentication Method Enrollment --", ""))
            for method, count in sorted(method_counts.items(), key=lambda x: -x[1]):
                rows.append((f"  {method}", str(count)))

        _tbl_label(doc, "MFA Registration Coverage")
        _table_2col(doc, rows, hdr=("Metric", "Value"), col_widths=(3.0, 1.5))

    # --- Conditional Access Policy Summary ---
    ca_sum = data.get("Identity", {}).get("ConditionalAccessPolicySummary", {}).get("Summary", {})
    if ca_sum:
        bool_disp = {True: "Yes", False: "No"}
        total_p = ca_sum.get("TotalPolicies", 0)
        enabled = ca_sum.get("EnabledPolicies", 0)
        report  = ca_sum.get("ReportOnlyPolicies", 0)

        ca_rows: list[tuple[str, str]] = [
            ("Total policies defined",                          str(total_p)),
            ("Enabled (actively enforced)",                     str(enabled)),
            ("Report-only (monitoring, not enforced)",          str(report)),
            ("Disabled",                                        str(max(0, total_p - enabled - report))),
            ("Policies with exclusions",                        str(ca_sum.get("PoliciesWithExclusions", 0))),
            ("Blocking legacy authentication",                  str(ca_sum.get("PoliciesBlockingLegacyAuth", 0))),
            ("Requiring compliant or Entra-joined device",      str(ca_sum.get("PoliciesRequiringCompliantDevice", 0))),
            ("Covering guest / external users",                 bool_disp.get(ca_sum.get("HasGuestCoverage"), "Unknown")),
            ("Covering privileged admin roles",                 bool_disp.get(ca_sum.get("HasPrivilegedRoleCoverage"), "Unknown")),
            ("Using sign-in or user risk signals",              bool_disp.get(ca_sum.get("HasRiskBasedCoverage"), "Unknown")),
        ]
        _tbl_label(doc, "Conditional Access Policy Summary")
        _table_2col(doc, ca_rows, hdr=("Conditional Access Signal", "Current State"), col_widths=(3.0, 1.5))

    # --- Security Defaults Status ---
    sec_defaults = data.get("Identity", {}).get("SecurityDefaultsPolicy", {}) or {}
    if sec_defaults:
        is_enabled = sec_defaults.get("isEnabled", False)
        has_ca     = bool(ca_sum and ca_sum.get("TotalPolicies", 0) > 0)
        sd_rows: list[tuple[str, str]] = [
            ("Security Defaults enabled",                "Yes" if is_enabled else "No"),
            ("Conditional Access policies defined",      "Yes" if has_ca else "No"),
            ("Assessment",
             ("Security Defaults active — CA policies not in use" if is_enabled
              else ("Disabled — tenant uses Conditional Access (recommended)" if has_ca
                    else "Disabled — no CA policies detected; verify authentication controls"))),
        ]
        _tbl_label(doc, "Security Defaults Status")
        tbl = doc.add_table(rows=1, cols=2)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(["Setting", "State"], [3.0, 2.5])):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        for setting, state in sd_rows:
            row = tbl.add_row().cells
            warn = (not is_enabled and not has_ca and setting == "Assessment")
            ok   = (not is_enabled and has_ca and setting == "Assessment")
            for i, val in enumerate([setting, state]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                if warn:
                    _set_cell_bg(row[i], "FFE9CC")
                    if i == 1:
                        run.font.color.rgb = _ORANGE; run.bold = True
                elif ok and i == 1:
                    run.font.color.rgb = RGBColor(0x37, 0x63, 0x2F)
        doc.add_paragraph()

    # --- Global Administrator Inventory ---
    admins_raw  = data.get("Identity", {}).get("Admins", []) or []
    admins_list = admins_raw if isinstance(admins_raw, list) else list(admins_raw.values())
    gas = sorted(
        [a for a in admins_list if a.get("RoleName") == "Global Administrator"],
        key=lambda a: a.get("displayName", ""),
    )
    # Build sign-in lookup from Users (admins batch doesn't return signInActivity)
    from datetime import datetime as _dt, timezone as _tz
    _now = _dt.now(_tz.utc)
    _user_signin: dict[str, str] = {}
    for upn, u in (users_raw.items() if isinstance(users_raw, dict) else []):
        # PS snapshots store SignInActivity as a string (PS object serialization) or a dict
        sia_raw = u.get("SignInActivity") or u.get("signInActivity")
        sia = sia_raw if isinstance(sia_raw, dict) else {}
        last = (sia.get("LastSignInDateTime") or sia.get("lastSignInDateTime") or
                sia.get("LastNonInteractiveSignInDateTime") or sia.get("lastNonInteractiveSignInDateTime") or "")
        _user_signin[upn.lower()] = last[:10] if last else ""
    if gas:
        _tbl_label(doc, "Global Administrator Inventory")
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(["Name", "User Principal Name", "Status", "Last Sign-in"], [2.0, 2.5, 0.75, 1.0])):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        for a in gas:
            enabled  = a.get("accountEnabled", True)
            upn      = a.get("userPrincipalName", "")
            last_in  = _user_signin.get(upn.lower(), "")
            status   = "Active" if enabled else "Disabled"
            try:
                days_ago = (_now - _dt.fromisoformat(last_in + "T00:00:00+00:00")).days if last_in else None
                stale    = days_ago is None or days_ago > 180
            except (ValueError, AttributeError):
                days_ago = None
                stale    = True
            signin_label = (f"{last_in} ({days_ago}d ago)" if days_ago is not None else "No sign-in data")
            row = tbl.add_row().cells
            bg  = "F2F2F2" if not enabled else ("FFF2CC" if stale else None)
            for i, val in enumerate([a.get("displayName", ""), upn, status, signin_label]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                if not enabled:
                    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
                if stale and i == 3 and enabled:
                    run.font.color.rgb = _ORANGE
                    run.bold = True
                if bg:
                    _set_cell_bg(row[i], bg)
        doc.add_paragraph(
            "Stale sign-in (>180 days or no data) is highlighted in amber. "
            "Disabled accounts are shown in grey."
        )
        doc.add_paragraph()

    # --- Conditional Access Policy Inventory ---
    caps = data.get("Identity", {}).get("ConditionalAccessPolicies", []) or []
    if isinstance(caps, dict):
        caps = list(caps.values())
    if caps:
        _state_label = {
            "enabled": "Enabled",
            "enabledForReportingButNotEnforced": "Report Only",
            "disabled": "Disabled",
        }
        _state_bg = {
            "enabled": "E2F0D9",
            "enabledForReportingButNotEnforced": "FFE9CC",
            "disabled": "F2F2F2",
        }
        _state_order = {"enabled": 0, "enabledForReportingButNotEnforced": 1, "disabled": 2}

        def _cap_name(p: dict) -> str:
            return p.get("DisplayName") or p.get("displayName") or ""

        def _cap_state(p: dict) -> str:
            return p.get("State") or p.get("state") or ""

        caps_sorted = sorted(caps, key=lambda p: (_state_order.get(_cap_state(p), 3), _cap_name(p)))
        _tbl_label(doc, "Conditional Access Policy Inventory")
        tbl = doc.add_table(rows=1, cols=2)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(["Policy", "State"], [4.0, 1.5])):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        for policy in caps_sorted:
            state = _cap_state(policy)
            row = tbl.add_row().cells
            bg = _state_bg.get(state)
            for i, val in enumerate([_cap_name(policy), _state_label.get(state, state)]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                if bg:
                    _set_cell_bg(row[i], bg)
        doc.add_paragraph()

    # --- MFA Enforcement Gap Users ---
    gap_users = data.get("Identity", {}).get("MfaEnforcementGapUsers", []) or []
    if isinstance(gap_users, list) and gap_users:
        _tbl_label(doc, f"MFA Enforcement Gap Users ({len(gap_users)} total)")
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(["User Principal Name", "Gap", "Status", "Last Sign-in"], [2.7, 1.5, 0.65, 1.0])):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        _gap_labels = {
            "NotRegisteredForMFA": "Not registered for MFA",
            "NoEnforcingPolicy":   "No enforcing CA policy",
        }
        for u in gap_users[:20]:
            upn = u.get("UserPrincipalName", "")
            usr_obj = users_raw.get(upn) if isinstance(users_raw, dict) else {}
            usr_obj = usr_obj or {}
            acct_enabled = usr_obj.get("accountEnabled")
            status_lbl = "Active" if acct_enabled else ("Disabled" if acct_enabled is False else "Unknown")
            sia = usr_obj.get("signInActivity") or {}
            last = (sia.get("lastSignInDateTime") or sia.get("lastNonInteractiveSignInDateTime") or "")[:10]
            row = tbl.add_row().cells
            _set_cell_bg(row[0], "FFD9D9"); _set_cell_bg(row[1], "FFD9D9")
            _set_cell_bg(row[2], "FFD9D9"); _set_cell_bg(row[3], "FFD9D9")
            for i, val in enumerate([upn, _gap_labels.get(u.get("GapType", ""), u.get("GapType", "")), status_lbl, last or "No data"]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
        if len(gap_users) > 20:
            row = tbl.add_row().cells
            p = row[0].paragraphs[0]; run = p.add_run(f"... and {len(gap_users) - 20} more"); run.font.size = Pt(9); run.italic = True
        doc.add_paragraph()

    # --- CA Policy Optimization Issues ---
    ca_opt = data.get("Identity", {}).get("ConditionalAccessOptimization", []) or []
    ca_opt_list = ca_opt if isinstance(ca_opt, list) else list(ca_opt.values())
    if ca_opt_list:
        _tbl_label(doc, "Conditional Access Policy Issues")
        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(["Policy", "Issue", "Severity"], [2.5, 2.5, 0.7])):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        _opt_sev_bg = {"High": "FFD9D9", "Medium": "FFE9CC", "Low": "E2F0D9", "Info": "EBF3FB"}
        _opt_sev_fg = {"High": RGBColor(0xC0, 0x00, 0x00), "Medium": _ORANGE, "Low": RGBColor(0x37, 0x63, 0x2F), "Info": _NAVY}
        for issue in ca_opt_list:
            sev = issue.get("Severity", "Info")
            row = tbl.add_row().cells
            bg = _opt_sev_bg.get(sev, "EBF3FB")
            for i, val in enumerate([issue.get("PolicyName", ""), issue.get("Signal", ""), sev]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                if i == 2:
                    run.font.color.rgb = _opt_sev_fg.get(sev, _NAVY); run.bold = True
                    _set_cell_bg(row[i], bg)
        doc.add_paragraph()

    # --- MFA-Enforcing Policy Scope Review ---
    mfa_scope = data.get("Identity", {}).get("MfaEnforcementScopeReview", []) or []
    if mfa_scope:
        _tbl_label(doc, "MFA-Enforcing Policy Scope Review")
        _scope_state_label = {
            "enabled": "Enforced",
            "enabledForReportingButNotEnforced": "Report Only",
            "disabled": "Disabled",
        }
        _scope_state_bg = {
            "enabled": "E2F0D9",
            "enabledForReportingButNotEnforced": "FFE9CC",
            "disabled": "F2F2F2",
        }
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(["Policy", "Enforcement", "Scope", "Exclusions"], [2.2, 1.0, 1.2, 0.8])):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        for pol in mfa_scope:
            if not isinstance(pol, dict):
                continue
            state   = pol.get("State", "")
            inc     = pol.get("IncludeUsers", []) or []
            scope   = "All Users" if "All" in inc else f"{len(inc)} user/group(s)"
            excl    = pol.get("ExcludeUserCount", 0) or 0
            bg      = _scope_state_bg.get(state)
            row = tbl.add_row().cells
            for i, val in enumerate([pol.get("PolicyName", ""), _scope_state_label.get(state, state), scope, str(excl)]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                if bg:
                    _set_cell_bg(row[i], bg)
        doc.add_paragraph()

    # --- Guest Access Configuration ---
    gac = data.get("Identity", {}).get("GuestAccessConfiguration", {}) or {}
    if gac and gac.get("allowInvitesFrom") is not None:
        _INVITE_LABELS = {
            "none":                      "Disabled (no guest invitations)",
            "adminsAndGuestInviters":    "Admins and Guest Inviters only",
            "adminsGuestInvitersAndMembers": "Admins, Guest Inviters, and Members",
            "everyone":                  "Everyone (any user can invite guests)",
        }
        invite_from  = gac.get("allowInvitesFrom", "")
        guest_role   = gac.get("guestUserRoleId", "")
        # Decode guest role: limited guest=2af84b1e, guest user=10dae51f, member=a0b1b346
        _ROLE_LABELS = {
            "2af84b1e-5c62-4b32-bcd4-b9d1e39f5e9e": "Restricted guest (limited access)",
            "10dae51f-b6af-4016-8d66-8c2a99b929b3": "Guest user (default)",
            "a0b1b346-4d3e-4e8b-98f8-753987be4970": "Member-equivalent guest (broad access)",
        }
        default_perms = gac.get("defaultUserRolePermissions", {}) or {}
        gac_rows: list[tuple[str, str]] = [
            ("Guest invitation permissions",     _INVITE_LABELS.get(invite_from, invite_from)),
            ("Guest user role",                  _ROLE_LABELS.get(guest_role, "Custom")),
            ("Members can create security groups", "Yes" if default_perms.get("allowedToCreateSecurityGroups") else "No"),
            ("Members can create M365 groups",    "Yes" if default_perms.get("allowedToCreateMicrosoft365Groups") else "No"),
            ("Members can register apps",         "Yes" if default_perms.get("allowedToCreateApps") else "No"),
            ("Users can use SSPR",                "Yes" if gac.get("allowedToUseSSPR") else "No"),
        ]
        _tbl_label(doc, "Guest Access Configuration")
        tbl = doc.add_table(rows=1, cols=2)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(["Setting", "Current State"], [3.0, 2.5])):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        for setting, state in gac_rows:
            row = tbl.add_row().cells
            risky = (setting == "Guest invitation permissions" and "everyone" in state.lower())
            for i, val in enumerate([setting, state]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                if risky:
                    _set_cell_bg(row[i], "FFD9D9")
                    if i == 1:
                        run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00); run.bold = True
        doc.add_paragraph()

    # --- AD Connect / Hybrid Sync Status ---
    adc = snapshot.get("Data", {}).get("Tenant", {}).get("AdConnectConfiguration", {}) or {}
    adc_sum = adc.get("Summary", {}) or {}
    if adc_sum.get("OnPremisesSyncEnabled"):
        from datetime import datetime as _dtadc, timezone as _tzadc
        _last_sync = adc_sum.get("LastSyncDateTime") or ""
        _last_pw   = adc_sum.get("LastPasswordSyncDateTime") or ""
        def _days_ago_adc(dt_str: str) -> str:
            if not dt_str:
                return "Never / Unknown"
            try:
                d = _dtadc.fromisoformat(dt_str.replace("Z", "+00:00"))
                days = (_dtadc.now(_tzadc.utc) - d).days
                return f"{dt_str[:10]} ({days} days ago)"
            except (ValueError, AttributeError):
                return dt_str[:10]
        sync_rows: list[tuple[str, str]] = [
            ("Azure AD Connect sync enabled",  "Yes"),
            ("Last directory sync",            _days_ago_adc(_last_sync)),
            ("Last password sync",             _days_ago_adc(_last_pw)),
        ]
        _tbl_label(doc, "AD Connect / Hybrid Sync Status")
        tbl = doc.add_table(rows=1, cols=2)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(["Setting", "Current State"], [3.0, 2.5])):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        for setting, state in sync_rows:
            row = tbl.add_row().cells
            stale = "days ago" in state and int(state.split("(")[1].split(" ")[0]) > 3 if "days ago" in state else False
            for i, val in enumerate([setting, state]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                if stale and i == 1:
                    _set_cell_bg(row[i], "FFD9D9"); run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00); run.bold = True
        doc.add_paragraph()

    # --- Enterprise Applications Summary ---
    app_sum = data.get("Identity", {}).get("EnterpriseApplicationSummary", {}) or {}
    apps_raw = data.get("Identity", {}).get("EnterpriseApplications", []) or []
    apps_list = apps_raw if isinstance(apps_raw, list) else list(apps_raw.values())
    if app_sum and app_sum.get("TotalApplications"):
        total_apps  = app_sum.get("TotalApplications", 0)
        with_perms  = app_sum.get("ApplicationsWithPermissions", 0)
        verified    = app_sum.get("VerifiedPublisherApps", 0)
        unverified  = app_sum.get("UnverifiedApps", 0)
        app_sum_rows: list[tuple[str, str]] = [
            ("Total registered service principals",   str(total_apps)),
            ("With configured permissions",           str(with_perms)),
            ("Verified publisher",                    str(verified)),
            ("Unverified publisher",                  str(unverified)),
        ]
        _tbl_label(doc, "Enterprise Application Summary")
        _table_2col(doc, app_sum_rows, hdr=("Metric", "Count"), col_widths=(3.0, 1.5))
        doc.add_paragraph()

    # Unverified apps with permissions — review candidates
    if apps_list:
        risky_apps = [
            a for a in apps_list
            if not (a.get("verifiedPublisher") or {}).get("verifiedPublisherId")
            and (a.get("oauth2PermissionScopes") or a.get("appRoles"))
            and str(a.get("displayName", "")).strip()
        ]
        risky_apps = sorted(risky_apps, key=lambda a: a.get("displayName", ""))
        if risky_apps:
            _tbl_label(doc, f"Unverified Applications with Permissions ({len(risky_apps)} total, top 20 shown)")
            tbl = doc.add_table(rows=1, cols=3)
            tbl.style = "Table Grid"
            for i, (label, w) in enumerate(zip(["Application Name", "App ID", "Requires Assignment"], [2.5, 2.5, 1.2])):
                _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
                tbl.rows[0].cells[i].width = Inches(w)
                p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
                run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
            for a in risky_apps[:20]:
                row = tbl.add_row().cells
                req_assign = "Yes" if a.get("appRoleAssignmentRequired") else "No"
                for i, val in enumerate([a.get("displayName", ""), a.get("appId", ""), req_assign]):
                    p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                    _set_cell_bg(row[i], "FFF2CC")
            if len(risky_apps) > 20:
                row = tbl.add_row().cells
                p = row[0].paragraphs[0]; run = p.add_run(f"... and {len(risky_apps) - 20} more"); run.font.size = Pt(9); run.italic = True
            doc.add_paragraph()


_RECIPIENT_TYPE_LABELS: dict[str, str] = {
    "SharedMailbox":              "Shared Mailboxes",
    "GroupMailbox":               "Microsoft 365 Groups",
    "UserMailbox":                "User Mailboxes",
    "MailUniversalSecurityGroup": "Mail-Enabled Security Groups",
    "MailUniversalDistributionGroup": "Distribution Groups",
    "SchedulingMailbox":          "Scheduling Mailboxes",
}


def _messaging_data_tables(doc: Document, snapshot: dict) -> None:
    from collections import Counter
    ex      = snapshot.get("Data", {}).get("Exchange", {})
    domains = snapshot.get("Data", {}).get("Tenant", {}).get("Domains", {}) or {}

    # --- Domain Configuration Table ---
    if domains:
        _tbl_label(doc, "Accepted Domain Inventory")
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(["Domain", "Default", "Verified", "Type"], [2.8, 0.7, 0.7, 1.2])):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        for dname, dv in sorted(domains.items(), key=lambda x: (not x[1].get("isDefault"), x[0])):
            is_verified = dv.get("isVerified", False)
            is_default  = dv.get("isDefault", False)
            auth_type   = str(dv.get("authenticationType") or "Managed").title()
            row = tbl.add_row().cells
            bg = "FFD9D9" if not is_verified and not dv.get("isInitial") else None
            for i, val in enumerate([dname, "Yes" if is_default else "", "Yes" if is_verified else "No", auth_type]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                if not is_verified and i == 2:
                    run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00); run.bold = True
                if bg:
                    _set_cell_bg(row[i], bg)
        doc.add_paragraph()

    # --- Email Authentication Records (SPF/DKIM/DMARC per domain) ---
    dns_auth = snapshot.get("Data", {}).get("Governance", {}).get("DomainAuthenticationRecords") or {}
    if dns_auth:
        _tbl_label(doc, "Email Authentication Records (SPF / DKIM / DMARC)")
        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(
            ["Domain", "SPF", "DKIM", "DMARC", "DMARC Policy"],
            [2.2, 0.8, 0.8, 0.8, 1.2],
        )):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)

        _AUTH_OK  = ("Configured", "Yes")
        for dname, rec in sorted(dns_auth.items()):
            spf_val   = rec.get("SPF",   "Not found")
            dkim_val  = rec.get("DKIM",  "Not found")
            dmarc_val = rec.get("DMARC", "Not found")
            policy    = rec.get("DMARCPolicy", "")
            row = tbl.add_row().cells
            for i, val in enumerate([dname, spf_val, dkim_val, dmarc_val, policy]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                if i in (1, 2, 3) and val not in _AUTH_OK:
                    _set_cell_bg(row[i], "FFD9D9")
                    run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00); run.bold = True
        doc.add_paragraph()

    # --- Group Type Summary (all group types from Identity + Collaboration) ---
    all_groups = snapshot.get("Data", {}).get("Identity", {}).get("EntraIDGroups", {}) or {}
    all_teams_ids = set((snapshot.get("Data", {}).get("Collaboration", {}).get("AllTeams") or {}).keys())

    if all_groups:
        _teams_backed  = 0
        _m365_no_teams = 0
        _security      = 0
        _mail_sec      = 0
        _distribution  = 0
        _dynamic       = 0
        _with_license  = 0

        for gid, g in all_groups.items():
            # Support PS snapshot (GroupType string) and Python collector (groupTypes array)
            gtype_str  = g.get("GroupType") or g.get("groupType") or ""
            gtypes_arr = g.get("groupTypes") or []
            mem_type   = g.get("MembershipType") or g.get("membershipType") or ""

            is_unified    = "Microsoft 365" in gtype_str or "Unified" in gtypes_arr
            is_dist       = "Distribution" in gtype_str or (
                               not gtype_str and (g.get("mailEnabled") or g.get("MailEnabled"))
                               and not (g.get("securityEnabled") or g.get("SecurityEnabled"))
                               and not is_unified)
            is_mail_sec_g = "Mail-enabled Security" in gtype_str or (
                               not gtype_str and (g.get("mailEnabled") or g.get("MailEnabled"))
                               and (g.get("securityEnabled") or g.get("SecurityEnabled"))
                               and not is_unified)
            is_security_g = "Security Group" == gtype_str or (
                               not gtype_str and not (g.get("mailEnabled") or g.get("MailEnabled"))
                               and (g.get("securityEnabled") or g.get("SecurityEnabled"))
                               and not is_unified)
            is_dynamic    = "Dynamic" in mem_type or "DynamicMembership" in gtypes_arr
            has_license   = bool(g.get("IsManagingLicenses") or g.get("isManagingLicenses")
                                 or (g.get("AssignedLicenseCount") or g.get("assignedLicenseCount") or 0) > 0
                                 or bool(g.get("assignedLicenses")))

            if is_unified and gid in all_teams_ids:
                _teams_backed += 1
            elif is_unified:
                _m365_no_teams += 1
            elif is_mail_sec_g:
                _mail_sec += 1
            elif is_dist:
                _distribution += 1
            elif is_security_g:
                _security += 1

            if is_dynamic:
                _dynamic += 1
            if has_license:
                _with_license += 1

        _tbl_label(doc, f"Group Type Summary ({len(all_groups)} total groups)")
        _table_2col(doc, [
            ("Teams-backed M365 Groups",      str(_teams_backed)),
            ("M365 Groups (without Teams)",   str(_m365_no_teams)),
            ("Security Groups",               str(_security)),
            ("Mail-Enabled Security Groups",  str(_mail_sec)),
            ("Distribution Lists",            str(_distribution)),
            ("Dynamic Membership Groups",     str(_dynamic)),
            ("Groups with License Assignment", str(_with_license)),
        ], hdr=("Group Type", "Count"), col_widths=(3.5, 1.0))
        doc.add_paragraph()

    # --- Mailbox Type Inventory (Graph: MailboxSettings.Read + Place.Read.All) ---
    shared_ids   = ex.get("SharedMailboxes")    # None = not collected; [] = none found
    equip_ids    = ex.get("EquipmentMailboxes")
    room_mailboxes = ex.get("RoomMailboxes") or []
    if shared_ids is not None or equip_ids is not None or room_mailboxes:
        mbx_summary_rows: list[tuple[str, str]] = []
        if shared_ids is not None:
            mbx_summary_rows.append(("Shared mailboxes detected", str(len(shared_ids))))
        if equip_ids is not None:
            mbx_summary_rows.append(("Equipment mailboxes detected", str(len(equip_ids))))
        if room_mailboxes:
            mbx_summary_rows.append(("Room mailboxes (Places API)", str(len(room_mailboxes))))
        if mbx_summary_rows:
            _tbl_label(doc, "Mailbox Type Summary")
            _table_2col(doc, mbx_summary_rows, hdr=("Mailbox Type", "Count"), col_widths=(3.0, 1.5))

        # Room mailboxes detail (Places API)
        if room_mailboxes:
            _tbl_label(doc, f"Room Mailboxes ({len(room_mailboxes)})")
            tbl = doc.add_table(rows=1, cols=4)
            tbl.style = "Table Grid"
            for i, (label, w) in enumerate(zip(
                ["Display Name", "Email Address", "Building", "Capacity"],
                [2.2, 2.2, 1.0, 0.7],
            )):
                _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
                tbl.rows[0].cells[i].width = Inches(w)
                p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
                run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
            for rm in sorted(room_mailboxes, key=lambda x: (x.get("displayName") or "").lower()):
                row = tbl.add_row().cells
                for i, val in enumerate([
                    rm.get("displayName", ""),
                    rm.get("emailAddress", ""),
                    rm.get("building", ""),
                    str(rm.get("capacity") or ""),
                ]):
                    p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
            doc.add_paragraph()

    # --- Inbox Rules with External Forwarding (Graph: MailboxSettings.Read) ---
    fwd_rules = ex.get("InboxRulesExternalForwarding")  # None = not collected; [] = none found
    if isinstance(fwd_rules, list) and fwd_rules:
        external_rules = [r for r in fwd_rules if r.get("ExternalAddresses")]
        all_fwd = fwd_rules  # show all forwarding rules; highlight external ones
        _tbl_label(doc, f"Inbox Rules with Forwarding ({len(all_fwd)} rule(s), "
                        f"{len(external_rules)} with external targets)")
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(
            ["User", "Rule Name", "Forwarding Targets", "External?"],
            [1.8, 1.4, 2.4, 0.8],
        )):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        identity_users = snapshot.get("Data", {}).get("Identity", {}).get("Users", {}) or {}
        for rule in sorted(all_fwd, key=lambda x: (not x.get("ExternalAddresses"), x.get("UserId", ""))):
            uid      = rule.get("UserId", "")
            u        = identity_users.get(uid, {}) or {}
            user_lbl = u.get("userPrincipalName") or uid
            targets  = rule.get("ForwardToAddresses", []) + rule.get("RedirectToAddresses", [])
            ext      = rule.get("ExternalAddresses") or []
            is_ext   = bool(ext)
            bg       = "FFD9D9" if is_ext else None
            row = tbl.add_row().cells
            for i, val in enumerate([
                user_lbl,
                rule.get("RuleName", ""),
                "; ".join(targets[:3]) + (" ..." if len(targets) > 3 else ""),
                "Yes" if is_ext else "No",
            ]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                if bg:
                    _set_cell_bg(row[i], bg)
                if i == 3 and is_ext:
                    run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00); run.bold = True
        doc.add_paragraph()

    # --- Top Email Senders (Graph reporting data — always available) ---
    senders_raw = ex.get("EmailActivityTopSenders", []) or []
    if isinstance(senders_raw, list) and senders_raw:
        period = senders_raw[0].get("Report Period", "30") if senders_raw else "30"
        active = [s for s in senders_raw if int(s.get("Send Count", 0) or 0) > 0]
        active_sorted = sorted(active, key=lambda s: -int(s.get("Send Count", 0) or 0))
        if active_sorted:
            _tbl_label(doc, f"Top Email Senders (last {period} days)")
            tbl = doc.add_table(rows=1, cols=4)
            tbl.style = "Table Grid"
            for i, (label, w) in enumerate(zip(["Name", "UPN", "Sent", "Received"], [1.8, 2.2, 0.8, 0.9])):
                _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
                tbl.rows[0].cells[i].width = Inches(w)
                p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
                run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
            for s in active_sorted[:10]:
                row = tbl.add_row().cells
                for i, val in enumerate([
                    s.get("Display Name", ""),
                    s.get("User Principal Name", ""),
                    str(s.get("Send Count", 0) or 0),
                    str(s.get("Receive Count", 0) or 0),
                ]):
                    p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
            doc.add_paragraph()

    # --- AllRecipients-based tables (PS collection only) ---
    ar_raw = ex.get("AllRecipients", {})
    if isinstance(ar_raw, dict):
        recipients = list(ar_raw.values())
    elif isinstance(ar_raw, list):
        recipients = ar_raw
    else:
        recipients = []

    if not recipients:
        return

    by_type   = Counter(r.get("RecipientTypeDetails", "") for r in recipients)
    by_domain: Counter = Counter()
    for r in recipients:
        smtp = r.get("PrimarySmtpAddress", "")
        if "@" in smtp:
            by_domain[smtp.split("@")[1].lower()] += 1

    _tbl_label(doc, "Exchange Recipient Breakdown")
    type_rows: list[tuple[str, str]] = []
    for rtype, cnt in sorted(by_type.items(), key=lambda x: -x[1]):
        label = _RECIPIENT_TYPE_LABELS.get(rtype, rtype) if rtype else "Unknown"
        type_rows.append((label, str(cnt)))
    type_rows.append(("Total", str(len(recipients))))
    _table_2col(doc, type_rows, hdr=("Recipient Type", "Count"), col_widths=(3.0, 1.5))

    if len(by_domain) > 1:
        _tbl_label(doc, "Recipients by Primary SMTP Domain")
        domain_rows: list[tuple[str, str]] = [
            (domain, str(cnt))
            for domain, cnt in sorted(by_domain.items(), key=lambda x: -x[1])[:10]
        ]
        domain_rows.append(("Total", str(len(recipients))))
        _table_2col(doc, domain_rows, hdr=("Domain", "Count"), col_widths=(3.0, 1.5))


def _collab_data_tables(doc: Document, snapshot: dict) -> None:
    collab   = snapshot.get("Data", {}).get("Collaboration", {})
    activity = collab.get("CollaborationActivitySummary", {}) or {}
    teams_raw = collab.get("TeamsGroupsCleanupCandidates", {})
    teams_list = list(teams_raw.values()) if isinstance(teams_raw, dict) else (teams_raw if isinstance(teams_raw, list) else [])

    # --- SharePoint Tenant Overview ---
    sp_settings = collab.get("SharePoint", {}) or {}
    sp_summary  = collab.get("SharePointSharingSummary", {}) or {}
    od_summary  = collab.get("OneDrive", {}) or {}
    _SHARING_LABELS = {
        "disabled":                         "Disabled (no external sharing)",
        "externalUserSharingOnly":          "Authenticated external users only",
        "externalUserAndGuestSharing":      "External users and anonymous guests",
        "existingExternalUserSharingOnly":  "Existing external users only",
    }
    sharing_cap = sp_settings.get("SharingCapability", "")
    if sp_settings or sp_summary:
        rows: list[tuple[str, str]] = []
        if sharing_cap:
            sharing_label = _SHARING_LABELS.get(sharing_cap, sharing_cap)
            rows.append(("Tenant external sharing level", sharing_label))
        anon_access = sp_settings.get("AllowAnonymousAccess")
        if anon_access is not None:
            rows.append(("Anonymous link access", "Enabled" if anon_access else "Disabled"))
        if sp_summary.get("TotalSites"):
            rows.append(("Total SharePoint sites", str(sp_summary["TotalSites"])))
            rows.append(("Sites with external sharing configured", str(sp_summary.get("SitesWithExternalSharing", 0))))
            rows.append(("Total SharePoint storage (GB)", f"{sp_summary.get('TotalStorageGB', 0):.2f}"))
        if od_summary.get("TotalAccounts"):
            rows.append(("OneDrive accounts provisioned", str(od_summary["TotalAccounts"])))
            rows.append(("OneDrive total storage (GB)", f"{od_summary.get('TotalStorageGB', 0):.2f}"))
        if rows:
            _tbl_label(doc, "SharePoint and OneDrive Tenant Overview")
            _table_2col(doc, rows, hdr=("Setting", "Current State"), col_widths=(3.0, 2.5))
            if sharing_cap in ("externalUserAndGuestSharing", "externalUserSharingOnly"):
                doc.add_paragraph(
                    "The tenant sharing level permits external sharing. "
                    "Confirm that site-level sharing restrictions are configured to limit exposure "
                    "to only those sites that require external collaboration."
                )
            doc.add_paragraph()

    # Group activity summary — use TotalUnifiedGroups as the key (actual snapshot field)
    total = activity.get("TotalUnifiedGroups") or activity.get("TotalGroups") or 0
    active_grps  = activity.get("GroupsWithActivityInPeriod") or activity.get("ActiveGroups") or 0
    active_teams = activity.get("ActiveTeamsUsersInPeriod") or 0
    if total or active_grps or active_teams:
        inactive     = max(0, total - active_grps)
        inactive_pct = f"{round(inactive / total * 100)}%" if total else "0%"
        active_pct   = f"{round(active_grps / total * 100)}%" if total else "0%"
        act_rows: list[tuple[str, str]] = [
            ("Total unified groups and Teams",     str(total)),
            ("Groups with activity in period",     f"{active_grps} ({active_pct})"),
            ("Groups with no activity in period",  f"{inactive} ({inactive_pct})"),
            ("Active Teams users in period",       str(active_teams)),
        ]
        _tbl_label(doc, "Collaboration Activity Summary")
        _table_2col(doc, act_rows, hdr=("Metric", "Value"), col_widths=(3.0, 1.5))
        rows = act_rows  # keep name for compat

    # Cleanup candidates detail
    if teams_list:
        ownerless  = [t for t in teams_list if "Ownerless"   in str(t.get("RiskSignal", ""))]
        no_members = [t for t in teams_list if "No members"  in str(t.get("RiskSignal", ""))]
        dormant    = [t for t in teams_list if "Dormant"     in str(t.get("RiskSignal", ""))]
        both       = [t for t in teams_list if "Ownerless"   in str(t.get("RiskSignal", "")) and "Dormant" in str(t.get("RiskSignal", ""))]
        cand_rows: list[tuple[str, str]] = [
            ("Total cleanup candidates",           str(len(teams_list))),
            ("Ownerless",                          str(len(ownerless))),
            ("No members",                         str(len(no_members))),
            ("Dormant (no activity 90+ days)",     str(len(dormant))),
            ("Ownerless and dormant",              str(len(both))),
            ("Teams (vs groups only)",             str(sum(1 for t in teams_list if t.get("IsTeam") or t.get("ObjectType") == "Team"))),
        ]
        _tbl_label(doc, "Cleanup Candidate Breakdown")
        _table_2col(doc, cand_rows, hdr=("Category", "Count"), col_widths=(3.0, 1.5))

    # Named teams/groups requiring review
    if teams_list:
        sorted_teams = sorted(teams_list, key=lambda x: ("Ownerless" not in str(x.get("RiskSignal","")), x.get("Name","")))
        shown = sorted_teams[:15]
        overflow = len(sorted_teams) - len(shown)
        _tbl_label(doc, f"Teams and Groups Requiring Review ({len(sorted_teams)} total)")
        tbl = doc.add_table(rows=1, cols=6)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(
            ["Name", "Risk Signal", "Last Activity", "Type", "Owners", "Members"],
            [2.0, 1.6, 1.0, 0.6, 0.6, 0.6],
        )):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        for t in shown:
            signal   = str(t.get("RiskSignal", ""))
            bg       = "FFD9D9" if "Ownerless" in signal else "FFE9CC"
            last     = (t.get("LastActivityDate") or "")[:10] or "No record"
            obj_type = "Team" if (t.get("IsTeam") or t.get("ObjectType") == "Team") else "Group"
            mc       = t.get("MemberCount")
            members  = str(mc) if mc is not None else "-"
            row = tbl.add_row().cells
            for i, val in enumerate([t.get("Name", t.get("DisplayName", "")), signal, last, obj_type, str(t.get("OwnerCount", 0)), members]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                _set_cell_bg(row[i], bg)
        if overflow > 0:
            row = tbl.add_row().cells
            p = row[0].paragraphs[0]
            run = p.add_run(f"... and {overflow} more -- see TeamsGroupsCleanupCandidates in Excel workbook")
            run.font.size = Pt(9); run.italic = True; run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
        doc.add_paragraph()

    # --- Teams User Activity ---
    teams_activity = collab.get("TeamsActivityTopUsers", []) or []
    if isinstance(teams_activity, dict):
        teams_activity = list(teams_activity.values())
    if teams_activity:
        def _ta_total(u: dict) -> int:
            return (int(u.get("Team Chat Message Count", 0) or 0)
                    + int(u.get("Private Chat Message Count", 0) or 0)
                    + int(u.get("Call Count", 0) or 0)
                    + int(u.get("Meeting Count", 0) or 0))
        active_ta = [u for u in teams_activity if _ta_total(u) > 0]
        inactive_ta = [u for u in teams_activity if _ta_total(u) == 0]
        active_ta_sorted = sorted(active_ta, key=lambda u: -_ta_total(u))
        period = (teams_activity[0].get("Report Period", "30") if teams_activity else "30")
        _tbl_label(doc, f"Teams User Activity - Top Active Users (last {period} days)")
        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(
            ["User", "Team Chat", "Private Chat", "Calls", "Meetings"], [2.4, 0.8, 0.9, 0.7, 0.8]
        )):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        for u in active_ta_sorted[:20]:
            row = tbl.add_row().cells
            upn = u.get("User Principal Name", "")
            for i, val in enumerate([
                upn,
                str(u.get("Team Chat Message Count", 0) or 0),
                str(u.get("Private Chat Message Count", 0) or 0),
                str(u.get("Call Count", 0) or 0),
                str(u.get("Meeting Count", 0) or 0),
            ]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
        if inactive_ta:
            row = tbl.add_row().cells
            p = row[0].paragraphs[0]
            run = p.add_run(f"{len(inactive_ta)} user(s) with no Teams activity in period")
            run.font.size = Pt(9); run.italic = True; run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
            _set_cell_bg(row[0], "F2F2F2")
        doc.add_paragraph()

    # --- Teams with Private Channels ---
    private_ch = collab.get("PrivateChannels", {}) or {}
    teams_all  = collab.get("AllTeams", {}) or {}
    if private_ch:
        teams_with_private = [(tid, chs) for tid, chs in private_ch.items() if chs]
        if teams_with_private:
            teams_with_private.sort(key=lambda x: -len(x[1]))
            _tbl_label(doc, f"Teams with Private Channels ({len(teams_with_private)})")
            tbl = doc.add_table(rows=1, cols=2)
            tbl.style = "Table Grid"
            for i, (label, w) in enumerate(zip(["Team", "Private Channel Count"], [4.0, 1.5])):
                _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
                tbl.rows[0].cells[i].width = Inches(w)
                p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
                run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
            for tid, channels in teams_with_private:
                _td = teams_all.get(tid, {}) or {}
                team_name = _td.get("DisplayName") or _td.get("displayName") or tid
                row = tbl.add_row().cells
                for i, val in enumerate([team_name, str(len(channels))]):
                    p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
            doc.add_paragraph()

    # --- Sensitivity Labels on M365 Groups ---
    unified_groups = collab.get("UnifiedGroups", {}) or {}
    labeled = [
        g for g in unified_groups.values()
        if g.get("assignedLabels") or g.get("AssignedLabels")
    ]
    if labeled:
        labeled.sort(key=lambda x: (x.get("DisplayName") or x.get("displayName") or "").lower())
        _tbl_label(doc, f"M365 Groups with Sensitivity Labels ({len(labeled)})")
        tbl = doc.add_table(rows=1, cols=2)
        tbl.style = "Table Grid"
        for i, (label_hdr, w) in enumerate(zip(["Group", "Sensitivity Label(s)"], [3.0, 2.5])):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label_hdr); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        for g in labeled:
            raw_labels = g.get("assignedLabels") or g.get("AssignedLabels") or []
            if isinstance(raw_labels, list):
                labels = "; ".join(
                    (lbl.get("displayName") or lbl.get("DisplayName") or lbl.get("labelId") or "")
                    for lbl in raw_labels if isinstance(lbl, dict)
                )
            else:
                labels = str(raw_labels)
            row = tbl.add_row().cells
            g_name = g.get("DisplayName") or g.get("displayName") or ""
            for i, val in enumerate([g_name, labels]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
        doc.add_paragraph()


def _endpoint_data_tables(doc: Document, snapshot: dict) -> None:
    from collections import Counter
    devices_raw = snapshot.get("Data", {}).get("Identity", {}).get("DeviceDetails", {}) or {}
    devices = list(devices_raw.values()) if isinstance(devices_raw, dict) else (devices_raw if isinstance(devices_raw, list) else [])
    if not devices:
        return

    def _dev(d: dict, key: str, pascal: str | None = None):
        return d.get(pascal or key.title().replace(" ", ""), d.get(key))

    # Summary table
    os_counts: Counter = Counter(
        d.get("OperatingSystem") or d.get("operatingSystem") or "Unknown" for d in devices
    )
    managed   = sum(1 for d in devices if d.get("IsManaged") or d.get("isManaged"))
    unmanaged = len(devices) - managed
    compliant = sum(1 for d in devices if d.get("IsCompliant") or d.get("isCompliant"))
    non_comp  = sum(1 for d in devices if (d.get("IsCompliant") if "IsCompliant" in d else d.get("isCompliant")) is False)
    unknown_c = len(devices) - compliant - non_comp

    summary_rows: list[tuple[str, str]] = [("Total registered devices", str(len(devices)))]
    for os_name, cnt in sorted(os_counts.items(), key=lambda x: -x[1]):
        summary_rows.append((f"  {os_name}", str(cnt)))
    summary_rows += [
        ("Managed (Intune / hybrid-joined)", str(managed)),
        ("Unmanaged",                        str(unmanaged)),
        ("Compliant",                        str(compliant)),
        ("Non-compliant",                    str(non_comp)),
        ("Compliance state unknown",         str(unknown_c)),
    ]
    _tbl_label(doc, "Device Inventory Summary")
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Table Grid"
    for i, (label, w) in enumerate(zip(["Category", "Count"], [3.5, 1.0])):
        _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
        tbl.rows[0].cells[i].width = Inches(w)
        p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
        run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
    for label, val in summary_rows:
        row = tbl.add_row().cells
        bg = "FFD9D9" if (label == "Unmanaged" and unmanaged > 0) or (label == "Non-compliant" and non_comp > 0) else None
        for i, v in enumerate([label, val]):
            p = row[i].paragraphs[0]; run = p.add_run(v); run.font.size = Pt(9)
            if bg:
                _set_cell_bg(row[i], bg)
    doc.add_paragraph()

    # Per-device table (non-compliant / unmanaged first)
    def _d_managed(d: dict):
        v = d.get("IsManaged")
        return v if v is not None else d.get("isManaged")

    def _d_compliant(d: dict):
        v = d.get("IsCompliant")
        return v if v is not None else d.get("isCompliant")

    def _mgmt_sort(d: dict) -> tuple:
        m = _d_managed(d)
        c = _d_compliant(d)
        return (0 if m is False else (1 if m is None else 2), 0 if c is False else 1,
                str(d.get("DisplayName") or d.get("displayName") or ""))

    flagged = sorted(
        [d for d in devices if _d_managed(d) is not True or _d_compliant(d) is False],
        key=_mgmt_sort,
    )
    if flagged:
        shown_devices = flagged[:10]
        device_overflow = len(flagged) - len(shown_devices)
        _tbl_label(doc, f"Devices Requiring Attention ({len(shown_devices)} of {len(flagged)} shown)")
        tbl2 = doc.add_table(rows=1, cols=5)
        tbl2.style = "Table Grid"
        for i, (label, w) in enumerate(zip(["Device Name", "OS", "Managed", "Compliant", "Last Sign-in"], [2.0, 0.85, 0.7, 0.7, 0.9])):
            _set_cell_bg(tbl2.rows[0].cells[i], "1F3864")
            tbl2.rows[0].cells[i].width = Inches(w)
            p = tbl2.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        for d in shown_devices:
            is_managed   = _d_managed(d)
            is_compliant = _d_compliant(d)
            last_in      = (d.get("ApproximateLastSignInDateTime") or d.get("approximateLastSignInDateTime") or "")[:10]
            row = tbl2.add_row().cells
            bg = "FFD9D9" if (is_managed is False or is_compliant is False) else "FFE9CC"
            for i, val in enumerate([
                d.get("DisplayName") or d.get("displayName") or "",
                d.get("OperatingSystem") or d.get("operatingSystem") or "",
                "Yes" if is_managed else ("No" if is_managed is False else "Unknown"),
                "Yes" if is_compliant else ("No" if is_compliant is False else "Unknown"),
                last_in or "No data",
            ]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                _set_cell_bg(row[i], bg)
        if device_overflow > 0:
            row = tbl2.add_row().cells
            p = row[0].paragraphs[0]
            run = p.add_run(f"... and {device_overflow} more -- see DeviceDetails in Excel workbook")
            run.font.size = Pt(9); run.italic = True; run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
        doc.add_paragraph()


def _security_data_tables(doc: Document, snapshot: dict) -> None:
    sec        = snapshot.get("Data", {}).get("Security", {})
    score_raw  = sec.get("SecuritySecureScore", {}) or {}
    controls   = sec.get("SecureScoreActions", []) or []
    if isinstance(controls, dict):
        controls = list(controls.values())

    # SecuritySecureScore may be a dict keyed by tenantId_date — unwrap if needed
    if score_raw and not any(k in score_raw for k in ("currentScore", "CurrentScore", "maxScore", "MaxScore")):
        score_data = next(iter(score_raw.values())) if score_raw else {}
    else:
        score_data = score_raw

    current = score_data.get("CurrentScore") or score_data.get("currentScore") or 0
    max_s   = score_data.get("MaxScore")     or score_data.get("maxScore")     or 0
    if current and max_s:
        pct  = f"{round(current / max_s * 100)}%"
        comp = score_data.get("AverageComparativeScores") or score_data.get("averageComparativeScores") or []
        avg_all = next((s.get("AverageScore") or s.get("averageScore") for s in comp
                        if s.get("Basis") == "AllTenants" or s.get("basis") == "AllTenants"), None)
        rows: list[tuple[str, str]] = [
            ("Current Secure Score", f"{current:.0f} / {max_s:.0f} ({pct})"),
        ]
        if avg_all:
            rows.append(("Cross-tenant average (all tenants)", f"{avg_all:.0f}"))
            delta = current - avg_all
            rows.append(("Vs. average", f"{delta:+.0f} points"))
        _tbl_label(doc, "Microsoft Secure Score")
        _table_2col(doc, rows, hdr=("Metric", "Value"), col_widths=(3.0, 2.0))

    # Top unimplemented controls
    def _ctrl_score(c: dict) -> float:
        return float(c.get("MaxScore") or c.get("maxScore") or c.get("ScoreGap") or 0)

    def _ctrl_status(c: dict) -> str:
        return c.get("Status") or c.get("implementationStatus") or ""

    def _ctrl_title(c: dict) -> str:
        return c.get("RecommendationTitle") or c.get("title") or c.get("ControlId") or ""

    not_impl = [c for c in controls if isinstance(c, dict) and _ctrl_status(c) in ("notImplemented", "Not Implemented", "Default")]
    top5 = sorted(not_impl, key=lambda c: -_ctrl_score(c))[:5]
    if top5:
        ctrl_rows: list[tuple[str, str]] = [
            (_ctrl_title(c), f"{_ctrl_score(c):.0f} pts") for c in top5
        ]
        _tbl_label(doc, "Top Unimplemented Controls (by point value)")
        _table_2col(doc, ctrl_rows, hdr=("Control", "Max Points"), col_widths=(3.5, 1.0))


def _governance_data_tables(doc: Document, snapshot: dict) -> None:
    data = snapshot.get("Data", {})

    # --- Group-Based Licensing ---
    gls = data.get("Identity", {}).get("GroupLicensingSummary", []) or []
    if isinstance(gls, list) and gls:
        _tbl_label(doc, "Group-Based Licensing Groups")
        gls_rows: list[tuple] = [
            (
                g.get("GroupName", ""),
                g.get("AssignedLicenseFriendlyNames") or g.get("LicenseSKUs") or g.get("AssignedLicenseSkuPartNumbers") or "",
                str(g.get("OwnerCount", 0)),
                "Yes" if (g.get("IsOwnerless") or g.get("OwnerCount", 1) == 0) else "No",
            )
            for g in gls
        ]
        tbl_g = doc.add_table(rows=1, cols=4)
        tbl_g.style = "Table Grid"
        for i, (lbl, w) in enumerate(zip(["Group Name", "License SKU(s)", "Owners", "Ownerless"], [2.5, 2.5, 0.6, 0.7])):
            _set_cell_bg(tbl_g.rows[0].cells[i], "1F3864")
            tbl_g.rows[0].cells[i].width = Inches(w)
            p = tbl_g.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(lbl); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        for g_name, g_skus, g_owners, g_ownerless in gls_rows:
            row = tbl_g.add_row().cells
            for i, val in enumerate([g_name, g_skus, g_owners, g_ownerless]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                if g_ownerless == "Yes":
                    _set_cell_bg(row[i], "FFF2CC")
        doc.add_paragraph()

    # --- License SKU Utilization ---
    lic_raw = data.get("Identity", {}).get("LicenseSKUs", {}) or {}
    skus = list(lic_raw.values()) if isinstance(lic_raw, dict) else lic_raw
    if skus:
        _tbl_label(doc, "License SKU Utilization")
        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Table Grid"
        for i, (label, w) in enumerate(zip(["License SKU", "SKU Part Number", "Consumed", "Available", "Status"], [2.0, 1.6, 0.7, 0.7, 0.8])):
            _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
            tbl.rows[0].cells[i].width = Inches(w)
            p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
            run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
        for sku in sorted(skus, key=lambda s: s.get("SkuPartNumber") or s.get("skuPartNumber") or ""):
            consumed  = sku.get("ConsumedUnits") or sku.get("consumedUnits") or 0
            prepaid   = (sku.get("PurchasedUnits") or sku.get("prepaidUnits") if not isinstance(sku.get("prepaidUnits"), dict)
                         else (sku.get("prepaidUnits") or {}).get("enabled") or 0) or 0
            available = sku.get("RemainingUnits") if sku.get("RemainingUnits") is not None else max(0, prepaid - consumed)
            pct       = round(consumed / prepaid * 100) if prepaid else 0
            status    = "At Capacity" if prepaid > 0 and consumed >= prepaid else f"{pct}% used"
            bg        = "FFD9D9" if prepaid > 0 and consumed >= prepaid else ("FFFCE8" if pct > 80 else None)
            row = tbl.add_row().cells
            part_num  = sku.get("SkuPartNumber") or sku.get("skuPartNumber") or ""
            name      = sku.get("SkuFriendlyName") or part_num
            for i, val in enumerate([name, part_num, str(consumed), str(available), status]):
                p = row[i].paragraphs[0]; run = p.add_run(val); run.font.size = Pt(9)
                if bg:
                    _set_cell_bg(row[i], bg)
        doc.add_paragraph()

    # --- License Optimization Candidates ---
    loc = data.get("Identity", {}).get("LicenseOptimizationCandidates", []) or []
    if isinstance(loc, dict):
        loc = list(loc.values())
    if loc:
        # Group by issue type — field is "Issue" (PS) or "Signal" (Python)
        from collections import Counter as _Counter
        issue_counts: dict[str, int] = {}
        for c in loc:
            issue = c.get("Issue") or c.get("Signal") or "Other"
            issue_counts[issue] = issue_counts.get(issue, 0) + 1
        opt_rows: list[tuple[str, str]] = [(iss, str(cnt)) for iss, cnt in sorted(issue_counts.items(), key=lambda x: -x[1])]
        _tbl_label(doc, f"License Optimization Candidates ({len(loc)} total)")
        _table_2col(doc, opt_rows, hdr=("Issue Type", "Count"), col_widths=(4.0, 0.8))
        p = doc.add_paragraph()
        run = p.add_run("Full candidate details are available in the LicenseOptimizationCandidates tab of the Excel workbook.")
        run.italic = True; run.font.size = Pt(9); run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
        doc.add_paragraph()

    # --- Password Lifecycle Configuration ---
    pwd = data.get("Governance", {}).get("PasswordLifecycleSummary", {}) or {}
    if not pwd:
        return

    validity      = pwd.get("PasswordValidityPeriodInDays")
    never_expires = pwd.get("PasswordNeverExpires") or validity == 2147483647
    sspr          = pwd.get("SelfServicePasswordResetEnabled")
    notif_days    = pwd.get("PasswordNotificationWindowInDays")

    bool_str = {True: "Yes", False: "No", None: "Unknown"}
    rows: list[tuple[str, str]] = [
        ("Password expiry policy",        "Never expires" if never_expires else (f"{validity} days" if validity else "Unknown")),
        ("Notification window",           f"{notif_days} days before expiry" if notif_days and not never_expires else "N/A"),
        ("Self-service password reset",   bool_str.get(sspr, "Unknown")),
    ]
    _tbl_label(doc, "Password Lifecycle Configuration")
    _table_2col(doc, rows, hdr=("Setting", "Value"), col_widths=(3.0, 2.0))


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
    doc.add_page_break()


# ---------------------------------------------------------------------------
# 13.0 Offboarding Process Recommendation
# ---------------------------------------------------------------------------

def _offboarding_lifecycle_table(doc: Document, snapshot: dict) -> None:
    """Data-driven lifecycle signals table sourced from the snapshot."""
    from datetime import datetime as _dt, timezone as _tz

    data    = snapshot.get("Data", {})
    identity = data.get("Identity", {})
    collab   = data.get("Collaboration", {})
    _now     = _dt.now(_tz.utc)

    # Stale privileged accounts (no sign-in 180+ days)
    priv_remediation = identity.get("PrivilegedAccessRemediationSummary") or []
    stale_privs = next(
        (int(r.get("Count", 0)) for r in priv_remediation if "Stale" in str(r.get("Signal", ""))),
        None,
    )

    # Guests
    guest_sum = identity.get("GuestSignInSummary") or {}
    guest_count   = guest_sum.get("GuestCount", 0)
    enabled_guests = guest_sum.get("EnabledGuestCount", 0)

    # Cleanup candidates
    cleanup = collab.get("TeamsGroupsCleanupCandidates") or {}
    cleanup_list = list(cleanup.values()) if isinstance(cleanup, dict) else (cleanup if isinstance(cleanup, list) else [])
    ownerless_count = sum(1 for c in cleanup_list if "Ownerless" in str(c.get("RiskSignal", "")))
    dormant_count   = sum(1 for c in cleanup_list if "Dormant"   in str(c.get("RiskSignal", "")))
    no_members_count = sum(1 for c in cleanup_list if "No members" in str(c.get("RiskSignal", "")))

    # License optimization candidates
    loc = identity.get("LicenseOptimizationCandidates") or []
    disabled_licensed = sum(1 for c in loc if "Disabled" in str(c.get("Signal", "")))
    inactive_licensed = sum(1 for c in loc if "Disabled" not in str(c.get("Signal", "")))

    # Stale devices (no sign-in 90+ days)
    devices_raw = identity.get("DeviceDetails") or {}
    devices = list(devices_raw.values()) if isinstance(devices_raw, dict) else (devices_raw if isinstance(devices_raw, list) else [])
    stale_devices = 0
    for d in devices:
        last = d.get("approximateLastSignInDateTime") or ""
        if not last:
            stale_devices += 1
            continue
        try:
            dt = _dt.fromisoformat(last.replace("Z", "+00:00"))
            if (_now - dt).days >= 90:
                stale_devices += 1
        except (ValueError, AttributeError):
            pass

    def _fmt(val, warn_threshold: int = 1) -> tuple[str, bool]:
        if val is None:
            return "Not available", False
        return str(val), int(val) >= warn_threshold

    rows_raw = [
        ("Stale privileged accounts (no sign-in 180+ days)", _fmt(stale_privs, 1)),
        ("Total guest accounts",                              _fmt(guest_count, 0)),
        ("Enabled guest accounts",                           _fmt(enabled_guests, 1)),
        ("Disabled users with assigned licenses",            _fmt(disabled_licensed, 1)),
        ("Inactive licensed users (no sign-in 90+ days)",    _fmt(inactive_licensed, 1)),
        ("Ownerless Teams or groups",                        _fmt(ownerless_count, 1)),
        ("No-member Teams or groups",                        _fmt(no_members_count, 1)),
        ("Dormant Teams or groups (no activity 90+ days)",   _fmt(dormant_count, 1)),
        ("Devices with no sign-in in 90+ days",              _fmt(stale_devices, 1)),
    ]

    _tbl_label(doc, "Supporting Observations from Environment Review")
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Table Grid"
    for i, (label, w) in enumerate(zip(["Lifecycle Signal", "Current State"], [3.5, 2.0])):
        _set_cell_bg(tbl.rows[0].cells[i], "1F3864")
        tbl.rows[0].cells[i].width = Inches(w)
        p = tbl.rows[0].cells[i].paragraphs[0]; p.clear()
        run = p.add_run(label); run.bold = True; run.font.color.rgb = _WHITE; run.font.size = Pt(9)
    for signal, (state, is_flagged) in rows_raw:
        row = tbl.add_row().cells
        for i, val in enumerate([signal, state]):
            p = row[i].paragraphs[0]
            run = p.add_run(val); run.font.size = Pt(9)
            if is_flagged:
                _set_cell_bg(row[i], "FFE9CC")
                if i == 1:
                    run.font.color.rgb = _ORANGE; run.bold = True
    doc.add_paragraph()


def _write_offboarding(doc: Document, tenant_name: str, snapshot: dict | None = None) -> None:
    _h(doc, "13.0 Offboarding Process Recommendation", 1)
    doc.add_paragraph(
        "A structured offboarding process reduces the risk of data exposure, "
        "unauthorized access, and license waste after an employee or contractor "
        "leaves the organization. The checklist below reflects Microsoft 365 "
        "best practice and should be adapted to your HR and IT workflow."
    )

    # --- Supporting Observations from Environment Review ---
    if snapshot:
        _offboarding_lifecycle_table(doc, snapshot)

    _h(doc, "13.1 Immediate Actions (Day of Departure)", 2)
    doc.add_paragraph(
        "These steps should be completed on the employee's last day or as soon as "
        "the departure is confirmed."
    )
    immediate_steps = [
        ("Disable the user account in Entra ID",
         "Prevents new sign-ins immediately. The account is preserved for mailbox and data access during transition."),
        ("Revoke all active sessions and tokens",
         "Use the 'Revoke sign-in sessions' action in Entra ID or via PowerShell to invalidate cached tokens."),
        ("Remove from privileged roles (Global Admin, etc.)",
         "Privileged role membership should be removed before or on the last day. Review other sensitive roles."),
        ("Reset account password",
         "Prevents re-authentication if the account is inadvertently re-enabled."),
        ("Remove from distribution groups and Teams channels as needed",
         "Prevents ongoing delivery to the user and removes presence from active channels."),
    ]
    for title, desc in immediate_steps:
        p = doc.add_paragraph(style="List Bullet")
        run_bold = p.add_run(title + ": ")
        run_bold.bold = True; run_bold.font.size = Pt(10)
        run_desc = p.add_run(desc)
        run_desc.font.size = Pt(10)
        p.paragraph_format.space_after = Pt(4)

    _h(doc, "13.2 Short-Term Actions (Within 30 Days)", 2)
    doc.add_paragraph(
        "These steps should be completed within 30 days to preserve data and clean up access."
    )
    short_term_steps = [
        ("Convert user mailbox to shared mailbox",
         "Allows the manager or delegate to access mail without consuming a licensed seat. "
         "Remove the license after conversion."),
        ("Transfer OneDrive files to the manager or a designated owner",
         "Set a OneDrive access delegation (manager access) so files are not lost when the account is eventually deleted."),
        ("Review and reassign owned SharePoint sites",
         "Ownerless sites become ungoverned. Assign a new primary administrator for each site the user owned."),
        ("Review app ownership and service accounts",
         "If the user owned enterprise applications or service principals, assign a replacement owner to avoid orphaned apps."),
        ("Remove assigned licenses",
         "Reclaim licenses to reduce spend. Shared mailboxes do not require a full M365 license for access by delegates."),
        ("Remove from Teams ownership",
         "Teams without owners are flagged as ownerless. Assign new owners or archive inactive teams."),
    ]
    for title, desc in short_term_steps:
        p = doc.add_paragraph(style="List Bullet")
        run_bold = p.add_run(title + ": ")
        run_bold.bold = True; run_bold.font.size = Pt(10)
        run_desc = p.add_run(desc)
        run_desc.font.size = Pt(10)
        p.paragraph_format.space_after = Pt(4)

    _h(doc, "13.3 Long-Term Cleanup (30-90 Days)", 2)
    doc.add_paragraph(
        "After the transition period, complete cleanup to prevent long-term access risk and license waste."
    )
    long_term_steps = [
        ("Delete the user account",
         "Deleting the account moves it to the Entra ID recycle bin for 30 days. "
         "After deletion, the UPN is released and the account is permanently removed after the retention period."),
        ("Verify shared mailbox is not licensed",
         "A shared mailbox does not require a license unless it exceeds 50 GB or uses advanced features. "
         "Audit shared mailboxes periodically for unnecessary licenses."),
        ("Archive or delete inactive Teams and groups",
         "Use the Microsoft 365 admin center or lifecycle policy to archive or delete groups with no active owners or activity."),
        ("Review MFA and authenticator app registrations",
         "Authenticator methods tied to a deleted account may linger in the tenant. "
         "Validate the authentication methods policy shows no orphaned registrations."),
    ]
    for title, desc in long_term_steps:
        p = doc.add_paragraph(style="List Bullet")
        run_bold = p.add_run(title + ": ")
        run_bold.bold = True; run_bold.font.size = Pt(10)
        run_desc = p.add_run(desc)
        run_desc.font.size = Pt(10)
        p.paragraph_format.space_after = Pt(4)

    _h(doc, "13.4 Recommended Automation", 2)
    doc.add_paragraph(
        "Manual offboarding is error-prone. Consider the following automation options to enforce a consistent process:"
    )
    automation_rows = [
        ("Lifecycle Workflows (Entra ID Governance)",
         "Automate offboarding tasks (disable account, send manager notification, remove groups) "
         "triggered by HR-driven attribute changes. Requires Entra ID Governance license (P2 or standalone)."),
        ("Microsoft Entra Access Reviews",
         "Schedule recurring reviews of privileged role membership and group access to surface "
         "accounts that should be offboarded but were not caught in time."),
        ("Power Automate HR integration",
         "Build a flow triggered by your HR system (e.g., Workday, BambooHR) to initiate the "
         "offboarding checklist automatically on the last day."),
        ("Group Expiration Policy",
         "Configure a Microsoft 365 Groups expiration policy so that groups and Teams without "
         "active owners are flagged for renewal or deletion automatically."),
    ]
    _table_2col(doc, automation_rows, hdr=("Solution", "Description"), col_widths=(1.8, 4.7))
