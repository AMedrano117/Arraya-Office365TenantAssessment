"""Deterministic rendering of the four operator deliverables.

Every document leads with the answer and the recommendations; detail follows.
Wording comes from item['Plain'] (see plain.py) and action numbers from
item['ActionNumber'], so all four documents say the same thing in the same words
and number each action the same way.
"""
from __future__ import annotations

from datetime import date
from io import BytesIO
import re
import zipfile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from .plain import (AREA_NAMES, DECISION_LABEL, EVIDENCE_LABEL, GROUP, GROUP_FILL, GROUP_MEANING, GROUP_ORDER,
                    STATEMENT_LABEL)
from .util import FIXED_TIME, ascii_clean, text as txt

EXEC = "Executive-Readiness-Summary.docx"
TECH = "Technical-Assessment-Review.docx"
PLAN = "Customer-Action-Plan.xlsx"
BRIEF = "Review-Meeting-Brief.docx"
DELIVERABLES = (EXEC, TECH, PLAN, BRIEF)
NAVY = "16324F"
GREY = "595959"
ANSWER_FILL = "EAF1FB"
NO_COMMITMENT = ("This document lists actions for the customer's responsible roles. It is not a statement of work and "
                 "does not commit Arraya to perform any action. Owners, dates and approvals are recorded only when "
                 "the customer supplies them.")
CONFIDENTIAL = ("Confidential: the list sheets name users, devices, sites and apps in your tenant. Share this workbook "
                "only with the people responsible for these actions.")
ACTION_HEADERS = ("#", "When", "Priority", "Action", "Owner", "What we found", "How to fix", "Affected items",
                  "Done when", "Status", "Ref")
TRACE_SHEET = "Review trace"
DETAIL_HEADERS = ("#", "Ref", "Action", "When", "Area", "Original report title", "Review result", "Evidence quality",
                  "Change from original report", "Microsoft guidance", "Evidence references")
# Column widths for the object list sheets; unlisted columns use 18.
DETAIL_WIDTH = {"Account": 34, "Name": 26, "Roles": 40, "Assignment": 20, "MFA methods": 30, "MFA strength": 22,
                "Issues": 34, "Action needed": 60, "Actions": 12, "Status": 13, "Policy": 40, "Applies to": 22,
                "Excluded": 40, "Requires": 26, "Risk detail": 24, "Recent detections": 40,
                "Where to investigate": 34, "Device": 24, "Operating system": 22, "Defender for Endpoint": 34,
                "Site": 30, "URL": 50, "Primary admin": 30, "Where the broad access is": 60, "Options to fix": 70,
                "Owner": 34, "App": 34, "Why flagged": 30, "High-impact permissions": 50, "Consent": 30,
                "Owners": 30, "Secrets and certificates": 34, "Incident": 10, "Alerts": 40,
                "What the alerts recommend": 60, "Link": 40, "Guest": 34}
STATUS_CHOICES = '"Not started,In progress,Done,Not doing"'
MONTHS = ("January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December")
WORK_GROUPS = GROUP_ORDER[:-1]


def qualification(results):
    if results["Publication"]["CustomerReady"]:
        return ("Prepared for operator review. Release to the customer only after an Arraya engineer has reviewed "
                "this document.")
    return ("ENGINEERING REVIEW ONLY - NOT CUSTOMER-READY. Package integrity is not established; see the package "
            "checks in the Technical Assessment Review.")


def customer_refs(item):
    shown = [r["EvidenceId"] for r in item["EvidenceReferences"] if r["Classification"] != "Restricted"]
    restricted = sum(1 for r in item["EvidenceReferences"] if r["Classification"] == "Restricted")
    value = ", ".join(shown) if shown else "None shareable"
    if restricted:
        value += f" (+{restricted} restricted, held by engineering)"
    return value


def _long_date(iso):
    try:
        d = date.fromisoformat(str(iso)[:10])
    except ValueError:
        return txt(iso)
    return f"{d.day} {MONTHS[d.month - 1]} {d.year}"


# ------------------------------------------------------------------ shared content

def actions(results):
    return sorted((i for i in results["Items"] if i["ActionNumber"]), key=lambda i: i["ActionNumber"])


def _in_groups(results, *groups):
    return [i for i in actions(results) if i["Group"] in groups]


def _decisions(results):
    return [i for i in actions(results) if i["CustomerDecisionRequired"]]


def top_recommendations(results, limit=8):
    """Every Fix before rollout item, then the leading pre-rollout check in each area."""
    chosen = _in_groups(results, "Fix before rollout")
    areas = {i["Area"] for i in chosen}
    for i in _in_groups(results, "Check before rollout"):
        if len(chosen) >= limit:
            break
        if i["Area"] not in areas:
            chosen.append(i)
            areas.add(i["Area"])
    for i in _in_groups(results, "Improve security"):
        if len(chosen) >= 5:
            break
        chosen.append(i)
    return sorted(chosen, key=lambda i: i["ActionNumber"])


def answer(results):
    """Two to four plain sentences that explain the readiness decision."""
    r = results["Readiness"]
    groups = results["Summary"]["ByGroup"]
    fix, check = groups.get("Fix before rollout", 0), groups.get("Check before rollout", 0)
    lines = []
    if fix:
        lines.append(f"{fix} confirmed {'gap affects' if fix == 1 else 'gaps affect'} every Copilot user and must "
                     "be fixed first" + (f", and {check} more must be checked before Copilot is rolled out further."
                                         if check else "."))
    elif check:
        lines.append(f"Nothing we found is a confirmed blocker, but {check} items must be checked or fixed before "
                     "Copilot is rolled out further.")
    else:
        lines.append("No item has to be fixed or checked before Copilot is rolled out.")
    if any(i["FindingId"] == "RVW-EXISTING-COPILOT-USE" for i in results["Items"]):
        lines.append("Copilot is already in use, so these gaps apply today.")
    changed = any(i["Plain"]["Change"] for i in results["Items"] if i["ReviewDecision"] != "DuplicateOrOverlapping")
    if r["AgreesWithScript"]:
        lines.append("This matches the original automated assessment"
                     + (", but we changed how serious some of its findings are." if changed else "."))
    else:
        lines.append(f"The original automated assessment said: {r['ScriptDecision']}. This review reached a "
                     "different answer; the Technical Assessment Review explains why.")
    return " ".join(lines)


def _area_rows(results):
    rows, fills = [], {}
    for area in AREA_NAMES:
        reviewed = [i for i in results["Items"] if i["Area"] == area]
        if not reviewed:
            continue
        work = [i for i in actions(results) if i["Area"] == area]
        if work:
            status = work[0]["Group"]
            point = work[0]["Plain"]["Headline"] + (f" (+{len(work) - 1} more)" if len(work) > 1 else "")
        else:
            status, point = "No action needed", "Reviewed; nothing to do."
        fills[(len(rows), 1)] = GROUP_FILL[status]
        rows.append((area, status, point))
    return rows, fills


# ------------------------------------------------------------------ docx helpers

def _shade(cell, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def _run(paragraph, value, bold=False, italic=False, size=None, color=None):
    raw = value if isinstance(value, str) else txt(value, "")
    run = paragraph.add_run(ascii_clean(raw) + (" " if raw.endswith(" ") else ""))
    run.bold, run.italic = bold, italic
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    return run


def _cell(cell, value, bold=False, size=9, color=None):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    _run(paragraph, value, bold=bold, size=size, color=color)


def _table(doc, headers, rows, widths, fills=None, bold_cols=(), size=9):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.autofit = False
    for i, header in enumerate(headers):
        _cell(table.rows[0].cells[i], header, bold=True, size=size, color="FFFFFF")
        _shade(table.rows[0].cells[i], NAVY)
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    table.rows[0]._tr.get_or_add_trPr().append(repeat)
    for r, row in enumerate(rows):
        added = table.add_row()
        added._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        cells = added.cells
        for c, value in enumerate(row):
            _cell(cells[c], value, bold=c in bold_cols, size=size)
            if fills and (r, c) in fills:
                _shade(cells[c], fills[(r, c)])
    for i, width in enumerate(widths):
        table.columns[i].width = Inches(width)
        for row in table.rows:
            row.cells[i].width = Inches(width)
    _space(doc)
    return table


def _space(doc, points=4):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(points)
    return paragraph


def _small(doc, value, italic=False, color=GREY, size=8.5):
    paragraph = doc.add_paragraph()
    _run(paragraph, value, italic=italic, size=size, color=color)
    return paragraph


def _kv(doc, label, value, size=None):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(2)
    _run(paragraph, label + ": ", bold=True, size=size)
    _run(paragraph, value, size=size)
    return paragraph


def _bullets(doc, values, empty="None.", style="List Bullet"):
    if not values:
        doc.add_paragraph(empty)
        return
    for value in values:
        doc.add_paragraph(ascii_clean(value), style=style)


def _document(results, title, audience):
    doc = Document()
    props = doc.core_properties
    props.author = props.last_modified_by = "Arraya Solutions"
    props.title = title
    props.subject = f"{results['CustomerName']} assessment review"
    props.created = props.modified = FIXED_TIME
    props.revision = 1
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.7)
    section.left_margin = section.right_margin = Inches(0.75)
    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = "Calibri", Pt(10.5)
    normal.paragraph_format.space_after = Pt(4)
    for name in ("Title", "Heading 1", "Heading 2", "Heading 3"):
        doc.styles[name].font.color.rgb = RGBColor.from_string(NAVY)
    doc.styles["Title"].font.size = Pt(24)
    doc.add_heading(title, 0)
    _run(doc.add_paragraph(), f"{results['CustomerName']}  |  {audience}  |  {_long_date(results['ReviewDate'])}",
         size=11, color=GREY)
    ready = results["Publication"]["CustomerReady"]
    _run(doc.add_paragraph(), qualification(results), bold=not ready, italic=ready, size=8.5,
         color=GREY if ready else "9C0006")
    footer = section.footer.paragraphs[0]
    footer.text = qualification(results)
    footer.runs[0].font.size = Pt(7.5)
    return doc


def _callout(doc, heading, body, fill=ANSWER_FILL):
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.rows[0].cells[0]
    cell.width = Inches(7.0)
    _shade(cell, fill)
    paragraph = cell.paragraphs[0]
    _run(paragraph, "The answer", size=9, color=GREY)
    _run(cell.add_paragraph(), heading, bold=True, size=15, color=NAVY)
    _run(cell.add_paragraph(), body, size=10.5)
    _space(doc)


def _glance(doc, results, legend=True):
    groups = results["Summary"]["ByGroup"]
    shown = [g for g in WORK_GROUPS if groups.get(g)]
    cells = [(str(groups[g]), g, GROUP_FILL[g]) for g in shown]
    cells.append((str(results["Summary"]["CustomerDecisions"]), "Decisions for you", "FFFFFF"))
    table = doc.add_table(rows=1, cols=len(cells))
    table.style = "Table Grid"
    width = 7.0 / len(cells)
    for i, (count, label, fill) in enumerate(cells):
        cell = table.rows[0].cells[i]
        cell.width = Inches(width)
        _shade(cell, fill)
        _run(cell.paragraphs[0], count, bold=True, size=18, color=NAVY)
        _run(cell.add_paragraph(), label, size=9)
    _space(doc, 2)
    if legend:
        for g in shown:
            paragraph = doc.add_paragraph()
            paragraph.paragraph_format.space_after = Pt(0)
            _run(paragraph, g + ": ", bold=True, size=8.5, color=GREY)
            _run(paragraph, GROUP_MEANING[g], size=8.5, color=GREY)
        _space(doc)


def _number(item):
    return f"#{item['ActionNumber']}" if item["ActionNumber"] else "-"


def _number_fills(items, col=0):
    return {(r, col): GROUP_FILL[i["Group"]] for r, i in enumerate(items)}


def _about(doc, results):
    ident = results["Identity"]
    doc.add_heading("About this review", 2)
    _small(doc, f"Arraya reviewed the results of the automated Copilot readiness assessment evaluated on "
                f"{_long_date(ident['EvaluationDate'])}. Each finding was checked against the data collected from "
                "the tenant and against current Microsoft guidance. Findings were corrected where the data or "
                "Microsoft documentation disagreed, and items the automated assessment missed were added.")
    _small(doc, "Limits of this review:")
    for limit in results["CustomerLimitations"]:
        paragraph = doc.add_paragraph(style="List Bullet")
        _run(paragraph, limit, size=8.5, color=GREY)
    _small(doc, f"Reference: assessment {ident['AssessmentId']}, run {ident['RunId']}, review version "
                f"{results['ReviewVersion']}, Microsoft guidance version {results['BenchmarkVersion']}.")
    _small(doc, NO_COMMITMENT)


# ------------------------------------------------------------------ executive summary

def executive(results):
    r = results["Readiness"]
    doc = _document(results, "Copilot Readiness: Executive Summary", "For leadership and business sponsors")
    _callout(doc, r["ReviewDecision"], answer(results))
    doc.add_heading("At a glance", 1)
    _glance(doc, results)
    doc.add_heading("Top recommendations", 1)
    doc.add_paragraph("Start with these. The numbers match the Customer Action Plan, which lists every action in order.")
    top = top_recommendations(results)
    _table(doc, ("#", "Do this", "What we found", "Owner"),
           [(_number(i), i["Plain"]["Headline"], i["Plain"]["Issue"], i["ResponsibleRole"]) for i in top],
           [0.45, 2.0, 3.25, 1.3], fills=_number_fills(top), bold_cols=(1,))
    doc.add_heading("How each area looks", 1)
    rows, fills = _area_rows(results)
    _table(doc, ("Area", "Status", "Main point"), rows, [1.8, 1.4, 3.8], fills=fills, bold_cols=(0,))
    decisions = _decisions(results)
    urgent = [i for i in decisions if i["Group"] in ("Fix before rollout", "Check before rollout")]
    if decisions:
        doc.add_heading("Decisions we need from you", 1)
        if urgent:
            doc.add_paragraph("These decisions are needed before Copilot is rolled out further:")
            _table(doc, ("#", "Decision", "Why it matters", "Who decides"),
                   [(_number(i), i["Plain"]["Headline"], i["Plain"]["Why"], i["ResponsibleRole"]) for i in urgent],
                   [0.45, 2.2, 3.05, 1.3], fills=_number_fills(urgent), bold_cols=(1,))
        later = len(decisions) - len(urgent)
        if later:
            doc.add_paragraph(f"{later} further policy {'decision is' if later == 1 else 'decisions are'} listed in "
                              "the Decisions sheet of the Customer Action Plan.")
    doc.add_heading("What happens next", 1)
    steps = ["Agree the decisions above at the review meeting.",
             "Work through the Customer Action Plan in number order, starting with the before-rollout items. Its "
             "list sheets name the specific accounts, sites, devices and apps to fix.",
             "Send Arraya the proof listed in the Done when column of the Action Plan.",
             "The assessment is run again to confirm readiness."]
    _bullets(doc, steps, style="List Number")
    _about(doc, results)
    return doc


# ------------------------------------------------------------------ technical review

def technical(results):
    r = results["Readiness"]
    bench = {b["BenchmarkId"]: b for b in results["Benchmarks"]}
    summary = results["Summary"]
    doc = _document(results, "Copilot Readiness: Technical Assessment Review", "For engineers")
    doc.add_heading("1. Summary", 1)
    _callout(doc, r["ReviewDecision"], answer(results))
    _kv(doc, "Original assessment", r["ScriptDecision"])
    doc.add_paragraph(r["Basis"])
    _kv(doc, "Scope", f"{summary['ItemCount']} items reviewed: {summary['ScriptFindings']} from the original "
                      f"assessment and {summary['AddedFindings']} added by this review. {summary['ActionCount']} need "
                      "action.")
    open_checks = [c for c in results["Package"]["Checks"] if c["Status"] != "Pass"]
    _kv(doc, "Package", ("Integrity established; release to the customer after operator review."
                         if results["Publication"]["CustomerReady"]
                         else "Customer release stopped: " + "; ".join(results["Publication"]["Reasons"]))
        + f" {len(open_checks)} package checks raised issues (Appendix A).")
    _glance(doc, results)

    doc.add_heading("2. Recommendations in order", 1)
    doc.add_paragraph("Work in this order. Within each group, items run by priority, then by dependency (foundation "
                      "before AI-specific controls). Section 4 has the detail for each.")
    for group in WORK_GROUPS:
        members = _in_groups(results, group)
        if not members:
            continue
        doc.add_heading(f"{group} ({len(members)})", 2)
        _small(doc, GROUP_MEANING[group])
        _table(doc, ("#", "Action", "Owner", "Priority", "Ref"),
               [(_number(i), i["Plain"]["Headline"], i["ResponsibleRole"], i["ReviewPriority"], i["FindingId"])
                for i in members], [0.45, 2.85, 1.6, 0.65, 1.45], fills=_number_fills(members), bold_cols=(1,))

    doc.add_heading("3. What changed from the original assessment", 1)
    changed = [i for i in results["Items"] if i["ItemType"] == "ScriptFinding"
               and i["ReviewDecision"] != "DuplicateOrOverlapping" and (i["Plain"]["Change"] or i["ReviewDecision"] in (
                   "CorrectFinding", "CorrectRecommendation", "DowngradeToEvidenceGap", "InsufficientEvidence"))]
    changed.sort(key=lambda i: (i["ActionNumber"] or 999, i["FindingId"]))
    if changed:
        doc.add_heading("Corrected or re-rated", 2)
        _table(doc, ("Ref", "Original said", "Review says", "Why"),
               [(i["FindingId"], _original_said(i), f"{DECISION_LABEL[i['ReviewDecision']]}; {i['Group']}",
                 i["Plain"]["Change"] or i["CorrectedInterpretation"]) for i in changed],
               [1.15, 1.45, 1.4, 3.0])
    added = [i for i in actions(results) if i["ItemType"] == "AddedFinding"]
    if added:
        doc.add_heading("Added by this review", 2)
        _table(doc, ("#", "Action", "What we found"),
               [(_number(i), i["Plain"]["Headline"], i["Plain"]["Issue"]) for i in added],
               [0.45, 2.4, 4.15], fills=_number_fills(added), bold_cols=(1,))
    duplicates = [i for i in results["Items"] if i["ReviewDecision"] == "DuplicateOrOverlapping"]
    if duplicates:
        doc.add_heading("Merged duplicates", 2)
        doc.add_paragraph("; ".join(f"{i['FindingId']} into {i['OverlapsWith']}" for i in duplicates) + ".")

    doc.add_heading("4. Findings in detail", 1)
    doc.add_paragraph("Grouped by area. Evidence notes are labeled by source: " + ", ".join(
        STATEMENT_LABEL[k] for k in STATEMENT_LABEL if k != "Script-derived conclusion") + ". The original report's "
        "own conclusion is shown separately.")
    for area in AREA_NAMES:
        members = [i for i in actions(results) if i["Area"] == area]
        if not members:
            continue
        doc.add_heading(area, 2)
        for item in members:
            _technical_item(doc, item, bench)

    doc.add_heading("5. Reviewed, no action needed", 1)
    quiet = [i for i in results["Items"] if not i["ActionNumber"]]
    _table(doc, ("Ref", "Item", "Result", "Note"),
           [(i["FindingId"], i["Plain"]["Headline"], DECISION_LABEL[i["ReviewDecision"]],
             i["Plain"]["Change"] or i["Plain"]["Issue"]) for i in quiet], [1.15, 2.0, 1.2, 2.65])

    _appendices(doc, results)
    doc.add_paragraph(NO_COMMITMENT)
    return doc


def _original_said(item):
    s = item.get("ScriptConclusion") or {}
    effect = s.get("ReadinessEffect")
    return f"{s.get('Status', 'Not recorded')}" + (f"; {effect}" if effect and effect != "None" else "")


def _technical_item(doc, item, bench):
    p = item["Plain"]
    doc.add_heading(f"{_number(item)}  {p['Headline']}  ({item['FindingId']})", 3)
    _small(doc, " | ".join((item["Group"], f"Priority {item['ReviewPriority']}", f"Owner: {item['ResponsibleRole']}",
                            f"Review result: {DECISION_LABEL[item['ReviewDecision']]}",
                            f"Evidence: {EVIDENCE_LABEL.get(item['EvidenceSufficiency'], item['EvidenceSufficiency'])}",
                            f"{item['FoundationalOrAISpecific']}, step {item['DependencyStep']}")))
    _kv(doc, "What we found", p["Issue"])
    if p["Why"]:
        _kv(doc, "Why it matters", p["Why"])
    _kv(doc, "What to do", p["HowToFix"])
    script_advice = (item.get("ScriptConclusion") or {}).get("Recommendation") or ""
    if item["RecommendedAction"] not in ("", p["HowToFix"], item["ClosureEvidence"]["ValidationProcedure"]) \
            and item["RecommendedAction"] not in script_advice:
        _kv(doc, "Implementation detail", item["RecommendedAction"])
    if item.get("DetailSheets"):
        _kv(doc, "Named list", "Customer Action Plan, " + _affected(item))
    _kv(doc, "Done when", f"{p['DoneWhen']} Proof: {item['ClosureEvidence']['RequiredEvidence']}")
    if p["Change"]:
        _kv(doc, "Change from original", p["Change"])
    _kv(doc, "Reviewer conclusion", item["CorrectedInterpretation"])
    s = item.get("ScriptConclusion")
    if s:
        _kv(doc, "Original report", f"{_original_said(item)}, priority {s['Priority']}. {s['Observation']}")
        if s.get("Recommendation"):
            _kv(doc, "Original advice", s["Recommendation"])
    notes = [st for st in item["Statements"] if st["Type"] != "Script-derived conclusion"]
    if notes:
        _run(doc.add_paragraph(), "Evidence notes:", bold=True)
        for st in notes:
            paragraph = doc.add_paragraph(style="List Bullet")
            _run(paragraph, STATEMENT_LABEL[st["Type"]] + ": ", bold=True, size=9.5)
            _run(paragraph, st["Text"], size=9.5)
    # Package-wide conflicts (C-INTERNAL) are listed once in Appendix C, not on every item.
    issues = [c["Detail"] for c in item["ConsistencyChecks"] if c["Result"] in ("Issue", "Note")
              and c["CheckId"] != "C-INTERNAL"]
    if issues:
        _kv(doc, "Checks raised", " ".join(issues), size=9.5)
    e = item["Evidence"]
    age = (f", {e['AgeDays']} {'day' if e['AgeDays'] == 1 else 'days'} old at review"
           if e.get("AgeDays") is not None else "")
    refs = ", ".join(f"{r['EvidenceId']} [{r['Classification']}]" for r in item["EvidenceReferences"]) or "None linked"
    state = ("see the evidence notes" if e.get("CollectionState") == "See statements"
             else f"collection {e.get('CollectionState')}")
    _kv(doc, "Evidence", f"{e.get('NativeSource')}; observed {e.get('ObservationDate')}{age}; {state}. "
                         f"References: {refs}", size=9.5)
    _kv(doc, "Microsoft guidance", "; ".join(f"{bench[b]['Title']} ({b})" for b in item["MicrosoftBenchmarkIds"]
                                             if b in bench) or "None mapped", size=9.5)


def _appendices(doc, results):
    pkg = results["Package"]
    ident = results["Identity"]
    doc.add_heading("Appendix A. Package and evidence", 1)
    _kv(doc, "Assessment", f"{ident['AssessmentId']}, run {ident['RunId']} ({ident['RunType']}, evaluated "
                           f"{ident['EvaluationDate']})")
    _kv(doc, "Environment", ident["EnvironmentId"])
    _kv(doc, "Methodology", f"{ident['MethodologyVersion']} (catalog {ident['CatalogVersion']})")
    _kv(doc, "Producing revision", f"{ident['ProducingCommit']} ({ident['ProducingSourceState']})")
    _kv(doc, "Review", f"Version {results['ReviewVersion']}, Microsoft guidance version {results['BenchmarkVersion']}, "
                       f"review date {results['ReviewDate']}")
    w = pkg["Windows"]
    period = w.get("EvidencePeriod") or {}
    _kv(doc, "Evidence period", f"{period.get('start')} to {period.get('end')}; collected {w.get('CollectedAt')}")
    _kv(doc, "Oldest observation", f"{w.get('OldestFindingObservation')} ({w.get('OldestFindingAgeDays')} days at "
                                   f"review; limit {w.get('FreshnessThresholdDays')} days)")
    doc.add_heading("Limits of this review", 2)
    _bullets(doc, results["Limitations"])
    doc.add_heading("Package checks", 2)
    checks = sorted(pkg["Checks"], key=lambda c: (c["Status"] == "Pass", c["CheckId"]))
    _table(doc, ("Check", "Result", "Severity", "Detail"),
           [(c["CheckId"], c["Status"], c["Severity"], c["Detail"]) for c in checks], [1.6, 0.6, 0.95, 3.85], size=8.5)
    doc.add_heading("Source classification", 2)
    _table(doc, ("Classification", "Sources"), sorted(pkg["SourceClassCounts"].items()), [2.5, 1.0], size=8.5)
    _bullets(doc, [f"{s['SourceId']}: {s['Classification']}" + (f" ({s['Reason']})" if s["Reason"] else "")
                   for s in pkg["NotableSources"]], "All sources are Complete or CompleteEmpty.")
    doc.add_heading("Input files and hashes", 2)
    _table(doc, ("Role", "File", "Hash status", "SHA-256 (first 16)"),
           [(f["Role"], f["RelativePath"] or "missing", f["HashStatus"], (f["SHA256"] or "")[:16])
            for f in pkg["Files"]], [1.4, 3.1, 1.0, 1.5], size=8)
    doc.add_heading("Appendix B. Control-by-control", 1)
    _table(doc, ("Control", "Original status", "Review result", "Findings", "Notes"),
           [(c["ControlId"], c["ScriptStatus"], _control_group(c["ReviewTreatment"]), ", ".join(c["FindingIds"]),
             c["ReviewNotes"]) for c in results["Controls"]], [1.3, 0.9, 1.2, 1.6, 2.0], size=8.5)
    doc.add_heading("Appendix C. Known conflicts and open questions", 1)
    _bullets(doc, [f"{c['ConflictId']} ({c['Area']}): {c['Detail']}" for c in results["Conflicts"]],
             "No conflicts recorded.")
    _bullets(doc, results["OpenQuestions"], "No open questions.")
    doc.add_heading("Appendix D. Microsoft guidance used", 1)
    _table(doc, ("ID", "Title", "Type", "Verified", "Link"),
           [(b["BenchmarkId"], b["Title"], b["RequirementType"], b["VerifiedOn"], b["Url"])
            for b in results["Benchmarks"]], [1.5, 1.9, 1.1, 0.75, 1.75], size=8)


def _control_group(treatment):
    return GROUP.get(treatment, "Not reviewed")


# ------------------------------------------------------------------ meeting brief

def brief(results):
    r = results["Readiness"]
    doc = _document(results, "Copilot Readiness: Review Meeting Brief", "For the review meeting")
    _callout(doc, r["ReviewDecision"], answer(results))
    doc.add_heading("Meeting goals", 1)
    _bullets(doc, ["Agree the readiness answer above.",
                   "Make the decisions in part 2, or agree who will make them and by when.",
                   "Agree an owner for each item in the before-rollout checklist (part 3).",
                   "Agree what information will be sent after the meeting (part 5)."])
    doc.add_heading("Agenda", 1)
    _bullets(doc, ["The readiness answer", "Decisions needed", "Before-rollout checklist",
                   "What we changed from the original assessment", "Information to send after the meeting"],
             style="List Number")

    doc.add_heading("1. The readiness answer", 1)
    doc.add_paragraph(r["Basis"])
    _glance(doc, results, legend=False)

    doc.add_heading("2. Decisions needed", 1)
    decisions = _decisions(results)
    if decisions:
        _table(doc, ("#", "Decision", "Why it matters", "Who decides"),
               [(_number(i), i["Plain"]["Headline"], i["Plain"]["Why"], i["ResponsibleRole"]) for i in decisions],
               [0.45, 2.2, 3.05, 1.3], fills=_number_fills(decisions), bold_cols=(1,))
    else:
        doc.add_paragraph("No customer decisions are needed.")

    doc.add_heading("3. Before-rollout checklist", 1)
    checklist = _in_groups(results, "Fix before rollout", "Check before rollout")
    if checklist:
        doc.add_paragraph("Each item must be fixed or confirmed before Copilot is rolled out further. Agree an owner "
                          "for each. The Customer Action Plan lists the specific accounts, sites, devices and apps.")
        _table(doc, ("#", "Check", "Owner", "Done when"),
               [(_number(i), i["Plain"]["Headline"], i["ResponsibleRole"], i["Plain"]["DoneWhen"]) for i in checklist],
               [0.45, 2.55, 1.6, 2.4], fills=_number_fills(checklist), bold_cols=(1,))
    else:
        doc.add_paragraph("Nothing has to be fixed or checked before rollout.")

    doc.add_heading("4. What we changed from the original assessment", 1)
    reviewed = sorted((i for i in results["Items"] if i["ItemType"] == "ScriptFinding"
                       and i["ReviewDecision"] != "DuplicateOrOverlapping"),
                      key=lambda i: (i["ActionNumber"] or 10 ** 6, i["FindingId"]))
    added = [i for i in actions(results) if i["ItemType"] == "AddedFinding"]
    items = [f"{i['Plain']['Headline']} (original item {i['FindingId']}): {i['Plain']['Correction']}"
             for i in reviewed if i["Plain"]["Correction"]]
    rerated = [i for i in reviewed if i["Plain"]["RatingChange"]]
    if rerated:
        items.insert(0, f"The original report called {len(rerated)} of these items blockers. After checking the "
                        "evidence we re-rated them: " + "; ".join(
                            f"{_number(i)} {i['Plain']['Headline']} (now {i['Group']})" for i in rerated) + ".")
    if added:
        items.append(f"We added {len(added)} items the original assessment missed: "
                     + "; ".join(f"{_number(i)} {i['Plain']['Headline']}" for i in added) + ".")
    _bullets(doc, items, "No changes.")

    doc.add_heading("5. Information to send after the meeting", 1)
    needed = _in_groups(results, "Information needed")
    if needed:
        _table(doc, ("#", "What to send", "Owner"),
               [(_number(i), f"{i['Plain']['Headline']}: {i['Plain']['DoneWhen']}", i["ResponsibleRole"])
                for i in needed], [0.45, 4.75, 1.8], fills=_number_fills(needed))
    else:
        doc.add_paragraph("No further information is needed.")
    doc.add_heading("Before the assessment is run again", 2)
    _bullets(doc, results["ReassessmentCriteria"])
    _small(doc, NO_COMMITMENT)
    return doc


# ------------------------------------------------------------------ workbook

def _xl(ws, row, col, value, bold=False, fill=None, color=None, italic=False, size=None, wrap=True):
    cell = ws.cell(row, col)
    if isinstance(value, int) and not isinstance(value, bool):
        cell.value = value
    else:
        value = ascii_clean(txt(value, ""))
        cell.value = value or None
        if value:
            cell.data_type = "s"
            if value[:1] in ("=", "+", "-", "@"):
                cell.quotePrefix = True
    cell.alignment = Alignment(vertical="top", wrap_text=wrap)
    if bold or color or italic or size:
        cell.font = Font(bold=bold, italic=italic, color=color, size=size)
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    return cell


def _sheet(ws, headers, rows, name, widths, freeze="B2", fills=None):
    for c, header in enumerate(headers, 1):
        _xl(ws, 1, c, header, bold=True, fill=NAVY, color="FFFFFF")
    for r, row in enumerate(rows, 2):
        for c, value in enumerate(row, 1):
            _xl(ws, r, c, value, fill=(fills or {}).get((r - 2, c - 1)))
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = width
    ws.freeze_panes = freeze
    if rows:
        table = Table(displayName=name, ref=f"A1:{ws.cell(len(rows) + 1, len(headers)).coordinate}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=False)
        ws.add_table(table)


def _start_sheet(ws, results, sheets):
    r = results["Readiness"]
    ident = results["Identity"]
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 110
    _xl(ws, 1, 1, "Customer Action Plan", bold=True, size=16, color=NAVY, wrap=False)
    _xl(ws, 2, 1, f"{results['CustomerName']}  |  Review date {_long_date(results['ReviewDate'])}", color=GREY,
        wrap=False)
    _xl(ws, 3, 1, qualification(results), italic=True, color=GREY, wrap=False)
    _xl(ws, 4, 1, CONFIDENTIAL, bold=True, color="9C0006", wrap=False)
    row = 6
    _xl(ws, row, 1, "Readiness answer", bold=True)
    _xl(ws, row, 2, r["ReviewDecision"], bold=True, size=13, color=NAVY)
    _xl(ws, row + 1, 1, "In short", bold=True)
    _xl(ws, row + 1, 2, answer(results))
    row += 3
    _xl(ws, row, 1, "Actions by group", bold=True, size=12, color=NAVY)
    row += 1
    for group in WORK_GROUPS:
        count = results["Summary"]["ByGroup"].get(group)
        if count:
            _xl(ws, row, 1, group, bold=True, fill=GROUP_FILL[group])
            _xl(ws, row, 2, f"{count} {'action' if count == 1 else 'actions'}. {GROUP_MEANING[group]}")
            row += 1
    _xl(ws, row, 1, "Decisions for you", bold=True)
    _xl(ws, row, 2, f"{results['Summary']['CustomerDecisions']} policy decisions (see the Decisions sheet).")
    row += 2
    _xl(ws, row, 1, "How to use this workbook", bold=True, size=12, color=NAVY)
    row += 1
    for label, value in (
            ("Action Plan", "Every action, in the order to work them. # is the action number used in all review "
                            "documents. Use the filter arrows on When, Priority or Owner to see your part. The "
                            "Affected items column names the sheet that lists the specific accounts, sites, devices "
                            "or apps."),
            ("Status columns", "Blank until you fill them in. Choose Not started, In progress, Done or Not doing."),
            ("Decisions", "Policy decisions only you can make. Record the decision, who made it and when.")):
        _xl(ws, row, 1, label, bold=True)
        _xl(ws, row, 2, value)
        row += 1
    row += 1
    _xl(ws, row, 1, "Lists for each action", bold=True, size=12, color=NAVY)
    row += 1
    for sheet in sheets:
        _xl(ws, row, 1, sheet.name, bold=True)
        count = f"{len(sheet.rows)} {'row' if len(sheet.rows) == 1 else 'rows'}. " if sheet.rows else ""
        _xl(ws, row, 2, count + sheet.purpose + (" " + sheet.note if sheet.note else ""))
        row += 1
    for label, value in (
            ("Review trace", "Reference for engineers: original report titles, review results, evidence and "
                             "Microsoft guidance for every item, including items that need no action."),
            ("Microsoft guidance", "The Microsoft Learn pages this review relied on.")):
        _xl(ws, row, 1, label, bold=True)
        _xl(ws, row, 2, value)
        row += 1
    row += 1
    _xl(ws, row, 1, "Reference", bold=True)
    _xl(ws, row, 2, f"Assessment {ident['AssessmentId']}, run {ident['RunId']}")
    _xl(ws, row + 1, 1, "Commitment", bold=True)
    _xl(ws, row + 1, 2, NO_COMMITMENT)


def _affected(item):
    return "; ".join(f"{d['Sheet']} sheet: {d['Summary']}" for d in item.get("DetailSheets") or []) or "None listed"


def _status_list(ws, column, rows):
    if rows:
        status = DataValidation(type="list", formula1=STATUS_CHOICES, allow_blank=True)
        ws.add_data_validation(status)
        letter = ws.cell(1, column).column_letter
        status.add(f"{letter}2:{letter}{rows + 1}")


def _detail_sheet(book, sheet, numbers):
    ws = book.create_sheet(sheet.name)
    headers = sheet.headers + ("Status",)
    rows = []
    for values, refs in zip(sheet.rows, sheet.refs):
        related = sorted({numbers[fid] for fid, _ in refs if numbers.get(fid)})
        rows.append(values + (", ".join(f"#{n}" for n in related) or "None", None))
    if not rows:
        for c, header in enumerate(headers, 1):
            _xl(ws, 1, c, header, bold=True, fill=NAVY, color="FFFFFF")
        _xl(ws, 2, 1, sheet.note or "Nothing to list.", italic=True, color=GREY, wrap=False)
        return
    _sheet(ws, headers, rows, sheet.table, [DETAIL_WIDTH.get(h, 18) for h in headers], freeze="B2")
    link = headers.index("Link") + 1 if "Link" in headers else None
    if link:
        for r in range(2, len(rows) + 2):
            cell = ws.cell(r, link)
            if str(cell.value or "").startswith("https://"):
                cell.hyperlink = cell.value
    _status_list(ws, len(headers), len(rows))


def workbook(results, sheets=()):
    book = Workbook()
    book.properties.creator = book.properties.lastModifiedBy = "Arraya Solutions"
    book.properties.title = "Customer Action Plan"
    book.properties.created = book.properties.modified = FIXED_TIME.replace(tzinfo=None)
    _start_sheet(book.active, results, sheets)
    book.active.title = "Start Here"

    plan = book.create_sheet("Action Plan")
    work = actions(results)
    _sheet(plan, ACTION_HEADERS,
           [(i["ActionNumber"], i["Group"], i["ReviewPriority"], i["Plain"]["Headline"], i["ResponsibleRole"],
             i["Plain"]["Issue"], i["Plain"]["HowToFix"], _affected(i), i["Plain"]["DoneWhen"], None, i["FindingId"])
            for i in work],
           "ActionPlan", [5, 18, 9, 40, 24, 50, 55, 36, 40, 13, 17], freeze="E2",
           fills={(r, 1): GROUP_FILL[i["Group"]] for r, i in enumerate(work)})
    _status_list(plan, ACTION_HEADERS.index("Status") + 1, len(work))

    decisions = book.create_sheet("Decisions")
    _sheet(decisions, ("#", "Decision needed", "Why it matters", "Who decides", "Decision", "Decided by", "Date"),
           [(i["ActionNumber"], i["Plain"]["Headline"], i["Plain"]["Why"], i["ResponsibleRole"], None, None, None)
            for i in _decisions(results)], "Decisions", [5, 40, 60, 28, 40, 20, 12])

    numbers = {i["FindingId"]: i["ActionNumber"] for i in results["Items"]}
    for sheet in sheets:
        _detail_sheet(book, sheet, numbers)

    trace = book.create_sheet(TRACE_SHEET)
    every = sorted(results["Items"], key=lambda i: (i["ActionNumber"] or 10 ** 6, i["DependencyStep"], i["FindingId"]))
    _sheet(trace, DETAIL_HEADERS,
           [(i["ActionNumber"], i["FindingId"], i["Plain"]["Headline"], i["Group"], i["Area"],
             i["Title"] if i["ItemType"] == "ScriptFinding" else "Added by this review", DECISION_LABEL[i["ReviewDecision"]],
             EVIDENCE_LABEL.get(i["EvidenceSufficiency"], i["EvidenceSufficiency"]), i["Plain"]["Change"],
             ", ".join(i["MicrosoftBenchmarkIds"]), customer_refs(i)) for i in every],
           "ReviewTrace", [5, 17, 40, 18, 22, 36, 20, 16, 55, 34, 30], freeze="C2")

    guidance = book.create_sheet("Microsoft guidance")
    _sheet(guidance, ("ID", "Guidance", "What Microsoft says", "Link"),
           [(b["BenchmarkId"], b["Title"], b["Statement"], b["Url"]) for b in results["Benchmarks"]],
           "Guidance", [28, 45, 80, 70])
    return book


# ------------------------------------------------------------------ packaging

def normalized_zip(data):
    """Pin OOXML container metadata so identical inputs produce identical bytes."""
    out = BytesIO()
    with zipfile.ZipFile(BytesIO(data)) as source, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as target:
        for name in source.namelist():
            content = source.read(name)
            if name == "docProps/core.xml":
                content = re.sub(rb"(<dcterms:(?:created|modified)\b[^>]*>)[^<]+", rb"\g<1>2000-01-01T00:00:00Z", content)
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type, info.create_system, info.external_attr = zipfile.ZIP_DEFLATED, 0, 0o600 << 16
            target.writestr(info, content)
    return out.getvalue()


def render_all(results, sheets=()):
    """Render the four deliverables. Only the workbook receives the object-level list sheets."""
    outputs = {}
    builders = ((EXEC, executive), (TECH, technical), (PLAN, lambda r: workbook(r, sheets)), (BRIEF, brief))
    for name, builder in builders:
        buffer = BytesIO()
        builder(results).save(buffer)
        outputs[name] = normalized_zip(buffer.getvalue())
    return outputs
