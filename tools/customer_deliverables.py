"""Deterministic rendering of approved projection 1.2.0; no assessment evaluation."""
from __future__ import annotations
import argparse
from datetime import datetime, timezone
from hashlib import sha256
from io import BytesIO
import json
import os
from pathlib import Path
import re
import tempfile
import zipfile
from xml.etree import ElementTree as ET

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

REVIEW = "agent-customer-review.json"
KNOWLEDGE = "agent-knowledge.md"
INDEX = "agent-evidence-index.json"
MANIFEST = "agent-demo-manifest.json"
INPUTS = (REVIEW, KNOWLEDGE, INDEX, MANIFEST)
OUTPUTS = ("Executive Readiness Summary.docx", "Customer Action Plan.xlsx", "Review Meeting Brief.docx")
QUALIFICATION = "ENGINEERING VALIDATION ONLY \u2014 NOT CUSTOMER-READY"
FIXED_TIME = datetime(2000, 1, 1, tzinfo=timezone.utc)
HEADERS = ("Finding ID", "Title", "Priority", "Pilot impact", "Observed condition",
           "Qualification", "Recommended action", "Responsible role", "Closure evidence",
           "Customer action status")
FIELDS = ("FindingId", "Title", "Priority", "PilotImpact", "CurrentConclusion",
          "Qualification", "RecommendedAction", "ResponsibleRole", "ClosureEvidence",
          "CustomerActionStatus")
STATE_FIELDS = ("RecordId", "FindingId", "CustomerFindingId", "PersistentFindingId",
                "SourceRecommendationId", "ActionId", "TechnicalStatus",
                "EvidenceSufficiency", "LifecycleState", "GovernanceTreatment",
                "CustomerActionStatus", "ActionType", "EvidenceLevel", "ObservationDate")
LIFECYCLE_FIELDS = ("EntityId", "EntityType", "BaselineEntityId", "CurrentEntityId", "State",
                   "BaselineAssessmentStatus", "CurrentAssessmentStatus", "Eligibility",
                   "ComparabilityOutcome", "RemainsOpen", "RemainingClosureRequirement",
                   "ClosureRequirementSupported")
GOVERNANCE_FIELDS = ("DecisionId", "TargetEntityId", "TargetEntityType", "DecisionType",
    "WorkflowState", "GovernanceState", "CurrentGovernanceState", "CurrentLifecycleState",
    "CurrentAssessmentStatus", "EffectiveForCurrentRun", "CustomerActionStatus",
    "CurrentRunQualification", "ApprovalRole", "AuthorityRole", "ReviewAt",
    "ExpirationAt", "ValidationResult", "ClosureCriteriaEvaluated", "ClosureOverride")
SENSITIVE = re.compile(
    r"https?://\S+|\b[^\s<>@]+@[^\s<>@]+\b|"
    r"\b[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}\b|"
    r"\b(?:\d{1,3}\.){3}\d{1,3}\b|[A-Za-z]:[\\/]|"
    r"\bARE-[0-9a-f]+\b|\b[\w.-]+\.sharepoint\.com\b|"
    r"(?<!\w)(?:\+\d[\d ()-]{8,}\d|\(\d{3}\)\s*\d{3}[- ]\d{4})(?!\w)",
    re.I)
PRIVATE_KEYS = {"displayname", "givenname", "surname", "userprincipalname", "email",
                "mail", "username", "upn", "sitename", "siteurl", "tenantname",
                "tenantid", "tenantdomain", "objectid", "appid", "applicationid",
                "ipaddress", "accountableowner", "decisionauthority", "mobilephone"}


def require(condition, message):
    if not condition:
        raise ValueError(message)


def canonical(value):
    return (json.dumps(value, ensure_ascii=False, sort_keys=True, indent=2,
                       allow_nan=False) + "\n").encode("utf-8")


def digest(value):
    return sha256(value).hexdigest()


def parse_json(data):
    def unique(pairs):
        result = {}
        for key, value in pairs:
            require(key not in result, "Duplicate JSON key.")
            result[key] = value
        return result
    return json.loads(data.decode("utf-8-sig"), object_pairs_hook=unique,
                      parse_constant=lambda value: (_ for _ in ()).throw(ValueError("Non-finite JSON value.")))


def scalar(value):
    require(value is None or type(value) in (str, bool, int, float), "Non-scalar exported field.")
    if value is None:
        return "Not recorded"
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value) if str(value) else "Not recorded"


def validate_package(directory):
    """Read only these four named files; the restricted index never enters rendering."""
    root = Path(directory).resolve(strict=True)
    payloads = {name: (root / name).read_bytes() for name in INPUTS}
    review, manifest, index = (parse_json(payloads[n]) for n in (REVIEW, MANIFEST, INDEX))
    require(digest(canonical({k: v for k, v in manifest.items() if k != "ManifestHash"}))
            == manifest.get("ManifestHash"), "Manifest self-hash mismatch.")
    rows = manifest.get("GeneratedFiles", [])
    require(len(rows) == 3 and {r["FileName"] for r in rows} == {REVIEW, KNOWLEDGE, INDEX},
            "Manifest must list exactly the three approved export files.")
    for row in rows:
        require(digest(payloads[row["FileName"]]) == row["SHA256"],
                "Manifest file hash mismatch: " + row["FileName"])
        expected = "RestrictedTechnical" if row["FileName"] == INDEX else "CustomerReview"
        require(row["AccessClassification"] == expected, "Manifest classification mismatch.")
    require(review.get("projection_version") == index.get("projection_version")
            == manifest.get("ProjectionVersion") == "1.2.0", "Projection version must be 1.2.0 throughout.")
    require(review.get("kind") == "customer-review" and index.get("kind") == "evidence-index",
            "Wrong export kind.")
    identity = review["identity"]
    for key in ("AssessmentId", "RunId"):
        value = identity[key]
        require(isinstance(value, str) and re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,150}", value),
                "Unsafe assessment/run identifier.")
        require(value == manifest.get(key) == index.get(key), key + " mismatch.")
    require(review["source_sha256"] == index["source_sha256"] == manifest["SourceHash"],
            "Source hash metadata mismatch.")
    require(review["output_hashes"] == {
        REVIEW: digest(canonical({k: v for k, v in review.items() if k not in ("output_hashes", "hash_contract")})),
        KNOWLEDGE: digest(payloads[KNOWLEDGE]), INDEX: digest(payloads[INDEX])},
        "Review hash contract mismatch.")
    require(manifest["PublicationQualification"] == QUALIFICATION
            and manifest["customerPublicationEligible"] is False
            and review["customerPublicationEligible"] is False
            and review["exporter_provenance"]["output_qualification"] == QUALIFICATION
            and review["demoUse"] == manifest["IntendedUse"] == "InternalReadOnlyEvaluation",
            "Engineering-only qualification is required.")
    require(not review["approved_documents"] and manifest["ApprovedLinksPresent"] is False,
            "This MVP does not export approved links.")
    # Structural reconciliation only: never evaluate evidence or recalculate a finding.
    actions, findings = review["actions"], review["findings"]
    for records in (actions, findings):
        ids = [r["FindingId"] for r in records]
        require(len(ids) == len(set(ids)) and all(ids), "Missing or duplicate finding ID.")
        require(all(r["FindingId"] == r["RecordId"] for r in records), "Finding/record ID mismatch.")
    action_map = {a["FindingId"]: a for a in actions}
    finding_map = {a["FindingId"]: a for a in findings}
    require(set(action_map) <= set(finding_map), "Action has no source finding.")
    for key, action in action_map.items():
        for field in set(FIELDS + STATE_FIELDS):
            require(action.get(field) == finding_map[key].get(field), "Action/finding disagreement: " + field)
    rec = review["reconciliation"]
    require(len(actions) == review["counts"]["actions"] == rec["actions"] == 20,
            "Expected exactly twenty source customer actions.")
    require(len(review["blockers_ids"]) == len(set(review["blockers_ids"])) == rec["blockers"] == 5,
            "Expected exactly five source blockers.")
    require(set(review["blockers_ids"]) <= set(action_map), "Unresolved blocker ID.")
    require(sorted(action_map) == sorted(rec["action_ids"]), "Action ID reconciliation mismatch.")
    require(sorted(finding_map) == sorted(rec["finding_ids"]) and len(findings) == rec["findings"],
            "Finding ID reconciliation mismatch.")
    for key in ("pilot_conditions_ids", "strengths_ids", "evidence_gaps"):
        require(set(review[key]) <= set(finding_map), "Unresolved " + key)
    # Detect a package replacement during validation, including Markdown and restricted index.
    hashes = {name: digest(data) for name, data in payloads.items()}
    assert_inputs_unchanged(root, hashes)
    return review, hashes


def assert_inputs_unchanged(root, hashes):
    require(all(digest((root / name).read_bytes()) == value for name, value in hashes.items()),
            "Inputs changed during processing; no package may be published.")


def private_terms(review):
    values = set()
    def walk(item):
        if isinstance(item, dict):
            for key, value in item.items():
                if re.sub("[^a-z]", "", key.lower()) in PRIVATE_KEYS and isinstance(value, str) and len(value) > 2:
                    values.add(value.casefold())
                walk(value)
        elif isinstance(item, list):
            for value in item:
                walk(value)
    walk(review)
    return values


def privacy_check(texts, review):
    terms = private_terms(review)
    for text in texts:
        require(not SENSITIVE.search(text), "Privacy check failed: restricted identity/link pattern.")
        lowered = text.casefold()
        require(not any(value in lowered for value in terms), "Privacy check failed: private source value.")


def document(title, review):
    doc = Document()
    props = doc.core_properties
    props.author = "Arraya Solutions"
    props.last_modified_by = "Arraya Solutions"
    props.title, props.subject = title, "Engineering validation assessment review"
    props.created = props.modified = FIXED_TIME
    props.revision = 1
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(.65)
    section.left_margin = section.right_margin = Inches(.7)
    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = "Calibri", Pt(10)
    normal.paragraph_format.space_after = Pt(5)
    for name in ("Title", "Heading 1", "Heading 2"):
        doc.styles[name].font.color.rgb = RGBColor.from_string("16324F")
    doc.add_heading(title, 0)
    doc.add_paragraph(QUALIFICATION, "Subtitle")
    i = review["identity"]
    doc.add_paragraph(f"Assessment: {i['AssessmentId']}\nRun: {i['RunId']}\n"
                      f"Source evaluation date: {scalar(review['evaluation_date'])}\n"
                      f"Methodology: {scalar(review['methodology_version'])}")
    section.footer.paragraphs[0].text = QUALIFICATION
    section.footer.paragraphs[0].runs[0].font.size = Pt(8)
    return doc


def para(doc, label, value):
    p = doc.add_paragraph()
    p.add_run(label + ": ").bold = True
    p.add_run(scalar(value))


def states(row):
    return "; ".join(f"{label}: {scalar(row.get(key))}" for label, key in (
        ("Technical status", "TechnicalStatus"), ("Evidence sufficiency", "EvidenceSufficiency"),
        ("Lifecycle state", "LifecycleState"), ("Governance treatment", "GovernanceTreatment"),
        ("Customer action status", "CustomerActionStatus")))


def finding(doc, row, detail=True):
    doc.add_heading(row["FindingId"] + " | " + scalar(row["Title"]), 2)
    para(doc, "Priority / pilot impact", scalar(row["Priority"]) + " / " + scalar(row["PilotImpact"]))
    if detail:
        para(doc, "Observed condition", row["CurrentConclusion"])
        para(doc, "Qualification", row.get("Qualification"))
    para(doc, "Recommended action", row["RecommendedAction"])
    para(doc, "Responsible role", row.get("ResponsibleRole"))
    para(doc, "Closure evidence", row["ClosureEvidence"])
    doc.add_paragraph(states(row))


def context_pairs(review):
    pairs = [(QUALIFICATION, "Internal engineering review only"),
        ("Current decision", review["decision"]), ("Evidence sufficiency", review["evidence_sufficiency"]),
        ("Source qualification", review["source_publication"]["output_qualification"]),
        ("Export qualification", review["exporter_provenance"]["output_qualification"]),
        ("Null/empty values", "Not recorded; no owner, date, closure, exception, or approval is inferred.")]
    for key in ("AssessmentId", "RunId", "RunType", "BaselineRunId", "EvaluatedAt"):
        pairs.append((key, review["identity"].get(key)))
    for key, value in review["counts"].items():
        pairs.append(("Source count: " + key, value))
    for key, value in review["controlCounts"].items():
        pairs.append(("Source control count: " + key, value))
    pairs += [("Source blocker count", review["reconciliation"]["blockers"]),
              ("Source blockers", ", ".join(review["blockers_ids"])),
              ("Source pilot conditions", ", ".join(review["pilot_conditions_ids"]))]
    for group, fields, record_fields in (
        ("lifecycle", ("recorded", "Enabled", "Outcome", "BaselineRunId"), LIFECYCLE_FIELDS),
        ("governance", ("recorded", "AsOf"), GOVERNANCE_FIELDS)):
        for key in fields:
            pairs.append((group + ": " + key, review[group].get(key)))
        for number, record in enumerate(review[group]["records"], 1):
            for key in record_fields:
                if key in record:
                    pairs.append((f"{group} record {number}: {key}", record[key]))
    return [(k, scalar(v)) for k, v in pairs]


def common_limitations(doc, review):
    doc.add_heading("Important evidence limitations", 1)
    for value in review["limitations"] + review.get("integrity_limitations", []):
        doc.add_paragraph(scalar(value), "List Bullet")
    para(doc, "Overall evidence sufficiency", review["evidence_sufficiency"])
    para(doc, "Lifecycle outcome", review["lifecycle"].get("Outcome"))
    para(doc, "Governance recorded", review["governance"].get("recorded"))
    doc.add_paragraph("This package preserves the completed assessment. It does not recalculate results, "
                      "close findings, accept risk, approve exceptions, assign people, or set dates.")


def executive(review):
    doc = document("Executive Readiness Summary", review)
    by_id = {r["FindingId"]: r for r in review["findings"]}
    doc.add_heading("Assessment purpose and scope", 1)
    doc.add_paragraph("Explain the completed M365 readiness and hardening assessment and its recorded "
                      "pilot conditions using the approved customer-review export. Scope is limited "
                      "to the controls and evidence qualifications represented in that export.")
    domains = sorted({c["domain_id"] for c in review["controls"]})
    para(doc, "Recorded assessment domains", ", ".join(domains))
    doc.add_heading("Current readiness decision", 1)
    doc.add_paragraph(review["decision"], "Heading 2")
    doc.add_paragraph(review["rationale"])
    para(doc, "Recorded customer actions", review["counts"]["actions"])
    doc.add_heading("Five pilot blockers", 1)
    for key in review["blockers_ids"]:
        finding(doc, by_id[key])
    doc.add_heading("Items requiring confirmation", 1)
    for row in review["actions"]:
        if row["ActionType"] == "Confirmation":
            finding(doc, row, detail=False)
    doc.add_heading("Supported configuration observations", 1)
    doc.add_paragraph("Configuration observations do not establish operational effectiveness.")
    for control in review["controls"]:
        if control["status"] == "Observed":
            para(doc, control["control_id"], control["title"])
            para(doc, "Recorded results", f"Configuration: {scalar(control['configuration_result'])}; "
                 f"operation: {scalar(control['operational_result'])}; status: {control['status']}")
            if control.get("statusExplanation"):
                doc.add_paragraph(control["statusExplanation"])
    doc.add_heading("Immediate next steps", 1)
    for line in (
        "Review the five listed blockers and their recorded recommended actions first.",
        "Use the Customer Action Plan to review all twenty source actions and their closure evidence.",
        "Confirm unresolved items and supply the requested scoped evidence through the existing review process.",
        "Request reassessment when the required closure evidence and outstanding confirmations are available."):
        doc.add_paragraph(line, "List Bullet")
    common_limitations(doc, review)
    return doc


def meeting(review):
    doc = document("Review Meeting Brief", review)
    by_id = {r["FindingId"]: r for r in review["findings"]}
    doc.add_heading("Meeting objective", 1)
    doc.add_paragraph("Review the recorded readiness decision, understand the blockers and evidence gaps, "
                      "and agree the evidence and follow-up needed for reassessment.")
    doc.add_heading("Current decision", 1)
    doc.add_paragraph(review["decision"])
    doc.add_paragraph(review["rationale"])
    doc.add_heading("Blockers to review", 1)
    for key in review["blockers_ids"]:
        finding(doc, by_id[key], detail=False)
    doc.add_heading("Evidence gaps", 1)
    for key in review["evidence_gaps"]:
        row = by_id[key]
        para(doc, key, row["Title"])
        para(doc, "Observed condition", row["CurrentConclusion"])
        para(doc, "Qualification", row.get("Qualification"))
        para(doc, "Evidence sufficiency", row["EvidenceSufficiency"])
    doc.add_heading("Customer decisions needed", 1)
    for line in (
        "Confirm the intended pilot scope against the scope represented in this assessment.",
        "Identify accountable people through the customer's process for the recorded responsible roles; none are assigned by this package.",
        "Agree how the specified closure evidence and outstanding confirmations will be supplied.",
        "Review recorded governance treatment where present; this meeting brief grants no approval or exception."):
        doc.add_paragraph(line, "List Bullet")
    doc.add_heading("Evidence requests", 1)
    for row in review["actions"]:
        para(doc, row["FindingId"] + " | " + row["Title"], row["ClosureEvidence"])
    doc.add_heading("Recommended discussion order", 1)
    for line in ("Confirm scope and engineering-only qualification.", "Review the current decision and its five blockers.",
                 "Review confirmation items and evidence gaps.", "Review the twenty actions and their unchanged priorities.",
                 "Agree evidence follow-up and the next reassessment condition."):
        doc.add_paragraph(line, "List Number")
    doc.add_heading("Next reassessment condition", 1)
    doc.add_paragraph("Request reassessment through the existing assessment process when the recorded closure "
                      "evidence for blockers and other open actions, plus outstanding confirmations, can be reviewed. "
                      "Evidence submission alone does not close a finding or change the current decision.")
    common_limitations(doc, review)
    return doc


def text_cell(ws, row, col, value):
    cell = ws.cell(row, col, scalar(value))
    cell.data_type = "s"  # Source text cannot become an executable Excel formula.
    if cell.value[:1] in "=+-@":
        cell.quotePrefix = True
    cell.alignment = Alignment(vertical="top", wrap_text=True)
    return cell


def tabulate(ws, headers, rows, name, widths):
    for r, values in enumerate([headers, *rows], 1):
        for c, value in enumerate(values, 1):
            text_cell(ws, r, c, value)
    for cell in ws[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = PatternFill("solid", fgColor="16324F")
    ws.row_dimensions[1].height = 32
    for r in range(2, len(rows) + 2):
        lines = max(sum(max(1, (len(part) + int(widths[c]) - 1) // int(widths[c]))
                        for part in scalar(v).splitlines()) for c, v in enumerate(rows[r-2]))
        ws.row_dimensions[r].height = min(400, max(45, 16 * (lines + 1)))
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = width
    ws.freeze_panes = "C2" if name == "CustomerActions" else "B2"
    if rows:
        table = Table(displayName=name, ref=ws.dimensions)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(table)
    ws.sheet_view.zoomScale = 75
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddFooter.center.text = QUALIFICATION
    ws.oddFooter.center.size = 8


def workbook(review):
    book = Workbook()
    book.properties.creator = book.properties.lastModifiedBy = "Arraya Solutions"
    book.properties.title = "Customer Action Plan"
    book.properties.subject = QUALIFICATION
    book.properties.created = book.properties.modified = FIXED_TIME.replace(tzinfo=None)
    ws = book.active
    ws.title = "Customer Actions"
    tabulate(ws, HEADERS, [[r.get(k) for k in FIELDS] for r in review["actions"]],
             "CustomerActions", [23, 43, 15, 25, 65, 60, 65, 35, 65, 26])
    states_ws = book.create_sheet("Source States")
    tabulate(states_ws, STATE_FIELDS, [[r.get(k) for k in STATE_FIELDS] for r in review["actions"]],
             "SourceStates", [25] * len(STATE_FIELDS))
    context = book.create_sheet("Assessment Context")
    tabulate(context, ["Source field", "Recorded value"], context_pairs(review),
             "AssessmentContext", [48, 115])
    return book


def normalized_zip(data):
    """Pin container metadata and OOXML timestamps, keeping content unchanged."""
    result = BytesIO()
    with zipfile.ZipFile(BytesIO(data)) as source, zipfile.ZipFile(result, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as target:
        for name in sorted(source.namelist()):
            content = source.read(name)
            if name == "docProps/core.xml":
                content = re.sub(rb"(<dcterms:(?:created|modified)\b[^>]*>)[^<]+",
                                 rb"\g<1>2000-01-01T00:00:00Z", content)
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type, info.create_system, info.external_attr = zipfile.ZIP_DEFLATED, 0, 0o600 << 16
            target.writestr(info, content)
    return result.getvalue()


def render(review):
    outputs = {}
    for name, obj in ((OUTPUTS[0], executive(review)), (OUTPUTS[1], workbook(review)),
                      (OUTPUTS[2], meeting(review))):
        data = BytesIO()
        obj.save(data)
        outputs[name] = normalized_zip(data.getvalue())
    validate_outputs(outputs, review)
    return outputs


def validate_outputs(outputs, review):
    require(set(outputs) == set(OUTPUTS), "Unexpected output files.")
    all_texts, word_text = [], {}
    for name, data in outputs.items():
        with zipfile.ZipFile(BytesIO(data)) as archive:
            require(archive.testzip() is None, "Corrupt Office container.")
            for part in archive.namelist():
                if part.endswith((".xml", ".rels")):
                    root = ET.fromstring(archive.read(part))
                    # Source text lives in text nodes. Namespace URIs are OOXML metadata.
                    all_texts.extend(t for t in root.itertext() if t.strip())
            require(not any("vbaProject" in p or "/embeddings/" in p for p in archive.namelist()),
                    "Unexpected active or embedded content.")
        if name.endswith(".docx"):
            doc = Document(BytesIO(data))
            text = "\n".join(p.text for p in doc.paragraphs)
            word_text[name] = text
            require(QUALIFICATION in text and review["decision"] in text, "Missing qualification/decision.")
            section_title, next_title = (("Five pilot blockers", "Items requiring confirmation")
                                         if name == OUTPUTS[0] else ("Blockers to review", "Evidence gaps"))
            blocker_section = text.split(section_title + "\n", 1)[1].split(next_title + "\n", 1)[0]
            for key in review["blockers_ids"]:
                require(blocker_section.count(key + " | ") == 1, "Blocker missing or duplicated in report.")
        else:
            book = load_workbook(BytesIO(data), data_only=False)
            ws = book["Customer Actions"]
            require(list(next(ws.values)) == list(HEADERS) and ws.max_row == 21, "Action table shape mismatch.")
            require("CustomerActions" in ws.tables and ws.tables["CustomerActions"].autoFilter is not None,
                    "Missing action filters.")
            for actual, expected in zip(list(ws.values)[1:], review["actions"]):
                require(list(actual) == [scalar(expected.get(k)) for k in FIELDS], "Action content differs from source.")
            expected_states = [[scalar(r.get(k)) for k in STATE_FIELDS] for r in review["actions"]]
            require([list(r) for r in list(book["Source States"].values)[1:]] == expected_states,
                    "Source state differs from JSON.")
            require(all(c.data_type != "f" for sheet in book for row in sheet for c in row),
                    "Unexpected Excel formula.")
            book.close()
    # IDs can appear again in evidence requests in the meeting brief; check its blocker section separately.
    privacy_check(all_texts, review)
    return {"blockers": review["reconciliation"]["blockers"], "actions": review["counts"]["actions"],
            "privacy": "Passed: allowlisted fields and identity-pattern checks; restricted index content excluded.",
            "opened": list(OUTPUTS)}


def generate(input_dir, output_root):
    source = Path(input_dir).resolve(strict=True)
    review, input_hashes = validate_package(source)
    first = render(review)
    require(first == render(review), "Repeated generation is not byte deterministic.")
    root = Path(output_root).resolve()
    destination = root / review["identity"]["AssessmentId"] / review["identity"]["RunId"]
    require(destination.resolve().is_relative_to(root), "Output path escapes output root.")
    require(not destination.resolve().is_relative_to(source) and not source.is_relative_to(destination.resolve()),
            "Output and input paths overlap.")
    assert_inputs_unchanged(source, input_hashes)
    if destination.exists():
        require(destination.is_dir() and {p.name for p in destination.iterdir()} == set(OUTPUTS),
                "Existing destination contains different files; no overwrite performed.")
        require(all((destination / n).read_bytes() == data for n, data in first.items()),
                "Existing deliverables differ; no overwrite performed.")
    else:
        destination.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix=".mvp-", dir=destination.parent) as temp:
            stage = Path(temp).resolve()
            require(stage.is_relative_to(root), "Staging path escapes output root.")
            for name, data in first.items():
                (stage / name).write_bytes(data)
            validate_outputs({n: (stage / n).read_bytes() for n in OUTPUTS}, review)
            assert_inputs_unchanged(source, input_hashes)
            # Exact three-file directory is exposed only after every check passes.
            os.replace(stage, destination)
    assert_inputs_unchanged(source, input_hashes)
    return {"status": "Generated and validated", "output": str(destination),
            "files": {n: digest(data) for n, data in first.items()},
            "manifest_hashes": "All three listed file hashes and manifest self-hash passed",
            "projection_version": "1.2.0", "input_hashes": input_hashes,
            "inputs_unchanged": True, "deterministic": True,
            **validate_outputs(first, review),
            "qualification": QUALIFICATION, "ready_for_operator_review": True}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-dir", type=Path, default=Path("inputs/assessment-package"))
    parser.add_argument("--output-root", type=Path, default=Path("generated"))
    args = parser.parse_args()
    try:
        result = generate(args.input_dir, args.output_root)
    except (ValueError, OSError, KeyError, TypeError, json.JSONDecodeError) as exc:
        print(json.dumps({"status": "Blocked", "reason": str(exc), "ready_for_operator_review": False}))
        return 1
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

