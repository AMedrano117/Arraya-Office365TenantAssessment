"""Results assembly, output validation, delivery manifest and writing."""
from __future__ import annotations

from collections import Counter
from io import BytesIO
import json
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET

from docx import Document
from openpyxl import load_workbook

from . import GENERATOR_VERSION, RESULTS_SCHEMA_VERSION, REVIEW_VERSION
from .engine import TREATMENTS, run_review
from .package import import_package
from .plain import DECISION_LABEL, GROUP_ORDER, order_key
from .details import collect
from .privacy import scan_secrets, scan_text
from .render import (ACTION_HEADERS, BRIEF, DELIVERABLES, DETAIL_HEADERS, EXEC, PLAN, TECH, TRACE_SHEET, render_all,
                     top_recommendations)
from .util import ReviewError, bytes_sha256, canonical_json, file_sha256

RESULTS = "review-results.json"
MANIFEST = "delivery-manifest.json"
AUDIENCE = {
    EXEC: ("Leadership and business sponsors", True, "Readiness decision, blockers, gaps and immediate actions"),
    TECH: ("Senior engineers (customer IT on request)", True, "Control-by-control review and evidence qualification"),
    PLAN: ("Customer responsible roles (confidential: names tenant users, devices, sites and apps)", True,
           "Filterable action plan with closure evidence and named lists of affected objects"),
    BRIEF: ("Review meeting participants", True, "Agenda, decisions required and evidence requests"),
    RESULTS: ("Arraya engineering (machine-readable)", False, "Traceable review record"),
}
SUMMARY_SHEETS = ("Start Here", "Action Plan", "Decisions", TRACE_SHEET, "Microsoft guidance")


PLAIN_LIMITS = {
    "PKG-SOURCE-PARTIAL": "Some data could not be fully collected, so conclusions in those areas are limited.",
    "PKG-SOURCE-FAILED": "Some required data could not be collected; those areas are treated as unknown, not as "
                         "passed or failed.",
    "PKG-SOURCE-TRUNCATED": "Some data sources returned only part of their records.",
    "PKG-EVIDENCE-TRUNCATED": "Some evidence records are incomplete.",
    "PKG-STALE": "Some evidence is older than the freshness limit and must be refreshed before decisions rely on it.",
    "PKG-IDENTITY-UNRESOLVED": "Some account counts may include accounts the tool could not fully identify.",
    "PKG-PRIV-POPULATION": "The list of administrator accounts may be incomplete.",
    "PKG-INTERNAL-CONSISTENCY": "The original assessment's outputs contain internal inconsistencies; this review lists "
                                "them for correction.",
    "PKG-DETAIL-LINKS": "Some links in the original assessment report do not open the intended detail.",
    "PKG-SOURCE-INTEGRITY-CHECKS": "The original assessment recorded a failed internal data check.",
    "PKG-DATE-MALFORMED": "Some dates in the evidence are unreadable and were not used.",
    "PKG-DATE-SENTINEL": "Some dates in the evidence are placeholders and were not used.",
    "PKG-DATE-FUTURE": "Some dates in the evidence are after the evaluation date and were not used.",
    "PKG-EVIDENCE-LINKS": "Some findings point to evidence that is not in the package.",
}
# Limits about the original assessment's own outputs matter to engineers, not to leadership.
ENGINEERING_LIMITS = ("PKG-INTERNAL-CONSISTENCY", "PKG-DETAIL-LINKS", "PKG-SOURCE-INTEGRITY-CHECKS",
                      "PKG-EVIDENCE-LINKS")


def _open_questions(items, pkg):
    questions = []
    for i in items:
        if i["EvidenceSufficiency"] == "Contradictory":
            questions.append(f"{i['FindingId']}: sources disagree; confirm which record reflects the tenant.")
    if any(f["HashStatus"] == "NotRecorded" and f["Role"].endswith(("workbook", "html")) for f in pkg.files):
        questions.append("Script deliverables have no recorded hashes; their integrity relies on the hashes computed "
                         "at review time.")
    return sorted(set(questions))


def build_results(pkg, ctx, items, controls, readiness, review_date):
    benchmarks = ctx.benchmarks
    clean = ctx.clean
    notable = [{"SourceId": k, "Classification": v["Classification"], "Reason": clean(v["Reason"])}
               for k, v in pkg.sources.items() if v["Classification"] not in ("Complete", "CompleteEmpty")]
    complete_empty = [k for k, v in pkg.sources.items() if v["Classification"] == "CompleteEmpty"]
    limitations, customer_limits = [], []
    for check in pkg.checks:
        if check.severity in ("Qualification", "Blocking") and check.status != "Pass":
            plain = PLAIN_LIMITS.get(check.check_id) or clean(check.detail)
            if plain not in limitations:
                limitations.append(plain)
                if check.check_id not in ENGINEERING_LIMITS:
                    customer_limits.append(plain)
    behavior = ("Settings show how controls are configured, not how they behave for users; tests are listed where "
                "needed.")
    limitations.append(behavior)
    customer_limits.append(behavior)
    immediate = [i for i in items if i["ReviewDecision"] != "DuplicateOrOverlapping"
                 and i["ReadinessTreatment"] in ("BlocksPilot", "ConfirmBeforePilot")]
    immediate.sort(key=lambda i: (TREATMENTS.index(i["ReadinessTreatment"]),
                                  {"High": 0, "Medium": 1, "Low": 2}.get(i["ReviewPriority"], 3),
                                  i["DependencyStep"], i["FindingId"]))
    threshold = pkg.windows.get("FreshnessThresholdDays") or 35
    criteria = [
        f"Every Check before rollout item is done, with dated proof from the last {threshold} days, or has an "
        "approved exception that records the owner, scope, compensating control and expiry date.",
        "The out-of-date or missing data listed in this review is refreshed, and the assessment is run again.",
        "The decisions listed in the meeting brief are recorded, with the role that made each one.",
    ]
    if readiness["ConfirmedBlockers"]:
        criteria.insert(0, "Every Fix before rollout item is fixed and the fix is confirmed in the tenant.")
    actionable = sorted((i for i in items if i["ReviewDecision"] != "DuplicateOrOverlapping"
                         and i["ReadinessTreatment"] != "NoAction"), key=order_key)
    for i in items:
        i["ActionNumber"] = None
    for number, i in enumerate(actionable, 1):
        i["ActionNumber"] = number
    decisions = Counter(i["ReviewDecision"] for i in items)
    treatments = Counter(i["ReadinessTreatment"] for i in items)
    groups = Counter(i["Group"] for i in actionable)
    files = [{k: f[k] for k in ("Role", "RelativePath", "Required", "Present", "SHA256", "Bytes",
                                "RecordedSHA256", "RecordedBy", "HashStatus")} for f in pkg.files]
    windows = {k: pkg.windows.get(k) for k in ("EvaluationDate", "ReviewDate", "EvidencePeriod", "Thresholds",
                                                "CollectedAt", "OldestFindingObservation", "OldestFindingAgeDays",
                                                "FreshnessThresholdDays", "DateIssues", "BrokenEvidenceReferences",
                                                "BrokenDetailLinks")}
    identity = {k: pkg.identity.get(k) for k in ("AssessmentId", "RunId", "EnvironmentId", "RunType", "EvaluationDate",
                                                  "CreatedAt", "MethodologyVersion", "CatalogVersion", "ProducingCommit",
                                                  "ProducingSourceState", "BaselineRunId")}
    results = {
        "SchemaVersion": RESULTS_SCHEMA_VERSION, "ReviewVersion": REVIEW_VERSION,
        "BenchmarkVersion": benchmarks["BenchmarkVersion"], "GeneratorVersion": GENERATOR_VERSION,
        "AssessmentId": identity["AssessmentId"], "RunId": identity["RunId"],
        "CustomerName": pkg.identity.get("CustomerName") or "Customer", "ReviewDate": review_date.isoformat(),
        "Identity": identity,
        "Publication": {"CustomerReady": pkg.customer_ready,
                        "Reasons": [clean(c.detail) for c in pkg.blocking] or ["Package integrity established."]},
        "Package": {"Files": files, "Checks": [c.as_dict() for c in pkg.checks],
                    "SourceClassCounts": pkg.windows.get("SourceClassCounts"), "NotableSources": notable,
                    "CompleteEmptySources": complete_empty, "Windows": windows},
        "Readiness": readiness, "Controls": controls, "Items": items,
        "Conflicts": [{**c, "Detail": clean(c["Detail"])} for c in pkg.conflicts],
        "OpenQuestions": _open_questions(items, pkg), "Limitations": limitations,
        "CustomerLimitations": customer_limits,
        "ImmediateActions": immediate[:10], "ReassessmentCriteria": criteria,
        "Summary": {"ItemCount": len(items), "ScriptFindings": sum(1 for i in items if i["ItemType"] == "ScriptFinding"),
                    "AddedFindings": sum(1 for i in items if i["ItemType"] == "AddedFinding"),
                    "ByDecision": dict(sorted(decisions.items())), "ByTreatment": dict(sorted(treatments.items())),
                    "ActionCount": len(actionable),
                    "ByGroup": {g: groups[g] for g in GROUP_ORDER if groups.get(g)},
                    "CustomerDecisions": sum(1 for i in actionable if i["CustomerDecisionRequired"])},
        "Benchmarks": benchmarks["Benchmarks"],
        "StatementTypes": ["Collector fact", "Script-derived conclusion", "Microsoft-documented guidance",
                           "Reviewer interpretation", "Suggested hardening practice", "Customer policy decision",
                           "Evidence gap"],
    }
    # ImmediateActions keeps only references to avoid duplicated content.
    results["ImmediateActions"] = [{"FindingId": i["FindingId"], "RecommendedAction": i["RecommendedAction"],
                                    "ResponsibleRole": i["ResponsibleRole"]} for i in immediate[:10]]
    return results


# ------------------------------------------------------------------ validation

def docx_text(data):
    with zipfile.ZipFile(BytesIO(data)) as z:
        for name in z.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                ET.fromstring(z.read(name))      # every XML part must parse
        xml = z.read("word/document.xml").decode("utf-8")
    Document(BytesIO(data))                       # must reopen with python-docx
    return re.sub(r"<[^>]+>", " ", re.sub(r"</w:p>", "\n", xml))


def xlsx_rows(data):
    with zipfile.ZipFile(BytesIO(data)) as z:
        for name in z.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                ET.fromstring(z.read(name))
    book = load_workbook(BytesIO(data))
    sheets = {ws.title: [[c for c in row] for row in ws.iter_rows(values_only=True)] for ws in book.worksheets}
    tables = {ws.title: list(ws.tables.keys()) for ws in book.worksheets}
    return sheets, tables


def validate_outputs(outputs, results, scan_terms):
    problems = []
    require_names = set(DELIVERABLES) | {RESULTS}
    if set(outputs) != require_names:
        problems.append(f"Unexpected output set: {sorted(set(outputs) ^ require_names)}")
    texts = {}
    for name in (EXEC, TECH, BRIEF):
        try:
            texts[name] = docx_text(outputs[name])
        except Exception as exc:
            problems.append(f"{name} does not open: {type(exc).__name__}")
    try:
        sheets, tables = xlsx_rows(outputs[PLAN])
    except Exception as exc:
        problems.append(f"{PLAN} does not open: {type(exc).__name__}")
        sheets, tables = {}, {}
    plan = sheets.get("Action Plan") or []
    if not plan or tuple(plan[0]) != ACTION_HEADERS:
        problems.append("Action Plan headers are missing or out of order.")
    if "ActionPlan" not in tables.get("Action Plan", []):
        problems.append("Action Plan is not a filterable Excel table.")
    items = results["Items"]
    ids = [i["FindingId"] for i in items]
    actionable = sorted((i for i in items if i["ActionNumber"]), key=lambda i: i["ActionNumber"])
    ref_col, number_col = ACTION_HEADERS.index("Ref"), ACTION_HEADERS.index("#")
    if [(row[number_col], row[ref_col]) for row in plan[1:]] != [(i["ActionNumber"], i["FindingId"]) for i in actionable]:
        problems.append("Action Plan rows do not reconcile with the numbered actions in review-results.json.")
    trace = sheets.get(TRACE_SHEET) or []
    if not trace or tuple(trace[0]) != DETAIL_HEADERS:
        problems.append("Review trace headers are missing or out of order.")
    trace_ref, trace_result = DETAIL_HEADERS.index("Ref"), DETAIL_HEADERS.index("Review result")
    if sorted(row[trace_ref] for row in trace[1:]) != sorted(ids):
        problems.append("Review trace item IDs do not reconcile with review-results.json.")
    expected = Counter({DECISION_LABEL[k]: v for k, v in results["Summary"]["ByDecision"].items()})
    if Counter(row[trace_result] for row in trace[1:]) != expected:
        problems.append("Review trace decision counts do not reconcile with review-results.json.")
    decision = results["Readiness"]["ReviewDecision"]
    for name, body in texts.items():
        if decision not in body:
            problems.append(f"{name} does not state the review readiness decision.")

    def joined(names):
        return " ".join(str(c) for n in names for row in sheets.get(n, []) for c in row if c is not None)
    # Summary sheets follow the same privacy rules as the Word documents; the list sheets name tenant objects by
    # design and are checked only for credentials and payloads.
    summary_names = [n for n in sheets if n in SUMMARY_SHEETS]
    workbook_text = joined(summary_names)
    list_text = joined([n for n in sheets if n not in SUMMARY_SHEETS])
    for issue in scan_secrets(list_text):
        problems.append(f"{PLAN} list sheets: {issue}")
    if decision not in workbook_text:
        problems.append("Action Plan does not state the review readiness decision.")
    if TECH in texts:
        missing = [i for i in ids if i not in texts[TECH]]
        if missing:
            problems.append(f"Technical review omits items: {missing[:5]}")
    gated = [i for i in actionable if i["ReadinessTreatment"] in ("BlocksPilot", "ConfirmBeforePilot")]
    for name, body in ((BRIEF, texts.get(BRIEF)), (PLAN, workbook_text)):
        if body is not None:
            missing = [i["FindingId"] for i in gated if i["Plain"]["Headline"] not in body]
            if missing:
                problems.append(f"{name} omits rollout-gating items: {missing}")
    if EXEC in texts:
        shown = top_recommendations(results)
        missing = [i["FindingId"] for i in actionable if i["ReadinessTreatment"] == "BlocksPilot" and i not in shown]
        missing += [i["FindingId"] for i in shown if i["Plain"]["Headline"] not in texts[EXEC]]
        if missing:
            problems.append(f"{EXEC} omits blockers or top recommendations: {missing}")
    qual_phrase = "ENGINEERING REVIEW ONLY" if not results["Publication"]["CustomerReady"] else "Prepared for operator review"
    for name, body in list(texts.items()) + [(PLAN, workbook_text)]:
        if qual_phrase not in body:
            problems.append(f"{name} lost the engineering qualification banner.")
    restricted = {r["EvidenceId"] for i in items for r in i["EvidenceReferences"] if r["Classification"] == "Restricted"}
    for name, body in ((EXEC, texts.get(EXEC, "")), (BRIEF, texts.get(BRIEF, "")), (PLAN, workbook_text)):
        leaked = sorted(ev for ev in restricted if ev in body)
        if leaked:
            problems.append(f"{name} exposes restricted evidence references: {leaked[:3]}")
    results_text = outputs.get(RESULTS, b"").decode("utf-8")
    for name, body in list(texts.items()) + [(PLAN, workbook_text), (RESULTS, results_text)]:
        for issue in scan_text(body, scan_terms, allow_terms=[results["CustomerName"]]):
            problems.append(f"{name} privacy: {issue}")
    for name, body in list(texts.items()) + [(PLAN, workbook_text)]:
        if re.search(r"(?i)arraya (will|shall|commits to) (implement|remediate|perform|deliver)", body):
            problems.append(f"{name} implies an Arraya implementation commitment.")
    return sorted(set(problems))


# ------------------------------------------------------------------ orchestration

def produce(package_dir, review_date, benchmarks=None):
    """Import, review, render and validate. Returns (outputs, results, context, package)."""
    pkg = import_package(package_dir, review_date)
    ctx, items, controls, readiness = run_review(pkg, review_date, benchmarks)
    results = build_results(pkg, ctx, items, controls, readiness, review_date)
    # Named lists go only into the workbook; results keep only the sheet name and counts per item.
    sheets = collect(pkg, results["Items"], review_date)
    outputs = render_all(results, sheets)
    outputs[RESULTS] = canonical_json(results)
    # Determinism: a second independent review and render must be byte-identical.
    ctx2, items2, controls2, readiness2 = run_review(pkg, review_date, benchmarks)
    results2 = build_results(pkg, ctx2, items2, controls2, readiness2, review_date)
    outputs2 = render_all(results2, collect(pkg, results2["Items"], review_date))
    outputs2[RESULTS] = canonical_json(results2)
    problems = [f"Non-deterministic output: {n}" for n in outputs if outputs[n] != outputs2.get(n)]
    problems += validate_outputs(outputs, results, ctx.scan_terms)
    return outputs, results, ctx, pkg, problems


def build_manifest(outputs, results, pkg):
    deliverables = []
    for name in DELIVERABLES + (RESULTS,):
        audience, customer_facing, purpose = AUDIENCE[name]
        deliverables.append({"File": name, "Audience": audience, "OperatorReceives": True,
                             "CustomerFacing": customer_facing, "Purpose": purpose,
                             "SHA256": bytes_sha256(outputs[name]), "Bytes": len(outputs[name])})
    manifest = {
        "ManifestVersion": "1.0.0", "AssessmentId": results["AssessmentId"], "RunId": results["RunId"],
        "CustomerName": results["CustomerName"], "ReviewVersion": results["ReviewVersion"],
        "BenchmarkVersion": results["BenchmarkVersion"], "GeneratorVersion": results["GeneratorVersion"],
        "ReviewDate": results["ReviewDate"], "ReadinessDecision": results["Readiness"]["ReviewDecision"],
        "Publication": results["Publication"],
        "OperatorInstructions": ("Review all four documents before release. Customer-facing documents may be shared "
                                 "only when Publication.CustomerReady is true and an Arraya engineer has reviewed them. "
                                 "review-results.json is the engineering record and is not sent to the customer."),
        "Deliverables": deliverables,
        "Inputs": [{"Role": f["Role"], "RelativePath": f["RelativePath"], "SHA256": f["SHA256"],
                    "HashStatus": f["HashStatus"]} for f in pkg.files if f["Present"]],
    }
    manifest["ManifestContentSHA256"] = bytes_sha256(canonical_json(manifest))
    return manifest


def write_outputs(outputs, manifest, out_dir):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    allowed = set(outputs) | {MANIFEST}
    unexpected = [p.name for p in out_dir.iterdir() if p.name not in allowed]
    if unexpected:
        raise ReviewError(f"Output folder contains unexpected files {unexpected}; refusing to mix outputs.")
    for name, data in outputs.items():
        (out_dir / name).write_bytes(data)
    (out_dir / MANIFEST).write_bytes(canonical_json(manifest))
    mismatched = [d["File"] for d in manifest["Deliverables"] if file_sha256(out_dir / d["File"]) != d["SHA256"]]
    if mismatched:
        raise ReviewError(f"Delivery manifest hashes do not match written files: {mismatched}")
    check = json.loads((out_dir / MANIFEST).read_text(encoding="utf-8"))
    stored = check.pop("ManifestContentSHA256")
    if bytes_sha256(canonical_json(check)) != stored:
        raise ReviewError("Delivery manifest self-hash does not verify.")
    return out_dir


def verify_input_unchanged(pkg):
    changed = [f["RelativePath"] for f in pkg.files
               if f["Present"] and f["SHA256"] and file_sha256(pkg.root / f["RelativePath"]) != f["SHA256"]]
    if changed:
        raise ReviewError(f"Package inputs changed during review: {changed}")
