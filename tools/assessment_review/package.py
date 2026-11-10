"""Import and validate one completed assessment package (read-only).

The importer never writes to the package. It establishes identity, integrity,
source classification, observation windows, and internal consistency, and
decides whether customer-ready publication is permitted.
"""
from __future__ import annotations

from dataclasses import dataclass, field
import html as html_lib
from pathlib import Path
import re
from urllib.parse import unquote

from openpyxl import load_workbook

from .util import (ReviewError, age_days, file_sha256, load_json,
                   parse_date, require)

SOURCE_CLASSES = ("Complete", "CompleteEmpty", "Partial", "Failed", "Unavailable",
                  "NotRequested", "NotApplicable", "Unknown")

# Sources whose failure undermines a security-gate conclusion.
REQUIRED_SOURCES = (
    "ca_policies", "security_defaults", "auth_methods", "role_assignments",
    "managed_devices", "authorization_policy", "oauth_grants", "consent_policies",
    "risky_users", "defender_incidents", "purview_dlp_policies",
    "purview_sensitivity_labels", "purview_label_policies", "purview_audit_config",
    "purview_retention_policies", "m365_license_coverage")

SNAPSHOT_KEYS = (
    "schema_version", "methodology_version", "evaluation_date", "report_completeness",
    "evidence_completeness", "tenant_id", "decision", "rationale", "customer_findings",
    "actions", "blockers", "pilot_conditions", "controls", "control_results", "counts",
    "evidence", "executive_summary", "rollout_progress", "readiness_review",
    "supplemental_reviews", "device_reconciliation", "privileged_assessment",
    "identity_taxonomy", "collection_source_states", "collection_thresholds",
    "evidence_period", "producing_revision", "identity", "run_context",
    "answered_controls", "required_controls")


@dataclass
class Check:
    check_id: str
    area: str
    status: str          # Pass | Fail | Warning | Info
    severity: str        # Blocking | Qualification | Info
    detail: str

    def as_dict(self):
        return {"CheckId": self.check_id, "Area": self.area, "Status": self.status,
                "Severity": self.severity, "Detail": self.detail}


@dataclass
class Package:
    root: Path
    files: list = field(default_factory=list)
    checks: list = field(default_factory=list)
    identity: dict = field(default_factory=dict)
    sources: dict = field(default_factory=dict)
    windows: dict = field(default_factory=dict)
    snapshot: dict = field(default_factory=dict)
    collection: dict = field(default_factory=dict)
    workbook: dict = field(default_factory=dict)
    technical_workbook: dict = field(default_factory=dict)
    html: dict = field(default_factory=dict)
    optional: dict = field(default_factory=dict)
    conflicts: list = field(default_factory=list)

    def add(self, check_id, area, status, severity, detail):
        self.checks.append(Check(check_id, area, status, severity, detail))

    @property
    def blocking(self):
        return [c for c in self.checks if c.severity == "Blocking" and c.status == "Fail"]

    @property
    def customer_ready(self):
        return not self.blocking

    def conflict(self, conflict_id, area, detail, refs=()):
        self.conflicts.append({"ConflictId": conflict_id, "Area": area, "Detail": detail,
                               "References": sorted(set(refs))})


# --------------------------------------------------------------------------- files

def _record_file(pkg, role, path, required, recorded=None, recorded_source=None):
    rel = path.relative_to(pkg.root).as_posix() if path is not None else None
    entry = {"Role": role, "RelativePath": rel, "Required": required,
             "Present": bool(path is not None and path.is_file()), "SHA256": None,
             "Bytes": None, "RecordedSHA256": recorded, "RecordedBy": recorded_source,
             "HashStatus": "NotRecorded"}
    if entry["Present"]:
        entry["SHA256"] = file_sha256(path)
        entry["Bytes"] = path.stat().st_size
        if recorded:
            entry["HashStatus"] = "Verified" if recorded.lower() == entry["SHA256"] else "Mismatch"
    else:
        entry["HashStatus"] = "Missing"
    pkg.files.append(entry)
    return entry


def _find_deliverables(root):
    folder = root / "Deliverables"
    found = {"customer_workbook": None, "technical_workbook": None,
             "customer_html": None, "readiness_html": None}
    if not folder.is_dir():
        return found
    for item in sorted(folder.iterdir()):
        name = item.name.lower()
        if item.suffix.lower() == ".xlsx":
            key = "technical_workbook" if name.startswith("technical evidence") else "customer_workbook"
        elif item.suffix.lower() in (".html", ".htm"):
            key = "readiness_html" if name.startswith("readiness summary") else "customer_html"
        else:
            continue
        if found[key] is None:
            found[key] = item
    return found


# ------------------------------------------------------------------ classification

def classify_source(state=None, complete=None, records=None, truncated=False, reason=""):
    state = (state or "").strip().lower()
    reason = (reason or "").lower()
    if state in ("not_requested", "skipped", "not requested"):
        return "NotRequested"
    if state in ("not_applicable", "not applicable"):
        return "NotApplicable"
    if state in ("failed", "error", "errored"):
        return "Failed"
    if state in ("unavailable", "denied", "access_denied", "forbidden"):
        return "Unavailable"
    if state in ("partial", "incomplete") or truncated:
        return "Partial"
    if state == "empty":
        return "CompleteEmpty" if complete is not False else "Partial"
    if state in ("available", "complete", "collected", "ready"):
        if complete is False:
            return "Partial"
        if isinstance(records, int) and records == 0:
            return "CompleteEmpty"
        return "Complete"
    if "denied" in reason or "permission" in reason:
        return "Unavailable"
    return "Unknown"


def _classify_sources(pkg):
    states = pkg.snapshot.get("collection_source_states") or {}
    sources = {}
    for key, value in states.items():
        if ":" in key:            # per-object enumeration records, summarized below
            continue
        value = value if isinstance(value, dict) else {}
        records = value.get("records_collected")
        records = records if isinstance(records, int) else None
        sources[key] = {
            "SourceId": key, "State": value.get("state"),
            "Classification": classify_source(value.get("state"), value.get("complete"),
                                              records, value.get("truncated", False),
                                              value.get("reason", "")),
            "Records": records, "Complete": value.get("complete"),
            "Reason": value.get("reason") or "", "Truncated": bool(value.get("truncated"))}
    # Collector-side status carries truncation and record counts the snapshot may lack.
    for service in ("entra_info", "purview_info", "defender_info"):
        attrs = _attrs(pkg.collection, service)
        for key, value in (attrs.get("collection_status") or {}).items():
            if ":" in key or not isinstance(value, dict):
                continue
            source_id = key if service == "entra_info" else (
                ("purview_" if service == "purview_info" else "defender_") + key)
            truncated = bool(value.get("truncated"))
            records = value.get("records_collected")
            records = records if isinstance(records, int) else None
            state = value.get("availability_status") or ("available" if value.get("available") else "unavailable")
            entry = sources.get(source_id)
            classification = classify_source(state, value.get("complete"), records, truncated,
                                             value.get("reason", ""))
            if entry is None:
                sources[source_id] = {"SourceId": source_id, "State": state,
                                      "Classification": classification, "Records": records,
                                      "Complete": value.get("complete"),
                                      "Reason": value.get("reason") or "", "Truncated": truncated}
            elif truncated and entry["Classification"] in ("Complete", "CompleteEmpty"):
                entry.update({"Classification": "Partial", "Truncated": True})
    pkg.sources = dict(sorted(sources.items()))


def _attrs(collection, service):
    info = (collection.get("service_results") or {}).get(service) or {}
    client = info.get("_client") or {}
    return client.get("$collection_attributes") or {}


# ------------------------------------------------------------------ workbooks/html

def _read_customer_workbook(path):
    result = {"sheets": [], "action_plan": [], "control_results": [], "start_here": {},
              "integrity_failed": 0, "read_error": None, "finding_titles": {}}
    try:
        book = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:  # pragma: no cover - surfaced as a check
        result["read_error"] = f"{type(exc).__name__}"
        return result
    try:
        result["sheets"] = list(book.sheetnames)
        result["action_plan"] = _sheet_rows(book, "Action Plan")
        result["control_results"] = _sheet_rows(book, "Control Results")
        for row in _sheet_rows(book, "Start Here", header=False):
            if len(row) >= 2 and row[0]:
                result["start_here"][str(row[0])] = row[1]
        for row in _sheet_rows(book, "Findings"):
            if row.get("ID") and row.get("Finding"):
                result["finding_titles"].setdefault(str(row["ID"]), str(row["Finding"]))
        for row in result["action_plan"]:
            if row.get("RecommendationId") and row.get("Finding"):
                result["finding_titles"][str(row["RecommendationId"])] = str(row["Finding"])
        for row in _sheet_rows(book, "Integrity Checks"):
            if str(row.get("Status", "")).lower() == "failed":
                result["integrity_failed"] += 1
    finally:
        book.close()
    return result


def _read_technical_workbook(path):
    result = {"sheets": [], "control_results": [], "read_error": None}
    try:
        book = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:  # pragma: no cover
        result["read_error"] = f"{type(exc).__name__}"
        return result
    try:
        result["sheets"] = list(book.sheetnames)
        result["control_results"] = _sheet_rows(book, "Control Results")
    finally:
        book.close()
    return result


def _sheet_rows(book, name, header=True):
    if name not in book.sheetnames:
        return []
    rows = list(book[name].iter_rows(values_only=True))
    if not header:
        return [list(r) for r in rows]
    if not rows:
        return []
    keys = [str(k) if k is not None else f"col{i}" for i, k in enumerate(rows[0])]
    out = []
    for row in rows[1:]:
        if row is None or all(v is None for v in row):
            continue
        out.append({keys[i]: row[i] for i in range(min(len(keys), len(row)))})
    return out


def _read_html(path):
    raw = path.read_text(encoding="utf-8", errors="replace")
    ids = set(re.findall(r"""\bid=["']([^"']+)["']""", raw))
    hrefs = re.findall(r"""\bhref=["']([^"']+)["']""", raw)
    body = re.sub(r"<script.*?</script>|<style.*?</style>", " ", raw, flags=re.S)
    plain = html_lib.unescape(re.sub(r"<[^>]+>", "\n", body))
    lines = [line.strip() for line in plain.splitlines() if line.strip()]
    return {"ids": ids, "hrefs": hrefs, "lines": lines}


# ------------------------------------------------------------------ main import

def import_package(root, review_date=None):
    root = Path(root).resolve()
    require(root.is_dir(), f"Package directory not found: {root}")
    pkg = Package(root=root)

    run_path = root / "assessment-run.json"
    coll_path = root / "collection.json"
    _record_file(pkg, "assessment-run", run_path if run_path.exists() else None, True)
    if not run_path.exists():
        pkg.add("PKG-FILE-RUN", "Files", "Fail", "Blocking", "assessment-run.json is missing.")
        raise_package(pkg, "Required file assessment-run.json is missing.")
    run = load_json(run_path)
    ident = run.get("Identity") or {}
    run_id = ident.get("RunId")
    locator = run.get("SnapshotLocator") or (f"assessment-results/{run_id}.json" if run_id else None)
    snap_path = root / locator if locator else None
    _record_file(pkg, "assessment-snapshot", snap_path if snap_path and snap_path.exists() else None,
                 True, run.get("SnapshotHash"), "assessment-run.json SnapshotHash")
    _record_file(pkg, "collection", coll_path if coll_path.exists() else None, True)

    deliverables = _find_deliverables(root)
    for role, required in (("customer_workbook", True), ("technical_workbook", True),
                           ("customer_html", True), ("readiness_html", False)):
        _record_file(pkg, role, deliverables[role], required)

    for name, role in (("run-seed.json", "run-seed"), ("assessment-history.json", "history"),
                       ("operator-log.jsonl", "operator-log")):
        path = root / name
        if path.exists():
            _record_file(pkg, role, path, False)
    if run_id and (root / "run-seeds" / f"{run_id}.json").exists():
        _record_file(pkg, "run-seed-archive", root / "run-seeds" / f"{run_id}.json", False)
    for opt in ("agent-customer-review.json", "agent-knowledge.md", "agent-evidence-index.json",
                "agent-demo-manifest.json"):
        for base in (root, root / "Deliverables"):
            if (base / opt).exists():
                _record_file(pkg, "optional:" + opt, base / opt, False)
                pkg.optional[opt] = base / opt
                break

    missing = [f for f in pkg.files if f["Required"] and not f["Present"]]
    for entry in missing:
        pkg.add("PKG-FILE-MISSING", "Files", "Fail", "Blocking",
                f"Required input missing: {entry['Role']}.")
    if any(f["Role"] in ("assessment-snapshot", "collection") for f in missing):
        raise_package(pkg, "Required JSON inputs are missing; review cannot proceed.")

    snapshot_full = load_json(snap_path)
    pkg.snapshot = {k: snapshot_full.get(k) for k in SNAPSHOT_KEYS}
    validation = snapshot_full.get("identity_validation") or []
    pkg.snapshot["_unresolved_identity_count"] = sum(
        1 for v in validation if isinstance(v, dict) and v.get("code") == "unresolved_alias")
    del snapshot_full, validation
    pkg.collection = load_json(coll_path)

    _check_hashes(pkg, run)
    _check_identity(pkg, run, ident, review_date)
    _classify_sources(pkg)
    _check_sources(pkg)
    _check_dates_and_windows(pkg, review_date)
    _check_evidence_links(pkg)
    _check_population_contradictions(pkg)

    if deliverables["customer_workbook"]:
        pkg.workbook = _read_customer_workbook(deliverables["customer_workbook"])
    if deliverables["technical_workbook"]:
        pkg.technical_workbook = _read_technical_workbook(deliverables["technical_workbook"])
    for role in ("customer_html", "readiness_html"):
        if deliverables[role]:
            pkg.html[role] = _read_html(deliverables[role])
            pkg.html[role]["name"] = deliverables[role].name
    _check_deliverable_consistency(pkg, deliverables)
    _check_publication(pkg)
    return pkg


def raise_package(pkg, message):
    err = ReviewError(message)
    err.package = pkg
    raise err


def _check_hashes(pkg, run):
    for entry in pkg.files:
        if entry["HashStatus"] == "Mismatch":
            pkg.add("PKG-HASH-MISMATCH", "Integrity", "Fail", "Blocking",
                    f"{entry['Role']} SHA-256 does not match the value recorded by {entry['RecordedBy']}.")
        elif entry["HashStatus"] == "Verified":
            pkg.add("PKG-HASH-OK", "Integrity", "Pass", "Info",
                    f"{entry['Role']} SHA-256 matches {entry['RecordedBy']}.")
    files = ((pkg.collection.get("package") or {}).get("files")) or []
    for item in files:
        rel = item.get("path")
        recorded = item.get("sha256")
        if not rel:
            continue
        path = pkg.root / rel
        entry = _record_file(pkg, "collected-input", path if path.exists() else None, False,
                             recorded, "collection.json package manifest")
        entry["RelativePath"] = rel
        if not entry["Present"]:
            pkg.add("PKG-INPUT-MISSING", "Integrity", "Fail", "Blocking",
                    f"Collected input listed in the package manifest is missing: {Path(rel).name}.")
        elif entry["HashStatus"] == "Mismatch":
            pkg.add("PKG-HASH-MISMATCH", "Integrity", "Fail", "Blocking",
                    f"Collected input hash mismatch: {Path(rel).name}.")
    verified = sum(1 for f in pkg.files if f["Role"] == "collected-input" and f["HashStatus"] == "Verified")
    if files:
        pkg.add("PKG-INPUT-HASHES", "Integrity", "Pass" if verified == len(files) else "Fail",
                "Info" if verified == len(files) else "Blocking",
                f"{verified} of {len(files)} collected input files match recorded hashes.")
    for entry in pkg.files:
        if entry["Role"] in ("customer_workbook", "technical_workbook", "customer_html",
                             "readiness_html") and entry["Present"]:
            pkg.add("PKG-DELIVERABLE-HASH", "Integrity", "Info", "Info",
                    f"{entry['Role']}: no recorded hash exists in the package; the computed hash is "
                    "retained in the delivery manifest.")


def _check_identity(pkg, run, ident, review_date):
    snap = pkg.snapshot
    coll = pkg.collection
    producing = run.get("ProducingRevision") or {}
    pkg.identity = {
        "AssessmentId": ident.get("AssessmentId"), "RunId": ident.get("RunId"),
        "EnvironmentId": ident.get("PrimaryEnvironmentId"), "RunType": ident.get("RunType"),
        "EvaluationDate": ident.get("EvaluatedAt"), "CreatedAt": ident.get("CreatedAt"),
        "MethodologyVersion": ident.get("MethodologyVersion"),
        "CatalogVersion": ident.get("CatalogVersion"),
        "ProducingCommit": producing.get("commit_sha"),
        "ProducingSourceState": producing.get("source_state"),
        "ProducingRepository": producing.get("repository_identifier"),
        "PublicationEligible": ((producing.get("publication") or {}).get("customer_publication_eligible")),
        "TenantId": coll.get("tenant_id"), "CustomerName": coll.get("tenant_name") or coll.get("customer_name"),
        "BaselineRunId": ident.get("BaselineRunId"),
    }
    require_ids = ("AssessmentId", "RunId", "EnvironmentId")
    for key in require_ids:
        if not pkg.identity.get(key):
            pkg.add("PKG-ID-MISSING", "Identity", "Fail", "Blocking", f"{key} is not recorded in assessment-run.json.")
    sources = {
        "collection.json": coll.get("identity") or {},
        "snapshot identity": snap.get("identity") or {},
        "snapshot run context": snap.get("run_context") or {},
    }
    root = pkg.root
    for name, path in (("run-seed.json", root / "run-seed.json"),
                       ("run-seeds archive", root / "run-seeds" / f"{pkg.identity['RunId']}.json")):
        if path.exists():
            try:
                sources[name] = load_json(path)
            except (ValueError, OSError):
                pkg.add("PKG-JSON", "Identity", "Fail", "Qualification", f"{name} is not valid JSON.")
    mismatches = []
    for name, data in sources.items():
        for key, other in (("AssessmentId", "AssessmentId"), ("RunId", "RunId"),
                           ("EnvironmentId", "PrimaryEnvironmentId")):
            value = data.get(other)
            if value and pkg.identity.get(key) and value != pkg.identity[key]:
                mismatches.append(f"{name} {other}")
    snap_name = Path(run.get("SnapshotLocator") or "").stem
    if snap_name and pkg.identity.get("RunId") and snap_name != pkg.identity["RunId"]:
        mismatches.append("snapshot file name RunId")
    if snap.get("tenant_id") and coll.get("tenant_id") and snap["tenant_id"] != coll["tenant_id"]:
        mismatches.append("snapshot tenant_id")
    history_path = root / "assessment-history.json"
    if history_path.exists():
        try:
            history = load_json(history_path)
            if history.get("AssessmentId") and history["AssessmentId"] != pkg.identity.get("AssessmentId"):
                mismatches.append("assessment-history AssessmentId")
            runs = [r.get("RunId") for r in history.get("Runs") or []]
            if runs and pkg.identity.get("RunId") not in runs:
                mismatches.append("assessment-history run list")
            for r in history.get("Runs") or []:
                if r.get("RunId") == pkg.identity.get("RunId") and r.get("SnapshotHash") and \
                        r["SnapshotHash"] != run.get("SnapshotHash"):
                    mismatches.append("assessment-history SnapshotHash")
        except (ValueError, OSError):
            pkg.add("PKG-JSON", "Identity", "Fail", "Qualification", "assessment-history.json is not valid JSON.")
    ev_mismatch = sum(1 for e in snap.get("evidence") or []
                      if (e.get("AssessmentId") and e["AssessmentId"] != pkg.identity.get("AssessmentId"))
                      or (e.get("RunId") and e["RunId"] != pkg.identity.get("RunId")))
    if ev_mismatch:
        mismatches.append(f"{ev_mismatch} evidence records")
    if mismatches:
        pkg.add("PKG-ID-MISMATCH", "Identity", "Fail", "Blocking",
                "Assessment, run, environment or tenant identity disagrees in: " + ", ".join(sorted(mismatches)) + ".")
    else:
        pkg.add("PKG-ID-OK", "Identity", "Pass", "Info",
                "AssessmentId, RunId, EnvironmentId and tenant agree across run record, collection, snapshot, "
                "seeds, history and evidence records.")
    if pkg.identity.get("PublicationEligible") is False:
        pkg.add("PKG-PRODUCER", "Identity", "Fail", "Blocking",
                "The producing revision is not eligible for customer publication.")


def _check_sources(pkg):
    counts = {c: 0 for c in SOURCE_CLASSES}
    for entry in pkg.sources.values():
        counts[entry["Classification"]] += 1
    pkg.windows["SourceClassCounts"] = counts
    failed_required = [s for s in REQUIRED_SOURCES
                       if pkg.sources.get(s, {}).get("Classification") in ("Failed", "Unavailable")]
    for src in failed_required:
        pkg.add("PKG-SOURCE-FAILED", "Collection", "Fail", "Qualification",
                f"Required source {src} is {pkg.sources[src]['Classification']}; dependent conclusions are "
                "treated as evidence gaps, not failures or passes.")
    partial = sorted(s for s, v in pkg.sources.items() if v["Classification"] == "Partial")
    if partial:
        pkg.add("PKG-SOURCE-PARTIAL", "Collection", "Warning", "Qualification",
                "Partial or truncated sources: " + ", ".join(partial) + ".")
    truncated = sorted(s for s, v in pkg.sources.items() if v.get("Truncated"))
    if truncated:
        pkg.add("PKG-SOURCE-TRUNCATED", "Collection", "Warning", "Qualification",
                "Truncated sources: " + ", ".join(truncated) + ".")
    trunc_ev = sum(1 for e in pkg.snapshot.get("evidence") or [] if e.get("truncated"))
    if trunc_ev:
        pkg.add("PKG-EVIDENCE-TRUNCATED", "Collection", "Warning", "Qualification",
                f"{trunc_ev} evidence records are marked truncated.")
    if pkg.snapshot.get("_unresolved_identity_count"):
        pkg.add("PKG-IDENTITY-UNRESOLVED", "Identity", "Warning", "Qualification",
                f"{pkg.snapshot['_unresolved_identity_count']} legacy identity aliases are unresolved; "
                "object counts can include unresolved identities.")
    diags = (pkg.snapshot.get("privileged_assessment") or {}).get("diagnostics") or []
    if diags:
        pkg.add("PKG-PRIV-POPULATION", "Identity", "Warning", "Qualification",
                "Effective privileged population is incomplete (group membership not fully resolved).")


def _check_dates_and_windows(pkg, review_date):
    snap = pkg.snapshot
    eval_date, problem = parse_date(pkg.identity.get("EvaluationDate") or snap.get("evaluation_date"))
    if problem or eval_date is None:
        pkg.add("PKG-DATE-EVAL", "Dates", "Fail", "Blocking", "Evaluation date is missing or malformed.")
    ref = review_date or eval_date
    pkg.windows.update({"EvaluationDate": eval_date.isoformat() if eval_date else None,
                        "ReviewDate": ref.isoformat() if ref else None,
                        "EvidencePeriod": snap.get("evidence_period") or {},
                        "Thresholds": snap.get("collection_thresholds") or {},
                        "CollectedAt": pkg.collection.get("collected_at")})
    malformed, sentinel, future = [], [], []
    for f in snap.get("customer_findings") or []:
        for key in ("ObservationDate", "SourceObservedAt"):
            value = f.get(key)
            parsed, issue = parse_date(value)
            label = f"{f.get('RecommendationId')}.{key}"
            if issue == "sentinel or out-of-range year":
                sentinel.append(label)
            elif issue:
                malformed.append(label)
            elif parsed and eval_date and (parsed - eval_date).days > 1:
                future.append(label)
    for e in snap.get("evidence") or []:
        for key in ("observed_at", "source_observed_at"):
            parsed, issue = parse_date(e.get(key))
            label = f"{e.get('evidence_id')}.{key}"
            if issue == "sentinel or out-of-range year":
                sentinel.append(label)
            elif issue:
                malformed.append(label)
            elif parsed and eval_date and (parsed - eval_date).days > 1:
                future.append(label)
    for name, items in (("PKG-DATE-MALFORMED", malformed), ("PKG-DATE-SENTINEL", sentinel),
                        ("PKG-DATE-FUTURE", future)):
        if items:
            pkg.add(name, "Dates", "Fail", "Qualification",
                    f"{len(items)} date values are {name.split('-')[-1].lower()} and are excluded from "
                    f"freshness calculations (first: {sorted(set(items))[0]}).")
    pkg.windows["DateIssues"] = {"Malformed": sorted(set(malformed)), "Sentinel": sorted(set(sentinel)),
                                 "Future": sorted(set(future))}
    # Oldest material evidence age.
    oldest = None
    for f in snap.get("customer_findings") or []:
        parsed, issue = parse_date(f.get("ObservationDate"))
        if parsed and not issue:
            oldest = parsed if oldest is None or parsed < oldest else oldest
    pkg.windows["OldestFindingObservation"] = oldest.isoformat() if oldest else None
    pkg.windows["OldestFindingAgeDays"] = age_days(oldest, ref)
    threshold = (snap.get("collection_thresholds") or {}).get("review_freshness_days")
    pkg.windows["FreshnessThresholdDays"] = threshold
    if oldest and threshold and ref and age_days(oldest, ref) > threshold:
        pkg.add("PKG-STALE", "Dates", "Warning", "Qualification",
                f"Oldest finding evidence is {age_days(oldest, ref)} days old at the review date, above the "
                f"{threshold}-day freshness threshold.")


def _check_evidence_links(pkg):
    snap = pkg.snapshot
    evidence_ids = {e.get("evidence_id") for e in snap.get("evidence") or []}
    broken = []
    for f in snap.get("customer_findings") or []:
        for ev in [f.get("EvidenceId")] + list(f.get("RelatedEvidenceIds") or []):
            if ev and ev not in evidence_ids:
                broken.append((f.get("RecommendationId"), ev, f.get("Disposition")))
    for c in snap.get("controls") or []:
        for ev in c.get("evidence_ids") or []:
            if ev not in evidence_ids:
                broken.append((c.get("control_id"), ev, "Control"))
    pkg.windows["BrokenEvidenceReferences"] = [{"Owner": o, "EvidenceId": e} for o, e, _ in broken]
    if broken:
        material = [b for b in broken if b[2] in ("Action", "Control")]
        pkg.add("PKG-EVIDENCE-LINKS", "Evidence", "Fail", "Blocking" if material else "Qualification",
                f"{len(broken)} evidence references do not resolve to retained evidence "
                f"({len(material)} on actions or controls).")
    else:
        pkg.add("PKG-EVIDENCE-LINKS", "Evidence", "Pass", "Info",
                "Every finding and control evidence reference resolves to a retained evidence record.")
    recs = {f.get("RecommendationId") for f in snap.get("customer_findings") or []}
    actions = snap.get("actions") or []
    orphan = [a.get("RecommendationId") for a in actions if a.get("RecommendationId") not in recs]
    if orphan:
        pkg.add("PKG-ACTION-LINKS", "Evidence", "Fail", "Blocking",
                f"{len(orphan)} actions do not resolve to a finding.")


def _check_population_contradictions(pkg):
    snap = pkg.snapshot
    counts = snap.get("counts") or {}
    actions = snap.get("actions") or []
    blockers = snap.get("blockers") or []
    rollout = (snap.get("rollout_progress") or {}).get("counts") or {}
    exec_blockers = (snap.get("executive_summary") or {}).get("Pilot blockers") or []
    if counts.get("actions") is not None and counts["actions"] != len(actions):
        pkg.conflict("CNF-COUNT-ACTIONS", "Script internal",
                     f"counts.actions is {counts['actions']} but {len(actions)} action records are retained.")
    if rollout.get("pilot_blockers") is not None and rollout["pilot_blockers"] != len(blockers):
        pkg.conflict("CNF-COUNT-BLOCKERS", "Script internal",
                     f"Rollout progress reports {rollout['pilot_blockers']} pilot blockers while the blocker list "
                     f"and decision rationale contain {len(blockers)}.",
                     [b.get("RecommendationId") for b in blockers])
    if exec_blockers and len(exec_blockers) != len(blockers):
        pkg.conflict("CNF-EXEC-BLOCKERS", "Script internal",
                     f"Executive summary lists {len(exec_blockers)} blockers; blocker list has {len(blockers)}.")
    remediation = sum(1 for a in actions if a.get("ActionType") == "Remediation")
    if counts.get("remediation") is not None and counts["remediation"] != remediation:
        pkg.conflict("CNF-COUNT-REMEDIATION", "Script internal",
                     f"counts.remediation is {counts['remediation']} but {remediation} remediation actions exist.")
    staged = []
    for f in snap.get("customer_findings") or []:
        stages = {o.get("ReadinessStage") for o in f.get("SourceOccurrences") or [] if o.get("ReadinessStage")}
        if f.get("ReadinessStage") and stages and f["ReadinessStage"] not in stages:
            staged.append(f.get("RecommendationId"))
    if staged:
        pkg.conflict("CNF-STAGE", "Script internal",
                     f"{len(staged)} findings carry a readiness stage that differs from the stage recorded on their "
                     "source occurrence (for example Before pilot versus Before broad rollout); the review uses "
                     "its own readiness treatment.", staged)
    # Control identifiers: findings sometimes carry a legacy control alias.
    listed = {}
    for c in snap.get("controls") or []:
        for rid in c.get("recommendation_ids") or []:
            listed.setdefault(rid, set()).add(c.get("control_id"))
    aliased = [f.get("RecommendationId") for f in snap.get("customer_findings") or []
               if f.get("RecommendationId") in listed and f.get("ControlId")
               and f["ControlId"] not in listed[f["RecommendationId"]]]
    if aliased:
        pkg.conflict("CNF-CONTROL-ALIAS", "Script internal",
                     f"{len(aliased)} findings carry a legacy control ID (for example IDENTITY-001) that differs from "
                     "the readiness control that lists them; the review reports the listing control.", aliased)
    if pkg.conflicts:
        pkg.add("PKG-INTERNAL-CONSISTENCY", "Consistency", "Warning", "Qualification",
                f"{len(pkg.conflicts)} internal inconsistencies in script outputs are listed as known conflicts.")


def _norm(value):
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def _check_deliverable_consistency(pkg, deliverables):
    snap = pkg.snapshot
    wb = pkg.workbook or {}
    if wb.get("read_error") or (pkg.technical_workbook or {}).get("read_error"):
        pkg.add("PKG-WORKBOOK-READ", "Deliverables", "Fail", "Blocking", "A required workbook could not be opened.")
    actions = {a.get("RecommendationId"): a for a in snap.get("actions") or []}
    rows = {r.get("RecommendationId"): r for r in wb.get("action_plan") or [] if r.get("RecommendationId")}
    if wb.get("action_plan") is not None and deliverables.get("customer_workbook"):
        if set(rows) != set(actions):
            pkg.conflict("CNF-WB-ACTIONS", "Workbook vs JSON",
                         f"Workbook Action Plan has {len(rows)} actions; JSON has {len(actions)}.",
                         sorted(set(rows) ^ set(actions)))
        for rid in sorted(set(rows) & set(actions)):
            for wkey, jkey in (("Priority", "Priority"), ("Pilot impact", "PilotImpact")):
                if wkey in rows[rid] and _norm(rows[rid][wkey]) != _norm(actions[rid].get(jkey)):
                    pkg.conflict(f"CNF-WB-{rid}-{jkey}", "Workbook vs JSON",
                                 f"{rid} {jkey}: workbook '{rows[rid][wkey]}' vs JSON '{actions[rid].get(jkey)}'.", [rid])
    for source, rows_ in (("customer", wb.get("control_results")),
                          ("technical", (pkg.technical_workbook or {}).get("control_results"))):
        if not rows_:
            continue
        json_rows = {r.get("Control ID"): r for r in snap.get("control_results") or []}
        for row in rows_:
            cid = row.get("Control ID")
            if cid in json_rows and _norm(row.get("Status")) != _norm(json_rows[cid].get("Status")):
                pkg.conflict(f"CNF-WB-CTRL-{source}-{cid}", "Workbook vs JSON",
                             f"{source} workbook control {cid} status '{row.get('Status')}' vs JSON "
                             f"'{json_rows[cid].get('Status')}'.", [cid])
    start = wb.get("start_here") or {}
    tenant = start.get("Tenant")
    if tenant and pkg.identity.get("CustomerName") and _norm(tenant) != _norm(pkg.identity["CustomerName"]):
        pkg.add("PKG-BOUNDARY-CUSTOMER", "Boundary", "Fail", "Blocking",
                "Customer workbook tenant name differs from the collection tenant name.")
    wb_eval = start.get("Evaluation date")
    if wb_eval and pkg.identity.get("EvaluationDate") and str(wb_eval)[:10] != pkg.identity["EvaluationDate"]:
        pkg.add("PKG-BOUNDARY-DATE", "Boundary", "Fail", "Blocking",
                "Customer workbook evaluation date differs from the run record.")
    if wb.get("integrity_failed"):
        pkg.add("PKG-SOURCE-INTEGRITY-CHECKS", "Deliverables", "Warning", "Qualification",
                f"The customer workbook records {wb['integrity_failed']} failed integrity check(s) from the engine.")
    # Detail links: in-page anchors, cross-document anchors, and workbook sheet references.
    broken = []
    sheets = {"customer": set(wb.get("sheets") or []),
              "technical": set((pkg.technical_workbook or {}).get("sheets") or [])}
    html_ids = {role: data["ids"] for role, data in pkg.html.items()}
    names = {role: (deliverables[role].name if deliverables.get(role) else None)
             for role in ("customer_html", "readiness_html", "customer_workbook", "technical_workbook")}
    for role, data in pkg.html.items():
        for href in data["hrefs"]:
            target = unquote(href)
            if target.startswith(("http://", "https://", "mailto:")):
                continue
            file_part, _, frag = target.partition("#")
            if not file_part:
                if frag and frag not in data["ids"]:
                    broken.append(f"{role}#{frag}")
                continue
            matched = [r for r, n in names.items() if n == file_part]
            if not matched:
                broken.append(f"{role}->{file_part}")
                continue
            target_role = matched[0]
            if not frag:
                continue
            if target_role in html_ids:
                if frag not in html_ids[target_role]:
                    broken.append(f"{role}->{file_part}#{frag}")
            else:
                sheet = frag.split("!")[0].strip("'")
                key = "technical" if target_role == "technical_workbook" else "customer"
                if sheet and sheet not in sheets[key]:
                    broken.append(f"{role}->{file_part}#{sheet}")
    for row in wb.get("action_plan") or []:
        ref = str(row.get("Evidence") or "")
        match = re.match(r"^'?([^'!]+)'?!", ref)
        if match and match.group(1) not in sheets["customer"]:
            broken.append(f"Action Plan {row.get('RecommendationId')} -> {match.group(1)}")
    pkg.windows["BrokenDetailLinks"] = sorted(set(broken))
    if broken:
        pkg.conflict("CNF-DETAIL-LINKS", "HTML vs workbook",
                     f"{len(set(broken))} detail links in the script deliverables point to worksheets or anchors that "
                     "do not exist: " + "; ".join(sorted({b.split('#')[-1] for b in broken})) + ".")
        pkg.add("PKG-DETAIL-LINKS", "Deliverables", "Fail", "Qualification",
                f"{len(set(broken))} detail links do not resolve.")
    elif pkg.html:
        pkg.add("PKG-DETAIL-LINKS", "Deliverables", "Pass", "Info",
                "HTML anchors, cross-document links and workbook sheet references resolve.")
    # Decision agreement between JSON and the readiness summary.
    decision = _norm(snap.get("decision"))
    for role, data in pkg.html.items():
        joined = " ".join(data["lines"]).lower()
        if "not ready" in decision and "not ready" not in joined:
            pkg.conflict(f"CNF-HTML-DECISION-{role}", "HTML vs JSON",
                         f"{role} does not state the JSON decision '{snap.get('decision')}'.")
        if decision.startswith("ready") and "not ready" in joined:
            pkg.conflict(f"CNF-HTML-DECISION-{role}", "HTML vs JSON",
                         f"{role} states not ready while JSON decision is '{snap.get('decision')}'.")
    customer = _norm(pkg.identity.get("CustomerName"))
    for role, data in pkg.html.items():
        if customer and data["lines"] and customer not in _norm(" ".join(data["lines"][:5])):
            pkg.add("PKG-BOUNDARY-HTML", "Boundary", "Fail", "Blocking",
                    f"{role} does not identify the package customer in its title.")


def _check_publication(pkg):
    if pkg.blocking:
        pkg.add("PKG-PUBLICATION", "Publication", "Fail", "Blocking",
                "Customer-ready publication is stopped: material package integrity is not established. "
                "Engineering review proceeds with explicit qualification.")
    else:
        pkg.add("PKG-PUBLICATION", "Publication", "Pass", "Info",
                "Package integrity is established; outputs may be released after operator review.")
