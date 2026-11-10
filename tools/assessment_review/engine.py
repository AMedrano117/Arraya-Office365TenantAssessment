"""Independent review of script conclusions (workflow steps 1 to 7).

Every statement the engine produces is typed so readers can tell collector
facts from script conclusions, Microsoft guidance, reviewer interpretation,
suggested hardening, customer decisions and evidence gaps.
"""
from __future__ import annotations

import json
from pathlib import Path
import re

from . import REVIEW_VERSION
from .facts import extract, private_terms
from .knowledge import (CLAIM_CHECKS, CLOSURE, CONTRADICTED_GAPS, DEPENDENCY_STEPS, EXCEPTION_DOC,
                        PRIORITY_DEFAULT, TOPICS, TREATMENT_SENTENCE, overlap_primary, topic_for)
from .plain import GROUP, describe
from .privacy import Sanitizer
from .util import parse_date, age_days, text as txt

FACT = "Collector fact"
SCRIPT = "Script-derived conclusion"
MSDOC = "Microsoft-documented guidance"
REVIEWER = "Reviewer interpretation"
HARDENING = "Suggested hardening practice"
CUSTOMER = "Customer policy decision"
GAP = "Evidence gap"
STATEMENT_TYPES = (FACT, SCRIPT, MSDOC, REVIEWER, HARDENING, CUSTOMER, GAP)

DECISIONS = ("Confirmed", "ConfirmedWithQualification", "CorrectFinding", "CorrectRecommendation",
             "DowngradeToEvidenceGap", "AddMissingFinding", "DuplicateOrOverlapping", "NotApplicable",
             "CustomerDecisionRequired", "InsufficientEvidence")
TREATMENTS = ("BlocksPilot", "ConfirmBeforePilot", "TenantHardening", "LongerTermImprovement",
              "InventoryHygiene", "EvidenceRequest", "NoAction")
SUFFICIENCY = ("Sufficient", "SufficientWithQualification", "Limited", "Stale", "Contradictory", "Missing")
STATEMENT_KIND = {"MicrosoftGuidance": MSDOC, "CustomerPolicyDecision": CUSTOMER,
                  "ReviewerInterpretation": REVIEWER}

# Finding evidence source names mapped to collector source identifiers.
SOURCE_HINTS = {
    "entra_auth_methods": "auth_methods", "auth_methods": "auth_methods", "ca_policies": "ca_policies",
    "authorization_policy": "authorization_policy", "defender_incidents": "incidents",
    "purview_dlp_policies": "purview_dlp_policies", "purview_label_policies": "purview_label_policies",
    "purview_sensitivity_labels": "purview_sensitivity_labels", "purview_audit_config": "purview_audit_config",
    "purview_retention_policies": "purview_retention_policies", "m365_license_coverage": "m365_license_coverage",
    "m365_external_connections": "m365_external_connections",
}
TOPIC_SOURCES = {"RISK": "risky_users", "ACCESS_REVIEW": "access_reviews", "GUEST": "guests",
                 "DEVICE": "managed_devices", "CONSENT": "oauth_grants", "PRIV": "role_assignments",
                 "LICENSING": "m365_license_coverage", "SHARING": "sharepoint_tenant_settings",
                 "LEGACY": "sharepoint_tenant_settings"}
RESTRICTED_SOURCE = re.compile(r"(?i)users|signin|risky|roleassign|role_assign|incident|alert|oauth2|grant|"
                               r"auditlog|registration|guest|privileg|admin_role|device")
SCRIPT_TREATMENT = {"Blocks pilot": "BlocksPilot", "Confirm before pilot": "ConfirmBeforePilot"}


class Context:
    def __init__(self, pkg, benchmarks, review_date):
        self.pkg = pkg
        self.benchmarks = benchmarks
        self.bench_by_id = {b["BenchmarkId"]: b for b in benchmarks["Benchmarks"]}
        self.review_date = review_date
        self.snap = pkg.snapshot
        self.facts = extract(pkg)
        redact, scan = private_terms(pkg)
        self.scan_terms = scan
        self.clean = Sanitizer(redact)
        self.evidence = {}
        for e in self.snap.get("evidence") or []:
            self.evidence.setdefault(e.get("evidence_id"), e)
        self.control_of = {}
        self.controls = {c.get("control_id"): c for c in self.snap.get("controls") or []}
        for c in self.snap.get("controls") or []:
            for rid in c.get("recommendation_ids") or []:
                self.control_of.setdefault(rid, c.get("control_id"))
        self.findings = self.snap.get("customer_findings") or []
        self.overlap = overlap_primary(self.findings)
        self.threshold = (self.snap.get("collection_thresholds") or {}).get("review_freshness_days") or 35
        self.identity = pkg.identity
        self.conflict_refs = {}
        for conflict in pkg.conflicts:
            for ref in conflict["References"]:
                self.conflict_refs.setdefault(ref, []).append(conflict)


# ------------------------------------------------------------------ item scaffold

def _new_item(ctx, finding_id, control_id, title, topic, item_type):
    step, domain, bench, role, subject = TOPICS[topic]
    kind, area = DEPENDENCY_STEPS[step]
    return {
        "AssessmentId": ctx.identity.get("AssessmentId"), "RunId": ctx.identity.get("RunId"),
        "ReviewVersion": REVIEW_VERSION, "BenchmarkVersion": ctx.benchmarks["BenchmarkVersion"],
        "ItemType": item_type, "FindingId": finding_id, "FindingKey": "", "ControlId": control_id, "Title": title,
        "Topic": topic, "DependencyStep": step, "DependencyArea": area,
        "FoundationalOrAISpecific": kind, "TechnicalDomain": domain,
        "ScriptConclusion": None, "Evidence": {}, "EvidenceReferences": [],
        "EvidenceSufficiency": "Missing", "ConsistencyChecks": [], "Statements": [],
        "ReviewDecision": None, "DecisionRationale": "", "ReadinessTreatment": None,
        "ReviewPriority": None, "CorrectedInterpretation": "", "PlainLanguage": "",
        "ConclusionChanges": [], "MicrosoftBenchmarkIds": list(bench), "RecommendedAction": "",
        "ClosureEvidence": {}, "CustomerDecisionRequired": False, "Limitations": [],
        "CustomerActionStatus": "Not recorded", "OverlapsWith": None, "ResponsibleRole": role,
        "_subject": subject,
    }


def _say(item, kind, text_value):
    assert kind in STATEMENT_TYPES, kind
    item["Statements"].append({"Type": kind, "Text": text_value})


def _check(item, check_id, result, detail):
    item["ConsistencyChecks"].append({"CheckId": check_id, "Result": result, "Detail": detail})


def _bench(item, *ids):
    for b in ids:
        if b not in item["MicrosoftBenchmarkIds"]:
            item["MicrosoftBenchmarkIds"].append(b)


# ------------------------------------------------------------------ step 1 and 2

def _source_for(ctx, f, topic):
    candidates = []
    src = f.get("EvidenceSource")
    if src:
        candidates += [SOURCE_HINTS.get(src, src), src, src.replace("entra_", ""), "purview_" + src]
    if topic in TOPIC_SOURCES:
        candidates.append(TOPIC_SOURCES[topic])
    for cand in candidates:
        if cand in ctx.pkg.sources:
            return cand, ctx.pkg.sources[cand]
    return (candidates[0] if candidates else None), None


def _native_source(f, ev):
    source_file = str((ev or {}).get("source_file") or f.get("SourceFile") or "")
    stype = (ev or {}).get("source_type") or f.get("SourceType") or ""
    if "graph.microsoft.com" in source_file:
        return "Microsoft Graph API (tenant collection)"
    if "api.security.microsoft.com" in source_file:
        return "Microsoft Defender API (tenant collection)"
    if stype == "portal_export":
        return "Administrator portal export supplied to the collector"
    if stype == "derived_evidence":
        return "Derived by the assessment engine from retained collection records"
    if stype == "assessment_requirement":
        return "Assessment requirement (no source collected)"
    if stype == "tenant_collection":
        return "Tenant collection"
    return "Not recorded"


def _reconstruct(ctx, item, f):
    ev = ctx.evidence.get(f.get("EvidenceId"))
    source_id, source = _source_for(ctx, f, item["Topic"])
    obs_date, problem = parse_date(f.get("ObservationDate"))
    age = age_days(obs_date, ctx.review_date) if not problem else None
    qualification = ctx.clean(f.get("Qualification") or (ev or {}).get("qualification") or "")
    missing = []
    if not f.get("EvidenceId"):
        missing.append("No evidence record is linked to this finding.")
    if "population is not recorded" in qualification.lower():
        missing.append("The assessed population is not recorded.")
    if "observation date is not recorded" in qualification.lower():
        missing.append("The original observation date is not recorded.")
    if "source association is" in qualification.lower():
        missing.append("The retained dataset behind the observation is not linked.")
    population = f.get("MFARegistrationPopulation") or {}
    item["Evidence"] = {
        "NativeSource": _native_source(f, ev),
        "LogicalDataset": txt(f.get("EvidenceKey")),
        "CollectorSourceId": source_id or "Not linked",
        "CollectionState": source["Classification"] if source else "Unknown",
        "CollectorRecords": source.get("Records") if source else None,
        "ObservationDate": obs_date.isoformat() if obs_date else "Not recorded",
        "AgeDays": age,
        "Freshness": f.get("Freshness") or (ev or {}).get("freshness") or "unknown",
        "Population": ctx.clean(f.get("EvidenceScope") or (ev or {}).get("population") or "") or "Not recorded",
        "RecordCount": (ev or {}).get("denominator") or population.get("total"),
        "DistinctEntityCount": population.get("denominator"),
        "EvidenceLevel": f.get("EvidenceLevel") or (ev or {}).get("evidence_level") or "Not recorded",
        "Qualification": qualification or "None recorded",
        "MissingEvidence": missing,
    }
    refs = []
    for ev_id in [f.get("EvidenceId")] + list(f.get("RelatedEvidenceIds") or []):
        if not ev_id or any(r["EvidenceId"] == ev_id for r in refs):
            continue
        record = ctx.evidence.get(ev_id)
        restricted = bool(record and (record.get("affected_objects") or
                                      RESTRICTED_SOURCE.search(str(record.get("metric_id") or "") +
                                                               str(record.get("source_file") or ""))))
        refs.append({"EvidenceId": ev_id, "Resolves": record is not None,
                     "Classification": "Restricted" if restricted else "Internal",
                     "Availability": (record or {}).get("availability", "missing"),
                     "Freshness": (record or {}).get("freshness", "unknown")})
    item["EvidenceReferences"] = refs
    status = f.get("EvidenceStatus")
    if status == "gap" or not refs:
        suff = "Missing"
    elif status == "stale" or (age is not None and age > ctx.threshold):
        suff = "Stale"
    elif status == "supported" and f.get("EvidenceComplete") and not qualification:
        suff = "Sufficient"
    elif status == "supported":
        suff = "SufficientWithQualification"
    elif source and source["Classification"] in ("Complete", "CompleteEmpty"):
        suff = "SufficientWithQualification"
    else:
        suff = "Limited"
    if source and source["Classification"] in ("Failed", "Unavailable") and status != "gap":
        suff = "Missing"
    item["EvidenceSufficiency"] = suff
    return source


def _script(ctx, item, f):
    item["ScriptConclusion"] = {
        "FindingId": f.get("RecommendationId"), "ScriptControlId": f.get("ControlId"),
        "Title": ctx.clean(f.get("Feature")), "Status": f.get("Status"), "Priority": f.get("Priority") or "None",
        "Disposition": f.get("Disposition"), "ActionType": f.get("ActionType") or "None",
        "ReadinessEffect": f.get("PilotImpact") or "None", "ReadinessStage": f.get("ReadinessStage"),
        "Observation": ctx.clean(f.get("Observation")), "Recommendation": ctx.clean(f.get("Recommendation")),
        "ResponsibleRole": f.get("OwnerRole") or "Not recorded",
        "ClosureEvidence": ctx.clean(f.get("CompletionEvidence")) or "Not recorded",
        "Confidence": f.get("Confidence"), "EvidenceStatus": f.get("EvidenceStatus"),
        "LifecycleState": f.get("LifecycleState") or "Not evaluated (no baseline run)",
        "GovernanceState": f.get("GovernanceState") or "No governance decision recorded",
        "SecurityGate": bool(f.get("SecurityGate")),
    }
    _say(item, SCRIPT, item["ScriptConclusion"]["Observation"] or "No observation recorded.")


# ------------------------------------------------------------------ step 3

def _numbers(value):
    return [int(n) for n in re.findall(r"(?<![\w.-])(\d{1,6})(?![\w.])", value or "")]


def _consistency(ctx, item, f, source):
    obs = f.get("Observation") or ""
    rec = f.get("Recommendation") or ""
    qual = (f.get("Qualification") or "").lower()
    topic = item["Topic"]
    # Counts reconcile with collector facts.
    expected = _expected_counts(ctx, topic, f)
    if expected:
        nums = _numbers(obs)
        missing = [f"{label} {value}" for label, value in expected if value is not None and value not in nums]
        if missing:
            _check(item, "C-COUNTS", "Issue", "Observation does not state the collector count(s): " + ", ".join(missing) + ".")
        else:
            _check(item, "C-COUNTS", "Pass", "Observation counts match collector counts.")
    else:
        _check(item, "C-COUNTS", "NotApplicable", "No independent collector count for this finding.")
    # Records versus entities.
    if re.search(r"(?i)\d+\s+(permissions?|links?|grants?|assignments?|policies|incidents?|alerts?|records?)", obs):
        if re.search(r"(?i)not a count of distinct|can overlap|semantic assignments|unique users", obs):
            _check(item, "C-RECORDS", "Pass", "Record counts are explicitly distinguished from distinct entities.")
        else:
            _check(item, "C-RECORDS", "Note", "Counts are records (policies, incidents, grants or links), not distinct "
                   "people or files.")
    else:
        _check(item, "C-RECORDS", "NotApplicable", "No record-type counts stated.")
    # Population and period comparability.
    age = item["Evidence"]["AgeDays"]
    if age is not None and age > ctx.threshold:
        _check(item, "C-PERIOD", "Issue", f"Evidence is {age} days old at the review date, beyond the "
               f"{ctx.threshold}-day freshness threshold; it is not comparable with current tenant collection.")
    elif "population is not recorded" in qual:
        _check(item, "C-PERIOD", "Issue", "The assessed population is not recorded, so scope cannot be compared.")
    else:
        _check(item, "C-PERIOD", "Pass", "Observation period and population are recorded.")
    # Configuration versus operation.
    level = str(f.get("EvidenceLevel") or "")
    if level in ("configuration", "policy_enforcement") and f.get("Status") == "Success":
        if re.search(r"(?i)configuration component only|operational confirmation|separate evidence|checked separately", qual + obs):
            _check(item, "C-CONFIG-OP", "Pass", "Configuration is explicitly separated from operation.")
        else:
            _check(item, "C-CONFIG-OP", "Issue", "A configuration observation is reported as Success without "
                   "stating that operation or effect was not tested.")
    else:
        _check(item, "C-CONFIG-OP", "NotApplicable", "Not a configuration-only success claim.")
    # Registration versus enforcement.
    if topic in ("MFA", "CA", "PRIV"):
        if re.search(r"(?i)registration (is separate|does not prove)|separate from policy enforcement|"
                     r"registration, enforcement|not actual sign-in use or enforcement|policy targets", obs):
            _check(item, "C-REG-ENF", "Pass", "Registration, enforcement and observed sign-in are kept separate.")
        else:
            _check(item, "C-REG-ENF", "Note", "The script text does not separate registration from enforcement; "
                   "the review states the distinction.")
    # Inventory versus entitlement.
    if topic == "LICENSING":
        if re.search(r"(?i)\d+\s+[\w ]{0,40}licen[sc]e", obs) and \
                not re.search(r"(?i)not unique people|entitlement", obs + qual):
            _check(item, "C-INV-ENT", "Issue", "License inventory is stated without separating seats from "
                   "user entitlement.")
        else:
            _check(item, "C-INV-ENT", "Pass", "Inventory and entitlement are not conflated.")
    # Missing evidence versus failure.
    if f.get("Disposition") == "Coverage" and str(f.get("Status")).lower() in ("fail", "failed", "action required"):
        _check(item, "C-GAP-NOT-FAIL", "Issue", "An evidence gap is presented as a failure.")
    elif f.get("Disposition") == "Coverage":
        _check(item, "C-GAP-NOT-FAIL", "Pass", "The gap is presented as not assessed, not as a failure.")
    # Customer-supplied evidence versus tenant telemetry.
    if f.get("SourceType") == "portal_export":
        if "tenant" in str(f.get("EvidenceBasis") or "").lower():
            _check(item, "C-CUSTOMER-EVIDENCE", "Issue", "A portal export supplied to the collector is labeled "
                   "as tenant evidence; it is a dated export, not live collector telemetry.")
        else:
            _check(item, "C-CUSTOMER-EVIDENCE", "Pass", "Export evidence is labeled as an export.")
    # Confidence versus evidence quality.
    conf = str(f.get("Confidence") or "")
    if conf == "High" and (f.get("EvidenceStatus") in ("unknown", "limited", "gap", "stale")
                           or "source association is" in qual):
        _check(item, "C-CONFIDENCE", "Issue", f"Confidence is High while evidence status is "
               f"'{f.get('EvidenceStatus')}'{' and the source association is not recorded' if 'source association is' in qual else ''}.")
    else:
        _check(item, "C-CONFIDENCE", "Pass", "Stated confidence is consistent with evidence status.")
    # Source association understated.
    if "source association is" in qual and source and source["Classification"] in ("Complete", "CompleteEmpty"):
        _check(item, "C-SOURCE-LINK", "Issue", f"The finding says its dataset is not linked, but the collector "
               f"retained source '{item['Evidence']['CollectorSourceId']}' as {source['Classification']}.")
    # Detail links and output agreement.
    rid = f.get("RecommendationId")
    broken = [b for b in ctx.pkg.windows.get("BrokenDetailLinks") or [] if rid and rid in b]
    if f.get("Disposition") == "Action":
        _check(item, "C-LINKS", "Issue" if broken else "Pass",
               "Detail link does not resolve." if broken else "Workbook detail links for this action resolve.")
    conflicts = ctx.conflict_refs.get(rid) or []
    output_conflicts = [c for c in conflicts if c["Area"] != "Script internal"]
    if output_conflicts:
        _check(item, "C-OUTPUTS-AGREE", "Issue", " ".join(c["Detail"] for c in output_conflicts))
    elif f.get("Disposition") == "Action":
        _check(item, "C-OUTPUTS-AGREE", "Pass", "Workbook, HTML and JSON agree for this action.")
    for c in conflicts:
        if c["Area"] == "Script internal":
            _check(item, "C-INTERNAL", "Note", f"{c['ConflictId']}: {c['Detail']}")


def _expected_counts(ctx, topic, f):
    facts = ctx.facts
    key = str(f.get("FindingKey") or "")
    if topic == "MFA" and "registration" in key and facts.get("mfa_registration"):
        reg = facts["mfa_registration"]
        return [("registered", reg["Registered"]), ("returned", reg["Returned"])]
    if topic == "RISK" and facts.get("risk"):
        return [("at risk", facts["risk"]["AtRisk"])]
    if topic == "INCIDENTS" and facts.get("incidents"):
        return [("active", facts["incidents"]["Active"]), ("active high", facts["incidents"]["ActiveHigh"])]
    if topic == "LICENSING" and "capacity" in key and facts.get("copilot_licenses"):
        lic = facts["copilot_licenses"]
        return [("enabled seats", lic["Enabled"]), ("assigned seats", lic["Assigned"])]
    if topic == "LABELS" and "sensitivity_labels" in key and facts.get("labels"):
        return [("labels", facts["labels"]["Count"])]
    return []


# ------------------------------------------------------------------ steps 4 to 7

def _claims(ctx, item, f):
    hits = []
    for claim_id, field, pattern, bench, correction, kind in CLAIM_CHECKS:
        if re.search(pattern, str(f.get(field) or "")):
            hits.append(claim_id)
            _bench(item, bench)
            _say(item, STATEMENT_KIND.get(kind, REVIEWER), correction)
            item["Limitations"].append(f"{claim_id}: script text corrected against {bench}.")
    return hits


def _default_decision(ctx, item, f, source):
    disp = f.get("Disposition")
    action_type = f.get("ActionType")
    suff = item["EvidenceSufficiency"]
    status = str(f.get("Status") or "")
    failed_source = source and source["Classification"] in ("Failed", "Unavailable")
    if failed_source and disp in ("Action", "Assurance", "Opportunity") and action_type != "Evidence":
        return ("DowngradeToEvidenceGap", "EvidenceRequest",
                f"The collector source this conclusion depends on is {source['Classification']}; the condition is "
                "an evidence gap, not a measured failure or pass.")
    if disp == "Coverage" or action_type == "Evidence":
        gated = f.get("PilotImpact") == "Confirm before pilot" and item["DependencyStep"] <= 8
        return ("Confirmed", "ConfirmBeforePilot" if gated else "EvidenceRequest",
                "The evidence gap is real: the check could not be answered from retained evidence.")
    if disp == "Assurance":
        if suff == "Sufficient" and not any(c["Result"] == "Issue" for c in item["ConsistencyChecks"]):
            return ("Confirmed", "NoAction", "Retained evidence supports the observed configuration.")
        return ("ConfirmedWithQualification", "NoAction",
                "The configuration is observed, but operation, scope or completeness is not established.")
    if disp == "Opportunity":
        if suff == "Missing":
            return ("InsufficientEvidence", "EvidenceRequest", "No retained evidence supports the observation.")
        return ("ConfirmedWithQualification", "LongerTermImprovement",
                "The observation is supported; it is an improvement rather than a rollout condition.")
    if disp == "Action":
        if action_type == "Remediation" and f.get("PilotImpact") == "Blocks pilot":
            return ("Confirmed", "BlocksPilot",
                    "The tenant-wide condition is confirmed and applies to every user who would use Copilot.")
        if suff == "Missing":
            return ("InsufficientEvidence", "ConfirmBeforePilot", "The condition cannot be confirmed from retained evidence.")
        return ("ConfirmedWithQualification", "ConfirmBeforePilot",
                "The condition is supported but needs confirmation against current, scoped evidence.")
    return ("InsufficientEvidence", "EvidenceRequest", "The finding type is not recognized by the review rules.")


def _apply(item, decision, treatment, rationale, priority=None):
    item["ReviewDecision"], item["ReadinessTreatment"], item["DecisionRationale"] = decision, treatment, rationale
    if priority:
        item["ReviewPriority"] = priority


# Topic handlers refine the default decision using collector facts and benchmarks.

def _h_ca(ctx, item, f):
    ca = ctx.facts.get("ca")
    if not ca or "sign_in" not in str(f.get("FindingKey") or ""):
        return
    sd = ctx.facts.get("security_defaults_enabled")
    _say(item, FACT, f"{ca['Total']} Conditional Access policies were returned: {ca['Enabled']} enabled, "
         f"{ca['ReportOnly']} report-only and {ca['Disabled']} disabled. Security defaults are "
         f"{'on' if sd else 'off' if sd is False else 'not recorded'}.")
    _say(item, FACT, f"Enforced all-user MFA policies: {ca['EnforcedAllUsersMfa']}. Report-only all-user MFA "
         f"policies: {ca['ReportOnlyAllUsersMfa']}. Enforced MFA policies targeting specific users or groups: "
         f"{ca['EnforcedTargetedMfa']}. Enforced legacy-authentication block for all users: "
         f"{ca['EnforcedLegacyBlockAllUsers']}.")
    _say(item, MSDOC, "Microsoft recommends a baseline policy requiring MFA for all users and all resources, "
         "excluding emergency access and synchronization accounts, and documents that report-only policies are "
         "evaluated but not enforced (MS-ENTRA-CA-MFA-01, MS-ENTRA-CA-REPORTONLY-01).")
    if ca["EnforcedAllUsersMfa"] or sd:
        _apply(item, "CorrectFinding", "NoAction",
               "Collector data shows tenant-wide MFA is enforced, contradicting the script conclusion.")
        item["CorrectedInterpretation"] = "Tenant-wide MFA enforcement is present in retained Conditional Access data."
        return
    universal = ca["EnforcedTargetedMfa"] == 0 and not sd
    treatment = "BlocksPilot"
    interp = ["No enforced policy requires MFA for all users; this tenant-wide gap is confirmed from configuration."]
    if universal:
        interp.append("No enforced MFA policy applies to any user.")
    else:
        interp.append(f"{ca['EnforcedTargetedMfa']} enforced policy targets specific users or groups; everyone "
                      "outside that scope can sign in without MFA.")
    if ca["EnforcedLegacyBlockAllUsers"]:
        interp.append("The legacy-authentication part of this control is met by an enforced all-user block policy.")
    if ca["ReportOnlyAllUsersMfa"]:
        exclusions = ", ".join(str(n) for n in ca["AllUsersMfaExclusionCounts"])
        interp.append(f"{ca['ReportOnlyAllUsersMfa']} all-user MFA policies already exist in report-only mode "
                      f"(exclusion counts: {exclusions}). The remediation is to review their report-only results "
                      "and exclusions and switch one to On, not to create a new policy.")
        _apply(item, "CorrectRecommendation", treatment,
               "The finding is confirmed; the recommendation to create a new policy is corrected because "
               "all-user MFA policies already exist in report-only mode.")
        item["RecommendedAction"] = ("Review report-only results for the existing all-user MFA policies, confirm "
                                     "each exclusion is an emergency access or synchronization account, then switch "
                                     "one policy to On and retire duplicates.")
        _say(item, CUSTOMER, "Which accounts are excluded as emergency access accounts is a customer designation.")
    else:
        _apply(item, "Confirmed", treatment, "The tenant-wide gap is confirmed from Conditional Access configuration.")
    item["CorrectedInterpretation"] = " ".join(interp)
    _bench(item, "MS-ENTRA-EMERGENCY-01")


def _h_mfa(ctx, item, f):
    reg = ctx.facts.get("mfa_registration")
    if "registration" not in str(f.get("FindingKey") or ""):
        if f.get("RecommendationId") == "ENT-007":
            _apply(item, "ConfirmedWithQualification", "LongerTermImprovement",
                   "Passwordless registration is registration evidence only; phishing-resistant methods are an "
                   "improvement after baseline MFA.", "Low")
            _say(item, HARDENING, "Prioritize phishing-resistant methods for administrators after baseline MFA "
                 "enforcement is in place.")
        return
    if not reg:
        return
    _say(item, FACT, f"The registration report returned {reg['Returned']} accounts ({reg['Members']} member, "
         f"{reg['Guests']} guest): {reg['Registered']} registered and {reg['NotRegistered']} not registered. "
         f"{reg['AdminsNotRegistered']} of {reg['AdminsReturned']} accounts flagged as administrators are not registered.")
    _say(item, REVIEWER, "Registration shows readiness to use MFA, not that MFA is enforced or used. The returned "
         "population includes guests and accounts that may be unlicensed, shared or inactive, so the percentage "
         "overstates the gap among active licensed users.")
    treatment = "BlocksPilot" if reg["Registered"] == 0 else "ConfirmBeforePilot"
    _apply(item, "ConfirmedWithQualification", treatment,
           "Low registration is confirmed; registration is the prerequisite for enforcing MFA, and the enforcement "
           "gap is carried by the Conditional Access review." if treatment != "BlocksPilot" else
           "No account is registered for MFA, so no user can satisfy an MFA requirement.")
    item["CorrectedInterpretation"] = (f"{reg['Registered']} of {reg['Returned']} returned accounts are registered for "
                                       "MFA. This is registration, not enforcement. Active users must register before "
                                       "MFA can be enforced through Conditional Access (see the IDENTITY.AUTH review).")


def _h_priv(ctx, item, f):
    key = str(f.get("FindingKey") or "")
    ca = ctx.facts.get("ca") or {}
    if "WithoutMFARegistration" in key:
        _say(item, FACT, f"Enforced Conditional Access policies requiring MFA for administrator roles: "
             f"{ca.get('EnforcedAdminMfa', 'not recorded')}; report-only or disabled admin MFA policies: "
             f"{ca.get('ReportOnlyOrDisabledAdminMfa', 'not recorded')}.")
        _say(item, MSDOC, "Microsoft recommends MFA for all administrator accounts and enforces MFA for some admin "
             "portals at the platform level; Microsoft Graph APIs are generally not in scope "
             "(MS-ENTRA-ROLES-01, MS-ENTRA-MANDATORY-MFA-01).")
        _apply(item, "ConfirmedWithQualification", "ConfirmBeforePilot",
               "Privileged identities with explicitly unregistered MFA are confirmed; population completeness is "
               "not established.", "High")
        no_admin_ca = not ca.get("EnforcedAdminMfa") and not ca.get("EnforcedAllUsersMfa")
        item["CorrectedInterpretation"] = (
            "Privileged identities without MFA registration are confirmed; the privileged population may be "
            "incomplete. " + ("No enforced Conditional Access policy targets administrator roles or all users, so "
                              "unless these accounts are in the group covered by the targeted MFA policy, only "
                              "Microsoft's platform-enforced admin-portal MFA applies. Treat this as a high-priority "
                              "foundation gap." if no_admin_ca else ""))
        item["RecommendedAction"] = ("Register MFA for each privileged identity or remove the privilege, enforce an "
                                     "administrator MFA policy (enable an existing report-only policy where present), "
                                     "and record any emergency access exception.")
    elif "DisabledIdentity" in key:
        _apply(item, "ConfirmedWithQualification", "TenantHardening",
               "Disabled identities retaining privilege are supported by retained assignments.", "Medium")
        _say(item, MSDOC, "Microsoft recommends least privilege and removing unneeded assignments (MS-ENTRA-ROLES-01).")
    elif "PermanentPrivilege" in key:
        _apply(item, "ConfirmedWithQualification", "TenantHardening",
               "Permanent privilege is supported by retained assignments; business need is a customer decision.", "Medium")
        _say(item, MSDOC, "Microsoft recommends Privileged Identity Management for just-in-time access, which "
             "requires Microsoft Entra ID P2 or Governance (MS-ENTRA-ROLES-01).")
    elif "ServicePurpose" in key:
        _apply(item, "ConfirmedWithQualification", "InventoryHygiene",
               "Service or synchronization purpose is unverified; this is record-keeping.", "Low")
    elif "ActivityUnknown" in key:
        _apply(item, "ConfirmedWithQualification", "EvidenceRequest",
               "Missing activity evidence informs privilege reduction; it is not a Copilot rollout condition.", "Medium")
        _say(item, GAP, "Successful privileged activity was not retained, so unused privilege cannot be identified.")


def _h_consent(ctx, item, f):
    key = str(f.get("FindingKey") or "")
    consent = ctx.facts.get("consent")
    if "user_consent_assignment" in key and consent:
        names = ", ".join(consent["AssignedUserPolicies"]) or "none"
        _say(item, FACT, f"Self-consent policies assigned to default users: {names}.")
        if consent["AllMicrosoftManaged"]:
            _say(item, MSDOC, "These are Microsoft-managed built-in policies; microsoft-user-default-recommended is "
                 "the default 'Let Microsoft manage your consent settings' setting, which excludes high-impact "
                 "Files, Sites, Mail, Calendar and Chat permissions (MS-ENTRA-CONSENT-02).")
            _say(item, CUSTOMER, "Whether to restrict consent further, for example to verified publishers with the "
                 "admin consent workflow, is a customer decision (MS-ENTRA-CONSENT-01).")
            _apply(item, "CorrectFinding", "TenantHardening",
                   "The assigned policies are Microsoft-managed defaults; the evidence does not support treating "
                   "user consent as a rollout blocker.", "Medium")
            item["CorrectedInterpretation"] = ("User consent is governed by Microsoft-managed default policies, not an "
                                               "unrestricted setting. This is a tenant hardening decision, not a "
                                               "demonstrated rollout blocker.")
            item["RecommendedAction"] = ("Decide whether the Microsoft-managed consent setting meets policy; if not, "
                                         "restrict to verified publishers and enable the admin consent workflow. "
                                         "Review existing grants separately (see application grant review).")
            item["CustomerDecisionRequired"] = True
        elif consent["LegacyAssigned"]:
            _say(item, MSDOC, "microsoft-user-default-legacy allows users to consent to any permission not requiring "
                 "admin consent; Microsoft recommends verified-publisher consent (MS-ENTRA-CONSENT-01).")
            _apply(item, "ConfirmedWithQualification", "ConfirmBeforePilot",
                   "Unrestricted user consent is confirmed.", "High")
        else:
            _apply(item, "CustomerDecisionRequired", "ConfirmBeforePilot",
                   "Custom consent policies are assigned; whether their conditions are acceptable is a customer decision.")
    elif "high_impact_grants" in key and consent:
        _say(item, FACT, f"Grant inventory flags {consent.get('HighPrivilegeApps')} applications with high-privilege "
             f"permissions and {consent.get('UnverifiedPublishers')} from unverified publishers.")
        if ctx.facts.get("app_permission_apps"):
            _say(item, FACT, f"{ctx.facts['app_permission_apps']} non-Microsoft applications hold tenant-wide "
                 "application permissions that read or change mail, files, sites or directory data; the script "
                 "count covers delegated grants only.")
        _say(item, REVIEWER, "Existing application grants are a tenant data-access risk independent of Copilot; "
             "Copilot does not extend application permissions.")
        _apply(item, "ConfirmedWithQualification", "TenantHardening",
               "High-impact grants are supported by the retained grant inventory; they do not change what Copilot "
               "can reach.", "High")


def _h_incidents(ctx, item, f):
    inc = ctx.facts.get("incidents")
    if not inc:
        return
    _say(item, FACT, f"{inc['Records']} incident records were returned; {inc['Active']} are active ({inc['ActiveHigh']} "
         f"high, {inc['ActiveMedium']} medium, {inc['ActiveLow']} low). {inc['ActiveNoUpdateSinceCreation']} active "
         f"incidents have no recorded update since creation and {inc['ActiveUnassigned']} have no owner. The oldest "
         f"active incident is {inc['OldestActiveAgeDays']} days old.")
    _say(item, MSDOC, "Incidents start Active with no owner; triage assigns an owner, severity and status, and "
         "resolution records a classification (MS-XDR-INCIDENTS-01).")
    _say(item, REVIEWER, "Incident counts are records; untriaged incidents show a handling gap, not a confirmed "
         "compromise. Triage decides whether any account or device is compromised.")
    treatment = "ConfirmBeforePilot" if inc["ActiveHigh"] else "TenantHardening"
    _apply(item, "ConfirmedWithQualification", treatment,
           "Active, untriaged incidents are confirmed; a blocker is not established until triage shows an active "
           "compromise.", "High" if inc["ActiveHigh"] else "Medium")
    item["CorrectedInterpretation"] = ("Untriaged active incidents indicate an incident-handling gap. Triage the "
                                       "high-severity incident before Copilot rollout to rule out an active "
                                       "compromise; work the lower-severity backlog as tenant hardening.")
    item["RecommendedAction"] = ("Assign owners, triage high severity first, classify and resolve or record an "
                                 "approved treatment.")


def _h_risk(ctx, item, f):
    risk = ctx.facts.get("risk")
    if not risk:
        return
    _say(item, FACT, f"Identity Protection returned {risk['Records']} risky-user records: {risk['AtRisk']} currently at "
         f"risk ({risk['High']} high, {risk['Medium']} medium, {risk['Low']} low), {risk['Remediated']} remediated and "
         f"{risk['Dismissed']} dismissed. A user-risk policy {'exists' if risk['UserRiskPolicy'] else 'was not found'}; "
         f"a sign-in risk policy {'exists' if risk['SignInRiskPolicy'] else 'was not found'}.")
    stated = _numbers(f.get("Observation"))
    if risk["AtRisk"] is not None and stated and risk["AtRisk"] not in stated:
        _apply(item, "CorrectFinding", "ConfirmBeforePilot",
               f"The script reports {stated[0]} risky users; the collector shows {risk['AtRisk']} accounts currently at "
               "risk, including low-risk accounts.", "Medium")
        item["CorrectedInterpretation"] = (f"{risk['AtRisk']} accounts are currently at risk ({risk['Medium']} medium, "
                                           f"{risk['Low']} low). The script count omits low-risk accounts. Each must be "
                                           "remediated or dismissed after investigation.")
    else:
        _apply(item, "ConfirmedWithQualification", "ConfirmBeforePilot",
               "Risky users are confirmed in the retained source.", "Medium")


def _h_legacy(ctx, item, f):
    ca = ctx.facts.get("ca") or {}
    sign = ctx.facts.get("signins") or {}
    if ca.get("EnforcedLegacyBlockAllUsers"):
        _say(item, FACT, "An enforced Conditional Access policy blocks Exchange ActiveSync and other legacy clients "
             f"for all users. The sign-in sample shows {sign.get('LegacyAttempts', 'no recorded')} legacy attempts in "
             f"{sign.get('Sampled', 'an unrecorded number of')} sampled sign-ins.")
        _say(item, HARDENING, "Disabling the SharePoint legacy authentication setting adds defense in depth after "
             "confirming no required client depends on it.")
        _apply(item, "ConfirmedWithQualification", "TenantHardening",
               "The setting is confirmed, but legacy clients are already blocked tenant-wide by Conditional Access.", "Low")
        item["CorrectedInterpretation"] = ("The SharePoint setting permits legacy protocols, but an enforced tenant-wide "
                                           "Conditional Access block applies. Residual risk is limited to excluded "
                                           "accounts; treat as defense in depth.")


def _h_sharing(ctx, item, f):
    _say(item, MSDOC, "Microsoft documents Anyone links as suited to friction-free sharing when most content is not "
         "sensitive, with expiration and view-only restrictions available, and Specific people as the most "
         "restrictive default (MS-SPO-SHARING-01).")
    _say(item, CUSTOMER, "Whether Anyone links and an Anyone default are acceptable for business content is a "
         "customer policy decision.")
    _apply(item, "CustomerDecisionRequired", SCRIPT_TREATMENT.get(f.get("PilotImpact"), "TenantHardening"),
           "The setting is a documented policy choice; the evidence does not show a misconfiguration.", "Medium")
    item["CustomerDecisionRequired"] = True


def _h_guest(ctx, item, f):
    guests = ctx.facts.get("guests") or {}
    _say(item, FACT, f"Guest invitation setting: {guests.get('InviteSetting', 'not recorded')}; guest accounts "
         f"returned: {guests.get('Total', 'not recorded')}.")
    _say(item, MSDOC, "Allowing all users to invite guests is the Microsoft default (MS-ENTRA-GUEST-01).")
    _apply(item, "CustomerDecisionRequired", "LongerTermImprovement",
           "Guest invitation rights are a collaboration policy decision.", "Low")
    item["CustomerDecisionRequired"] = True


def _h_permissions(ctx, item, f):
    key = str(f.get("FindingKey") or "")
    seats = ctx.facts.get("copilot_licenses") or {}
    if seats.get("Enabled"):
        _say(item, MSDOC, "SharePoint Advanced Management, which provides Data Access Governance reports, is "
             "included with the Microsoft Copilot license held by this tenant (MS-COPILOT-FOUNDATION-01).")
    if "broad_internal" in key:
        _say(item, REVIEWER, "The counts are permissions and links that can overlap, not distinct files or people, "
             "and come from a dated portal export rather than live collection.")
        _apply(item, "ConfirmedWithQualification", "ConfirmBeforePilot",
               "Broad grants are supported by a stale export; current scope must be confirmed.", "High")
    elif "not_assessed" in key:
        _apply(item, "Confirmed", "ConfirmBeforePilot",
               "Current oversharing evidence is missing for the listed report types.", "High")


def _h_labels(ctx, item, f):
    labels = ctx.facts.get("labels") or {}
    if labels.get("Count") is not None:
        _say(item, FACT, f"Purview returned {labels['Count']} sensitivity label definitions and "
             f"{ctx.facts.get('label_policies', 'an unrecorded number of')} label publishing policies.")
    _say(item, REVIEWER, "Label definitions and publishing do not show that business content is labeled.")


def _h_audit(ctx, item, f):
    audit = ctx.facts.get("audit") or {}
    if "coverage.operation" in str(f.get("FindingKey") or ""):
        _say(item, MSDOC, "Copilot interactions are logged automatically under Audit (Standard) when auditing is "
             "enabled (MS-PURVIEW-AUDIT-COPILOT-01); the open question is whether records can be retrieved.")
        _apply(item, "ConfirmedWithQualification", "ConfirmBeforePilot",
               "Audit is enabled by configuration; retrieval of Copilot records has not been demonstrated.", "Medium")
    _say(item, FACT, f"Unified audit logging is {'enabled' if audit.get('UnifiedAuditEnabled') else 'not confirmed'}. "
         f"The Copilot interaction audit query is {audit.get('CopilotInteractionSource', 'Unknown')}"
         + (f": {audit['CopilotInteractionReason']}" if audit.get("CopilotInteractionReason") else "."))


def _h_dlp(ctx, item, f):
    if "coverage.operation" in str(f.get("FindingKey") or ""):
        _say(item, HARDENING, "Testing DLP behavior before Copilot rollout is an Arraya methodology requirement and a "
             "suggested practice; Microsoft supports policy simulation mode but does not require a test "
             "(MS-PURVIEW-DLP-COPILOT-01).")
        _bench(item, "MS-PURVIEW-DLP-COPILOT-01")
        _apply(item, "ConfirmedWithQualification", "ConfirmBeforePilot",
               "Operational DLP evidence is missing; configuration alone is observed.", "Medium")


def _h_copilot_dlp(ctx, item, f):
    rows = ctx.facts.get("copilot_dlp_rows") or 0
    monitor = ctx.facts.get("copilot_dlp_monitor_only") or 0
    if rows:
        _say(item, FACT, f"The Copilot admin review lists {rows} Copilot-related DLP policy rows; {monitor} report "
             "monitoring-only mode.")
    _say(item, MSDOC, "DLP for the Microsoft 365 Copilot location can stop processing of labeled items and of prompts "
         "containing sensitive information types (MS-PURVIEW-DLP-COPILOT-01).")


def _h_device(ctx, item, f):
    intune = ctx.facts.get("intune") or {}
    mde = ctx.facts.get("defender_devices") or {}
    _say(item, FACT, f"Intune returned {intune.get('ManagedDevices', 'no recorded')} managed devices (source "
         f"{intune.get('Source', 'Unknown')}) while {intune.get('CompliancePolicies', 'an unrecorded number of')} "
         "compliance policies exist.")
    if mde:
        _say(item, FACT, f"Defender for Endpoint returned {mde['Records']} device(s): {mde['Onboarded']} onboarded, "
             f"{mde['Active']} active, {mde['Inactive']} inactive; {mde['ReportedManagedByIntune']} reported as "
             "managed by Intune.")
    contradictory = mde.get("ReportedManagedByIntune") and intune.get("ManagedDevices") == 0
    if contradictory:
        item["EvidenceSufficiency"] = "Contradictory"
        _say(item, REVIEWER, "Defender reports a device managed by Intune while the Intune query returned none. The "
             "empty Intune result may reflect enrollment state, permissions or timing; confirm before concluding "
             "that no devices are managed.")
    _say(item, CUSTOMER, "The endpoint control model (Intune, another MDM, or browser and session controls) is a "
         "customer decision.")
    _apply(item, "ConfirmedWithQualification", "ConfirmBeforePilot",
           "The device-posture gap stands; the collector source is complete-empty, not missing" +
           (", and sources disagree." if contradictory else "."), "High")
    item["CustomerDecisionRequired"] = True
    item["CorrectedInterpretation"] = ("Device posture for Copilot users is not established. The Intune result is a "
                                       "complete but empty response, not missing evidence"
                                       + ("; it conflicts with Defender device data." if contradictory else "."))


def _h_shadow(ctx, item, f):
    mde = ctx.facts.get("defender_devices") or {}
    if f.get("RecommendationId") == "PUR-016" or "endpoint" in str(f.get("Observation") or "").lower():
        _say(item, FACT, f"Defender for Endpoint shows {mde.get('Active', 0)} active onboarded device(s).")
        _say(item, GAP, "Purview device onboarding status was not collected; only the Defender device inventory "
             "is available.")
        _say(item, REVIEWER, "Endpoint DLP policies apply only to onboarded devices; with no active onboarded "
             "device in the retained inventory, the policies provide no demonstrated external-AI egress control.")
        if not mde.get("Active"):
            _apply(item, "CorrectFinding", "LongerTermImprovement",
                   "The Success status is not supported by retained device evidence: the policies exist, but no "
                   "active onboarded device is observed. Device onboarding is a foundation prerequisite.", "Low")
            item["CorrectedInterpretation"] = ("Endpoint DLP is configured but not shown to be effective; complete device "
                                               "onboarding (foundation step 3) before relying on it for AI egress control.")


def _h_ai_data_risk(ctx, item, f):
    _say(item, MSDOC, "DSPM for AI runs a default weekly data risk assessment for the top 100 SharePoint sites by "
         "usage (MS-PURVIEW-DSPMAI-01).")
    _say(item, REVIEWER, "A DSPM assessment supplements, and does not replace, the foundational permissions review "
         "of SharePoint sites.")
    if item["ReviewDecision"] in ("Confirmed", "ConfirmedWithQualification"):
        _apply(item, "ConfirmedWithQualification", "EvidenceRequest",
               "The evidence gap is confirmed. Rollout gating for oversharing is carried by the foundational permissions "
               "review (step 4); the AI data-risk assessment is requested evidence, not a substitute gate.", "Medium")


def _h_licensing(ctx, item, f):
    key = str(f.get("FindingKey") or "")
    seats = ctx.facts.get("copilot_licenses")
    usage = ctx.facts.get("copilot_usage") or {}
    if "capacity" in key and seats:
        _say(item, FACT, f"{seats['Enabled']} paid Copilot seats exist; {seats['Assigned']} are assigned.")
        if usage:
            _say(item, FACT, "Copilot usage reports show "
                 + ", ".join(f"{v} {k.lower().replace('paid copilot ', '')}" for k, v in sorted(usage.items()))
                 + " in the reporting window.")
        _say(item, REVIEWER, "Seat inventory is not user entitlement; it does not show which users are licensed.")
        _apply(item, "ConfirmedWithQualification", "NoAction",
               "License inventory is supported; which users hold the licenses is not established.")
    elif "license.apps" in key:
        _say(item, MSDOC, "Microsoft 365 Apps are required for in-app Copilot experiences; Copilot Chat on the web "
             "does not depend on desktop app channels (MS-COPILOT-REQ-01).")
        item["ReviewPriority"] = "Medium"
    elif "license.assignment" in key:
        item["ReviewPriority"] = "Medium"


def _h_simple(decision, treatment, priority, customer=False, statement=None):
    def handler(ctx, item, f):
        if statement:
            _say(item, statement[0], statement[1])
        _apply(item, decision, treatment, item["DecisionRationale"] or "Reviewed against retained evidence.", priority)
        if customer:
            item["CustomerDecisionRequired"] = True
    return handler


HANDLERS = {
    "CA": _h_ca, "MFA": _h_mfa, "PRIV": _h_priv, "CONSENT": _h_consent, "INCIDENTS": _h_incidents,
    "RISK": _h_risk, "LEGACY": _h_legacy, "SHARING": _h_sharing, "GUEST": _h_guest,
    "PERMISSIONS": _h_permissions, "LABELS": _h_labels, "AUDIT": _h_audit, "DLP": _h_dlp,
    "COPILOT_DLP": _h_copilot_dlp, "DEVICE": _h_device, "SHADOW_AI": _h_shadow,
    "AI_DATA_RISK": _h_ai_data_risk, "LICENSING": _h_licensing,
    "LOCKBOX": _h_simple("CustomerDecisionRequired", "LongerTermImprovement", "Low", True,
                         (MSDOC, "Customer Lockbox requires approval before Microsoft engineers access content, covers "
                          "Copilot interactions and is included with E5 (MS-PURVIEW-LOCKBOX-01).")),
    # Pilot design and adoption planning are outside the scope of this security review.
    "ADOPTION": _h_simple("NotApplicable", "NoAction", "None", False,
                          (CUSTOMER, "Use cases, baseline and success measures are business sponsor decisions "
                           "outside the scope of this security review.")),
    "ACCESS_REVIEW": _h_simple("ConfirmedWithQualification", "LongerTermImprovement", "Low", False,
                               (MSDOC, "Microsoft recommends recurring access reviews; they require Microsoft Entra ID "
                                "Governance or Microsoft Entra Suite (MS-ENTRA-ROLES-01).")),
}


def _contradicted_gap(ctx, item, f):
    key = str(f.get("FindingKey") or "")
    for gap_key, (section, field), template in CONTRADICTED_GAPS:
        if key != gap_key:
            continue
        value = (ctx.facts.get(section) or {}).get(field)
        if value:
            source = (ctx.facts.get(section) or {}).get("Source") or "retained collection"
            message = template.format(value=value, source=source)
            _say(item, FACT, message)
            _apply(item, "CorrectFinding", item["ReadinessTreatment"] or "ConfirmBeforePilot",
                   "The gap statement is contradicted by retained collection data. " + message)
            item["CorrectedInterpretation"] = ("The data was collected. " + message + " The remaining question is "
                                               "whether it meets the requirement for the users who will use Copilot.")
            item["EvidenceSufficiency"] = "Contradictory"
            return True
    return False


def _closure(ctx, item):
    evidence, validation, population, reassess = CLOSURE[item["Topic"]]
    if item["ReadinessTreatment"] == "NoAction":
        evidence = "No closure evidence required; retain the current evidence reference."
    item["ClosureEvidence"] = {
        "RequiredEvidence": evidence, "ValidationProcedure": validation,
        "ResponsibleRole": item["ResponsibleRole"], "PopulationAndScope": population,
        "EvidenceFreshness": f"Dated within {ctx.threshold} days before reassessment.",
        "ReassessmentRequirement": reassess,
        "CustomerDecisionRequirement": ("Yes: record the customer's decision and the deciding role."
                                        if item["CustomerDecisionRequired"] else "No"),
        "AcceptableExceptionDocumentation": EXCEPTION_DOC,
        "ClosureRule": "Closure is never automatic; the operator confirms the evidence and records it for "
                       "the next assessment run.",
    }


def _finish(ctx, item, script_treatment=None):
    if item["ReviewPriority"] is None:
        item["ReviewPriority"] = PRIORITY_DEFAULT[item["ReadinessTreatment"]]
    if not item["RecommendedAction"]:
        script = item["ScriptConclusion"] or {}
        rec = script.get("Recommendation") or ""
        item["RecommendedAction"] = rec if rec else "Provide the closure evidence listed for this item."
    if not item["CorrectedInterpretation"]:
        item["CorrectedInterpretation"] = item["DecisionRationale"]
    if item["ItemType"] == "ScriptFinding" and item["ReviewDecision"] != "DuplicateOrOverlapping":
        def gate(value):
            return value if value in ("BlocksPilot", "ConfirmBeforePilot") else "NotPilotGating"
        before, after = gate(script_treatment), gate(item["ReadinessTreatment"])
        if before != after:
            item["ConclusionChanges"].append(f"Pilot effect changed from {before} to {after}.")
        if item["ReviewDecision"] in ("CorrectFinding", "CorrectRecommendation", "DowngradeToEvidenceGap",
                                      "InsufficientEvidence"):
            item["ConclusionChanges"].append(f"{item['ReviewDecision']}: {item['DecisionRationale']}")
    subject = item["_subject"]
    gap = (item.get("ScriptConclusion") or {}).get("Disposition") == "Coverage" or \
        (item.get("ScriptConclusion") or {}).get("ActionType") == "Evidence"
    lead = {
        "Confirmed": (f"The assessment could not check {subject}; the review agrees this is missing evidence, not "
                      "a failure." if gap else f"The assessment's finding about {subject} is supported by the evidence."),
        "ConfirmedWithQualification": f"The assessment's finding about {subject} is supported, with limits on what the evidence shows.",
        "CorrectFinding": f"The review corrected the assessment's conclusion about {subject}.",
        "CorrectRecommendation": f"The finding about {subject} stands, but the recommended fix was corrected.",
        "DowngradeToEvidenceGap": f"There is not enough evidence to judge {subject}.",
        "AddMissingFinding": f"The assessment did not cover {subject}; the review added it.",
        "DuplicateOrOverlapping": f"This repeats another item about {subject}.",
        "NotApplicable": f"This item about {subject} does not apply.",
        "CustomerDecisionRequired": f"Your organization needs to decide its policy on {subject}.",
        "InsufficientEvidence": f"There is not enough evidence to judge {subject}.",
    }[item["ReviewDecision"]]
    describe(ctx, item, item["FindingKey"], script_treatment)
    item["PlainLanguage"] = item["Plain"]["Issue"] + " " + TREATMENT_SENTENCE[item["ReadinessTreatment"]]
    item["ReviewNote"] = lead
    if item["EvidenceSufficiency"] in ("Missing", "Limited", "Stale", "Contradictory"):
        item["Limitations"].append(f"Evidence sufficiency is {item['EvidenceSufficiency']}.")
    if item["Evidence"].get("MissingEvidence"):
        item["Limitations"].extend(item["Evidence"]["MissingEvidence"])
    _closure(ctx, item)


def review_finding(ctx, f):
    rid = f.get("RecommendationId")
    topic = topic_for(f)
    control = ctx.control_of.get(rid) or f.get("ControlId")
    item = _new_item(ctx, rid, control, _title(ctx, f), topic, "ScriptFinding")
    item["FindingKey"] = str(f.get("FindingKey") or "")
    if f.get("OwnerRole"):
        item["ResponsibleRole"] = f["OwnerRole"]
    source = _reconstruct(ctx, item, f)                     # step 1
    _script(ctx, item, f)                                   # step 2
    _consistency(ctx, item, f, source)                      # step 3
    claims = _claims(ctx, item, f)                          # step 4
    decision, treatment, rationale = _default_decision(ctx, item, f, source)
    _apply(item, decision, treatment, rationale)
    handler = HANDLERS.get(topic)
    if handler and decision != "DowngradeToEvidenceGap":
        handler(ctx, item, f)                               # step 5 and 6 refinements
    _contradicted_gap(ctx, item, f)
    if claims and item["ReviewDecision"] in ("Confirmed", "ConfirmedWithQualification"):
        item["ReviewDecision"] = "CorrectRecommendation"
        item["DecisionRationale"] = ("The observation is supported; script recommendation text conflicts with "
                                     "Microsoft documentation or prescribes customer policy (" + ", ".join(claims) + ").")
        item["CorrectedInterpretation"] = ("The observation stands. The recommendation is corrected: "
                                           + "; ".join(CLAIM_SUMMARY[c] for c in claims) + ".")
        item["RecommendedAction"] = item["RecommendedAction"] or _corrected_recommendation(item)
    if item["ReviewDecision"] == "CorrectFinding" and not item["RecommendedAction"]:
        item["RecommendedAction"] = CLOSURE[topic][1]
    if not item["RecommendedAction"]:
        cleaned = customer_action(item["ScriptConclusion"]["Recommendation"])
        item["RecommendedAction"] = cleaned or CLOSURE[topic][1]
    if rid in ctx.overlap:
        item["OverlapsWith"] = ctx.overlap[rid]
        item["ReviewDecision"] = "DuplicateOrOverlapping"
        item["ReadinessTreatment"] = "NoAction"
        item["ReviewPriority"] = "None"
        item["DecisionRationale"] = f"Covered by {ctx.overlap[rid]}, which carries the review decision and action."
    if any(c["Result"] == "Issue" and c["CheckId"] == "C-SOURCE-LINK" for c in item["ConsistencyChecks"]):
        item["Limitations"].append("Script qualification understates retained evidence (source link missing).")
    script_treatment = SCRIPT_TREATMENT.get(f.get("PilotImpact"))
    if f.get("Disposition") == "Coverage" and item["ReviewDecision"] in ("Confirmed", "ConfirmedWithQualification"):
        _say(item, GAP, ctx.clean(f.get("Observation")) or "The check could not be answered.")
    _finish(ctx, item, script_treatment)
    return item


LICENSE_LABEL = re.compile(r"(?i)^(microsoft|office|azure|windows)\b.*\b(P1|P2|E3|E5|Plan [A-Z0-9]+|license)\)?$")
TOOL_SENTENCE = re.compile(r"(?i)--[a-z]|\b[A-Z]{3,}_[A-Z_]{3,}\b|\bthis (tool|assessment)\b|\bthe collector\b|"
                           r"\brerun (this|the) assessment\b|\bGraph [A-Z][\w.]+\.Read")
CLAIM_SUMMARY = {
    "CLM-TRAINING": "retention is not an AI training control",
    "CLM-AUDIT-10Y": "ten-year audit retention needs an add-on license",
    "CLM-AUDIT-EVENTS": "the documented Copilot audit record type is CopilotInteraction",
    "CLM-AI-OUTPUT-LABEL": "labeling rules for Copilot output are customer policy decisions",
    "CLM-ASSURANCE": "the promised protection outcome is not established by the evidence",
    "CLM-EVENTS-LOGGED": "Copilot logging follows from enabled auditing but tenant retrieval was not verified",
}


TOPIC_TITLE = {"DEVICE": "Device management coverage", "MFA": "MFA registration coverage",
               "CONSENT": "Application consent coverage", "PRIV": "Privileged access review",
               "ACCESS_REVIEW": "Access review coverage", "GUEST": "Guest invitation policy",
               "RISK": "Risky user remediation", "LICENSING": "Copilot licensing"}


def _title(ctx, f):
    feature = ctx.clean(f.get("Feature")) or f.get("RecommendationId")
    if LICENSE_LABEL.match(feature or "") or feature == f.get("OriginalCompatibilityLabel"):
        alt = ctx.clean(((ctx.pkg.workbook or {}).get("finding_titles") or {}).get(f.get("RecommendationId")))
        if alt and not LICENSE_LABEL.match(alt):
            return alt
        return TOPIC_TITLE.get(topic_for(f), feature)
    return feature


def customer_action(text_value):
    """Drop collector-specific instructions from a script recommendation."""
    text_value = re.sub(r"(?i)^(re)?run with [^,]*,\s*or\s+", "", (text_value or "").strip())
    text_value = text_value[:1].upper() + text_value[1:]
    sentences = re.split(r"(?<=[.!?])\s+", text_value)
    kept = [s for s in sentences if s and not TOOL_SENTENCE.search(s)]
    return " ".join(kept).strip()


def _corrected_recommendation(item):
    topic = item["Topic"]
    return {
        "AUDIT": "Confirm Copilot interactions appear as CopilotInteraction records in an audit search for Copilot "
                 "users, and set audit retention to the duration the customer requires within its licensing.",
        "RETENTION": "Decide retention periods as a records policy, including a policy for the Copilot and AI apps "
                     "location; do not frame retention as AI training control.",
        "LABELS": "Confirm the label taxonomy, publishing scope, and any default, mandatory or automatic labeling "
                  "decisions with the information protection owner.",
        "ENCRYPTION": "Identify which labels apply encryption and confirm they are applied to sensitive content.",
    }.get(topic, "Apply the corrected interpretation and provide the closure evidence listed.")


# ------------------------------------------------------------------ added findings

def _added(ctx, items):
    facts = ctx.facts
    added = []

    def add(fid, topic, title, statements, treatment, priority, interpretation, action, customer=False,
            sufficiency="Limited", benchmarks=()):
        item = _new_item(ctx, fid, "Not assigned", title, topic, "AddedFinding")
        for kind, value in statements:
            _say(item, kind, value)
        _apply(item, "AddMissingFinding", treatment,
               "No script finding covers this condition; the review adds it from retained evidence.", priority)
        item["EvidenceSufficiency"] = sufficiency
        item["CorrectedInterpretation"] = interpretation
        item["RecommendedAction"] = action
        item["CustomerDecisionRequired"] = customer
        item["Evidence"] = {"NativeSource": "Retained collection (aggregate counts)", "LogicalDataset": "See statements",
                            "CollectorSourceId": "Multiple", "CollectionState": "See statements",
                            "ObservationDate": ctx.identity.get("EvaluationDate"), "AgeDays": None,
                            "Freshness": "current", "Population": "Tenant", "RecordCount": None,
                            "DistinctEntityCount": None, "EvidenceLevel": "configuration",
                            "Qualification": "Reviewer-added finding", "MissingEvidence": []}
        _bench(item, *benchmarks)
        _finish(ctx, item)
        added.append(item)

    script_decision = str(ctx.snap.get("decision") or "")
    seats = facts.get("copilot_licenses") or {}
    usage = facts.get("copilot_usage") or {}
    if (seats.get("Assigned") or usage.get("Paid Copilot active users")) and "not ready" in script_decision.lower():
        add("RVW-EXISTING-COPILOT-USE", "SCOPE", "Copilot is already licensed and in use before a readiness decision",
            [(FACT, f"{seats.get('Assigned', 0)} paid Copilot seats are assigned"
              + (f"; usage reports show {usage.get('Paid Copilot active users')} active users"
                 if usage.get("Paid Copilot active users") is not None else "") + "."),
             (SCRIPT, f"The script decision is '{script_decision}'."),
             (CUSTOMER, "Whether current Copilot use continues while foundation gaps are addressed is a customer "
              "decision.")],
            "ConfirmBeforePilot", "High",
            "The readiness decision applies to people already using Copilot, not only to a future rollout.",
            "Decide whether existing Copilot users continue, and confirm they are covered by MFA enforcement and the "
            "oversharing review before broader use.", True, "SufficientWithQualification", ("MS-COPILOT-ZT-01",))
    ga = facts.get("global_admins") or {}
    if (ga.get("DistinctPrincipals") or 0) >= 5:
        add("RVW-GLOBAL-ADMIN-COUNT", "PRIV", "Global Administrator assignments exceed Microsoft guidance",
            [(FACT, f"{ga['ActiveAssignments']} active Global Administrator assignments to {ga['DistinctPrincipals']} "
              "distinct principals were retained."),
             (MSDOC, "Microsoft recommends fewer than five Global Administrators (MS-ENTRA-ROLES-01).")],
            "TenantHardening", "High",
            "Global Administrator assignments are well above Microsoft's recommended maximum; population "
            "completeness is not established, so the true count may be higher.",
            "Reduce Global Administrator assignments to the minimum, keep two emergency access accounts, and move "
            "others to least-privileged roles.", False, "SufficientWithQualification", ("MS-ENTRA-ROLES-01",))
    ca = facts.get("ca")
    if ca is not None:
        exclusions = ", ".join(str(n) for n in ca["AllUsersMfaExclusionCounts"]) or "none recorded"
        add("RVW-EMERGENCY-ACCESS", "PRIV", "Emergency access accounts are not designated in the evidence",
            [(GAP, "The package does not identify emergency access accounts or show their monitoring."),
             (FACT, f"All-user MFA policies record user exclusion counts of: {exclusions}."),
             (MSDOC, "Microsoft recommends at least two cloud-only emergency access accounts with phishing-resistant "
              "authentication, excluded from blocking policies and validated every 90 days (MS-ENTRA-EMERGENCY-01).")],
            "EvidenceRequest", "Medium",
            "Policy exclusions cannot be judged until emergency access accounts are designated.",
            "Designate emergency access accounts, confirm each Conditional Access exclusion, and record monitoring.",
            True, "Missing", ("MS-ENTRA-EMERGENCY-01",))
        has_priv_mfa = any(i["Topic"] == "PRIV" and re.search(
            r"(?i)\bmfa\b|multifactor", i["Title"] + " " + ((i["ScriptConclusion"] or {}).get("Observation") or ""))
            for i in items)
        if not ca["EnforcedAdminMfa"] and not ca["EnforcedAllUsersMfa"] and not has_priv_mfa:
            add("RVW-ADMIN-MFA", "PRIV", "No enforced MFA policy for administrator roles",
                [(FACT, f"Enforced administrator MFA policies: 0; report-only or disabled: {ca['ReportOnlyOrDisabledAdminMfa']}.")],
                "ConfirmBeforePilot", "High", "Administrators are not covered by an enforced MFA policy.",
                "Enforce MFA for administrator roles.", False, "SufficientWithQualification", ("MS-ENTRA-ROLES-01",))
    mde = facts.get("defender_devices")
    recon = facts.get("device_reconciliation") or {}
    if mde is not None and (mde["Active"] == 0 or (recon.get("known_active") or 0) > mde["Onboarded"]):
        add("RVW-ENDPOINT-PROTECTION", "MDE", "Endpoint protection coverage is not established",
            [(FACT, f"Defender for Endpoint returned {mde['Records']} device(s): {mde['Onboarded']} onboarded, "
              f"{mde['Active']} active, {mde['Inactive']} inactive; newest last-seen age "
              f"{mde['NewestLastSeenAgeDays']} days."),
             (FACT, f"Device reconciliation counts {recon.get('known_active')} known active devices, "
              f"{recon.get('onboarded')} onboarded and {recon.get('reporting')} reporting."),
             (MSDOC, "Devices not sending signals for more than seven days are Inactive (MS-MDE-HEALTH-01).")],
            "ConfirmBeforePilot", "High",
            "Endpoint protection health for devices used with Copilot is not demonstrated.",
            "Confirm the device estate, then onboard and verify the devices in Defender for Endpoint.",
            False, "Limited", ("MS-MDE-HEALTH-01", "MS-COPILOT-ZT-01"))
    inc = facts.get("incidents")
    if inc and inc["ActiveNoUpdateSinceCreation"]:
        add("RVW-INCIDENT-EXCEPTION-PROCESS", "EXCEPTIONS", "Incident triage and exception handling are not evidenced",
            [(FACT, f"{inc['ActiveNoUpdateSinceCreation']} of {inc['Active']} active incidents have no recorded update "
              f"since creation; the oldest is {inc['OldestActiveAgeDays']} days old."),
             (GAP, "No exception register or approved treatment is recorded in the package."),
             (CUSTOMER, "The incident ownership model and exception approval process are customer decisions.")],
            "TenantHardening", "Medium",
            "Security operations do not show routine triage, and accepted risks have no recorded exception process.",
            "Assign incident queue ownership and a triage cadence, and create an exception register with role, scope, "
            "compensating control, expiry and approval reference.", True, "SufficientWithQualification",
            ("MS-XDR-INCIDENTS-01",))
    pp = facts.get("power_platform") or {}
    if pp and pp.get("Availability") not in ("available",) and pp.get("CopilotStudioLicensed"):
        add("RVW-AGENT-INVENTORY", "AGENTS", "AI agent and Copilot Studio inventory is not established",
            [(FACT, f"Copilot Studio licensing is present; Power Platform inventory is {pp.get('Availability')}."),
             (MSDOC, "Admins manage which agents are allowed in the Agent Registry; users can only access agents the "
              "admin allows (MS-AGENTS-ADMIN-01).")],
            "EvidenceRequest", "Medium", "Agents and AI applications in use cannot be reviewed or approved.",
            "Export the agent inventory from the Microsoft 365 admin center and Power Platform with an administrator "
            "role, and record approval status.", True, "Missing", ("MS-AGENTS-ADMIN-01",))
    shadow = facts.get("shadow_ai") or {}
    if shadow.get("Source") in ("CompleteEmpty", "Partial", "Unknown", "NotRequested"):
        add("RVW-SHADOW-AI-VISIBILITY", "SHADOW_AI", "Shadow AI visibility is limited",
            [(FACT, f"AI app discovery source is {shadow.get('Source')}."
              + (" " + ctx.clean(shadow.get("Reason")) if shadow.get("Reason") else "")),
             (MSDOC, "Visibility into third-party AI sites requires the Purview browser extension and device "
              "onboarding (MS-PURVIEW-DSPMAI-01).")],
            "EvidenceRequest", "Low",
            "An empty discovery result does not establish that unapproved AI services are not used.",
            "After device onboarding, enable AI activity discovery and record its coverage.", False, "Limited",
            ("MS-PURVIEW-DSPMAI-01",))
    covered_steps = {i["DependencyStep"] for i in items + added}
    audit = facts.get("audit") or {}
    if 13 not in covered_steps:
        add("RVW-AI-INVESTIGATION", "AI_INVESTIGATION", "AI interaction investigation requirements are not recorded",
            [(FACT, f"The Copilot interaction audit query is {audit.get('CopilotInteractionSource', 'Unknown')}."),
             (MSDOC, "Copilot interactions are recorded as CopilotInteraction audit records, and retention for Copilot "
              "and AI apps is a separate location (MS-PURVIEW-AUDIT-COPILOT-01, MS-PURVIEW-RET-COPILOT-01)."),
             (CUSTOMER, "Who may investigate prompts and responses, and for how long they are retained, are customer "
              "decisions.")],
            "EvidenceRequest", "Medium", "Investigation of Copilot interactions has not been tested or governed.",
            "Run a CopilotInteraction audit search for current Copilot users and record the investigation and "
            "retention decision.",
            True, "Missing", ("MS-PURVIEW-AUDIT-COPILOT-01",))
    covered_steps = {i["DependencyStep"] for i in items + added}
    # Step 15 (adoption and pilot measurement) is outside the scope of this security review.
    generic = {"AGENT_LIFECYCLE": 14, "AGENTS": 10, "COPILOT_DLP": 9, "AI_DATA_RISK": 12,
               "EXCEPTIONS": 8, "CONNECTORS": 7, "AUDIT": 6, "LABELS": 5, "PERMISSIONS": 4, "DEVICE": 3,
               "CA": 2, "LICENSING": 1, "SHADOW_AI": 11, "AI_INVESTIGATION": 13}
    for topic, step in sorted(generic.items(), key=lambda kv: kv[1]):
        if step in covered_steps:
            continue
        area = DEPENDENCY_STEPS[step][1]
        add(f"RVW-STEP-{step:02d}", topic, f"{area} not assessed",
            [(GAP, f"No script finding or retained evidence addresses {area.lower()}.")],
            "EvidenceRequest", "Low", f"{area} has not been assessed.",
            "Supply evidence or record a decision for this area before reassessment.", False, "Missing")
        covered_steps.add(step)
    return added


# ------------------------------------------------------------------ controls and readiness

def _controls(ctx, items):
    by_id = {i["FindingId"]: i for i in items}
    results = {r.get("Control ID"): r for r in ctx.snap.get("control_results") or []}
    out = []
    order = ["BlocksPilot", "ConfirmBeforePilot", "TenantHardening", "EvidenceRequest", "LongerTermImprovement",
             "InventoryHygiene", "NoAction"]
    for cid, control in sorted(ctx.controls.items()):
        members = [by_id[r] for r in control.get("recommendation_ids") or [] if r in by_id]
        treatments = [m["ReadinessTreatment"] for m in members if m["ReviewDecision"] != "DuplicateOrOverlapping"]
        review_status = next((t for t in order if t in treatments), "NoAction") if treatments else "NotReviewed"
        result = results.get(cid) or {}
        notes = []
        if str(result.get("Status")) == "Fail" and members and all(
                m["ReviewDecision"] in ("DowngradeToEvidenceGap", "InsufficientEvidence") for m in members):
            notes.append("Script failure rests only on missing evidence.")
        corrected = [m["FindingId"] for m in members if m["ReviewDecision"] in ("CorrectFinding", "CorrectRecommendation")]
        if corrected:
            notes.append("Corrected findings: " + ", ".join(corrected) + ".")
        aliases = sorted({m["ScriptConclusion"]["ScriptControlId"] for m in members if m["ScriptConclusion"]
                          and m["ScriptConclusion"]["ScriptControlId"] not in (None, cid)})
        if aliases:
            notes.append("Findings carry legacy control IDs: " + ", ".join(aliases) + ".")
        out.append({"ControlId": cid, "Title": control.get("title"), "SecurityGate": bool(control.get("security_gate")),
                    "ScriptStatus": result.get("Status") or control.get("status"),
                    "ScriptConfigurationResult": control.get("configuration_result"),
                    "ScriptOperationalResult": control.get("operational_result"),
                    "FindingIds": [m["FindingId"] for m in members],
                    "ReviewTreatment": review_status, "ReviewNotes": " ".join(notes) or "None"})
    return out


def _readiness(ctx, items):
    active = [i for i in items if i["ReviewDecision"] != "DuplicateOrOverlapping"]
    blockers = [i["FindingId"] for i in active if i["ReadinessTreatment"] == "BlocksPilot"]
    confirm = [i["FindingId"] for i in active if i["ReadinessTreatment"] == "ConfirmBeforePilot"]
    if blockers:
        decision = "Not ready: fix the blocking items first"
    elif confirm:
        decision = "Not ready until the pre-rollout checks are done"
    elif any(i["ReadinessTreatment"] == "EvidenceRequest" for i in active):
        decision = "Ready with conditions"
    else:
        decision = "Ready"
    script_decision = ctx.snap.get("decision")
    script_blockers = [b.get("RecommendationId") for b in ctx.snap.get("blockers") or []]
    agrees = ("not ready" in str(script_decision).lower()) == decision.lower().startswith("not ready")
    by_id = {i["FindingId"]: i for i in items}
    kept = [b for b in script_blockers if b in by_id and by_id[b]["ReadinessTreatment"] == "BlocksPilot"]
    changed = [f"{b} is now {GROUP[by_id[b]['ReadinessTreatment']]}" for b in script_blockers
               if b in by_id and by_id[b]["ReadinessTreatment"] != "BlocksPilot"]
    basis = (f"The review {'agrees' if agrees else 'disagrees'} with the original assessment that the tenant is "
             f"{'not ready' if decision.lower().startswith('not ready') else 'ready'} for Copilot. ")
    if script_blockers:
        basis += (f"The original assessment named {len(script_blockers)} blockers; the review confirms "
                  f"{len(kept)}" + (f" ({', '.join(kept)})" if kept else "") + ".")
    if changed:
        basis += " The others were re-rated after checking the evidence: " + "; ".join(changed) + "."
    return {"ScriptDecision": script_decision, "ScriptRationale": ctx.clean(ctx.snap.get("rationale")),
            "ScriptBlockers": script_blockers, "ReviewDecision": decision, "AgreesWithScript": agrees,
            "Basis": basis.strip(), "ConfirmedBlockers": blockers, "ConfirmBeforePilot": confirm}


def load_benchmarks(path=None):
    path = Path(path) if path else Path(__file__).with_name("benchmarks.json")
    data = json.loads(path.read_text(encoding="utf-8"))
    required = ("BenchmarkId", "Title", "Url", "VerifiedOn", "Product", "RequirementType", "Statement",
                "MappingRationale", "LicensingDependency", "Limitation")
    for b in data["Benchmarks"]:
        missing = [k for k in required if not b.get(k)]
        if missing:
            raise ValueError(f"Benchmark {b.get('BenchmarkId')} missing {missing}")
        if not b["Url"].startswith("https://learn.microsoft.com/"):
            raise ValueError(f"Benchmark {b['BenchmarkId']} is not a Microsoft Learn URL")
        if b["RequirementType"] not in data["RequirementTypes"]:
            raise ValueError(f"Benchmark {b['BenchmarkId']} has unknown type")
    return data


def run_review(pkg, review_date, benchmarks=None):
    benchmarks = benchmarks or load_benchmarks()
    ctx = Context(pkg, benchmarks, review_date)
    items = [review_finding(ctx, f) for f in ctx.findings]
    items += _added(ctx, items)
    for item in items:
        unknown = [b for b in item["MicrosoftBenchmarkIds"] if b not in ctx.bench_by_id]
        if unknown:
            raise ValueError(f"{item['FindingId']} references unknown benchmarks {unknown}")
        item.pop("_subject", None)
    items.sort(key=lambda i: (i["DependencyStep"], TREATMENTS.index(i["ReadinessTreatment"]), i["FindingId"]))
    return ctx, items, _controls(ctx, items), _readiness(ctx, items)
