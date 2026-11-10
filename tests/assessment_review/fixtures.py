"""Builds fictional completed assessment packages for automated tests.

No customer data is used. Names, tenants and identifiers are invented.
"""
from __future__ import annotations

import copy
from hashlib import sha256
import json
from pathlib import Path

from openpyxl import Workbook

CUSTOMERS = {
    "northwind": {"name": "Northwind Traders", "tenant": "11111111-2222-4333-8444-555555555555",
                  "assessment": "AST-" + "a1" * 16, "run": "RUN-" + "b2" * 16, "env": "ENV-" + "c3" * 32,
                  "domain": "northwindtraders.example", "people": ["Jordan Avery", "Priya Natarajan", "Sam Okafor"],
                  "incident": "Suspicious inbox rule on finance mailbox", "policy": "Northwind Finance Restricted"},
    "tailspin": {"name": "Tailspin Toys", "tenant": "66666666-7777-4888-9999-000000000000",
                 "assessment": "AST-" + "d4" * 16, "run": "RUN-" + "e5" * 16, "env": "ENV-" + "f6" * 32,
                 "domain": "tailspintoys.example", "people": ["Morgan Lee", "Alex Romero", "Taylor Chen"],
                 "incident": "Credential phishing campaign targeting sales", "policy": "Tailspin Design Vault"},
}


GA = "62e90394-69f5-4237-9190-012177145e10"
COPILOT_SKU = "639dec6b-bb19-468b-871c-c5c441c4b0cb"
# A service principal secret value that must never reach any output.
SECRET_VALUE = "Zq8Q~fixtureSecretValueThatMustNeverAppear01"


def _sha(path):
    return sha256(Path(path).read_bytes()).hexdigest()


def _write_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2), encoding="utf-8")


def _finding(rid, key, control, status, disposition, observation, recommendation, *, priority="High",
             action_type=None, pilot=None, evidence_id=None, evidence_status="supported", level="configuration",
             source=None, source_type="tenant_collection", freshness="current", obs_date="2026-03-01",
             qualification="", confidence="High", feature=None, complete=True, scope="Assessed tenant",
             basis="Tenant evidence", completion="A dated validation of the affected population."):
    finding = {
        "RecommendationId": rid, "FindingKey": key, "ControlId": control, "Status": status,
        "Disposition": disposition, "Priority": priority, "Observation": observation,
        "Recommendation": recommendation, "Feature": feature or rid, "OwnerRole": "Identity and access administrator",
        "EvidenceId": evidence_id, "RelatedEvidenceIds": [evidence_id] if evidence_id else [],
        "EvidenceStatus": evidence_status, "EvidenceLevel": level, "EvidenceSource": source,
        "SourceType": source_type, "Freshness": freshness, "ObservationDate": obs_date,
        "SourceObservedAt": obs_date + "T10:00:00+00:00" if obs_date and len(obs_date) == 10 else obs_date,
        "Qualification": qualification, "Confidence": confidence, "EvidenceComplete": complete,
        "EvidenceScope": scope, "EvidenceBasis": basis, "CompletionEvidence": completion,
        "SourceOccurrences": [], "EvidenceKey": (key or rid).split(".")[0] + "_detail",
    }
    if action_type:
        finding["ActionType"] = action_type
        finding["PilotImpact"] = pilot or ""
        finding["ReadinessStage"] = "Before pilot"
    return finding


def build_package(root, customer="northwind", *, pilot_scope=False, enforced_targeted_mfa=1,
                  consent_policies=("microsoft-user-default-recommended", "microsoft-user-default-allow-consent-apps"),
                  ca_state="available", managed_devices_state="empty", truncated_signins=False,
                  extra_finding_mutation=None):
    c = CUSTOMERS[customer]
    root = Path(root)
    root.mkdir(parents=True, exist_ok=True)
    eval_date = "2026-03-02"
    ident = {"SchemaVersion": "1.0.0", "State": "complete", "AssessmentId": c["assessment"], "RunId": c["run"],
             "PrimaryEnvironmentId": c["env"], "MethodologyVersion": "4.0.0", "CatalogVersion": "catalog-test",
             "EvaluatedAt": eval_date, "CreatedAt": eval_date + "T09:00:00+00:00", "RunType": "Initial",
             "BaselineRunId": None, "Entities": [], "Aliases": [], "References": []}
    run_context = {k: ident[k] for k in ("AssessmentId", "RunId", "PrimaryEnvironmentId", "RunType", "BaselineRunId",
                                        "EvaluatedAt", "CreatedAt", "MethodologyVersion", "CatalogVersion")}
    producing = {"commit_sha": "0" * 40, "source_state": "CleanCommitted", "repository_identifier": "example/engine",
                 "publication": {"customer_publication_eligible": True}}
    people = c["people"]
    upn = lambda n: n.lower().replace(" ", ".") + "@" + c["domain"]

    evidence = []

    def ev(eid, metric, source_file="", affected=None, observed=eval_date):
        evidence.append({"evidence_id": eid, "metric_id": metric, "source_file": source_file,
                         "affected_objects": affected or [], "observed_at": observed,
                         "source_observed_at": observed + "T10:00:00+00:00", "availability": "available",
                         "freshness": "current", "AssessmentId": c["assessment"], "RunId": c["run"],
                         "qualification": "", "truncated": False})
        return eid

    findings = [
        _finding("ENT-005", "entra.authentication.mfa_registration", "IDENTITY-001", "Action Required", "Action",
                 "4 of 10 users with known registration flags (40.0%) are MFA registered. Registration is separate "
                 "from policy enforcement and observed authentication.",
                 "Arrange registration for unregistered users. Verify MFA enforcement separately.",
                 action_type="Remediation", pilot="Blocks pilot",
                 evidence_id=ev("EV-0000000000000001", "entra.mfa", "https://graph.microsoft.com/v1.0/reports/x",
                                ["user:" + upn(people[0])]),
                 source="entra_auth_methods", level="configuration", feature="Microsoft Entra ID P1"),
        _finding("ENT-020", "baseline.identity.sign_in", "IDENTITY.AUTH", "Action Required", "Action",
                 "No enforced Conditional Access policy requires MFA for all users and cloud apps, and security "
                 "defaults are off.", "Create a Conditional Access policy that requires MFA for all users.",
                 action_type="Remediation", pilot="Blocks pilot",
                 evidence_id=ev("EV-0000000000000002", "baseline.identity.sign_in",
                                "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"),
                 source="ca_policies", level="policy_enforcement",
                 feature="Require MFA and block legacy sign-in for all users"),
        _finding("ENT-017", "entra.apps.user_consent_assignment", "APPS-001", "Action Required", "Action",
                 "The authorization policy assigns 2 self-consent policies to default users.",
                 "Review the assigned consent policy conditions.", action_type="Remediation", pilot="Blocks pilot",
                 evidence_id=ev("EV-0000000000000003", "entra.consent"), source="authorization_policy",
                 feature="Microsoft Entra ID P2"),
        _finding("DEF-013", "defender.incidents.current", "THREAT.INCIDENTS", "Action Required", "Action",
                 "The incident query returned 3 active incident(s), including 1 with high severity. Incident: "
                 + c["incident"] + ".", "Review and resolve the active incidents.",
                 action_type="Remediation", pilot="Blocks pilot",
                 evidence_id=ev("EV-0000000000000004", "defender.incidents", "https://api.security.microsoft.com/x"),
                 source="defender_incidents", feature="Review active security incidents"),
        _finding("PUR-040", "baseline.dlp.enforced", "DATA.DLP", "Success", "Assurance",
                 "2 of 3 DLP policies are configured in enforcement mode; 1 enforced policy applies to Microsoft 365 "
                 "Copilot ('" + c["policy"] + " Copilot').", "", priority="",
                 evidence_id=ev("EV-0000000000000005", "purview.dlp"), source="purview_dlp_policies",
                 level="policy_enforcement",
                 qualification="Configuration component only; operational confirmation remains unresolved.",
                 feature="DLP policies configured in enforcement mode"),
        _finding("M365-129", "baseline.license.capacity", "LICENSE.ASSIGNMENT", "Success", "Assurance",
                 "The tenant has 4 Microsoft 365 Copilot license(s): 2 assigned and 2 unassigned.", "", priority="",
                 evidence_id=ev("EV-0000000000000006", "m365.license"), source="m365_license_coverage",
                 evidence_status="limited", feature="Copilot licenses are available"),
        _finding("GAP-LICENSE-ASSIGNMENT", "coverage.license.assignment", "LICENSE.ASSIGNMENT", "Not Assessed",
                 "Coverage", "Copilot license information was not collected.",
                 "Collect it with --reports-dir or check the admin center.", priority="Medium",
                 action_type="Evidence", pilot="Confirm before pilot", evidence_status="gap",
                 source_type="assessment_requirement", obs_date="",
                 qualification="This check could not be answered from the collected evidence; it is not a measured "
                               "control failure.", confidence="Unknown", complete=False,
                 feature="Confirm Copilot licenses are available"),
        _finding("DEX-001", "data_exposure.broad_internal_access", "DATA-001", "Action Required", "Action",
                 "The selected reports record 5 permissions granted to Everyone; 12 links accessible to people in "
                 "the organization. These measures can overlap and are not a count of distinct files or users.",
                 "Confirm whether this condition still applies. Have site owners remove broad permissions.",
                 action_type="Confirmation", pilot="Confirm before pilot",
                 evidence_id=ev("EV-0000000000000007", "sam.everyone", "report.csv", ["site:abc"], "2026-01-10"),
                 evidence_status="stale", source_type="portal_export", freshness="stale", obs_date="2026-01-10",
                 qualification="The observation is 51 days old and requires confirmation.",
                 feature="Confirm broad group permissions"),
        _finding("M365-122", "sharepoint.sharing.anyone_enabled", "CONTENT.SHARING", "Attention Required", "Action",
                 "The tenant permits Anyone sharing. Anyone links are the default link type.",
                 "Confirm that Anyone sharing is required.", priority="Medium", action_type="Confirmation",
                 pilot="Confirm before pilot", evidence_id=ev("EV-0000000000000008", "spo.sharing"),
                 evidence_status="unknown", feature="Confirm access restrictions for business content"),
        _finding("PUR-022", "purview.retention_labels.deployed", "GOV-001", "Success", "Opportunity",
                 "2 retention policies configured.", "Retain Teams meeting transcripts 3 years (preserve Copilot "
                 "training context).", priority="Low", evidence_id=ev("EV-0000000000000009", "purview.retention"),
                 source="purview_retention_policies", evidence_status="limited",
                 feature="Information Governance - Retention Strategy"),
        _finding("PUR-029", "purview.audit.state", "GOV-001", "Success", "Opportunity",
                 "Unified Audit Log is ENABLED - Copilot events are being logged",
                 "Check audit log retention set to 10 years and confirm PromptSubmitted events are captured.",
                 priority="Low", evidence_id=ev("EV-0000000000000010", "purview.audit"),
                 source="purview_audit_config", evidence_status="limited",
                 feature="Microsoft 365 Advanced Auditing"),
        _finding("PUR-033", "purview.sensitivity_labels.deployed", "GOV-001", "Attention Required", "Opportunity",
                 "Purview returned 5 sensitivity label definitions.", "Set default label to General for all Copilot "
                 "outputs.", priority="Medium", evidence_id=ev("EV-0000000000000011", "purview.labels"),
                 source="purview_sensitivity_labels", evidence_status="limited", feature="Label deployment"),
        _finding("GAP-DATA-LABELS", "coverage.data.labels", "DATA.LABELS", "Not Assessed", "Coverage",
                 "Sensitivity label definitions were not collected.", "Grant permissions or confirm in the portal.",
                 priority="Medium", action_type="Evidence", pilot="Confirm before pilot", evidence_status="gap",
                 source_type="assessment_requirement", obs_date="", confidence="Unknown", complete=False,
                 feature="Confirm sensitivity label definitions"),
        _finding("GAP-DATA-AUDIT", "coverage.operation.data.audit", "DATA.AUDIT", "Not Assessed", "Coverage",
                 "Operational behavior is not established by configuration alone.", "Provide operational records.",
                 priority="Medium", action_type="Evidence", pilot="Confirm before pilot", evidence_status="gap",
                 source_type="assessment_requirement", obs_date="", confidence="Unknown", complete=False,
                 feature="Confirm audit coverage for investigation"),
        _finding("ENT-013", None, "IDENTITY-001", "Action Required", "Action",
                 "2 risky users detected in the tenant (0 high-risk, 2 medium-risk)",
                 "Review the risky users and confirm remediation.", priority="Medium", action_type="Confirmation",
                 pilot="Confirm before pilot", evidence_id=ev("EV-0000000000000012", "entra.risk"),
                 evidence_status="unknown",
                 qualification="Source association is missing; confirm the retained dataset for this observation.",
                 feature="Confirm remediation of identity risk detections"),
        _finding("ENT-002", None, "EGRESS-001", "Not Assessed", "Coverage",
                 "The Intune query returned no managed devices.", "Confirm the tenant's endpoint control model.",
                 priority="Medium", action_type="Evidence", pilot="Confirm before pilot",
                 evidence_id=ev("EV-0000000000000013", "intune.devices"), evidence_status="gap",
                 feature="Microsoft Intune (Plan A)"),
        _finding("DEX-003", "data_exposure.coverage.sensitive-data_exposure_evidence", "DATA-001", "Not Assessed",
                 "Coverage", "Microsoft Purview DSPM data-risk assessment evidence was not available.",
                 "Export a completed DSPM data risk assessment.", action_type="Evidence",
                 pilot="Confirm before pilot", evidence_id=ev("EV-0000000000000014", "dspm"),
                 evidence_status="gap", source_type="portal_export", feature="Sensitive-data exposure evidence"),
    ]
    if extra_finding_mutation:
        extra_finding_mutation(findings, evidence)
    actions = [copy.deepcopy(f) for f in findings if f.get("ActionType")]
    blockers = [a for a in actions if a.get("PilotImpact") == "Blocks pilot"]
    controls = [
        {"control_id": "IDENTITY.MFA", "title": "MFA registration", "status": "Action required", "security_gate": True,
         "configuration_result": "not_established", "operational_result": "not_applicable",
         "evidence_ids": ["EV-0000000000000001"], "recommendation_ids": ["ENT-005"]},
        {"control_id": "IDENTITY.AUTH", "title": "Require MFA", "status": "Action required", "security_gate": True,
         "configuration_result": "not_established", "operational_result": "unknown",
         "evidence_ids": ["EV-0000000000000002"], "recommendation_ids": ["ENT-020"]},
        {"control_id": "APPS.CONSENT", "title": "Consent", "status": "Action required", "security_gate": True,
         "configuration_result": "not_established", "operational_result": "not_applicable",
         "evidence_ids": [], "recommendation_ids": ["ENT-017"]},
        {"control_id": "DATA.LABELS", "title": "Labels", "status": "Not established", "security_gate": True,
         "configuration_result": "not_established", "operational_result": "not_applicable",
         "evidence_ids": [], "recommendation_ids": ["PUR-033", "GAP-DATA-LABELS"]},
    ]
    control_results = [{"Control ID": c_["control_id"], "Status": "Fail" if "Action" in c_["status"] else "Unknown",
                        "Result": "fail"} for c_ in controls]
    states = {
        "ca_policies": {"state": ca_state, "complete": ca_state == "available", "records_collected": 3, "reason": ""},
        "auth_methods": {"state": "available", "complete": True, "records_collected": 10, "reason": ""},
        "authorization_policy": {"state": "available", "complete": True, "records_collected": 1, "reason": ""},
        "risky_users": {"state": "available", "complete": True, "records_collected": 3, "reason": ""},
        "managed_devices": {"state": managed_devices_state, "complete": True, "records_collected": 0, "reason": ""},
        "purview_sensitivity_labels": {"state": "available", "complete": True, "records_collected": 5, "reason": ""},
        "purview_audit_config": {"state": "available", "complete": True, "records_collected": 1, "reason": ""},
        "m365_copilot_interaction_audit": {"state": "partial", "complete": False, "records_collected": 0,
                                           "reason": "The audit query did not finish."},
        "shadow_ai_usage": {"state": "empty", "complete": True, "records_collected": 0,
                            "reason": "Returned streams contain no classified AI candidates."},
        "copilot_dashboard": {"state": "not_requested", "complete": False, "records_collected": 0, "reason": ""},
    }
    readiness_review = {"available": False, "status": "not_supplied"}
    if pilot_scope:
        readiness_review = {"available": True, "status": "valid",
                            "pilot_scope": {"id": "pilot", "description": "Finance pilot group", "population_count": 8}}
    snapshot = {
        "schema_version": "1.1.0", "methodology_version": "4.0.0", "evaluation_date": eval_date,
        "tenant_id": c["tenant"], "decision": "Not ready for pilot",
        "rationale": f"{len(blockers)} high-priority tenant conditions must be fixed before a pilot starts.",
        "customer_findings": findings, "actions": actions, "blockers": blockers, "pilot_conditions": [],
        "controls": controls, "control_results": control_results,
        "counts": {"actions": len(actions), "remediation": sum(1 for a in actions if a["ActionType"] == "Remediation")},
        "evidence": evidence, "executive_summary": {"Pilot blockers": [b["Feature"] for b in blockers]},
        "rollout_progress": {"counts": {"pilot_blockers": len(blockers)}},
        "readiness_review": readiness_review,
        "supplemental_reviews": {"copilot_admin_review": {"rows": [
            {"Area": "adoption", "Topic": "Paid Copilot enabled users", "Value": 2},
            {"Area": "adoption", "Topic": "Paid Copilot active users", "Value": 1},
            {"Area": "data_protection", "Topic": "Copilot DLP", "Value": "Monitoring only"}]}},
        "device_reconciliation": {"known_active": 4, "onboarded": 1, "reporting": 0, "active_days": 30, "reporting_days": 7},
        "privileged_assessment": {"diagnostics": []}, "identity_taxonomy": {"diagnostics": []},
        "collection_source_states": states, "collection_thresholds": {"review_freshness_days": 35},
        "evidence_period": {"start": "2026-01-10", "end": eval_date}, "producing_revision": producing,
        "identity": ident, "run_context": run_context, "identity_validation": [],
    }
    snap_path = root / "assessment-results" / f"{c['run']}.json"
    _write_json(snap_path, snapshot)
    ca_policies = [
        {"displayName": c["policy"] + " MFA all users", "state": "enabledForReportingButNotEnforced",
         "conditions": {"users": {"includeUsers": ["All"], "excludeUsers": ["x"]}, "applications": {"includeApplications": ["All"]},
                        "clientAppTypes": ["all"]}, "grantControls": {"builtInControls": ["mfa"]}},
        {"displayName": "Block legacy", "state": "enabled",
         "conditions": {"users": {"includeUsers": ["All"]}, "applications": {"includeApplications": ["All"]},
                        "clientAppTypes": ["exchangeActiveSync", "other"]}, "grantControls": {"builtInControls": ["block"]}},
    ]
    if enforced_targeted_mfa:
        ca_policies.append({"displayName": "Finance MFA", "state": "enabled",
                            "conditions": {"users": {"includeGroups": ["g1"]}, "applications": {"includeApplications": ["All"]},
                                           "clientAppTypes": ["all"]}, "grantControls": {"builtInControls": ["mfa"]}})
    uid = lambda i: f"00000000-0000-4000-8000-{i:012d}"
    registration = [{"id": uid(i), "userDisplayName": people[i % 3], "userPrincipalName": upn(people[i % 3]),
                     "userType": "member", "isMfaRegistered": i < 4, "isAdmin": i in (0, 5),
                     "methodsRegistered": ["mobilePhone"] if i < 2 else ["microsoftAuthenticatorPush"] if i < 4 else []}
                    for i in range(10)]
    users = [{"id": uid(i), "displayName": people[i % 3], "userPrincipalName": upn(people[i % 3]),
              "userType": "Member", "accountEnabled": i != 7,
              "assignedLicenses": [{"skuId": COPILOT_SKU}] if i in (0, 1) else []} for i in range(10)]

    def user_principal(i):
        return {"@odata.type": "#microsoft.graph.user", "id": uid(i), "userPrincipalName": upn(people[i % 3]),
                "displayName": people[i % 3], "accountEnabled": i != 7, "userType": "Member"}
    sync_app = {"@odata.type": "#microsoft.graph.servicePrincipal", "id": "sp-admin", "appId": "app-admin",
                "displayName": c["name"] + " Directory Sync", "accountEnabled": True,
                "passwordCredentials": [{"endDateTime": "2027-01-01T00:00:00Z", "hint": "Xy1",
                                         "secretText": SECRET_VALUE}], "keyCredentials": []}
    role_assignments = [{"roleDefinitionId": GA, "principalId": uid(i), "principal": user_principal(i)}
                        for i in (0, 5, 7, 8, 9)]
    role_assignments.append({"roleDefinitionId": GA, "principalId": "sp-admin", "principal": sync_app})
    inputs = root / "Collected" / "1" / "inputs"
    dag = inputs / "1" / "Permissioned_users_count_SharePoint_report.csv"
    dag.parent.mkdir(parents=True, exist_ok=True)
    dag.write_text("Tenant ID,Site ID,Site name,Site URL,Site template,Primary admin,Primary admin email,"
                   "External sharing,Site privacy,Site sensitivity,Number of users having access,PeopleInYourOrg link "
                   "count,Anyone link count,EEEU permission count,Everyone permission count,Report date\n"
                   f"{c['tenant']},site-a,{c['name']} Finance,https://{c['domain'].split('.')[0]}.sharepoint.com/"
                   f"sites/Finance,Team site,Finance Owners,{upn(people[0])},true,Public,,40,3,0,1,0,"
                   "2026-01-10T10:00:00Z\n", encoding="utf-8")
    eeeu = inputs / "2" / "Detailed_eeeu_report.csv"
    eeeu.parent.mkdir(parents=True, exist_ok=True)
    eeeu.write_text("Tenant ID,Site ID,Item type,Item url,Role definition,Recipient,Parent group name,"
                    "Parent group type,Report date\n"
                    f"{c['tenant']},site-a,Web,sites/Finance,Edit,Everyone except external users,Finance Members,"
                    "SharePointGroup,2026-01-10T10:00:00Z\n", encoding="utf-8")
    sites = [{"Url": f"https://{c['domain'].split('.')[0]}.sharepoint.com/sites/Finance", "Title": c["name"] + " Finance",
              "Template": "GROUP#0", "Owner": "a1b2c3d4-0000-4000-8000-000000000001_o", "SiteId": "site-a",
              "SharingCapability": 1, "LastContentModifiedDate": "2026-02-20T00:00:00"},
             {"Url": f"https://{c['domain'].split('.')[0]}.sharepoint.com/sites/Archive", "Title": c["name"] + " Archive",
              "Template": "STS#3", "Owner": upn(people[1]), "SiteId": "site-b", "SharingCapability": 2,
              "LastContentModifiedDate": "2024-01-15T00:00:00"}]
    usage = [{"Site Id": "site-a", "Last Activity Date": "2026-02-28", "Storage Used (Byte)": "5242880",
              "File Count": "12"},
             {"Site Id": "site-b", "Last Activity Date": "", "Storage Used (Byte)": "1073741824", "File Count": "900"}]
    service_principals = [
        {"id": "sp-graph", "appId": "00000003-0000-0000-c000-000000000000", "displayName": "Microsoft Graph",
         "appOwnerOrganizationId": "f8cdef31-a31e-4b4a-93e4-5f571e91255a", "servicePrincipalType": "Application",
         "verifiedPublisher": {}, "appRoles": [{"id": "role-mail", "value": "Mail.ReadWrite"}]},
        {"id": "sp-mail", "appId": "app-mail", "displayName": c["name"] + " Mail Archiver",
         "appOwnerOrganizationId": c["tenant"], "servicePrincipalType": "Application", "verifiedPublisher": {},
         "appRoles": []},
        {"id": "sp-helper", "appId": "app-helper", "displayName": "Unverified Helper Tool",
         "appOwnerOrganizationId": "99999999-0000-4000-8000-000000000000", "servicePrincipalType": "Application",
         "verifiedPublisher": {}, "appRoles": []}]
    collection = {
        "format": "m365-readiness-collection", "tenant_id": c["tenant"], "tenant_name": c["name"],
        "collected_at": eval_date + "T08:00:00+00:00", "identity": ident,
        "package": {"files": [{"path": f"Collected/1/inputs/{p.parent.name}/{p.name}", "sha256": _sha(p)}
                              for p in (dag, eeeu)]},
        "service_results": {
            "m365_result": [{"licenses": [{"sku_part_number": "Microsoft_365_Copilot", "sku_id": COPILOT_SKU,
                                           "enabled": 4, "consumed": 2}],
                             "_client": {"$collection_attributes": {
                                 "users": users,
                                 "sharepoint_governance": {"sites": {"available": True, "records_collected": 2,
                                                                     "items": sites}},
                                 "assessment_datasets": {"sharepoint_usage": [{"records": usage, "source": {}}]}}}}],
            "entra_info": {"verified_domains": [c["domain"]], "_client": {"$collection_attributes": {
                "ca_policies": ca_policies, "security_defaults": {"available": True, "is_enabled": False},
                "auth_methods_registration": registration,
                "service_principals": service_principals,
                "oauth_permission_grants": [{"clientId": "sp-helper", "consentType": "AllPrincipals",
                                             "resourceId": "sp-graph", "scope": "User.Read Files.ReadWrite.All"}],
                "guest_users": [{"id": "guest-1", "displayName": "Partner Reviewer",
                                 "userPrincipalName": "partner_reviewer#EXT#@" + c["domain"],
                                 "createdDateTime": "2025-01-05T00:00:00Z"}],
                "assessment_datasets": {
                    "application_permissions": [{"records": [{"principalId": "sp-mail", "resourceId": "sp-graph",
                                                              "appRoleId": "role-mail",
                                                              "resourceDisplayName": "Microsoft Graph"}],
                                                 "source": {}}],
                    "user_signin_activity": [{"records": [
                        {"id": uid(i), "signInActivity": {"lastSignInDateTime": "2026-02-20T00:00:00Z"}}
                        for i in (0, 1, 2, 5)], "source": {}}],
                    "directory_devices": [{"records": [
                        {"displayName": "fin-laptop-07", "deviceId": "dev-1", "operatingSystem": "Windows",
                         "trustType": "AzureAd", "isManaged": False, "isCompliant": False, "accountEnabled": True,
                         "approximateLastSignInDateTime": "2026-02-25T00:00:00Z"}], "source": {}}]},
                "consent_summary": {"assigned_user_consent_policies": ["ManagePermissionGrantsForSelf." + p for p in consent_policies],
                                    "high_privilege_apps": 2, "unverified_publishers": 1},
                "risk_summary": {"at_risk": 3, "risky_users_high": 0, "risky_users_medium": 2, "risky_users_low": 1,
                                 "risky_users_total": 3, "remediated": 0, "dismissed": 0,
                                 "user_risk_policy_exists": False, "signin_risk_policy_exists": True},
                "risky_users": [{"id": uid(1), "userDisplayName": people[1], "userPrincipalName": upn(people[1]),
                                 "riskLevel": "medium", "riskState": "atRisk", "riskDetail": "none",
                                 "riskLastUpdatedDateTime": "2026-02-20T00:00:00Z"}],
                "signin_logs": [{"userDisplayName": people[2], "userPrincipalName": upn(people[2]), "ipAddress": "203.0.113.7"}],
                "managed_devices": [], "device_summary": {"compliance_policies_total": 2},
                "role_definitions": [{"id": GA, "displayName": "Global Administrator", "templateId": GA}],
                "role_assignments": role_assignments,
                "b2b_summary": {"total_guests": 2, "guest_invite_restrictions": "everyone"},
                "signin_summary": {"total_signins_sampled": 50, "legacy_auth_attempts": 0},
                "collection_status": {"signin_logs": {"availability_status": "available", "records_collected": 50,
                                                      "truncated": truncated_signins, "complete": True}},
            }}},
            "purview_info": {"_client": {"$collection_attributes": {
                "sensitivity_labels": {"available": True, "total_labels": 5, "labels": [{"name": c["policy"] + " Label"}]},
                "label_policies": {"available": True, "total_policies": 1},
                "audit_config": {"available": True, "unified_audit_enabled": True},
                "audit_logs": {"available": False}, "retention_policies": {"policies": [{"name": c["policy"] + " Retention"}]},
            }}},
            "defender_info": {"_client": {"$collection_attributes": {
                "security_incidents": [
                    {"id": "101", "displayName": c["incident"], "status": "active", "severity": "high",
                     "assignedTo": upn(people[0]), "incidentWebUrl": "https://security.microsoft.com/incidents/101",
                     "createdDateTime": "2026-02-01T10:00:00Z", "lastUpdateDateTime": "2026-02-01T10:00:00Z"},
                    {"id": "102", "displayName": "Low alert one", "status": "active", "severity": "low",
                     "assignedTo": None, "incidentWebUrl": "https://security.microsoft.com/incidents/102",
                     "createdDateTime": "2026-02-10T10:00:00Z", "lastUpdateDateTime": "2026-02-11T10:00:00Z"},
                    {"id": "103", "displayName": "Low alert two", "status": "active", "severity": "low",
                     "assignedTo": None, "incidentWebUrl": "https://security.microsoft.com/incidents/103",
                     "createdDateTime": "2026-02-12T10:00:00Z", "lastUpdateDateTime": "2026-02-12T10:00:00Z"}],
                "security_alerts": [{"incidentId": "101", "title": "Suspicious inbox manipulation rule",
                                     "recommendedActions": "Review the inbox rules and reset the account password."}],
                "defender_devices": [{"healthStatus": "Inactive", "onboardingStatus": "Onboarded", "managedBy": "Intune",
                                      "lastSeen": "2026-02-01T00:00:00Z", "computerDnsName": "fin-laptop-07"}],
            }}},
            "power_platform_info": {"availability_status": "unavailable"},
            "copilot_studio_info": {"has_copilot_studio": True, "total_licenses": 1},
        },
    }
    _write_json(root / "collection.json", collection)
    run = {"Identity": ident, "ProducingRevision": producing,
           "SnapshotLocator": f"assessment-results/{c['run']}.json", "SnapshotHash": _sha(snap_path)}
    _write_json(root / "assessment-run.json", run)
    _write_json(root / "run-seed.json", ident)
    _write_json(root / "assessment-history.json", {"AssessmentId": c["assessment"],
                                                   "Runs": [{"RunId": c["run"], "SnapshotHash": run["SnapshotHash"]}]})
    _deliverables(root, c, actions, controls, control_results, eval_date)
    return root


def _deliverables(root, c, actions, controls, control_results, eval_date):
    folder = root / "Deliverables"
    folder.mkdir(parents=True, exist_ok=True)
    stem = f"AI Readiness and M365 Hardening - {c['name']}"
    book = Workbook()
    ws = book.active
    ws.title = "Start Here"
    for row in (("Open", "Purpose"), ("Tenant", c["name"]), ("Evaluation date", eval_date)):
        ws.append(row)
    plan = book.create_sheet("Action Plan")
    plan.append(("Priority", "Pilot impact", "RecommendationId", "Finding", "Evidence"))
    for a in actions:
        plan.append((a["Priority"], a.get("PilotImpact"), a["RecommendationId"], a["Feature"], "'Configuration'!A2:B3"))
    book.create_sheet("Configuration").append(("x",))
    findings = book.create_sheet("Findings")
    findings.append(("ID", "Finding"))
    for a in actions:
        findings.append((a["RecommendationId"], a["Feature"]))
    cr = book.create_sheet("Control Results")
    cr.append(("Control ID", "Status"))
    for r in control_results:
        cr.append((r["Control ID"], r["Status"]))
    integrity = book.create_sheet("Integrity Checks")
    integrity.append(("Status", "Issue"))
    integrity.append(("Passed", "ok"))
    book.save(folder / f"{stem}.xlsx")
    tech = Workbook()
    tech.active.title = "Control Results"
    tech.active.append(("Control ID", "Status"))
    for r in control_results:
        tech.active.append((r["Control ID"], r["Status"]))
    tech.save(folder / f"Technical Evidence - {c['name']}.xlsx")
    html = (f"<html><head><title>{c['name']} - AI Readiness</title></head><body><h1>{c['name']}</h1>"
            f"<p>Not ready for a pilot yet</p><div id='action-1'>Action</div>"
            f"<a href='#action-1'>a</a><a href='{stem.replace(' ', '%20')}.xlsx#%27Configuration%27!A2:B3'>w</a>"
            "</body></html>")
    (folder / f"{stem}.html").write_text(html, encoding="utf-8")
    summary = (f"<html><head><title>{c['name']} - readiness summary</title></head><body><h1>{c['name']}</h1>"
               f"<p>Not ready for a pilot yet</p><a href='{stem.replace(' ', '%20')}.html#action-1'>d</a></body></html>")
    (folder / f"Readiness Summary - {c['name']}.html").write_text(summary, encoding="utf-8")


def refresh_snapshot_hash(root):
    """Re-record the snapshot hash after a deliberate snapshot edit."""
    root = Path(root)
    run = json.loads((root / "assessment-run.json").read_text(encoding="utf-8"))
    snap = root / run["SnapshotLocator"]
    run["SnapshotHash"] = _sha(snap)
    (root / "assessment-run.json").write_text(json.dumps(run, indent=2), encoding="utf-8")
    history = json.loads((root / "assessment-history.json").read_text(encoding="utf-8"))
    history["Runs"][0]["SnapshotHash"] = run["SnapshotHash"]
    (root / "assessment-history.json").write_text(json.dumps(history, indent=2), encoding="utf-8")


def edit_json(path, mutate):
    path = Path(path)
    data = json.loads(path.read_text(encoding="utf-8"))
    mutate(data)
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
