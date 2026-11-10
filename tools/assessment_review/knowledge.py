"""Data-driven review knowledge: topics, dependency order, claim checks, closure.

Mappings key on the assessment engine catalog (RecommendationId, FindingKey,
ControlId) and fall back to text patterns, so a second customer package from the
same engine needs no code changes.
"""
from __future__ import annotations

import re

DEPENDENCY_STEPS = {
    1: ("Foundational", "Licensing, prerequisites and approved AI services"),
    2: ("Foundational", "Identity, MFA, Conditional Access and privileged access"),
    3: ("Foundational", "Device posture and endpoint protection"),
    4: ("Foundational", "Content ownership, permissions, sharing and oversharing"),
    5: ("Foundational", "Data classification and sensitivity labels"),
    6: ("Foundational", "Audit and retention"),
    7: ("Foundational", "Application consent and connector review"),
    8: ("Foundational", "Incident handling and exception management"),
    9: ("AI-specific", "Copilot DLP and supported AI interaction controls"),
    10: ("AI-specific", "AI application and agent approval"),
    11: ("AI-specific", "Shadow AI discovery and monitoring"),
    12: ("AI-specific", "AI data-risk assessments"),
    13: ("AI-specific", "Prompt, response and AI interaction investigation"),
    14: ("AI-specific", "Agent and connector lifecycle"),
    15: ("AI-specific", "AI adoption, training and support (outside review scope)"),
}

# TopicId: (step, technical domain, benchmark ids, default role, plain-language subject)
TOPICS = {
    "SCOPE": (1, "Scope and governance", ["MS-COPILOT-REQ-01", "MS-COPILOT-ZT-01"],
              "Business sponsor and Microsoft 365 administrator",
              "current Copilot use and which AI services are approved"),
    "LICENSING": (1, "Licensing and prerequisites", ["MS-COPILOT-REQ-01"], "Microsoft 365 administrator",
                  "Copilot licenses and the Microsoft 365 apps people need to use it"),
    "MFA": (2, "Identity and access", ["MS-ENTRA-CA-MFA-01", "MS-COPILOT-ZT-01"],
            "Identity and access administrator", "how many accounts are set up for a second sign-in check"),
    "CA": (2, "Identity and access", ["MS-ENTRA-CA-MFA-01", "MS-ENTRA-CA-REPORTONLY-01", "MS-ENTRA-CA-LEGACY-01",
                                      "MS-COPILOT-ZT-01"],
           "Identity and access administrator", "whether every sign-in must pass a second check"),
    "LEGACY": (2, "Identity and access", ["MS-ENTRA-CA-LEGACY-01"], "Identity and access administrator",
               "older sign-in methods that cannot use a second check"),
    "PRIV": (2, "Privileged access", ["MS-ENTRA-ROLES-01", "MS-ENTRA-MANDATORY-MFA-01", "MS-ENTRA-EMERGENCY-01"],
             "Identity and access administrator", "administrator accounts and how they are protected"),
    "RISK": (2, "Identity and access", ["MS-ENTRA-IDP-01"], "Identity and access administrator",
             "accounts that Microsoft flagged as possibly compromised"),
    "ACCESS_REVIEW": (2, "Identity governance", ["MS-ENTRA-ROLES-01"], "Identity and access administrator",
                      "regular reviews of who still needs access"),
    "DEVICE": (3, "Devices and endpoints", ["MS-INTUNE-COMPLIANCE-01", "MS-COPILOT-ZT-01"],
               "Endpoint and security operations owners", "whether company devices are managed and healthy"),
    "MDE": (3, "Devices and endpoints", ["MS-MDE-HEALTH-01", "MS-COPILOT-ZT-01"],
            "Endpoint and security operations owners", "whether devices report to the endpoint protection service"),
    "SHARING": (4, "Content and sharing", ["MS-SPO-SHARING-01", "MS-COPILOT-FOUNDATION-01"],
                "SharePoint administrator and content owners", "how files can be shared by link"),
    "GUEST": (4, "Content and sharing", ["MS-ENTRA-GUEST-01"], "Identity and access administrator",
              "who may invite people from outside the organization"),
    "PERMISSIONS": (4, "Content and sharing", ["MS-SPO-DAG-01", "MS-COPILOT-FOUNDATION-01", "MS-COPILOT-PRIVACY-01"],
                    "SharePoint administrator and content owners", "content that is open to very large groups"),
    "OWNERSHIP": (4, "Content and sharing", ["MS-COPILOT-FOUNDATION-01"], "SharePoint administrator and content owners",
                  "whether sites have accountable owners"),
    "LABELS": (5, "Data protection", ["MS-COPILOT-ZT-01", "MS-COPILOT-PRIVACY-01"],
               "Information protection and compliance owner", "sensitivity labels that mark and protect business content"),
    "ENCRYPTION": (5, "Data protection", ["MS-COPILOT-PRIVACY-01"], "Information protection and compliance owner",
                   "encryption and usage rights on sensitive files"),
    "DLP": (5, "Data protection", ["MS-COPILOT-ZT-01"], "Information protection and compliance owner",
            "rules that stop sensitive information from leaving"),
    "AUDIT": (6, "Audit and retention", ["MS-PURVIEW-AUDIT-COPILOT-01", "MS-PURVIEW-AUDIT-RET-01"],
              "Information protection and compliance owner", "activity records used for investigations"),
    "RETENTION": (6, "Audit and retention", ["MS-PURVIEW-RET-COPILOT-01"], "Information protection and compliance owner",
                  "how long content and Copilot conversations are kept"),
    "LOCKBOX": (6, "Audit and retention", ["MS-PURVIEW-LOCKBOX-01"], "Information protection and compliance owner",
                "approval before Microsoft support can access content"),
    "CONSENT": (7, "Applications and connectors", ["MS-ENTRA-CONSENT-01", "MS-ENTRA-CONSENT-02"],
                "Application and security administrators", "which apps users may approve to read company data"),
    "CONNECTORS": (7, "Applications and connectors", ["MS-COPILOT-PRIVACY-01"], "Application and security administrators",
                   "outside data sources connected to Microsoft 365 search and Copilot"),
    "INCIDENTS": (8, "Security operations", ["MS-XDR-INCIDENTS-01", "MS-COPILOT-ZT-01"],
                  "Endpoint and security operations owners", "open security alerts that have not been worked"),
    "EXCEPTIONS": (8, "Security operations", ["MS-XDR-INCIDENTS-01"], "Security operations owner",
                   "how security incidents and approved exceptions are tracked"),
    "COPILOT_DLP": (9, "AI data protection", ["MS-PURVIEW-DLP-COPILOT-01"], "Information protection and compliance owner",
                    "rules that stop Copilot from using sensitive content or prompts"),
    "AGENTS": (10, "AI governance", ["MS-AGENTS-ADMIN-01"], "Microsoft 365 administrator",
               "which AI apps and agents are approved"),
    "SHADOW_AI": (11, "AI governance", ["MS-PURVIEW-DSPMAI-01"], "Security operations owner",
                  "use of AI services the organization has not approved"),
    "AI_DATA_RISK": (12, "AI data protection", ["MS-PURVIEW-DSPMAI-01", "MS-SPO-DAG-01"],
                     "Information protection and compliance owner", "where sensitive data may be overshared to AI"),
    "AI_INVESTIGATION": (13, "AI governance", ["MS-PURVIEW-AUDIT-COPILOT-01", "MS-PURVIEW-RET-COPILOT-01"],
                         "Information protection and compliance owner",
                         "how Copilot prompts and responses can be investigated"),
    "AGENT_LIFECYCLE": (14, "AI governance", ["MS-AGENTS-ADMIN-01"], "Microsoft 365 administrator",
                        "how agents and connectors are owned, reviewed and retired"),
    "ADOPTION": (15, "AI adoption", ["MS-COPILOT-REQ-01"], "Business sponsor and adoption lead",
                 "Copilot use cases and how success will be measured"),
}

# Ordered mapping rules: (field, regex, topic). First match wins.
TOPIC_RULES = [
    ("FindingKey", r"^entra\.privileged\.", "PRIV"),
    ("RecommendationId", r"^ENT-011$", "PRIV"),
    ("FindingKey", r"mfa_registration|registration", "MFA"),
    ("RecommendationId", r"^ENT-00[57]$", "MFA"),
    ("FindingKey", r"^baseline\.identity\.sign_in", "CA"),
    ("RecommendationId", r"^ENT-(004|006|020)$", "CA"),
    ("FindingKey", r"legacy", "LEGACY"),
    ("RecommendationId", r"^ENT-013$", "RISK"),
    ("RecommendationId", r"^ENT-01[29]$", "ACCESS_REVIEW"),
    ("RecommendationId", r"^ENT-015$", "GUEST"),
    ("RecommendationId", r"^ENT-002$", "DEVICE"),
    ("FindingKey", r"^defender\.incidents", "INCIDENTS"),
    ("FindingKey", r"^sharepoint\.sharing", "SHARING"),
    ("FindingKey", r"^sharepoint\.dag|^data_exposure\.(broad|freshness)", "PERMISSIONS"),
    ("FindingKey", r"ownership", "OWNERSHIP"),
    ("FindingKey", r"sensitive-data_exposure|data_exposure\.coverage", "AI_DATA_RISK"),
    ("FindingKey", r"retention", "RETENTION"),
    ("FindingKey", r"label", "LABELS"),
    ("FindingKey", r"rms|encryption", "ENCRYPTION"),
    ("FindingKey", r"lockbox", "LOCKBOX"),
    ("FindingKey", r"audit", "AUDIT"),
    ("FindingKey", r"retention", "RETENTION"),
    ("RecommendationId", r"^PUR-016$", "SHADOW_AI"),
    ("FindingKey", r"dlp", "DLP"),
    ("FindingKey", r"consent|app_consent", "CONSENT"),
    ("FindingKey", r"connector", "CONNECTORS"),
    ("FindingKey", r"license\.apps", "LICENSING"),
    ("FindingKey", r"license", "LICENSING"),
    ("FindingKey", r"adoption", "ADOPTION"),
    ("ControlId", r"^IDENTITY\.MFA$", "MFA"),
    ("ControlId", r"^IDENTITY\.AUTH$", "CA"),
    ("ControlId", r"^IDENTITY\.ADMIN$", "PRIV"),
    ("ControlId", r"^CONTENT\.SHARING$", "SHARING"),
    ("ControlId", r"^CONTENT\.PERMISSIONS$", "PERMISSIONS"),
    ("ControlId", r"^CONTENT\.OWNERSHIP$", "OWNERSHIP"),
    ("ControlId", r"^DATA\.(LABELS|PUBLISHING)$", "LABELS"),
    ("ControlId", r"^DATA\.DLP$", "DLP"),
    ("ControlId", r"^DATA\.AUDIT$", "AUDIT"),
    ("ControlId", r"^DATA\.RETENTION$", "RETENTION"),
    ("ControlId", r"^DATA\.EXPOSURE$", "AI_DATA_RISK"),
    ("ControlId", r"^APPS\.CONSENT$", "CONSENT"),
    ("ControlId", r"^APPS\.CONNECTIONS$", "CONNECTORS"),
    ("ControlId", r"^ENDPOINT\.", "DEVICE"),
    ("ControlId", r"^THREAT\.", "INCIDENTS"),
    ("ControlId", r"^LICENSE\.", "LICENSING"),
    ("ControlId", r"^ADOPTION\.", "ADOPTION"),
    ("Feature", r"(?i)copilot.*dlp|dlp.*copilot", "COPILOT_DLP"),
    ("Feature", r"(?i)intune|device", "DEVICE"),
    ("Feature", r"(?i)defender", "MDE"),
]

# A DLP finding that explicitly measures Copilot-scoped policies is AI-specific.
COPILOT_DLP_PATTERN = re.compile(r"(?i)apply to Microsoft 365 Copilot|Copilot location|policies? .*Copilot")

# Overlap groups: the first present member is primary; later members overlap it.
OVERLAP_GROUPS = [
    ("sharepoint.dag.not_assessed", "data_exposure.freshness.sam"),
    ("ENT-012", "ENT-019"),
    ("entra.privileged.PermanentPrivilegeNeedsReview", "ENT-011"),
    ("baseline.dlp.enforced", "dlp-enforced"),
    ("baseline.labels.published", "sensitivity-labels-published"),
    ("purview.audit.state", "unified-audit-enabled"),
    ("purview.rms.state", "azure-rms-enabled"),
]

# Gap findings contradicted by a retained source: (gap FindingKey, fact path, description)
CONTRADICTED_GAPS = [
    ("coverage.data.labels", ("labels", "Count"),
     "Purview returned {value} sensitivity label definitions (source {source})."),
    ("coverage.license.assignment", ("copilot_licenses", "Enabled"),
     "The license inventory returned {value} paid Copilot seats (source {source})."),
    ("coverage.content.ownership", ("site_inventory", "Count"),
     "The {source} returned owners and last-change dates for {value} sites."),
]

# Claim checks on script text: (id, field, regex, benchmark, correction, statement type)
CLAIM_CHECKS = [
    ("CLM-TRAINING", "Recommendation",
     r"(?i)\btrain(ing|ed)?\b[^.]{0,40}\b(copilot|context)\b|\bcopilot\b[^.]{0,60}\btrain(ing|ed)?\b",
     "MS-COPILOT-PRIVACY-01",
     "The recommendation frames retention around Copilot training. Microsoft documents that prompts, responses "
     "and Microsoft Graph data are not used to train foundation models; retention periods are a records "
     "decision, not an AI training control.", "MicrosoftGuidance"),
    ("CLM-AUDIT-10Y", "Recommendation", r"(?i)10 ?years?|ten years",
     "MS-PURVIEW-AUDIT-RET-01",
     "Ten-year audit retention requires the 10-Year Audit Log Retention add-on in addition to an E5-level "
     "license for each user generating records; Audit (Standard) keeps 180 days. Retention length is a customer "
     "decision.", "MicrosoftGuidance"),
    ("CLM-AUDIT-EVENTS", "Recommendation", r"PromptSubmitted|SensitiveDataAccessed",
     "MS-PURVIEW-AUDIT-COPILOT-01",
     "The event names PromptSubmitted and SensitiveDataAccessed are not Microsoft-documented audit types. "
     "Copilot interactions are recorded as the CopilotInteraction record type and operation.", "MicrosoftGuidance"),
    ("CLM-AI-OUTPUT-LABEL", "Recommendation",
     r"(?i)(created|drafted|generated)\s+(via|by|through)\s+(copilot|ai)|copilot outputs|default label to",
     "MS-COPILOT-ZT-01",
     "Specific labeling outcomes for Copilot-generated content (default labels, mandatory auto-labeling rules) "
     "are customer policy decisions; Microsoft guidance is to define the taxonomy, publish labels and choose "
     "default or automatic labeling deliberately.", "CustomerPolicyDecision"),
    ("CLM-ASSURANCE", "Recommendation", r"(?i)this ensures that even if|guarantees?",
     "MS-COPILOT-PRIVACY-01",
     "The recommendation promises an outcome the evidence cannot establish. Rights management protects content "
     "only where labels or permissions are actually applied.", "ReviewerInterpretation"),
    ("CLM-EVENTS-LOGGED", "Observation", r"(?i)copilot events are being logged",
     "MS-PURVIEW-AUDIT-COPILOT-01",
     "Microsoft documents that Copilot interactions are logged automatically when auditing is enabled, so the "
     "statement is consistent with product behavior. Tenant retrieval of Copilot audit records was not "
     "verified in this package.", "MicrosoftGuidance"),
]

# Closure templates per topic: (evidence, validation, population, reassessment)
CLOSURE = {
    "SCOPE": ("Recorded decision on current Copilot use and the approved AI services.",
              "Operator confirms the record and links it into the next assessment run.",
              "Current Copilot users and approved services", "Operator review"),
    "LICENSING": ("Dated license assignment export and Microsoft 365 Apps readiness export.",
                  "Compare assigned Copilot licenses and app channel or version with the intended users.",
                  "Users assigned Copilot", "Rerun collector with the readiness export"),
    "MFA": ("Dated authentication methods registration report for enabled member accounts.",
            "Recalculate registration for active users and list documented exceptions.",
            "Administrators first, then all enabled member accounts", "Rerun collector"),
    "CA": ("Export of enforced Conditional Access policies with assignments, exclusions and state; sign-in "
           "log sample showing MFA applied.", "Confirm policy state is On (not report-only) and that only emergency "
           "access and synchronization accounts are excluded; review report-only impact before switching.",
           "All users", "Rerun collector"),
    "LEGACY": ("Enforced Conditional Access legacy-block policy export and a legacy sign-in report for the "
               "review window.", "Confirm zero successful legacy sign-ins outside documented exceptions.",
               "All users", "Rerun collector"),
    "PRIV": ("Dated privileged role assignment export with purpose, owner role and assignment type; MFA "
             "registration for privileged identities.", "Confirm each active assignment is required, time-bound "
             "where PIM applies, and MFA-registered; record exceptions.", "All privileged identities",
             "Rerun collector"),
    "RISK": ("Dated risky users report showing remediation, dismissal or confirmed compromise for each at-risk "
             "account.", "Confirm no unremediated at-risk accounts remain.", "At-risk accounts", "Rerun collector"),
    "ACCESS_REVIEW": ("Access review definitions or an approved decision not to use them.",
                      "Confirm reviews exist for privileged roles and guests or a documented alternative.",
                      "Privileged roles and guests", "Operator review"),
    "DEVICE": ("Device management model statement and a dated enrolled-device and compliance report.",
               "Confirm devices used with Copilot are enrolled and compliant, or record the approved alternative "
               "control.", "Devices used with Copilot", "Rerun collector"),
    "MDE": ("Dated Defender for Endpoint device inventory with health and onboarding state.",
            "Confirm devices used with Copilot are onboarded and active.", "Devices used with Copilot",
            "Rerun collector"),
    "SHARING": ("Approved sharing policy decision and current tenant and site sharing settings export.",
                "Confirm tenant and site settings match the decision.", "Tenant and all sites", "Rerun collector"),
    "GUEST": ("Approved guest invitation policy decision and external collaboration settings export.",
              "Confirm the setting matches the decision.", "Tenant", "Operator review"),
    "PERMISSIONS": ("Current Data Access Governance reports (site permissions, Everyone and EEEU, sharing links) "
                    "dated within the freshness threshold, plus site access review outcomes.",
                    "Confirm broad grants on the listed sites were removed, restricted or approved.",
                    "Sites with broad access first", "Rerun collector with current reports"),
    "OWNERSHIP": ("Active sites report with owners and activity, with decisions for ownerless or inactive sites.",
                  "Confirm every active site has an accountable owner and inactive sites are archived or deleted.",
                  "All SharePoint sites", "Rerun collector"),
    "LABELS": ("Label taxonomy decision, publishing policy export and label usage evidence.",
               "Confirm labels exist for business content and are published to Copilot users.",
               "Copilot users and business content", "Rerun collector"),
    "ENCRYPTION": ("Label configuration showing which labels apply encryption.",
                   "Confirm encrypting labels are applied where required.", "Sensitive content", "Operator review"),
    "DLP": ("DLP policy export and dated test or simulation results for in-scope rules.",
            "Run policy simulation or a controlled test and record outcomes.", "Users in policy scope",
            "Operator review"),
    "AUDIT": ("Dated audit search showing records for the review period, including CopilotInteraction records.",
              "Run the audit search and record counts and retention settings.", "Copilot users",
              "Rerun collector"),
    "RETENTION": ("Retention policy export showing locations, including the Copilot and AI apps location, and "
                  "the approved retention decision.", "Confirm policy coverage matches the decision.", "Tenant",
                  "Rerun collector"),
    "LOCKBOX": ("Recorded decision on Customer Lockbox.", "Confirm setting matches the decision.", "Tenant",
                "Operator review"),
    "CONSENT": ("User consent settings export, admin consent workflow state and an approved consent policy "
                "decision.", "Confirm settings match the decision.", "Tenant", "Rerun collector"),
    "CONNECTORS": ("Connector and integrated app inventory for Copilot.", "Confirm each connector has an owner and "
                   "approved data scope.", "Tenant", "Operator review"),
    "INCIDENTS": ("Dated incident queue export showing owner, status and classification for each active incident.",
                  "Confirm high-severity incidents are triaged and no active compromise remains.",
                  "Active incidents", "Rerun collector"),
    "EXCEPTIONS": ("Exception register with role, scope, compensating control, expiry and approval reference.",
                   "Confirm each accepted exception is recorded before reassessment.", "All accepted exceptions",
                   "Operator review"),
    "COPILOT_DLP": ("Copilot DLP policy export with mode and dated simulation or test results.",
                    "Confirm intended prompts or labeled items are blocked for Copilot users.", "Copilot users",
                    "Operator review"),
    "AGENTS": ("Agent and AI application inventory with approval status.", "Confirm only approved agents are "
               "available to users.", "All users", "Rerun collector with admin access"),
    "SHADOW_AI": ("Dated AI app discovery or DSPM for AI activity report with stated coverage.",
                  "Confirm the monitored population and networks.", "Managed devices and networks",
                  "Operator review"),
    "AI_DATA_RISK": ("Completed DSPM data risk assessment export with scope and date.",
                     "Confirm the assessed sites and that findings were addressed.", "SharePoint sites",
                     "Rerun collector with the export"),
    "AI_INVESTIGATION": ("Recorded decision on who may investigate Copilot prompts and responses, with retention "
                         "and eDiscovery approach.", "Confirm a test search can retrieve Copilot interactions.",
                         "Copilot users", "Operator review"),
    "AGENT_LIFECYCLE": ("Agent and connector ownership and review schedule.", "Confirm each agent has an owner "
                        "and review date.", "Approved agents", "Operator review"),
    "ADOPTION": ("None required; adoption planning is outside the scope of this review.",
                 "Not applicable.", "Not applicable", "Not applicable"),
}

EXCEPTION_DOC = ("Record the accountable role, affected scope, business reason, compensating control, review or "
                 "expiry date and approval reference. A role, not a named person, is recorded here; the customer "
                 "supplies the approver.")

TREATMENT_SENTENCE = {
    "BlocksPilot": "This must be fixed before Copilot is rolled out.",
    "ConfirmBeforePilot": "This must be checked and confirmed before Copilot is rolled out.",
    "TenantHardening": "This is a security improvement for the whole tenant; it does not by itself stop a rollout.",
    "LongerTermImprovement": "This is a longer-term improvement.",
    "InventoryHygiene": "This is housekeeping to keep records accurate.",
    "EvidenceRequest": "More evidence is needed before anyone can judge this.",
    "NoAction": "No action is needed from this review.",
}

PRIORITY_DEFAULT = {"BlocksPilot": "High", "ConfirmBeforePilot": "High", "TenantHardening": "Medium",
                    "EvidenceRequest": "Medium", "LongerTermImprovement": "Low", "InventoryHygiene": "Low",
                    "NoAction": "None"}


def topic_for(finding):
    for field, pattern, topic in TOPIC_RULES:
        value = str(finding.get(field) or "")
        if value and re.search(pattern, value):
            if topic == "DLP" and COPILOT_DLP_PATTERN.search(str(finding.get("Observation") or "")):
                return "COPILOT_DLP"
            return topic
    return "SCOPE"


def overlap_primary(findings):
    """Return {finding id: primary finding id} for overlapping findings."""
    def keys(f):
        return {str(f.get("FindingKey") or ""), str(f.get("RecommendationId") or "")}
    present = {}
    for f in findings:
        for k in keys(f):
            if k:
                present.setdefault(k, f.get("RecommendationId"))
    result = {}
    for group in OVERLAP_GROUPS:
        members = [present[k] for k in group if k in present]
        if len(members) >= 2:
            primary = members[0]
            for other in members[1:]:
                if other != primary:
                    result[other] = primary
    return result
