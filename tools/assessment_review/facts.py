"""Aggregate collector facts derived from retained collection data.

Only counts and states leave this module. Object names, user principal names,
addresses and payloads are collected solely into private deny-lists used to
redact script text and to verify outputs; they are never rendered.
"""
from __future__ import annotations

import re

from .util import parse_date, age_days

GLOBAL_ADMIN_TEMPLATE = "62e90394-69f5-4237-9190-012177145e10"
MICROSOFT_MANAGED_CONSENT = {
    "microsoft-user-default-recommended", "microsoft-user-default-allow-consent-apps",
    "microsoft-user-default-low"}
LEGACY_CONSENT = "microsoft-user-default-legacy"
FIRST_PARTY_PREFIX = re.compile(
    r"^(microsoft|office|azure|windows|sharepoint|teams|exchange|power|dynamics|skype|outlook|"
    r"onedrive|viva|intune|defender|purview|entra|copilot|bing|yammer|graph)\b", re.I)


def _attrs(collection, service):
    info = (collection.get("service_results") or {}).get(service) or {}
    return ((info.get("_client") or {}).get("$collection_attributes")) or {}


def _service(collection, service):
    return (collection.get("service_results") or {}).get(service) or {}


def _is_mfa(policy):
    grant = policy.get("grantControls") or {}
    controls = [str(c).lower() for c in grant.get("builtInControls") or []]
    return "mfa" in controls or bool(grant.get("authenticationStrength"))


def _is_block(policy):
    grant = policy.get("grantControls") or {}
    return "block" in [str(c).lower() for c in grant.get("builtInControls") or []]


def _narrowing_conditions(policy):
    c = policy.get("conditions") or {}
    return bool(c.get("locations") or c.get("platforms") or c.get("signInRiskLevels")
                or c.get("userRiskLevels") or c.get("devices"))


def conditional_access_facts(policies):
    out = {"Total": len(policies), "Enabled": 0, "ReportOnly": 0, "Disabled": 0,
           "EnforcedAllUsersMfa": 0, "ReportOnlyAllUsersMfa": 0, "EnforcedTargetedMfa": 0,
           "EnforcedAdminMfa": 0, "ReportOnlyOrDisabledAdminMfa": 0,
           "EnforcedLegacyBlockAllUsers": 0, "AllUsersMfaExclusionCounts": [],
           "EnforcedCompliantDevice": 0}
    for p in policies:
        state = str(p.get("state") or "").lower()
        if state == "enabled":
            out["Enabled"] += 1
        elif state == "enabledforreportingbutnotenforced":
            out["ReportOnly"] += 1
        else:
            out["Disabled"] += 1
        cond = p.get("conditions") or {}
        users = cond.get("users") or {}
        apps = cond.get("applications") or {}
        all_users = "All" in (users.get("includeUsers") or [])
        all_apps = "All" in (apps.get("includeApplications") or [])
        roles = users.get("includeRoles") or []
        clients = [str(c).lower() for c in cond.get("clientAppTypes") or []]
        grant_controls = [str(c).lower() for c in (p.get("grantControls") or {}).get("builtInControls") or []]
        if state == "enabled" and ("compliantdevice" in grant_controls or "domainjoineddevice" in grant_controls):
            out["EnforcedCompliantDevice"] += 1
        if _is_mfa(p):
            if all_users and all_apps and not _narrowing_conditions(p):
                if state == "enabled":
                    out["EnforcedAllUsersMfa"] += 1
                elif state == "enabledforreportingbutnotenforced":
                    out["ReportOnlyAllUsersMfa"] += 1
                out["AllUsersMfaExclusionCounts"].append(
                    len(users.get("excludeUsers") or []) + len(users.get("excludeGroups") or []))
            elif roles:
                if state == "enabled":
                    out["EnforcedAdminMfa"] += 1
                else:
                    out["ReportOnlyOrDisabledAdminMfa"] += 1
            elif state == "enabled":
                out["EnforcedTargetedMfa"] += 1
        if _is_block(p) and state == "enabled" and all_users and clients and \
                set(clients) <= {"exchangeactivesync", "other"}:
            out["EnforcedLegacyBlockAllUsers"] += 1
    out["AllUsersMfaExclusionCounts"].sort()
    return out


def extract(pkg):
    coll = pkg.collection
    snap = pkg.snapshot
    entra = _attrs(coll, "entra_info")
    purview = _attrs(coll, "purview_info")
    defender = _attrs(coll, "defender_info")
    review_date = parse_date(pkg.windows.get("ReviewDate"))[0]
    facts = {}

    policies = entra.get("ca_policies")
    facts["ca"] = conditional_access_facts(policies) if isinstance(policies, list) else None
    sd = entra.get("security_defaults") or {}
    facts["security_defaults_enabled"] = sd.get("is_enabled") if sd.get("available", True) else None

    reg = entra.get("auth_methods_registration")
    if isinstance(reg, list):
        members = [r for r in reg if str(r.get("userType")).lower() == "member"]
        facts["mfa_registration"] = {
            "Returned": len(reg), "Members": len(members), "Guests": len(reg) - len(members),
            "Registered": sum(1 for r in reg if r.get("isMfaRegistered") is True),
            "NotRegistered": sum(1 for r in reg if r.get("isMfaRegistered") is False),
            "AdminsNotRegistered": sum(1 for r in reg if r.get("isAdmin") and r.get("isMfaRegistered") is False),
            "AdminsReturned": sum(1 for r in reg if r.get("isAdmin")),
        }
    else:
        facts["mfa_registration"] = None

    consent = entra.get("consent_summary") or {}
    assigned = consent.get("assigned_user_consent_policies")
    if assigned is None:
        perms = ((entra.get("authorization_policy") or {}).get("defaultUserRolePermissions") or {})
        assigned = [p for p in perms.get("permissionGrantPoliciesAssigned") or []
                    if p.lower().startswith("managepermissiongrantsforself.")]
    names = sorted({p.split(".", 1)[-1] for p in assigned or []})
    facts["consent"] = {
        "AssignedUserPolicies": names,
        "AllMicrosoftManaged": bool(names) and all(n in MICROSOFT_MANAGED_CONSENT for n in names),
        "LegacyAssigned": LEGACY_CONSENT in names,
        "HighPrivilegeApps": consent.get("high_privilege_apps"),
        "UnverifiedPublishers": consent.get("unverified_publishers"),
        "GrantInventoryAvailable": consent.get("grant_inventory_available"),
    } if (assigned is not None or consent) else None

    risk = entra.get("risk_summary")
    facts["risk"] = ({"AtRisk": risk.get("at_risk"), "High": risk.get("risky_users_high"),
                      "Medium": risk.get("risky_users_medium"), "Low": risk.get("risky_users_low"),
                      "Records": risk.get("risky_users_total"), "Remediated": risk.get("remediated"),
                      "Dismissed": risk.get("dismissed"),
                      "UserRiskPolicy": risk.get("user_risk_policy_exists"),
                      "SignInRiskPolicy": risk.get("signin_risk_policy_exists")}
                     if isinstance(risk, dict) else None)

    incidents = defender.get("security_incidents")
    if isinstance(incidents, list):
        active = [i for i in incidents if str(i.get("status")).lower() in ("active", "inprogress", "in_progress")]
        untouched = [i for i in active if i.get("createdDateTime") and
                     i.get("createdDateTime") == i.get("lastUpdateDateTime")]
        created = [parse_date(i.get("createdDateTime"))[0] for i in active]
        created = [c for c in created if c]
        facts["incidents"] = {
            "Records": len(incidents), "Active": len(active),
            "ActiveHigh": sum(1 for i in active if str(i.get("severity")).lower() == "high"),
            "ActiveMedium": sum(1 for i in active if str(i.get("severity")).lower() == "medium"),
            "ActiveLow": sum(1 for i in active if str(i.get("severity")).lower() == "low"),
            "ActiveNoUpdateSinceCreation": len(untouched),
            "ActiveUnassigned": sum(1 for i in active if not i.get("assignedTo")),
            "OldestActiveAgeDays": age_days(min(created), review_date) if created else None,
        }
    else:
        facts["incidents"] = None

    managed = entra.get("managed_devices")
    device_summary = entra.get("device_summary") or {}
    facts["intune"] = {"ManagedDevices": len(managed) if isinstance(managed, list) else None,
                       "CompliancePolicies": device_summary.get("compliance_policies_total"),
                       "Source": pkg.sources.get("managed_devices", {}).get("Classification", "Unknown")}
    machines = defender.get("defender_devices")
    if isinstance(machines, list):
        seen = [parse_date(m.get("lastSeen"))[0] for m in machines]
        seen = [s for s in seen if s]
        facts["defender_devices"] = {
            "Records": len(machines),
            "Onboarded": sum(1 for m in machines if str(m.get("onboardingStatus")).lower() == "onboarded"),
            "Active": sum(1 for m in machines if str(m.get("healthStatus")).lower() == "active"),
            "Inactive": sum(1 for m in machines if str(m.get("healthStatus")).lower() == "inactive"),
            "ReportedManagedByIntune": sum(1 for m in machines if str(m.get("managedBy")).lower() == "intune"),
            "NewestLastSeenAgeDays": age_days(max(seen), review_date) if seen else None,
        }
    else:
        facts["defender_devices"] = None
    recon = snap.get("device_reconciliation") or {}
    facts["device_reconciliation"] = {k: recon.get(k) for k in ("known_active", "onboarded", "reporting",
                                                                "active_days", "reporting_days")}

    assignments = entra.get("role_assignments")
    definitions = entra.get("role_definitions") or []
    ga_ids = {d.get("id") for d in definitions
              if d.get("templateId") == GLOBAL_ADMIN_TEMPLATE or d.get("displayName") == "Global Administrator"}
    ga_ids.add(GLOBAL_ADMIN_TEMPLATE)
    if isinstance(assignments, list):
        ga = [a for a in assignments if a.get("roleDefinitionId") in ga_ids]
        facts["global_admins"] = {"ActiveAssignments": len(ga),
                                  "DistinctPrincipals": len({a.get("principalId") for a in ga})}
    else:
        facts["global_admins"] = None
    pim = entra.get("pim_summary") or {}
    facts["pim"] = {k: pim.get(k) for k in ("total_active_assignments", "total_eligible_assignments",
                                            "permanent_assignments", "permanent_global_admins")} if pim else None
    ar = entra.get("access_review_summary")
    facts["access_reviews"] = ar.get("total_definitions") if isinstance(ar, dict) else None
    b2b = entra.get("b2b_summary") or {}
    facts["guests"] = {"Total": b2b.get("total_guests"), "InviteSetting": b2b.get("guest_invite_restrictions")} if b2b else None
    sign = entra.get("signin_summary") or {}
    facts["signins"] = {"Sampled": sign.get("total_signins_sampled"),
                        "LegacyAttempts": sign.get("legacy_auth_attempts"),
                        "ReportOnlyEvaluated": sign.get("ca_report_only")} if sign else None

    copilot_seats = None
    for item in (_service(coll, "m365_result") or []):
        if not isinstance(item, dict):
            continue
        for lic in item.get("licenses") or []:
            sku = str(lic.get("sku_part_number") or "")
            if "copilot" in sku.lower() and "ccibots" not in sku.lower():
                copilot_seats = copilot_seats or {"Enabled": 0, "Assigned": 0}
                copilot_seats["Enabled"] += int(lic.get("enabled") or 0)
                copilot_seats["Assigned"] += int(lic.get("consumed") or 0)
    facts["copilot_licenses"] = copilot_seats
    usage = {}
    admin_rows = (((snap.get("supplemental_reviews") or {}).get("copilot_admin_review") or {}).get("rows")) or []
    for row in admin_rows:
        topic = str(row.get("Topic") or "")
        if topic in ("Paid Copilot enabled users", "Paid Copilot active users") and isinstance(row.get("Value"), int):
            usage[topic] = row["Value"]
    facts["copilot_usage"] = usage or None
    facts["copilot_dlp_rows"] = sum(1 for r in admin_rows if r.get("Area") == "data_protection")
    facts["copilot_dlp_monitor_only"] = sum(1 for r in admin_rows if r.get("Area") == "data_protection"
                                            and "monitor" in str(r.get("Value")).lower())

    labels = purview.get("sensitivity_labels") or {}
    facts["labels"] = {"Available": labels.get("available"), "Count": labels.get("total_labels"),
                       "Source": pkg.sources.get("purview_sensitivity_labels", {}).get("Classification")} if labels else None
    lp = purview.get("label_policies") or {}
    facts["label_policies"] = lp.get("total_policies") if lp.get("available") else None
    audit = purview.get("audit_config") or {}
    audit_logs = purview.get("audit_logs") or {}
    facts["audit"] = {"UnifiedAuditEnabled": audit.get("unified_audit_enabled") if audit.get("available") else None,
                      "LogQueryAvailable": audit_logs.get("available"),
                      "CopilotInteractionSource": pkg.sources.get("m365_copilot_interaction_audit", {}).get("Classification", "Unknown"),
                      "CopilotInteractionReason": pkg.sources.get("m365_copilot_interaction_audit", {}).get("Reason", "")}
    facts["shadow_ai"] = {"Source": pkg.sources.get("shadow_ai_usage", {}).get("Classification", "Unknown"),
                          "Reason": pkg.sources.get("shadow_ai_usage", {}).get("Reason", "")}
    pp = _service(coll, "power_platform_info")
    cs = _service(coll, "copilot_studio_info")
    facts["power_platform"] = {"Availability": pp.get("availability_status") or ("available" if pp.get("available") else "unavailable"),
                               "CopilotStudioLicensed": bool(cs.get("has_copilot_studio") or cs.get("total_licenses"))} if (pp or cs) else None
    m365 = _m365_attrs(coll)
    gov_sites = ((m365.get("sharepoint_governance") or {}).get("sites")) or {}
    facts["site_inventory"] = ({"Count": gov_sites.get("records_collected") or len(gov_sites.get("items") or []),
                                "Source": "SharePoint admin site inventory"} if gov_sites.get("available") else None)
    facts["app_permission_apps"] = _apps_with_application_permissions(entra, coll.get("tenant_id"))
    return facts


def _m365_attrs(collection):
    for r in (collection.get("service_results") or {}).get("m365_result") or []:
        if isinstance(r, dict) and r.get("_client"):
            return r["_client"].get("$collection_attributes") or {}
    return {}


HIGH_APPLICATION = re.compile(r"ReadWrite|FullControl|^Mail\.(Read|Send)|^Files\.Read\.All|^Sites\.Read\.All|"
                              r"^Directory\.Read|RoleManagement|AppRoleAssignment")
MICROSOFT_TENANTS = {"f8cdef31-a31e-4b4a-93e4-5f571e91255a", "72f988bf-86f1-41af-91ab-2d7cd011db47"}


def _apps_with_application_permissions(entra, tenant_id):
    """Count non-Microsoft apps holding tenant-wide, high-impact application permissions."""
    pages = (entra.get("assessment_datasets") or {}).get("application_permissions")
    if not pages:
        return None
    by_id = {s.get("id"): s for s in entra.get("service_principals") or []}
    names = {(s.get("id"), r.get("id")): r.get("value") for s in by_id.values() for r in s.get("appRoles") or []}
    flagged = set()
    for page in pages:
        for rec in (page or {}).get("records") or []:
            owner = str((by_id.get(rec.get("principalId")) or {}).get("appOwnerOrganizationId") or "").lower()
            if owner in MICROSOFT_TENANTS:
                continue
            if HIGH_APPLICATION.search(str(names.get((rec.get("resourceId"), rec.get("appRoleId"))) or "")):
                flagged.add(rec.get("principalId"))
    return len(flagged)


# ------------------------------------------------------------------ deny-lists

EMAIL = re.compile(r"[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")


def private_terms(pkg):
    """Return (redact_terms, scan_terms): object names and identifiers from the package."""
    coll = pkg.collection
    entra = _attrs(coll, "entra_info")
    purview = _attrs(coll, "purview_info")
    defender = _attrs(coll, "defender_info")
    people, objects = set(), set()

    def add_person(value):
        if isinstance(value, str) and len(value.strip()) >= 4:
            people.add(value.strip())

    def add_object(value):
        if isinstance(value, str) and len(value.strip()) >= 5:
            objects.add(value.strip())

    for src in ("auth_methods_registration", "risky_users", "guest_users", "signin_logs"):
        for row in entra.get(src) or []:
            if isinstance(row, dict):
                for key in ("userDisplayName", "displayName", "userPrincipalName", "mail", "ipAddress"):
                    add_person(row.get(key))
    for row in defender.get("security_incidents") or []:
        add_object(row.get("displayName"))
        add_person(row.get("assignedTo"))
    for row in defender.get("security_alerts") or []:
        add_object(row.get("title"))
    for row in defender.get("defender_devices") or []:
        add_person(row.get("computerDnsName"))
        add_person(row.get("lastIpAddress"))
        add_person(row.get("lastExternalIpAddress"))
    for row in entra.get("ca_policies") or []:
        add_object(row.get("displayName"))
    for key in ("sensitivity_labels", "label_policies", "retention_policies", "retention_labels", "dlp_policies"):
        block = purview.get(key) or {}
        rows = block.get("policies") or block.get("labels") or []
        for row in rows:
            if isinstance(row, dict):
                add_object(row.get("name") or row.get("Name") or row.get("displayName") or row.get("DisplayName"))
    dlp_rules = (purview.get("dlp_rules") or {}).get("rules") or []
    for row in dlp_rules:
        if isinstance(row, dict):
            add_object(row.get("name") or row.get("Name"))
    for domain in (_service(coll, "entra_info").get("verified_domains") or []):
        add_person(domain)
    tenant = (pkg.identity.get("CustomerName") or "").strip().lower()
    people = {t for t in people if _multiword_or_identifier(t) and t.strip().lower() != tenant}
    # Object names (policies, labels, incidents) must be distinctive so ordinary words such as
    # "Block legacy" in a control title are not mistaken for tenant data.
    objects = {t for t in objects if _distinctive_object(t) and t.strip().lower() != tenant}
    redact = sorted(people | objects, key=lambda s: (-len(s), s))
    scan = sorted(people | {t for t in objects if not FIRST_PARTY_PREFIX.match(t)}, key=lambda s: (-len(s), s))
    return redact, scan


def _distinctive_object(value):
    value = value.strip()
    if len(value.split()) >= 3:
        return True
    return bool(re.search(r"\d|_|[a-z][A-Z]", value)) and len(value) >= 6


def _multiword_or_identifier(value):
    value = value.strip()
    if EMAIL.fullmatch(value) or re.fullmatch(r"[\d.:a-fA-F]{7,}", value) or "." in value and " " not in value:
        return True
    tokens = value.split()
    if len(tokens) >= 2:
        return True
    return bool(re.search(r"\d|_|[a-z][A-Z]", value)) and len(value) >= 6
