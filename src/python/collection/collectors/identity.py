"""
Identity section collector.
Produces all Data.Identity datasets via Microsoft Graph.
"""

from __future__ import annotations

import logging
from collections import Counter, defaultdict
from datetime import datetime, timezone
from typing import Any

from ..graph_client import GraphClient

log = logging.getLogger(__name__)

_STALE_DAYS = 180  # inactive sign-in threshold


def collect(client: GraphClient) -> dict[str, Any]:
    log.info("Collecting Identity section...")
    data: dict[str, Any] = {}

    # Core objects (run sequentially; some downstream derived steps depend on them)
    data["Users"]         = _collect_users(client)
    data["Admins"]        = _collect_admins(client)
    data["EntraIDGroups"] = _collect_groups(client)
    data["DeviceDetails"] = _collect_devices(client)

    # Intune managed devices (requires DeviceManagementManagedDevices.Read.All)
    data["IntuneDevices"] = _collect_intune_devices(client)

    # Group member counts — merged directly into EntraIDGroups[gid]["MemberCount"]
    # (requires GroupMember.Read.All; skipped gracefully if permission absent)
    group_ids = list(data["EntraIDGroups"].keys())
    member_counts = _collect_group_member_counts(client, group_ids)
    for gid, cnt in member_counts.items():
        if gid in data["EntraIDGroups"]:
            data["EntraIDGroups"][gid]["MemberCount"] = cnt

    # Disabled member detection for license-assigning groups
    license_group_ids = [gid for gid, g in data["EntraIDGroups"].items() if g.get("assignedLicenses")]
    if license_group_ids:
        user_id_to_enabled = {u["id"]: u.get("accountEnabled", True)
                              for u in data["Users"].values() if "id" in u}
        lic_members = _collect_license_group_members(client, license_group_ids)
        for gid, member_ids in lic_members.items():
            if gid in data["EntraIDGroups"]:
                data["EntraIDGroups"][gid]["DisabledMemberCount"] = sum(
                    1 for mid in member_ids if not user_id_to_enabled.get(mid, True)
                )

    # Auth / CA
    data["ConditionalAccessPolicies"]  = _collect_ca_policies(client)
    data["ConditionalAccessPolicySummary"] = _build_ca_summary(data["ConditionalAccessPolicies"])
    data["NamedLocations"]             = _collect_named_locations(client)
    data["AuthenticationStrengthPolicies"] = _collect_auth_strength_policies(client)
    data["AuthenticationMethods"]      = _collect_auth_methods_policy(client)
    data["SecurityDefaultsPolicy"]     = _collect_security_defaults(client)
    data["GuestAccessConfiguration"]   = _collect_guest_access_config(client)
    data["ExternalIdentityRestrictions"]      = _collect_external_identity(client)
    data["CrossTenantPartnerConfigurations"]  = _collect_cross_tenant_partners(client)
    data["RiskyUsers"]                 = _collect_risky_users(client)

    # MFA registration (requires Reports.Read.All or AuditLog.Read.All)
    data["MfaRegistrationDetails"]     = _collect_mfa_registration(client)

    # Licensing
    data["LicenseSKUs"]                = _collect_license_skus(client)

    # Applications
    data["EnterpriseApplications"]     = _collect_enterprise_apps(client)

    # Derived
    data["MfaEnrollmentSummary"]       = _build_mfa_enrollment_summary(data["MfaRegistrationDetails"])
    data["MfaEnforcementSummary"]      = _build_mfa_enforcement_summary(data["ConditionalAccessPolicies"])
    data["MfaMethodPostureSummary"]    = _build_mfa_method_posture(data["MfaRegistrationDetails"])
    data["PrivilegedAccessSummary"]    = _build_priv_access_summary(data["Admins"])
    data["PrivilegedAccessRemediationSummary"] = _build_priv_access_remediation(data["Admins"], data["Users"])
    data["GroupLicensingSummary"]      = _build_group_licensing_summary(data["EntraIDGroups"], data["LicenseSKUs"])
    data["EnterpriseApplicationSummary"] = _build_app_summary(data["EnterpriseApplications"])
    data["GuestSignInSummary"]         = _build_guest_signin_summary(data["Users"])
    data["DeviceManagementSummary"]    = _build_device_management_summary(data["DeviceDetails"], data["IntuneDevices"])
    data["LicenseOptimizationCandidates"] = _build_license_optimization(data["Users"], data["LicenseSKUs"])
    data["MfaEnforcementGapUsers"]     = _build_mfa_enforcement_gaps(
        data["Users"], data["MfaRegistrationDetails"], data["ConditionalAccessPolicies"]
    )
    data["MfaEnforcementScopeReview"]  = _build_mfa_scope_review(data["ConditionalAccessPolicies"])
    data["ConditionalAccessOptimization"] = _build_ca_optimization(data["ConditionalAccessPolicies"])
    data["AuthenticationSSOApplications"] = []  # populated from EnterpriseApplications if SSO data available
    data["AuthenticationConfig"]       = _build_auth_config(
        data["AuthenticationMethods"], data["SecurityDefaultsPolicy"], data["ConditionalAccessPolicies"]
    )

    log.info("Identity: %d users, %d groups, %d devices, %d CA policies, %d SKUs",
             len(data["Users"]), len(data["EntraIDGroups"]), len(data["DeviceDetails"]),
             len(data["ConditionalAccessPolicies"]) if isinstance(data["ConditionalAccessPolicies"], list) else 0,
             len(data["LicenseSKUs"]))
    return data


# ---------------------------------------------------------------------------
# Raw collectors
# ---------------------------------------------------------------------------

def _collect_users(client: GraphClient) -> dict:
    select = ("id,displayName,userPrincipalName,userType,accountEnabled,"
              "mail,mailNickname,department,jobTitle,country,usageLocation,"
              "createdDateTime,lastPasswordChangeDateTime,assignedLicenses,"
              "signInActivity,onPremisesSyncEnabled,onPremisesImmutableId")
    # $top=500 is required when signInActivity is in $select (Graph enforces max 500 per page)
    raw = client.get(f"users?$select={select}&$top=500")
    return {u["userPrincipalName"]: u for u in (raw if isinstance(raw, list) else [])}


def _collect_admins(client: GraphClient) -> list:
    roles = client.get("directoryRoles?$select=id,displayName,roleTemplateId")
    if not isinstance(roles, list):
        return []

    # signInActivity omitted here — unreliable on sub-resources; resolved from Users dict downstream
    batch_reqs = [
        {"id": str(i), "method": "GET", "url": f"/directoryRoles/{r['id']}/members?$select=id,displayName,userPrincipalName,userType,accountEnabled"}
        for i, r in enumerate(roles)
    ]
    responses = client.batch(batch_reqs)

    admins: list[dict] = []
    for role, resp in zip(roles, responses):
        members = (resp.get("body") or {}).get("value", [])
        for member in members:
            admins.append({**member, "RoleName": role["displayName"], "RoleId": role["id"]})
    return admins


def _collect_groups(client: GraphClient) -> dict:
    select = ("id,displayName,groupTypes,mailEnabled,securityEnabled,mail,"
              "membershipRule,createdDateTime,visibility,expirationDateTime,"
              "onPremisesSyncEnabled,assignedLabels,assignedLicenses")
    raw = client.get(f"groups?$select={select}&$expand=owners($select=id,displayName,userPrincipalName)")
    return {g["id"]: g for g in (raw if isinstance(raw, list) else [])}


def _collect_devices(client: GraphClient) -> dict:
    select = ("id,displayName,operatingSystem,operatingSystemVersion,complianceState,"
              "registrationDateTime,approximateLastSignInDateTime,isManaged,isCompliant,"
              "enrollmentType,managementType,deviceOwnership,trustType,profileType")
    raw = client.get(f"devices?$select={select}")
    return {d["id"]: d for d in (raw if isinstance(raw, list) else [])}


def _collect_intune_devices(client: GraphClient) -> dict:
    select = ("id,azureADDeviceId,deviceName,operatingSystem,osVersion,"
              "complianceState,managementAgent,enrolledDateTime,lastSyncDateTime,"
              "userPrincipalName,manufacturer,model,deviceEnrollmentType,isEncrypted")
    try:
        raw = client.get(f"deviceManagement/managedDevices?$select={select}")
        return {d["id"]: d for d in (raw if isinstance(raw, list) else [])}
    except Exception as exc:
        log.debug("IntuneDevices unavailable (need DeviceManagementManagedDevices.Read.All): %s", exc)
        return {}


def _collect_group_member_counts(client: GraphClient, group_ids: list[str]) -> dict[str, int]:
    """Batch-fetch member counts for all groups via ?$count=true. Returns {} on 403."""
    if not group_ids:
        return {}
    requests_list = [
        {
            "id": gid,
            "method": "GET",
            "url": f"/groups/{gid}/members?$count=true&$top=1&$select=id",
            "headers": {"ConsistencyLevel": "eventual"},
        }
        for gid in group_ids
    ]
    counts: dict[str, int] = {}
    try:
        responses = client.batch(requests_list)
        for resp in responses:
            gid    = resp.get("id", "")
            status = resp.get("status", 0)
            if status == 403:
                log.debug("Group member counts unavailable: GroupMember.Read.All not granted")
                return {}
            if status == 200:
                body = resp.get("body", {}) or {}
                odata_count = body.get("@odata.count")
                try:
                    counts[gid] = int(odata_count) if odata_count is not None else -1
                except (TypeError, ValueError):
                    counts[gid] = -1
            else:
                counts[gid] = -1
    except Exception as exc:
        log.debug("Group member counts batch unavailable: %s", exc)
        counts = {gid: -1 for gid in group_ids}
    return counts


def _collect_license_group_members(client: GraphClient, license_group_ids: list[str]) -> dict[str, list[str]]:
    """Batch-fetch member IDs for license-assigning groups (up to 999 per group). Returns {} on 403."""
    if not license_group_ids:
        return {}
    requests_list = [
        {
            "id": gid,
            "method": "GET",
            "url": f"/groups/{gid}/members?$select=id&$top=999",
        }
        for gid in license_group_ids
    ]
    members: dict[str, list[str]] = {}
    try:
        responses = client.batch(requests_list)
        for resp in responses:
            gid    = resp.get("id", "")
            status = resp.get("status", 0)
            if status == 403:
                log.debug("License group members unavailable: GroupMember.Read.All not granted")
                return {}
            if status == 200:
                body = resp.get("body", {}) or {}
                members[gid] = [m.get("id", "") for m in body.get("value", [])]
            else:
                members[gid] = []
    except Exception as exc:
        log.debug("License group members batch unavailable: %s", exc)
    return members


def _collect_ca_policies(client: GraphClient) -> list:
    raw = client.get("identity/conditionalAccess/policies")
    return raw if isinstance(raw, list) else []


def _collect_auth_methods_policy(client: GraphClient) -> dict:
    try:
        return client.get("policies/authenticationMethodsPolicy") or {}
    except Exception as exc:
        log.debug("AuthenticationMethodsPolicy unavailable: %s", exc)
        return {}


def _collect_security_defaults(client: GraphClient) -> dict:
    try:
        return client.get("policies/identitySecurityDefaultsEnforcementPolicy") or {}
    except Exception as exc:
        log.debug("SecurityDefaultsPolicy unavailable: %s", exc)
        return {}


def _collect_guest_access_config(client: GraphClient) -> dict:
    try:
        return client.get("policies/authorizationPolicy") or {}
    except Exception as exc:
        log.debug("AuthorizationPolicy unavailable: %s", exc)
        return {}


def _collect_external_identity(client: GraphClient) -> dict:
    try:
        return client.get("policies/crossTenantAccessPolicy") or {}
    except Exception as exc:
        log.debug("CrossTenantAccessPolicy unavailable: %s", exc)
        return {}


def _collect_cross_tenant_partners(client: GraphClient) -> list:
    try:
        raw = client.get("policies/crossTenantAccessPolicy/partners")
        return raw if isinstance(raw, list) else []
    except Exception as exc:
        log.debug("CrossTenantAccessPolicy/partners unavailable: %s", exc)
        return []


def _collect_named_locations(client: GraphClient) -> list:
    try:
        raw = client.get("identity/conditionalAccess/namedLocations")
        return raw if isinstance(raw, list) else []
    except Exception as exc:
        log.debug("Named locations unavailable: %s", exc)
        return []


def _collect_auth_strength_policies(client: GraphClient) -> list:
    try:
        raw = client.get("policies/authenticationStrengthPolicies")
        return raw if isinstance(raw, list) else []
    except Exception as exc:
        log.debug("AuthenticationStrengthPolicies unavailable: %s", exc)
        return []


def _collect_risky_users(client: GraphClient) -> list:
    select = "id,userPrincipalName,riskLevel,riskState,riskLastUpdatedDateTime"
    try:
        raw = client.get(f"identityProtection/riskyUsers?$select={select}")
        return raw if isinstance(raw, list) else []
    except Exception as exc:
        log.debug("RiskyUsers unavailable (403 = no P2 license): %s", exc)
        return []


def _collect_mfa_registration(client: GraphClient) -> list:
    # No $select — the endpoint returns ~16 fields by default; maintaining a select
    # risks 400s when Graph renames properties (e.g. defaultMfaMethod ->
    # userPreferredMethodForSecondaryAuthentication).
    try:
        raw = client.get("reports/authenticationMethods/userRegistrationDetails")
        return raw if isinstance(raw, list) else []
    except Exception as exc:
        log.warning("MFA registration details unavailable (need Reports.Read.All): %s", exc)
        return []


def _collect_license_skus(client: GraphClient) -> dict:
    raw = client.get("subscribedSkus?$select=id,skuId,skuPartNumber,capabilityStatus,consumedUnits,prepaidUnits,servicePlans")
    return {s["skuPartNumber"]: s for s in (raw if isinstance(raw, list) else [])}


def _collect_enterprise_apps(client: GraphClient) -> list:
    select = ("id,displayName,appId,servicePrincipalType,signInAudience,"
              "appRoleAssignmentRequired,tags,createdDateTime,publisherName,"
              "verifiedPublisher,oauth2PermissionScopes,appRoles")
    try:
        raw = client.get(f"servicePrincipals?$filter=servicePrincipalType eq 'Application'&$select={select}")
        return raw if isinstance(raw, list) else []
    except Exception as exc:
        log.debug("Enterprise apps (servicePrincipals) unavailable: %s", exc)
        return []


# ---------------------------------------------------------------------------
# Derived builders
# ---------------------------------------------------------------------------

def _build_mfa_enrollment_summary(details: list) -> dict:
    if not details:
        return {}
    total     = len(details)
    registered = sum(1 for u in details if u.get("isMfaRegistered") or u.get("isMfaCapable"))
    not_reg   = total - registered
    method_counts: dict[str, int] = Counter()
    for u in details:
        for m in (u.get("methodsRegistered") or []):
            method_counts[m] += 1
    return {
        "TotalUsers":           total,
        "RegisteredUsers":      registered,
        "NotRegisteredUsers":   not_reg,
        "RegistrationPercent":  round(registered / total * 100, 1) if total else 0.0,
        "MethodCounts":         dict(method_counts),
    }


def _build_mfa_enforcement_summary(ca_policies: list) -> dict:
    if not ca_policies:
        return {}
    enabled = [p for p in ca_policies if p.get("state") == "enabled"]
    mfa_required = [
        p for p in enabled
        if any(
            g.get("builtInControls") and "mfa" in [c.lower() for c in g.get("builtInControls", [])]
            for g in [p.get("grantControls") or {}]
        )
    ]
    return {
        "TotalEnabledPolicies":        len(enabled),
        "PoliciesRequiringMfa":        len(mfa_required),
        "MFAConditionalAccessPolicies": len(mfa_required),
    }


def _build_mfa_method_posture(details: list) -> dict:
    if not details:
        return {}
    weak    = {"sms", "voiceMobile", "voiceAlternateMobile", "voiceOffice"}
    strong  = {"microsoftAuthenticatorApp", "softwareOneTimePasscode", "email"}
    phish   = {"windowsHelloForBusiness", "fido2", "passKeyDeviceBound"}

    weak_only    = 0
    strong_count = 0
    phish_count  = 0
    for u in details:
        methods = set(u.get("methodsRegistered") or [])
        if methods & phish:
            phish_count += 1
        elif methods & strong:
            strong_count += 1
        elif methods & weak:
            weak_only += 1
    return {
        "UsersWithWeakMethodsOnly":       weak_only,
        "UsersWithStrongMethods":         strong_count,
        "UsersWithPhishingResistantMethods": phish_count,
    }


def _build_priv_access_summary(admins: list) -> dict:
    role_counts: dict[str, int] = Counter(a.get("RoleName", "") for a in admins)
    return {"RoleMemberCounts": dict(role_counts), "TotalPrivilegedAccounts": len(set(a.get("id") for a in admins))}


def _build_priv_access_remediation(admins: list, users: dict) -> list:
    ga_members = [a for a in admins if "Global Administrator" in a.get("RoleName", "")]
    stale_ga   = _count_stale_accounts(ga_members)
    return [
        {"Signal": "Global Administrator Count",    "Count": len(ga_members),  "RiskLevel": "High" if len(ga_members) > 3 else "Low"},
        {"Signal": "Stale privileged accounts",     "Count": stale_ga,         "RiskLevel": "High" if stale_ga > 0 else "Low"},
        {"Signal": "Privileged accounts reviewed",  "Count": len(set(a.get("id") for a in admins)), "RiskLevel": "Info"},
    ]


def _build_guest_signin_summary(users: dict) -> dict:
    user_list = list(users.values())
    guests   = [u for u in user_list if u.get("userType") == "Guest"]
    enabled  = [u for u in guests if u.get("accountEnabled")]
    return {"GuestCount": len(guests), "EnabledGuestCount": len(enabled)}


def _build_device_management_summary(devices: dict, intune_devices: dict | None = None) -> dict:
    device_list = list(devices.values())
    total      = len(device_list)
    managed    = sum(1 for d in device_list if d.get("isManaged"))
    compliant  = sum(1 for d in device_list if d.get("isCompliant"))
    result = {
        "TotalDevices":     total,
        "ManagedDevices":   managed,
        "CompliantDevices": compliant,
        "UnmanagedDevices": total - managed,
    }
    if intune_devices:
        intune_list = list(intune_devices.values())
        result["IntuneEnrolledCount"]    = len(intune_list)
        result["IntuneComplianceStates"] = dict(Counter(d.get("complianceState", "unknown") for d in intune_list))
        result["IntuneManagementAgents"] = dict(Counter(d.get("managementAgent", "unknown") for d in intune_list))
        result["IntuneEncryptedCount"]   = sum(1 for d in intune_list if d.get("isEncrypted"))
    return result


def _build_group_licensing_summary(groups: dict, skus: dict) -> list:
    sku_name_map = {s.get("skuId", ""): s.get("skuPartNumber", s.get("skuId", "")) for s in skus.values()} if isinstance(skus, dict) else {}
    result = []
    for g in groups.values():
        assigned = g.get("assignedLicenses") or []
        if not assigned:
            continue
        sku_names = ", ".join(sku_name_map.get(lic.get("skuId", ""), lic.get("skuId", "")) for lic in assigned)
        owners = g.get("owners") or []
        result.append({
            "GroupName":           g.get("displayName", ""),
            "LicenseSKUs":         sku_names,
            "OwnerCount":          len(owners),
            "IsOwnerless":         len(owners) == 0,
            "MemberCount":         g.get("MemberCount"),         # None if GroupMember.Read.All not granted
            "DisabledMemberCount": g.get("DisabledMemberCount"), # None if not collected
        })
    return sorted(result, key=lambda x: x.get("GroupName", ""))


def _build_app_summary(apps: list) -> dict:
    total      = len(apps)
    has_perms  = sum(1 for a in apps if a.get("oauth2PermissionScopes") or a.get("appRoles"))
    verified   = sum(1 for a in apps if (a.get("verifiedPublisher") or {}).get("verifiedPublisherId"))
    return {
        "TotalApplications":         total,
        "ApplicationsWithPermissions": has_perms,
        "VerifiedPublisherApps":     verified,
        "UnverifiedApps":            total - verified,
    }


def _build_license_optimization(users: dict, skus: dict) -> list:
    from datetime import datetime, timezone
    _now = datetime.now(timezone.utc)
    candidates = []
    for u in users.values():
        if not u.get("assignedLicenses"):
            continue
        upn = u.get("userPrincipalName", "")
        if not u.get("accountEnabled"):
            candidates.append({"UserPrincipalName": upn, "Signal": "Disabled user with assigned license"})
            continue
        # Inactive licensed users: enabled but no sign-in for 90+ days
        sia = u.get("signInActivity") or {}
        last = sia.get("lastSignInDateTime") or sia.get("lastNonInteractiveSignInDateTime")
        if last:
            try:
                dt = datetime.fromisoformat(last.replace("Z", "+00:00"))
                if (_now - dt).days >= 90:
                    candidates.append({"UserPrincipalName": upn, "Signal": f"No sign-in for {(_now - dt).days} days (licensed)"})
            except (ValueError, AttributeError):
                pass
        else:
            # No sign-in record at all
            candidates.append({"UserPrincipalName": upn, "Signal": "No sign-in record (licensed)"})
    candidates.sort(key=lambda c: c.get("Signal", ""))
    return candidates[:100]


def _build_mfa_enforcement_gaps(users: dict, mfa_details: list, ca_policies: list) -> list:
    if not mfa_details:
        return []
    not_registered = {u.get("userPrincipalName", ""): u for u in mfa_details
                      if not (u.get("isMfaRegistered") or u.get("isMfaCapable"))}
    return [{"UserPrincipalName": upn, "GapType": "NotRegisteredForMFA"} for upn in list(not_registered)[:200]]


def _build_mfa_scope_review(ca_policies: list) -> list:
    mfa_policies = []
    for p in ca_policies:
        gc = p.get("grantControls") or {}
        if "mfa" in [c.lower() for c in (gc.get("builtInControls") or [])]:
            users_incl = (p.get("conditions", {}).get("users", {}).get("includeUsers") or [])
            users_excl = (p.get("conditions", {}).get("users", {}).get("excludeUsers") or [])
            mfa_policies.append({
                "PolicyName":      p.get("displayName", ""),
                "State":           p.get("state", ""),
                "IncludeUsers":    users_incl,
                "ExcludeUserCount": len(users_excl),
            })
    return mfa_policies


def _build_ca_optimization(ca_policies: list) -> list:
    issues = []
    for p in ca_policies:
        if p.get("state") == "disabled":
            issues.append({"PolicyName": p.get("displayName", ""), "Signal": "Policy is disabled", "Severity": "Info"})
        excl = len((p.get("conditions", {}).get("users", {}).get("excludeUsers") or []))
        excl += len((p.get("conditions", {}).get("users", {}).get("excludeGroups") or []))
        if excl > 10:
            issues.append({"PolicyName": p.get("displayName", ""), "Signal": f"{excl} exclusions", "Severity": "Medium"})
    return issues


def _build_auth_config(auth_methods: dict, sec_defaults: dict, ca_policies: list) -> dict:
    return {
        "SecurityDefaultsEnabled":    sec_defaults.get("isEnabled"),
        "AuthenticationMethodPolicy": auth_methods.get("id", ""),
        "CAPolicyCount":              len(ca_policies),
    }


def _build_ca_summary(ca_policies: list) -> dict:
    if not ca_policies:
        return {"Summary": {}}
    total    = len(ca_policies)
    enabled  = sum(1 for p in ca_policies if p.get("state") == "enabled")
    report   = sum(1 for p in ca_policies if p.get("state") == "enabledForReportingButNotEnforced")
    excl     = sum(1 for p in ca_policies
                   if (p.get("conditions", {}).get("users", {}).get("excludeUsers") or
                       p.get("conditions", {}).get("users", {}).get("excludeGroups")))

    def _has_control(target: str) -> bool:
        for p in ca_policies:
            if p.get("state") not in ("enabled", "enabledForReportingButNotEnforced"):
                continue
            gc = p.get("grantControls") or {}
            if target in [c.lower() for c in (gc.get("builtInControls") or [])]:
                return True
        return False

    def _targets_guests() -> bool:
        for p in ca_policies:
            users = p.get("conditions", {}).get("users", {})
            if "GuestsOrExternalUsers" in (users.get("includeGuestsOrExternalUsers") or {}).get("guestOrExternalUserTypes", ""):
                return True
        return False

    def _has_risk() -> bool:
        for p in ca_policies:
            cond = p.get("conditions", {})
            if cond.get("signInRiskLevels") or cond.get("userRiskLevels"):
                return True
        return False

    priv_roles = [
        p for p in ca_policies
        if any(r.get("includeRoles") for r in [(p.get("conditions") or {}).get("users") or {}])
    ]

    return {"Summary": {
        "TotalPolicies":                      total,
        "EnabledPolicies":                    enabled,
        "ReportOnlyPolicies":                 report,
        "PoliciesWithExclusions":             excl,
        "PoliciesBlockingLegacyAuth":         sum(1 for p in ca_policies if _policy_blocks_legacy(p)),
        "PoliciesRequiringCompliantDevice":   sum(1 for p in ca_policies if _policy_requires_compliant(p)),
        "PoliciesUsingRiskSignals":           sum(1 for p in ca_policies if _policy_uses_risk(p)),
        "HasGuestCoverage":                   _targets_guests(),
        "HasPrivilegedRoleCoverage":          bool(priv_roles),
        "HasLegacyAuthProtection":            any(_policy_blocks_legacy(p) for p in ca_policies),
        "HasCompliantDeviceRequirement":      any(_policy_requires_compliant(p) for p in ca_policies),
        "HasRiskBasedCoverage":               _has_risk(),
    }}


def _policy_blocks_legacy(p: dict) -> bool:
    for app in (p.get("conditions", {}).get("clientAppTypes") or []):
        if app.lower() in ("exchangeactivesync", "other"):
            gc = p.get("grantControls") or {}
            if gc.get("operator") == "OR" and "block" in [c.lower() for c in (gc.get("builtInControls") or [])]:
                return True
    return False


def _policy_requires_compliant(p: dict) -> bool:
    gc = p.get("grantControls") or {}
    return "compliantDevice" in [c.lower() for c in (gc.get("builtInControls") or [])]


def _policy_uses_risk(p: dict) -> bool:
    cond = p.get("conditions") or {}
    return bool(cond.get("signInRiskLevels") or cond.get("userRiskLevels"))


def _count_stale_accounts(accounts: list) -> int:
    now = datetime.now(timezone.utc)
    stale = 0
    for a in accounts:
        last = a.get("signInActivity", {})
        if isinstance(last, dict):
            last_sign_in = last.get("lastSignInDateTime") or last.get("lastNonInteractiveSignInDateTime")
        else:
            last_sign_in = None
        if not last_sign_in:
            stale += 1
            continue
        try:
            dt = datetime.fromisoformat(last_sign_in.replace("Z", "+00:00"))
            if (now - dt).days >= _STALE_DAYS:
                stale += 1
        except (ValueError, AttributeError):
            stale += 1
    return stale
