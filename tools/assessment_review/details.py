"""Object-level lists for the Customer Action Plan.

These sheets name the accounts, devices, sites, apps and incidents behind each
action, with a per-row action, so the customer can work through them. They are
built only from the retained collection and verified collected inputs. They go
into the Customer Action Plan only: never into the Word documents or
review-results.json. Secret and certificate values are never read; only their
presence and expiry dates.
"""
from __future__ import annotations

from collections import Counter
import csv
from dataclasses import dataclass, field
import re

from .facts import GLOBAL_ADMIN_TEMPLATE, HIGH_APPLICATION, MICROSOFT_TENANTS
from .util import parse_date, age_days, ascii_clean

STALE_DAYS = 90
INACTIVE_SITE_DAYS = 180
# Delegated scopes the assessment engine counts as high privilege.
HIGH_DELEGATED = ("mail.readwrite", "files.readwrite", "directory.readwrite")
OWNERS_URL = re.compile(r"servicePrincipals/([0-9a-fA-F-]{36})/owners")
PHISHING_RESISTANT = {"windowsHelloForBusiness", "fido2SecurityKey", "passKeyDeviceBound",
                      "passKeyDeviceBoundAuthenticator", "passKeyDeviceBoundWindowsHello",
                      "x509CertificateMultiFactor", "platformCredential"}
APP_METHODS = {"microsoftAuthenticatorPush", "microsoftAuthenticatorPasswordless", "softwareOneTimePasscode",
               "hardwareOneTimePasscode"}
PHONE_METHODS = {"mobilePhone", "alternateMobilePhone", "officePhone", "sms", "voice", "email"}
METHOD_LABEL = {
    "windowsHelloForBusiness": "Windows Hello", "fido2SecurityKey": "Security key",
    "passKeyDeviceBound": "Passkey", "passKeyDeviceBoundAuthenticator": "Passkey (Authenticator)",
    "passKeyDeviceBoundWindowsHello": "Passkey (Windows Hello)", "x509CertificateMultiFactor": "Certificate",
    "platformCredential": "Platform credential", "microsoftAuthenticatorPush": "Authenticator push",
    "microsoftAuthenticatorPasswordless": "Authenticator phone sign-in",
    "softwareOneTimePasscode": "Authenticator code", "hardwareOneTimePasscode": "Hardware token",
    "mobilePhone": "Text or call (mobile)", "alternateMobilePhone": "Text or call (alternate)",
    "officePhone": "Call (office phone)", "email": "Email", "sms": "Text", "voice": "Call"}
SHARING = {0: "Only people in the organization", 1: "New and existing guests", 2: "Anyone (no sign-in)",
           3: "Existing guests only"}
TEMPLATE = {"GROUP": "Team site (Microsoft 365 group)", "SITEPAGEPUBLISHING": "Communication site",
            "STS": "Team site (no group)", "CONTENTCTR": "Content center", "BDR": "Document center",
            "TEAMCHANNEL": "Teams channel site", "APPCATALOG": "App catalog"}


@dataclass
class Sheet:
    name: str
    table: str
    purpose: str
    headers: tuple
    rows: list = field(default_factory=list)
    refs: list = field(default_factory=list)      # per row: [(finding id, label), ...]
    note: str = ""

    def add(self, values, refs=()):
        self.rows.append(tuple(_cell(v) for v in values))
        self.refs.append([r for r in refs if r[0]])


def _cell(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, (int, float)):
        return value
    return ascii_clean(str(value))


def _attrs(collection, service):
    info = (collection.get("service_results") or {}).get(service) or {}
    if isinstance(info, list):
        info = next((r for r in info if isinstance(r, dict) and r.get("_client")), {})
    return ((info.get("_client") or {}).get("$collection_attributes")) or {}


def _records(dataset):
    """assessment_datasets entries are lists of {records, source} pages."""
    out = []
    for page in dataset or []:
        if isinstance(page, dict) and "records" in page:
            out.extend(page.get("records") or [])
        elif isinstance(page, dict):
            out.append(page)
    return out


def _day(value):
    parsed, problem = parse_date(str(value)[:10] if value else None)
    return None if problem else parsed


def _when(value, review_date):
    day = _day(value)
    if not day:
        return "Never recorded", None
    return day.isoformat(), age_days(day, review_date)


def _size(num_bytes):
    try:
        value = float(num_bytes)
    except (TypeError, ValueError):
        return ""
    for unit in ("bytes", "KB", "MB", "GB", "TB"):
        if value < 1024 or unit == "TB":
            return f"{value:.0f} {unit}" if unit == "bytes" else f"{value:.1f} {unit}"
        value /= 1024
    return ""


class Finder:
    """Locates the review item a detail row supports, by topic, finding key or id."""

    def __init__(self, items):
        self.items = [i for i in items if i["ReviewDecision"] != "DuplicateOrOverlapping"]

    def __call__(self, topic=None, key=None, fid=None):
        for i in self.items:
            if fid and i["FindingId"] == fid:
                return i["FindingId"]
            if fid:
                continue
            if topic and i["Topic"] != topic:
                continue
            if key and key not in (i.get("FindingKey") or ""):
                continue
            return i["FindingId"]
        return None


class Context:
    def __init__(self, pkg, items, review_date):
        self.pkg = pkg
        self.review_date = review_date
        self.find = Finder(items)
        coll = pkg.collection or {}
        self.entra = _attrs(coll, "entra_info")
        self.defender = _attrs(coll, "defender_info")
        self.m365 = _attrs(coll, "m365_result")
        datasets = self.entra.get("assessment_datasets") or {}
        self.signin = {}
        for r in _records(datasets.get("user_signin_activity")):
            activity = r.get("signInActivity") or {}
            last = max([v for v in (activity.get("lastSignInDateTime"), activity.get("lastNonInteractiveSignInDateTime"))
                        if v] or [None], key=lambda v: v or "")
            self.signin[r.get("id")] = last
        self.users = {u.get("id"): u for u in self.m365.get("users") or []}
        self.users_by_upn = {str(u.get("userPrincipalName") or "").lower(): u for u in self.users.values()}
        self.registration = {r.get("id"): r for r in self.entra.get("auth_methods_registration") or []}
        services = coll.get("service_results") or {}
        m365_info = next((r for r in services.get("m365_result") or [] if isinstance(r, dict)), {})
        self.copilot_skus = {
            lic.get("sku_id") for lic in (m365_info.get("licenses") or []) + ((services.get("entra_info") or {})
                                                                               .get("licenses") or [])
            if "copilot" in str(lic.get("sku_part_number")).lower()
            and "ccibots" not in str(lic.get("sku_part_number")).lower()}

    def name_of(self, object_id):
        user = self.users.get(object_id) or self.registration.get(object_id)
        if user:
            return user.get("userPrincipalName") or user.get("displayName") or object_id
        return object_id

    def csv_inputs(self):
        """Rows of verified collected CSV inputs, keyed by report kind."""
        out = {"permissions": [], "items": [], "links": []}
        for f in self.pkg.files:
            if f["Role"] != "collected-input" or f["HashStatus"] != "Verified":
                continue
            path = self.pkg.root / f["RelativePath"]
            if path.suffix.lower() != ".csv":
                continue
            with open(path, newline="", encoding="utf-8-sig") as handle:
                rows = list(csv.DictReader(handle))
            if not rows:
                continue
            headers = set(rows[0].keys())
            if "Number of users having access" in headers:
                out["permissions"].extend(rows)
            elif "Recipient" in headers and "Item type" in headers:
                out["items"].extend(rows)
            elif "Links created" in headers:
                for r in rows:
                    r["_kind"] = path.stem
                out["links"].extend(rows)
        return out


# ------------------------------------------------------------------ identity

def _strength(methods, registered):
    found = set(methods or [])
    if found & PHISHING_RESISTANT:
        return "Phishing-resistant"
    if found & APP_METHODS:
        return "Authenticator app or code"
    if found & PHONE_METHODS:
        return "Phone or text only"
    return "Registered (method not listed)" if registered else "No MFA method"


def _methods(methods):
    return ", ".join(sorted(METHOD_LABEL.get(m, m) for m in methods or [])) or "None"


def admin_accounts(c):
    e = c.entra
    roles = {}
    for d in e.get("role_definitions") or []:
        for key in (d.get("id"), d.get("templateId")):
            if key:
                roles[key] = d.get("displayName") or key
    schedules = {(s.get("principalId"), s.get("roleDefinitionId")): s for s in e.get("role_assignment_schedules") or []}
    principals = {}
    for a in e.get("role_assignments") or []:
        pid = a.get("principalId")
        entry = principals.setdefault(pid, {"principal": a.get("principal") or {}, "roles": [], "kinds": set(),
                                            "global": False})
        rid = a.get("roleDefinitionId")
        entry["roles"].append(roles.get(rid, "Unrecognized role"))
        entry["global"] |= rid == GLOBAL_ADMIN_TEMPLATE or roles.get(rid) == "Global Administrator"
        expiration = (((schedules.get((pid, rid)) or {}).get("scheduleInfo") or {}).get("expiration") or {})
        end = expiration.get("endDateTime")
        entry["kinds"].add(f"Active until {str(end)[:10]}" if end else "Permanent")
    for s in e.get("role_eligibility_schedules") or []:
        pid = s.get("principalId")
        entry = principals.setdefault(pid, {"principal": s.get("principal") or {}, "roles": [], "kinds": set(),
                                            "global": False})
        entry["roles"].append(roles.get(s.get("roleDefinitionId"), "Unrecognized role") + " (eligible)")
        entry["kinds"].add("Eligible (PIM)")
    global_count = sum(1 for p in principals.values() if p["global"])
    sp_signin = {x.get("appId"): x for x in e.get("service_principal_signin_activities") or []}
    sheet = Sheet("Admin accounts", "AdminAccounts",
                  "Every account and app that holds an administrator role, its sign-in methods and what to fix.",
                  ("Account", "Name", "Type", "Roles", "Assignment", "Enabled", "MFA methods", "MFA strength",
                   "Last sign-in", "Days since sign-in", "Issues", "Action needed", "Actions"))
    find = c.find
    mfa_item = find("PRIV", "WithoutMFARegistration") or find(fid="RVW-ADMIN-MFA")
    strength_item = find(fid="ENT-007") or find("MFA", "passwordless")
    rows = []
    for pid, entry in principals.items():
        p = entry["principal"]
        kind = str(p.get("@odata.type") or "").rsplit(".", 1)[-1]
        issues, actions, refs = [], [], []
        roles_text = "; ".join(sorted(set(entry["roles"])))
        assignment = ", ".join(sorted(entry["kinds"]))
        if kind == "servicePrincipal":
            last = ((sp_signin.get(p.get("appId")) or {}).get("lastSignInActivity") or {}).get("lastSignInDateTime")
            when, days = _when(last, c.review_date)
            creds = _credentials(p)
            issues.append("App identity with an admin role")
            actions.append("Confirm the app's owner and purpose and that it needs this role; check its secrets and "
                           f"certificates in App registrations ({creds})")
            refs.append((find("PRIV", "ServicePurpose"), "apps with admin roles"))
            row = (p.get("displayName") or pid, p.get("displayName"), "App (service principal)", roles_text, assignment,
                   p.get("accountEnabled") is not False, "Not applicable", "Not applicable", when,
                   days if days is not None else "")
        elif kind == "group":
            issues.append("Role assigned to a group")
            actions.append("Review the group's members; everyone in it holds these roles")
            refs.append((find("PRIV", "PermanentPrivilege"), "role-assignable groups"))
            row = (p.get("displayName") or pid, p.get("displayName"), "Group", roles_text, assignment, True,
                   "Not applicable", "Not applicable", "Not applicable", "")
            days = None
        else:
            reg = c.registration.get(pid)
            methods = (reg or {}).get("methodsRegistered") or []
            strength = _strength(methods, reg.get("isMfaRegistered")) if reg else "Not in registration report"
            enabled = p.get("accountEnabled")
            when, days = _when(c.signin.get(pid), c.review_date)
            guest = str(p.get("userType") or "").lower() == "guest"
            if enabled is False:
                issues.append("Account disabled")
                actions.append("Remove the admin roles from this disabled account")
                refs.append((find("PRIV", "DisabledIdentity"), "disabled accounts"))
            else:
                if strength == "No MFA method":
                    issues.append("No MFA method")
                    actions.append("Register MFA now (prefer a passkey, Windows Hello or a security key), or remove "
                                   "the admin roles")
                    eligible_only = entry["kinds"] == {"Eligible (PIM)"}
                    refs.append((mfa_item, "eligible (PIM) admins with no MFA method" if eligible_only
                                 else "active admins with no MFA method"))
                elif strength == "Phone or text only":
                    issues.append("Phone or text MFA only")
                    actions.append("Replace phone or text MFA with a phishing-resistant method (passkey, Windows "
                                   "Hello or security key)")
                    refs.append((mfa_item, "with phone or text MFA only"))
                    refs.append((strength_item, "admins without phishing-resistant MFA"))
                elif strength == "Not in registration report":
                    issues.append("MFA state unknown")
                    actions.append("The registration report did not include this account: check its methods in "
                                   "Microsoft Entra admin center > Users > Authentication methods")
                elif strength != "Phishing-resistant":
                    issues.append("No phishing-resistant method")
                    actions.append("Add a phishing-resistant method (passkey, Windows Hello or security key) for "
                                   "admin sign-in")
                    refs.append((strength_item, "admins without phishing-resistant MFA"))
                if days is None or days > STALE_DAYS:
                    issues.append("No recent sign-in")
                    actions.append("Not signed in " + (f"for {days} days" if days is not None else "on record")
                                   + ": remove the roles if the account is not needed")
                    refs.append((find("PRIV", "ActivityUnknown"), "admins with no recent sign-in"))
                if guest:
                    issues.append("Guest account")
                    actions.append("Guest with an admin role: confirm this is intended, or remove the role")
            row = (p.get("userPrincipalName") or pid, p.get("displayName"), "Guest" if guest else "User", roles_text,
                   assignment, enabled is not False, _methods(methods), strength, when,
                   days if days is not None else "")
        if entry["global"] and global_count >= 5:
            issues.append("Global Administrator")
            actions.append(f"One of {global_count} Global Administrators (Microsoft recommends fewer than five): "
                           "keep only if needed, otherwise move to a narrower role")
            refs.append((find(fid="RVW-GLOBAL-ADMIN-COUNT"), "Global Administrators"))
        if "Permanent" in assignment and kind == "user":
            issues.append("Permanent assignment")
            actions.append("Make the role time-limited with Privileged Identity Management (Entra ID P2), or record "
                           "why it must be permanent")
            refs.append((find("PRIV", "PermanentPrivilege"), "permanent assignments"))
        rows.append((row + ("; ".join(issues) or "None", "; ".join(actions) or "No action"), refs, len(issues)))
    rows.sort(key=lambda r: (-r[2], str(r[0][0]).lower()))
    for values, refs, _ in rows:
        sheet.add(values, refs)
    if not principals:
        sheet.note = "Role assignments were not collected."
    return sheet


def _credentials(principal):
    secrets = principal.get("passwordCredentials") or []
    certs = principal.get("keyCredentials") or []
    if not secrets and not certs:
        return "none on the service principal; app registration credentials were not collected"
    parts = []
    for label, values in (("secret", secrets), ("certificate", certs)):
        if values:
            ends = sorted(str(v.get("endDateTime"))[:10] for v in values if v.get("endDateTime"))
            parts.append(f"{len(values)} {label}{'s' if len(values) > 1 else ''}"
                         + (f", latest expiry {ends[-1]}" if ends else ""))
    return "; ".join(parts)


def users_without_mfa(c):
    sheet = Sheet("Users without MFA", "UsersWithoutMfa",
                  "Accounts that have not registered any MFA method.",
                  ("Account", "Name", "Type", "Administrator", "Enabled", "Licensed", "Last sign-in",
                   "Days since sign-in", "Action needed", "Actions"))
    ref = c.find("MFA", "registration")
    rows = []
    for r in c.entra.get("auth_methods_registration") or []:
        if r.get("isMfaRegistered") is not False:
            continue
        user = c.users.get(r.get("id")) or {}
        enabled = user.get("accountEnabled")
        licensed = bool(user.get("assignedLicenses")) if user else None
        when, days = _when(c.signin.get(r.get("id")), c.review_date)
        guest = str(r.get("userType") or "").lower() == "guest"
        if enabled is False:
            action = "Account is disabled: delete it if it is no longer needed"
        elif guest:
            action = "Guest: confirm the guest is still needed; require MFA for guests with Conditional Access"
        elif days is None and licensed is False:
            action = ("Unlicensed with no sign-in on record (often a shared mailbox, room or service account): block "
                      "sign-in if nobody signs in to it, otherwise register MFA")
        elif days is None or days > STALE_DAYS:
            action = ("Not signed in " + (f"for {days} days" if days is not None else "on record")
                      + ": disable the account if it is no longer needed, otherwise register MFA")
        else:
            action = "Ask the user to register MFA at mysignins.microsoft.com/security-info"
        rows.append(((r.get("userPrincipalName"), r.get("userDisplayName"), "Guest" if guest else "Member",
                      bool(r.get("isAdmin")), "Unknown" if enabled is None else enabled,
                      "Unknown" if licensed is None else licensed, when, days if days is not None else "", action),
                     (not r.get("isAdmin"), days is None, days or 0, str(r.get("userPrincipalName")).lower())))
    rows.sort(key=lambda r: r[1])
    for values, _ in rows:
        sheet.add(values, [(ref, "accounts without MFA")])
    return sheet


def copilot_users(c):
    sheet = Sheet("Copilot users", "CopilotUsers",
                  "People who hold a Microsoft 365 Copilot license today, with their sign-in and MFA state.",
                  ("Account", "Name", "Enabled", "Administrator", "Last sign-in", "Days since sign-in",
                   "MFA methods", "MFA strength", "Action needed", "Actions"))
    find = c.find
    refs = [(find(fid="RVW-EXISTING-COPILOT-USE"), "licensed Copilot users"),
            (find("LICENSING", "license.assignment"), "licensed Copilot users")]
    rows = []
    for u in c.users.values():
        if not any(lic.get("skuId") in c.copilot_skus for lic in u.get("assignedLicenses") or []):
            continue
        reg = c.registration.get(u.get("id"))
        methods = (reg or {}).get("methodsRegistered") or []
        strength = _strength(methods, reg.get("isMfaRegistered")) if reg else "Not in registration report"
        when, days = _when(c.signin.get(u.get("id")), c.review_date)
        actions = []
        if u.get("accountEnabled") is False:
            actions.append("Account disabled: remove the Copilot license")
        else:
            if days is None or days > STALE_DAYS:
                actions.append("Not signed in " + (f"for {days} days" if days is not None else "on record")
                               + ": reassign the license to someone who will use it")
            if strength == "No MFA method":
                actions.append("Register MFA before this person keeps using Copilot")
        rows.append(((u.get("userPrincipalName"), u.get("displayName"), u.get("accountEnabled") is not False,
                      bool((reg or {}).get("isAdmin")), when, days if days is not None else "", _methods(methods),
                      strength, "; ".join(actions) or "No action"), str(u.get("userPrincipalName")).lower()))
    for values, _ in sorted(rows, key=lambda r: (r[0][-1] == "No action", r[1])):
        sheet.add(values, refs)
    if c.copilot_skus and not rows:
        sheet.note = "No user in the collected user list holds a Copilot license."
    return sheet


def sign_in_policies(c):
    sheet = Sheet("Sign-in policies", "SignInPolicies",
                  "Conditional Access policies: which are on, who they cover and who is excluded.",
                  ("Policy", "State", "Applies to", "Excluded", "Apps", "Requires", "Action needed", "Actions"))
    ca_item = c.find("CA")
    emergency = c.find(fid="RVW-EMERGENCY-ACCESS")
    admin_item = c.find("PRIV", "WithoutMFARegistration") or c.find(fid="RVW-ADMIN-MFA")
    order = {"enabledForReportingButNotEnforced": 0, "disabled": 1, "enabled": 2}
    for p in sorted(c.entra.get("ca_policies") or [], key=lambda p: (order.get(p.get("state"), 3),
                                                                     str(p.get("displayName")).lower())):
        cond = p.get("conditions") or {}
        users = cond.get("users") or {}
        apps = cond.get("applications") or {}
        grant = p.get("grantControls") or {}
        controls = [str(x) for x in grant.get("builtInControls") or []]
        if grant.get("authenticationStrength"):
            controls.append("Authentication strength: " + str((grant["authenticationStrength"] or {}).get(
                "displayName") or "custom"))
        state = {"enabled": "On", "enabledForReportingButNotEnforced": "Report-only", "disabled": "Off"}.get(
            p.get("state"), str(p.get("state")))
        all_users = "All" in (users.get("includeUsers") or [])
        scope = ("All users" if all_users else ", ".join(filter(None, (
            _count(users.get("includeUsers"), "user"), _count(users.get("includeGroups"), "group"),
            _count(users.get("includeRoles"), "admin role"))))) or "Nobody"
        excluded_users = [c.name_of(x) for x in users.get("excludeUsers") or []]
        excluded = "; ".join(filter(None, ("; ".join(excluded_users[:20]),
                                           _count(users.get("excludeGroups"), "group"),
                                           _count(users.get("excludeRoles"), "admin role")))) or "None"
        is_mfa = "mfa" in [x.lower() for x in controls] or bool(grant.get("authenticationStrength"))
        refs = []
        if state == "Report-only" and is_mfa:
            action = ("Review the report-only results and the exclusions (only emergency access accounts should be "
                      "excluded), then switch the policy On")
            refs = [(ca_item if all_users else admin_item, "MFA policies in report-only mode"),
                    (emergency, "policies with exclusions to confirm")]
        elif state == "Off":
            action = "Off: delete the policy or record why it is kept"
        elif is_mfa and all_users and users.get("excludeUsers"):
            action = "On: confirm every excluded account is an emergency access or sync account"
            refs = [(emergency, "policies with exclusions to confirm")]
        else:
            action = "No change"
        sheet.add((p.get("displayName"), state, scope, excluded,
                   "All apps" if "All" in (apps.get("includeApplications") or []) else
                   _count(apps.get("includeApplications"), "app") or "None",
                   ", ".join(controls) or "Session controls only", action), refs)
    if not c.entra.get("ca_policies"):
        sheet.note = "Conditional Access policies were not collected."
    return sheet


def _count(values, noun):
    values = [v for v in values or [] if v not in ("All", "None")]
    return f"{len(values)} {noun}{'s' if len(values) != 1 else ''}" if values else ""


def risky_accounts(c):
    sheet = Sheet("Risky accounts", "RiskyAccounts",
                  "Accounts Microsoft Entra ID Protection flags as possibly compromised.",
                  ("Account", "Name", "Risk level", "Risk state", "Risk detail", "Last updated", "Recent detections",
                   "Where to investigate", "Action needed", "Actions"))
    ref = c.find("RISK")
    detections = {}
    for d in c.entra.get("risk_detections") or []:
        detections.setdefault(d.get("userId"), []).append(
            f"{d.get('riskEventType')} ({str(d.get('detectedDateTime'))[:10]}, {d.get('riskLevel')})")
    level = {"high": 0, "medium": 1, "low": 2}
    rows = [u for u in c.entra.get("risky_users") or []
            if str(u.get("riskState")) in ("atRisk", "confirmedCompromised")]
    for u in sorted(rows, key=lambda u: (level.get(str(u.get("riskLevel")), 3), str(u.get("userPrincipalName")))):
        sheet.add((u.get("userPrincipalName"), u.get("userDisplayName"), u.get("riskLevel"), u.get("riskState"),
                   u.get("riskDetail"), str(u.get("riskLastUpdatedDateTime") or "")[:10],
                   "; ".join(detections.get(u.get("id"), [])) or "None retained",
                   "Microsoft Entra admin center > ID Protection > Risky users",
                   "Review the user's risky sign-ins. If compromised: reset the password, revoke sessions and confirm "
                   "MFA methods. If not: dismiss the risk."), [(ref, "accounts at risk")])
    if c.entra.get("risky_users") is None:
        sheet.note = "Risky users were not collected."
    return sheet


def guests(c):
    sheet = Sheet("Guests", "Guests", "Guest accounts and when they last signed in.",
                  ("Guest", "Name", "Invited", "Last sign-in", "Days since sign-in", "Action needed", "Actions"))
    ref = c.find("GUEST")
    for g in sorted(c.entra.get("guest_users") or [], key=lambda g: str(g.get("userPrincipalName")).lower()):
        when, days = _when(c.signin.get(g.get("id")), c.review_date)
        action = ("No sign-in " + (f"for {days} days" if days is not None else "on record")
                  + ": remove the guest if no longer needed") if days is None or days > STALE_DAYS else \
            "Confirm the guest still needs access"
        sheet.add((g.get("userPrincipalName"), g.get("displayName"), str(g.get("createdDateTime") or "")[:10], when,
                   days if days is not None else "", action), [(ref, "guest accounts")])
    return sheet


# ------------------------------------------------------------------ devices, incidents, apps

def devices(c):
    sheet = Sheet("Devices", "Devices",
                  "Devices registered in Microsoft Entra ID, their management state and Defender for Endpoint status.",
                  ("Device", "Operating system", "Join type", "Ownership", "Managed", "Compliant", "Last sign-in",
                   "Days since sign-in", "Defender for Endpoint", "Action needed", "Actions"))
    find = c.find
    device_item, mde_item = find("DEVICE"), find("MDE")
    endpoint_dlp = find(fid="PUR-016")
    defender = {}
    for m in c.defender.get("defender_devices") or []:
        for key in (m.get("aadDeviceId"), str(m.get("computerDnsName") or "").split(".")[0].lower()):
            if key:
                defender[key] = m
    join = {"AzureAd": "Entra joined", "ServerAd": "Hybrid joined", "Workplace": "Registered (personal)"}
    rows = []
    for d in _records((c.entra.get("assessment_datasets") or {}).get("directory_devices")):
        if d.get("accountEnabled") is False:
            continue
        when, days = _when(d.get("approximateLastSignInDateTime"), c.review_date)
        m = defender.get(d.get("deviceId")) or defender.get(str(d.get("displayName") or "").lower())
        mde = (f"{m.get('onboardingStatus')}, {m.get('healthStatus')}, last seen {str(m.get('lastSeen'))[:10]}"
               if m else "Not found")
        actions, refs = [], []
        windows = str(d.get("operatingSystem") or "").lower().startswith("windows")
        if days is None or days > STALE_DAYS:
            actions.append("Not seen " + (f"for {days} days" if days is not None else "on record")
                           + ": remove the stale device record")
            refs.append((device_item, "stale device records"))
        else:
            if not d.get("isManaged"):
                actions.append("Not managed: enroll it in Intune (or your device management tool), or block "
                               "unmanaged devices with Conditional Access")
                refs.append((device_item, "unmanaged devices"))
            elif d.get("isCompliant") is False:
                actions.append("Managed but not compliant: fix the failing compliance settings")
                refs.append((device_item, "non-compliant devices"))
            if windows and not m:
                actions.append("Onboard it to Defender for Endpoint")
                refs += [(mde_item, "Windows devices not in Defender"), (endpoint_dlp, "devices not onboarded")]
            elif m and str(m.get("healthStatus")).lower() != "active":
                actions.append("Defender shows it is not reporting: check the sensor")
                refs.append((mde_item, "devices not reporting to Defender"))
        rows.append(((d.get("displayName"), f"{d.get('operatingSystem') or ''} {d.get('operatingSystemVersion') or ''}",
                      join.get(d.get("trustType"), d.get("trustType") or "Unknown"), d.get("deviceOwnership") or "",
                      "Unknown" if d.get("isManaged") is None else d.get("isManaged"),
                      "Unknown" if d.get("isCompliant") is None else d.get("isCompliant"), when,
                      days if days is not None else "", mde, "; ".join(actions) or "No action"), refs,
                     (not actions, days if days is not None else 10 ** 6)))
    rows.sort(key=lambda r: r[2])
    for values, refs, _ in rows:
        sheet.add(values, refs)
    if not rows:
        sheet.note = "The Entra device list was not collected."
    return sheet


def incidents(c):
    sheet = Sheet("Security incidents", "SecurityIncidents",
                  "Open Microsoft Defender XDR incidents. Open each link to see the full detail in the Defender portal.",
                  ("Incident", "Name", "Severity", "Status", "Created", "Days open", "Last updated", "Owner",
                   "Alerts", "What the alerts recommend", "Link", "Action needed", "Actions"))
    find = c.find
    alerts = {}
    for a in c.defender.get("security_alerts") or []:
        alerts.setdefault(str(a.get("incidentId")), []).append(a)
    severity = {"high": 0, "medium": 1, "low": 2, "informational": 3}
    open_items = [i for i in c.defender.get("security_incidents") or []
                  if str(i.get("status")).lower() in ("active", "inprogress")]
    for i in sorted(open_items, key=lambda i: (severity.get(str(i.get("severity")).lower(), 4),
                                               str(i.get("createdDateTime")))):
        created, days = _when(i.get("createdDateTime"), c.review_date)
        related = alerts.get(str(i.get("id")), [])
        titles = sorted({str(a.get("title")) for a in related if a.get("title")})
        advice = next((str(a.get("recommendedActions")) for a in related if a.get("recommendedActions")), "")
        high = str(i.get("severity")).lower() == "high"
        action = ("Triage now: assign an owner, investigate the affected accounts and devices, then resolve and "
                  "classify" if high else
                  "Assign an owner; resolve and classify it, or close it if it is a known benign activity")
        refs = [(find("INCIDENTS"), "open incidents")]
        if not i.get("assignedTo"):
            refs.append((find("EXCEPTIONS"), "incidents with no owner"))
        sheet.add((i.get("id"), i.get("displayName"), i.get("severity"), i.get("status"), created,
                   days if days is not None else "", str(i.get("lastUpdateDateTime") or "")[:10],
                   i.get("assignedTo") or "Nobody", f"{len(related)}: " + "; ".join(titles[:5]) if related else "0",
                   advice[:500], i.get("incidentWebUrl") or "", action), refs)
    return sheet


def apps(c):
    sheet = Sheet("Apps with high access", "HighAccessApps",
                  "Apps with broad permissions to company data, or from publishers Microsoft has not verified.",
                  ("App", "Publisher", "Why flagged", "High-impact permissions", "Other permissions", "Consent",
                   "Last used", "Owners", "Secrets and certificates", "Action needed", "Actions"))
    e = c.entra
    find = c.find
    tenant = str(c.pkg.collection.get("tenant_id") or "").lower()
    by_id = {s.get("id"): s for s in e.get("service_principals") or []}
    role_names = {}
    for s in by_id.values():
        for r in s.get("appRoles") or []:
            role_names[(s.get("id"), r.get("id"))] = r.get("value") or r.get("displayName")
    delegated, consent = {}, {}
    for g in e.get("oauth_permission_grants") or []:
        client = g.get("clientId")
        resource = (by_id.get(g.get("resourceId")) or {}).get("displayName") or "API"
        for scope in str(g.get("scope") or "").split():
            delegated.setdefault(client, set()).add((resource, scope))
        entry = consent.setdefault(client, {"admin": False, "users": set()})
        if g.get("consentType") == "AllPrincipals":
            entry["admin"] = True
        elif g.get("principalId"):
            entry["users"].add(g.get("principalId"))
    application = {}
    datasets = e.get("assessment_datasets") or {}
    for rec in _records(datasets.get("application_permissions")):
        name = role_names.get((rec.get("resourceId"), rec.get("appRoleId"))) or rec.get("appRoleId")
        application.setdefault(rec.get("principalId"), set()).add((rec.get("resourceDisplayName") or "API", name))
    owners = {}
    for page in datasets.get("application_owners") or []:
        match = OWNERS_URL.search(str(((page or {}).get("source") or {}).get("request_url") or ""))
        for rec in (page or {}).get("records") or []:
            parent = rec.get("ParentId") or (match.group(1) if match else None)
            owners.setdefault(parent, []).append(rec.get("userPrincipalName") or rec.get("displayName")
                                                 or rec.get("appDisplayName"))
    signin = {x.get("appId"): x for x in e.get("service_principal_signin_activities") or []}
    rows = []
    for sp_id in sorted(set(delegated) | set(application)):
        sp = by_id.get(sp_id) or {}
        owner_tenant = str(sp.get("appOwnerOrganizationId") or "").lower()
        verified = (sp.get("verifiedPublisher") or {}).get("displayName")
        microsoft = owner_tenant in MICROSOFT_TENANTS
        internal = bool(tenant) and owner_tenant == tenant
        managed = str(sp.get("servicePrincipalType")).lower() == "managedidentity"
        publisher = ("Microsoft" if microsoft else verified or ("Your organization" if internal else
                                                               "Managed identity" if managed else "Unverified"))
        high_delegated = sorted(f"{scope} ({res}, delegated)" for res, scope in delegated.get(sp_id, ())
                                if any(m in scope.lower() for m in HIGH_DELEGATED))
        # Microsoft's own service principals hold application permissions by design.
        high_app = [] if microsoft else sorted(f"{name} ({res}, application)" for res, name in
                                               application.get(sp_id, ()) if HIGH_APPLICATION.search(str(name)))
        unverified = publisher == "Unverified"
        if not (high_delegated or high_app or unverified):
            continue
        why = []
        refs = []
        grants_item = find("CONSENT", "high_impact_grants")
        if high_delegated:
            why.append("Users' delegated access can change mail, files or directory data")
            refs.append((grants_item, "apps with high-impact delegated permissions"))
        if high_app:
            why.append("Tenant-wide application permissions (no user needed)")
            refs.append((grants_item, "apps with tenant-wide application permissions"))
        if unverified:
            why.append("Publisher not verified")
            refs.append((grants_item, "apps from unverified publishers"))
        everything = len(delegated.get(sp_id, ())) + len(application.get(sp_id, ()))
        other = everything - len(high_delegated) - len(high_app)
        c_entry = consent.get(sp_id) or {}
        consent_text = "; ".join(filter(None, (
            "Admin consent for all users" if c_entry.get("admin") else "",
            f"{len(c_entry.get('users') or ())} users consented" if c_entry.get("users") else "",
            "Application permissions granted by an admin" if application.get(sp_id) else "")))
        if c_entry.get("users"):
            refs.append((find("CONSENT", "user_consent_assignment"), "apps users approved themselves"))
        activity = signin.get(sp.get("appId")) or {}
        last = max([v for v in ((activity.get(k) or {}).get("lastSignInDateTime") for k in (
            "lastSignInActivity", "delegatedClientSignInActivity", "applicationAuthenticationClientSignInActivity"))
                    if v] or [None], key=lambda v: v or "")
        when, days = _when(last, c.review_date)
        actions = []
        if days is None or days > STALE_DAYS:
            actions.append("Not used " + (f"for {days} days" if days is not None else "on record")
                           + ": remove its permissions or delete the app")
        if unverified:
            actions.append("Find out what the app is and who approved it; remove it if nobody owns it")
        if high_delegated or high_app:
            actions.append("Confirm it needs these permissions; reduce to read-only or narrower permissions, or "
                           "remove them")
        rows.append(((sp.get("displayName") or sp_id, publisher, "; ".join(why),
                      "; ".join(high_delegated + high_app) or "None", other, consent_text or "Not recorded", when,
                      ", ".join(sorted(filter(None, owners.get(sp_id, [])))) or "Not recorded",
                      "Not collected (the assessment reads service principals, not app registrations)",
                      "; ".join(actions)), refs, (not (high_delegated or high_app), str(sp.get("displayName")).lower())))
    rows.sort(key=lambda r: r[2])
    for values, refs, _ in rows:
        sheet.add(values, refs)
    if not e.get("oauth_permission_grants") and not e.get("service_principals"):
        sheet.note = "App permission grants were not collected."
    return sheet


# ------------------------------------------------------------------ sites

def overshared_sites(c, inputs):
    sheet = Sheet("Overshared sites", "OversharedSites",
                  "Sites and OneDrives where content is open to everyone in the organization or shared by broad links, "
                  "from the SharePoint Data Access Governance reports.",
                  ("Site", "URL", "Kind", "Primary admin", "People with access", "Everyone", "Everyone except external",
                   "Company-wide links", "Anyone links", "Privacy", "Sensitivity label", "Where the broad access is",
                   "Report date", "Options to fix", "Actions"))
    find = c.find
    perm_items = [find("PERMISSIONS", "broad_internal"), find("PERMISSIONS", "not_assessed")]
    sharing_item = find("SHARING")
    grants = {}
    for r in inputs["items"]:
        where = r.get("Item url") or "(whole site)"
        via = f" via {r['Parent group name']}" if r.get("Parent group name") else ""
        grants.setdefault(r.get("Site ID"), []).append(
            (f"{r.get('Item type')} {where}: {r.get('Role definition')} for {r.get('Recipient')}{via}", r))
    links = Counter()
    for r in inputs["links"]:
        try:
            links[(r.get("Site ID"), r["_kind"])] += int(r.get("Links created") or 0)
        except ValueError:
            pass
    rows = []
    for r in inputs["permissions"]:
        site_id = r.get("Site ID")
        nums = {k: _int(r.get(k)) for k in ("EEEU permission count", "Everyone permission count",
                                            "PeopleInYourOrg link count", "Anyone link count")}
        found = grants.get(site_id, [])
        if not (any(nums.values()) or found):
            continue
        onedrive = "-my.sharepoint.com" in str(r.get("Site URL") or "").lower()
        options = []
        for _, g in found:
            if g.get("Parent group type") == "SharePointGroup" and str(g.get("Parent group name")).endswith("Members"):
                options.append("Remove Everyone except external users from the site's Members group (public "
                               "Microsoft 365 group sites add it automatically; making the site Private removes it)")
                break
        if nums["Everyone permission count"]:
            options.append("Remove the Everyone permission from the listed items")
        if any(g.get("Item type") not in ("Web", "Site") for _, g in found):
            options.append("Break the broad grants on the listed files and folders")
        if nums["PeopleInYourOrg link count"]:
            options.append("Replace company-wide links with Specific people links")
        if nums["Anyone link count"]:
            options.append("Remove Anyone links that are not needed")
        if onedrive:
            options.append("Ask the owner to stop sharing these items broadly (Restricted Content Discovery does not "
                           "apply to OneDrive)")
        else:
            options.append("If the site must stay open while it is reviewed, turn on Restricted Content Discovery to "
                           "keep it out of Copilot, or limit the site to a group with restricted access control")
        refs = [(i, "sites with broad access") for i in perm_items]
        if nums["Anyone link count"]:
            refs.append((sharing_item, "sites with Anyone links"))
        detail = [text for text, _ in found]
        rows.append(((r.get("Site name"), r.get("Site URL"), "OneDrive" if onedrive else "SharePoint site",
                      r.get("Primary admin email") or r.get("Primary admin"), _int(r.get("Number of users having access")),
                      nums["Everyone permission count"], nums["EEEU permission count"],
                      nums["PeopleInYourOrg link count"], nums["Anyone link count"], r.get("Site privacy") or "",
                      r.get("Site sensitivity") or "None",
                      "; ".join(detail[:6]) + (f"; +{len(detail) - 6} more" if len(detail) > 6 else "") or
                      "See counts", str(r.get("Report date") or "")[:10], "; ".join(options)), refs,
                     -(sum(nums.values()) + len(found))))
    rows.sort(key=lambda r: (r[2], str(r[0][0]).lower()))
    for values, refs, _ in rows:
        sheet.add(values, refs)
    dates = sorted({v[12] for v in sheet.rows if v[12]})
    if not inputs["permissions"]:
        sheet.note = ("The Data Access Governance site permission reports were not supplied, so overshared sites "
                      "cannot be listed. Run them in the SharePoint admin center > Reports > Data access governance.")
    elif dates:
        age = age_days(_day(dates[0]), c.review_date)
        sheet.note = (f"From reports dated {dates[0]}" + (f" ({age} days before this review)" if age else "")
                      + ". Refresh the reports before relying on the counts.")
    return sheet


def _int(value):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0


def sites(c):
    sheet = Sheet("Sites", "Sites",
                  "SharePoint sites with their owner, last activity and storage, to find ownerless, inactive or "
                  "openly shared sites.",
                  ("Site", "URL", "Type", "Owner", "Last activity", "Last content change", "Days inactive",
                   "Storage used", "Files", "External sharing", "Sensitivity label", "Action needed", "Actions"))
    gov = ((c.m365.get("sharepoint_governance") or {}).get("sites") or {})
    usage = {r.get("Site Id"): r for r in _records((c.m365.get("assessment_datasets") or {}).get("sharepoint_usage"))}
    find = c.find
    owner_item, sharing_item = find("OWNERSHIP"), find("SHARING")
    rows = []
    for s in gov.get("items") or []:
        template = str(s.get("Template") or "")
        if template.startswith("REDIRECTSITE"):
            continue
        u = usage.get(s.get("SiteId")) or {}
        owner_raw = str(s.get("Owner") or "")
        actions, refs = [], []
        if "@" in owner_raw:
            account = c.users_by_upn.get(owner_raw.lower())
            owner = owner_raw + (" (account disabled)" if account and account.get("accountEnabled") is False else "")
            if account and account.get("accountEnabled") is False:
                actions.append("The owner's account is disabled: assign a new owner")
                refs.append((owner_item, "sites whose owner is disabled"))
        elif owner_raw.endswith("_o"):
            owner = "Microsoft 365 group owners (check in Teams and groups)"
        elif owner_raw:
            owner = "System account (no named owner)"
            actions.append("No named owner: assign at least two owners")
            refs.append((owner_item, "sites with no named owner"))
        else:
            owner = "None"
            actions.append("No owner: assign at least two owners")
            refs.append((owner_item, "sites with no owner"))
        activity = u.get("Last Activity Date") or ""
        changed = str(s.get("LastContentModifiedDate") or "")[:10]
        last = _day(activity) or _day(changed)
        days = age_days(last, c.review_date) if last else None
        if days is None or days >= INACTIVE_SITE_DAYS:
            actions.append("Inactive " + (f"for {days} days" if days is not None else "(no activity recorded)")
                           + ": confirm with the owner, then archive it (Microsoft 365 Archive) or delete it")
            refs.append((owner_item, "inactive sites"))
        elif days >= STALE_DAYS:
            actions.append(f"Little activity for {days} days: confirm the site is still needed")
            refs.append((owner_item, "sites with little activity"))
        if s.get("SharingCapability") == 2:
            actions.append("Anyone links are allowed on this site: restrict sharing if they are not needed")
            refs.append((sharing_item, "sites that allow Anyone links"))
        kind = TEMPLATE.get(template.split("#")[0], template or "Unknown")
        rows.append(((s.get("Title"), s.get("Url"), kind, owner, activity or "None recorded", changed,
                      days if days is not None else "", _size(u.get("Storage Used (Byte)")), _int(u.get("File Count")),
                      SHARING.get(s.get("SharingCapability"), str(s.get("SharingCapability"))),
                      s.get("SensitivityLabel") or "None", "; ".join(actions) or "No action"), refs,
                     (not actions, -(days if days is not None else 10 ** 6))))
    rows.sort(key=lambda r: r[2])
    for values, refs, _ in rows:
        sheet.add(values, refs)
    if not gov.get("items"):
        sheet.note = "The SharePoint site inventory was not collected."
    return sheet


# ------------------------------------------------------------------ assembly

def collect(pkg, items, review_date):
    """Build every detail sheet and record, per review item, which sheet lists its objects."""
    c = Context(pkg, items, review_date)
    inputs = c.csv_inputs()
    sheets = [copilot_users(c), admin_accounts(c), users_without_mfa(c), sign_in_policies(c), risky_accounts(c),
              devices(c), overshared_sites(c, inputs), sites(c), guests(c), apps(c), incidents(c)]
    sheets = [s for s in sheets if s.rows or s.note]
    pointers = {}
    for sheet in sheets:
        counts = Counter((fid, label) for refs in sheet.refs for fid, label in dict.fromkeys(refs))
        for (fid, label), n in sorted(counts.items()):
            pointers.setdefault(fid, {}).setdefault(sheet.name, []).append(f"{n} {label}")
    for item in items:
        found = pointers.get(item["FindingId"]) or {}
        item["DetailSheets"] = [{"Sheet": name, "Summary": "; ".join(parts)} for name, parts in found.items()]
    return sheets
