"""Customer-readable wording for each review item.

Every item gets five short fields used by all four deliverables:

- Headline: the recommendation as a short instruction ("Turn on MFA for all users").
- Issue: what was found, in plain words, with the collected counts.
- Why: why it matters for Copilot.
- HowToFix: the practical next step.
- DoneWhen: what shows the item is finished.

Correction explains how the review corrects the original report, RatingChange
how it re-rates the readiness effect, and Change joins the two (each empty when
nothing changed). Templates key on the engine catalog; unknown
items fall back to the script text so a new catalog entry still renders.
"""
from __future__ import annotations

import re

from .util import ascii_clean

# Customer timing groups, in the order customers should work them.
GROUP = {"BlocksPilot": "Fix before rollout", "ConfirmBeforePilot": "Check before rollout",
         "TenantHardening": "Improve security", "EvidenceRequest": "Information needed",
         "LongerTermImprovement": "Plan for later", "InventoryHygiene": "Plan for later",
         "NoAction": "No action needed"}
GROUP_ORDER = ["Fix before rollout", "Check before rollout", "Improve security", "Information needed", "Plan for later",
               "No action needed"]
GROUP_MEANING = {
    "Fix before rollout": "A confirmed gap that affects every Copilot user. Fix these first.",
    "Check before rollout": "Could affect Copilot users. Check or fix before Copilot is rolled out further.",
    "Improve security": "Makes the whole tenant safer. Does not by itself stop a rollout.",
    "Information needed": "We could not judge this from the data collected. Send us the information listed.",
    "Plan for later": "Worth doing, but not needed before rollout.",
    "No action needed": "Reviewed; nothing to do.",
}
# Light fills for the timing groups (Word and Excel).
GROUP_FILL = {"Fix before rollout": "F4C7C3", "Check before rollout": "FCE8B2", "Improve security": "D2E3FC",
              "Information needed": "E6E6E6", "Plan for later": "D9EAD3", "No action needed": "FFFFFF"}

# Customer areas in dependency order; topics are listed most important first.
AREAS = [
    ("Licensing and current use", ("SCOPE", "LICENSING", "ADOPTION")),
    ("Sign-in and MFA", ("CA", "MFA", "RISK", "LEGACY", "ACCESS_REVIEW")),
    ("Administrator accounts", ("PRIV",)),
    ("Devices", ("DEVICE", "MDE")),
    ("Files and sharing", ("PERMISSIONS", "SHARING", "OWNERSHIP", "GUEST")),
    ("Data protection", ("LABELS", "DLP", "ENCRYPTION", "COPILOT_DLP")),
    ("Audit and records", ("AUDIT", "AI_INVESTIGATION", "RETENTION", "LOCKBOX")),
    ("Apps and connectors", ("CONSENT", "CONNECTORS")),
    ("Security operations", ("INCIDENTS", "EXCEPTIONS")),
    ("AI tools and agents", ("AGENTS", "AI_DATA_RISK", "SHADOW_AI", "AGENT_LIFECYCLE")),
]
AREA_OF = {topic: area for area, topics in AREAS for topic in topics}
AREA_NAMES = [area for area, _ in AREAS] + ["Other"]
TOPIC_RANK = {topic: n for n, topic in enumerate(t for _, topics in AREAS for t in topics)}
PRIORITY_RANK = {"High": 0, "Medium": 1, "Low": 2}


def order_key(item):
    """Customer work order: timing group, priority, dependency step, then topic importance."""
    return (GROUP_ORDER.index(item["Group"]), PRIORITY_RANK.get(item["ReviewPriority"], 3), item["DependencyStep"],
            TOPIC_RANK.get(item["Topic"], 99), item["FindingId"])

DECISION_LABEL = {
    "Confirmed": "Confirmed", "ConfirmedWithQualification": "Confirmed with limits",
    "CorrectFinding": "Finding corrected", "CorrectRecommendation": "Advice corrected",
    "DowngradeToEvidenceGap": "Missing data, not a failure", "InsufficientEvidence": "Not enough evidence",
    "AddMissingFinding": "Added by review", "DuplicateOrOverlapping": "Duplicate",
    "CustomerDecisionRequired": "Customer decision", "NotApplicable": "Not applicable",
}
EVIDENCE_LABEL = {"Sufficient": "Good", "SufficientWithQualification": "Good, with limits", "Limited": "Limited",
                  "Stale": "Out of date", "Missing": "Missing", "Contradictory": "Sources disagree"}
STATEMENT_LABEL = {
    "Collector fact": "Collected data", "Script-derived conclusion": "Original report",
    "Microsoft-documented guidance": "Microsoft guidance", "Reviewer interpretation": "Reviewer view",
    "Suggested hardening practice": "Suggested practice", "Customer policy decision": "Your decision",
    "Evidence gap": "Missing evidence",
}

WHY = {
    "SCOPE": "People already using Copilot rely on the protections in this review today.",
    "LICENSING": "People need a Copilot license and supported Microsoft 365 apps before they can take part.",
    "ADOPTION": "Adoption planning is handled outside this security review.",
    "MFA": "Copilot can reach everything a user can. Without MFA, one stolen password gives an attacker that same "
           "reach.",
    "CA": "Without MFA required at sign-in, one stolen password can expose everything that person and Copilot "
          "can see.",
    "LEGACY": "Older sign-in methods skip MFA, so they are a common way into accounts.",
    "RISK": "A compromised account using Copilot could quickly find and collect sensitive information.",
    "ACCESS_REVIEW": "Regular reviews remove access that people no longer need.",
    "PRIV": "Administrator accounts can change any setting. One compromised admin account puts the whole tenant "
            "at risk.",
    "DEVICE": "Copilot shows company data on whatever device the user is on. Unmanaged devices are harder to "
              "protect.",
    "MDE": "Devices that do not report to Defender cannot be protected or investigated.",
    "SHARING": "Sharing settings decide who outside a site can open its files.",
    "GUEST": "Guests can see whatever they are given access to, including through Copilot if they are licensed.",
    "PERMISSIONS": "Copilot shows people what they already have access to. Files open to everyone in the company "
                   "become easy to find through Copilot.",
    "OWNERSHIP": "Owners decide who should see a site. Sites without owners tend to keep old, broad access.",
    "LABELS": "Sensitivity labels mark sensitive files, and Copilot honors their protection.",
    "ENCRYPTION": "Encryption keeps sensitive files protected even if they are shared with the wrong people.",
    "DLP": "Rules that are never tested can miss sensitive data or block normal work.",
    "COPILOT_DLP": "These rules can stop Copilot from using sensitive files or prompts.",
    "AUDIT": "If something goes wrong, you need records of who asked Copilot what.",
    "RETENTION": "Retention decides how long content and Copilot conversations are kept or deleted.",
    "LOCKBOX": "Customer Lockbox matters only if contracts or regulations require you to approve Microsoft support "
               "access.",
    "AI_INVESTIGATION": "If Copilot is misused, someone must be allowed and able to look at the conversations.",
    "CONSENT": "Apps that users approve can read company data outside Copilot's protections.",
    "CONNECTORS": "Connected outside sources add content that Copilot can search.",
    "INCIDENTS": "An unworked incident may mean an account or device is already compromised.",
    "EXCEPTIONS": "Without an owner and an exceptions list, security problems and accepted risks go untracked.",
    "AGENTS": "You can only control the AI agents you know about.",
    "SHADOW_AI": "People may paste company data into AI tools you have not approved.",
    "AI_DATA_RISK": "It shows where sensitive files are overshared before Copilot users can find them.",
    "AGENT_LIFECYCLE": "Agents and connectors without an owner keep their access long after anyone needs it.",
}

DONE = {
    "SCOPE": "A recorded decision on current Copilot use.",
    "LICENSING": "A license and app report for the users assigned Copilot.",
    "ADOPTION": "The business sponsor has approved the use cases and success measures.",
    "MFA": "A current registration report shows every active user registered.",
    "CA": "The MFA policy shows On for all users, with only emergency access accounts excluded.",
    "LEGACY": "The setting is off, and no legacy sign-ins appear outside approved exceptions.",
    "RISK": "No flagged account is left uninvestigated.",
    "ACCESS_REVIEW": "Reviews are set up, or the decision not to use them is recorded.",
    "PRIV": "An updated administrator list with a reason for each assignment.",
    "DEVICE": "Devices used with Copilot are managed and compliant, or the alternative control is recorded.",
    "MDE": "Devices used with Copilot show as onboarded and active in Defender.",
    "SHARING": "The sharing policy decision is recorded, and the settings match it.",
    "GUEST": "The guest invitation decision is recorded, and the setting matches it.",
    "PERMISSIONS": "A current report shows broad access on the listed sites removed, restricted or approved.",
    "OWNERSHIP": "Every active site has an owner, and inactive sites are archived or deleted.",
    "LABELS": "Labels are published to Copilot users, and the labeling approach is agreed.",
    "ENCRYPTION": "The labels that encrypt are known and applied where needed.",
    "DLP": "Test or simulation results are saved for the key rules.",
    "COPILOT_DLP": "Test results show the Copilot rules work.",
    "AUDIT": "An audit search returns Copilot activity.",
    "RETENTION": "Retention policies match the recorded decision, including for Copilot.",
    "LOCKBOX": "The Customer Lockbox decision is recorded, and the setting matches it.",
    "AI_INVESTIGATION": "The investigation decision is recorded, and a test search works.",
    "CONSENT": "The consent decision is recorded, and the settings match it.",
    "CONNECTORS": "Each connector has an owner and an approved scope.",
    "INCIDENTS": "Every open incident has an owner, and high-severity incidents are resolved.",
    "EXCEPTIONS": "An incident owner is named, and an exceptions register exists.",
    "AGENTS": "An agent list with approval status for each.",
    "SHADOW_AI": "AI app discovery is on, and its coverage is recorded.",
    "AI_DATA_RISK": "A completed data risk assessment covers the main SharePoint sites.",
    "AGENT_LIFECYCLE": "Each agent and connector has an owner and a review date.",
}

STEP_HEADLINE = {
    "LICENSING": "Confirm Copilot licensing and prerequisites",
    "CA": "Confirm MFA is required at sign-in",
    "DEVICE": "Confirm how devices are managed",
    "PERMISSIONS": "Check who can see SharePoint and OneDrive content",
    "LABELS": "Confirm sensitivity labels are in place",
    "AUDIT": "Confirm audit logging works",
    "CONNECTORS": "Review connected apps and data sources",
    "EXCEPTIONS": "Set up incident ownership and an exceptions list",
    "COPILOT_DLP": "Set up data loss prevention for Copilot",
    "AGENTS": "List and approve the AI agents in use",
    "SHADOW_AI": "Turn on discovery of unapproved AI tools",
    "AI_DATA_RISK": "Run a Purview data risk assessment for AI",
    "AI_INVESTIGATION": "Decide who can investigate Copilot conversations",
    "AGENT_LIFECYCLE": "Set owners and review dates for agents and connectors",
}


def _num(value, default="an unknown number of"):
    return default if value is None else str(value)


def _plural(count, word, plural=None):
    if count is None:
        return f"an unknown number of {plural or word + 's'}"
    return f"{count} {word if count == 1 else (plural or word + 's')}"


def _first_number(text_value):
    match = re.search(r"(?<![\w.-])(\d{1,6})(?![\w.])", text_value or "")
    return int(match.group(1)) if match else None


def _sentences(text_value, limit=2):
    """First sentences of script text, or '' when it carries redaction markers or tool jargon."""
    if not text_value or "[" in text_value or re.search(r"(?i)_[a-z]+_|raw_data|\w+\.\w+=", text_value):
        return ""
    parts = re.split(r"(?<=[.!?])\s+", text_value.strip())
    return " ".join(parts[:limit]).replace("(s)", "s").strip()


# ------------------------------------------------------------------ item templates

def _scope(ctx, item, obs):
    facts = ctx.facts
    if item["FindingId"] == "RVW-EXISTING-COPILOT-USE":
        seats = facts.get("copilot_licenses") or {}
        active = (facts.get("copilot_usage") or {}).get("Paid Copilot active users")
        issue = f"{_plural(seats.get('Assigned'), 'person', 'people')} already have a Copilot license"
        issue += f", and {active} used it recently." if active is not None else "."
        return dict(Headline="Decide whether current Copilot users keep access",
                    Issue=issue + " The gaps in this review apply to them today.",
                    Why="Current users already rely on the protections reviewed here, so every gap applies to them now.",
                    HowToFix="Decide whether current users continue while gaps are fixed, and confirm they are "
                             "covered by MFA and the file-sharing review.",
                    DoneWhen="A recorded decision for current Copilot users.")
    return None


def _licensing(ctx, item, obs, key):
    seats = ctx.facts.get("copilot_licenses") or {}
    if "license.apps" in key:
        return dict(Headline="Confirm Copilot users have supported Microsoft 365 apps",
                    Issue="The assessment cannot see which Microsoft 365 app versions people use.",
                    Why="Copilot in Word, Excel, Outlook and Teams needs supported Microsoft 365 Apps. Copilot Chat "
                        "on the web does not.",
                    HowToFix="Check that people assigned Copilot run Microsoft 365 Apps on a supported update channel.",
                    DoneWhen="An app version report for Copilot users.")
    if "license.assignment" in key:
        if seats.get("Enabled"):
            issue = (f"The original report said license data was missing, but {seats['Enabled']} Copilot licenses "
                     f"were found ({seats.get('Assigned', 0)} assigned). Who holds them was not checked against an intended "
                     "list.")
        else:
            issue = "Copilot license information was not collected."
        return dict(Headline="Confirm Copilot licenses go to the intended users", Issue=issue,
                    HowToFix="Review who holds the licenses and reassign any that are unused or held by the wrong "
                             "people.",
                    DoneWhen="A license report reviewed by the license owner.",
                    Change="The original report said license data was missing; it was collected.")
    if seats:
        free = (seats.get("Enabled") or 0) - (seats.get("Assigned") or 0)
        return dict(Headline="Copilot licenses are in place",
                    Issue=f"{_plural(seats.get('Enabled'), 'Copilot license')} exist: {seats.get('Assigned')} "
                          f"assigned and {free} free.")
    return None


def _mfa(ctx, item, obs, key):
    reg = ctx.facts.get("mfa_registration")
    if item["FindingId"] == "ENT-007" or "passwordless" in key:
        return dict(Headline="Plan phishing-resistant sign-in for administrators",
                    Issue=_sentences(obs, 1) or "Few accounts use passwordless or phishing-resistant sign-in.",
                    Why="Phishing-resistant methods stop most password-phishing attacks.",
                    HowToFix="Start with administrators, then expand to high-risk users once MFA is required for "
                             "everyone.",
                    DoneWhen="A plan, and a registration report for administrators.")
    if reg:
        pct = round(100 * reg["Registered"] / reg["Returned"]) if reg["Returned"] else 0
        issue = (f"Only {reg['Registered']} of {reg['Returned']} accounts ({pct}%) have set up MFA. "
                 f"{reg['AdminsNotRegistered']} of {reg['AdminsReturned']} administrator accounts have not.")
    else:
        issue = _sentences(obs) or "MFA registration could not be measured."
    return dict(Headline="Get every user set up for MFA", Issue=issue,
                Change="The original report called this a blocker. Registration is the first step toward "
                       "requiring MFA; the blocker is the MFA policy item." if item["ReadinessTreatment"]
                != "BlocksPilot" else "",
                HowToFix="Ask unregistered users to set up MFA, starting with administrators and current Copilot users. "
                         "Record anyone who is exempt and why. Setting up MFA is separate from requiring it (see "
                         "the MFA policy item).")


def _ca(ctx, item, obs, key):
    ca = ctx.facts.get("ca")
    if not ca:
        return None
    if item["ReadinessTreatment"] == "NoAction":
        return dict(Headline="MFA is required for all users",
                    Issue="An active policy requires MFA for every user.")
    parts = ["No active policy requires MFA for every user."]
    if ca["ReportOnlyAllUsersMfa"]:
        parts.append(f"{_plural(ca['ReportOnlyAllUsersMfa'], 'all-user MFA policy', 'all-user MFA policies')} "
                     "already exist but run only in test (report-only) mode.")
    if ca["EnforcedTargetedMfa"]:
        parts.append(f"{_plural(ca['EnforcedTargetedMfa'], 'active policy covers', 'active policies cover')} "
                     "specific groups only.")
    if ca["EnforcedLegacyBlockAllUsers"]:
        parts.append("Legacy sign-in is already blocked.")
    if ca["ReportOnlyAllUsersMfa"]:
        return dict(Headline="Turn on the existing MFA policy for all users", Issue=" ".join(parts),
                    HowToFix="Review the test results of the existing all-user MFA policy, check that its "
                             "exclusions are only emergency or sync accounts, then switch it On and remove "
                             "duplicates.",
                    Change="The original advice was to create a new MFA policy. Two already exist in test mode, so "
                           "the fix is to switch one on.")
    return dict(Headline="Require MFA for all users", Issue=" ".join(parts),
                HowToFix="Create a Conditional Access policy that requires MFA for all users, test it in "
                         "report-only mode, then switch it On. Exclude only emergency access accounts.")


def _legacy(ctx, item, obs, key):
    ca = ctx.facts.get("ca") or {}
    sign = ctx.facts.get("signins") or {}
    issue = "SharePoint still allows older sign-in methods that cannot use MFA."
    if ca.get("EnforcedLegacyBlockAllUsers"):
        issue += " A tenant-wide policy already blocks them"
        issue += (f", and no legacy sign-ins appeared in {sign['Sampled']} sampled sign-ins."
                  if sign.get("LegacyAttempts") == 0 and sign.get("Sampled") else ".")
    return dict(Headline="Turn off legacy sign-in for SharePoint", Issue=issue,
                Why="Turning the setting off removes a fallback route if the blocking policy is ever changed.",
                HowToFix="Confirm no business app needs legacy sign-in, then turn the SharePoint setting off.")


def _priv(ctx, item, obs, key):
    ca = ctx.facts.get("ca") or {}
    count = _first_number(obs)
    fid = item["FindingId"]
    if fid == "RVW-GLOBAL-ADMIN-COUNT":
        ga = ctx.facts.get("global_admins") or {}
        return dict(Headline="Reduce the number of Global Administrators",
                    Issue=f"{_plural(ga.get('DistinctPrincipals'), 'account')} are Global Administrators. "
                          "Microsoft recommends fewer than five.",
                    Why="Global Administrator is the most powerful role. Every extra holder is another way in.",
                    HowToFix="Keep two emergency access accounts plus the minimum you need, and move everyone else "
                             "to narrower admin roles.",
                    DoneWhen="Fewer than five Global Administrators, with a reason recorded for each.")
    if fid == "RVW-EMERGENCY-ACCESS":
        counts = [str(n) for n in ca.get("AllUsersMfaExclusionCounts") or []]
        if len(counts) > 1:
            exclusions = (f" The all-user MFA policies exclude {', '.join(counts[:-1])} and {counts[-1]} accounts "
                          "respectively.")
        elif counts:
            exclusions = f" The all-user MFA policy excludes {counts[0]} accounts."
        else:
            exclusions = ""
        return dict(Headline="Name your emergency access accounts",
                    Issue="We could not tell which accounts are the emergency (break-glass) accounts." + exclusions,
                    Why="Emergency accounts let you get back in if sign-in policies lock everyone out. Any other "
                        "exclusion is a gap.",
                    HowToFix="Name two cloud-only emergency accounts, make them the only policy exclusions, and set "
                             "up alerts when they are used.",
                    DoneWhen="The emergency accounts are recorded and are the only exclusions.")
    if fid == "RVW-ADMIN-MFA":
        return dict(Headline="Require MFA for administrator roles",
                    Issue="No active policy requires MFA for administrator roles.",
                    HowToFix="Switch on a Conditional Access policy that requires MFA for all admin roles.")
    if "WithoutMFARegistration" in key:
        issue = f"{_plural(count, 'administrator account')} have not set up MFA"
        issue += (", and no active policy requires MFA for admin roles." if not ca.get("EnforcedAdminMfa")
                  and not ca.get("EnforcedAllUsersMfa") else ".")
        existing = (" An admin MFA policy already exists but is not switched on; review it and turn it on."
                    if ca.get("ReportOnlyOrDisabledAdminMfa") else " Then require MFA for admin roles with a "
                    "Conditional Access policy.")
        return dict(Headline="Protect administrator accounts with MFA", Issue=issue,
                    HowToFix="Set up MFA on each admin account, or remove the admin role." + existing,
                    DoneWhen="Every admin account has MFA, and an admin MFA policy is On.")
    if "PermanentPrivilege" in key or fid == "ENT-011":
        return dict(Headline="Review permanent administrator access",
                    Issue=f"{_plural(count, 'account')} hold administrator rights permanently, with no business "
                          "reason recorded.",
                    Why="Standing admin rights give attackers more targets. Time-limited access reduces that.",
                    HowToFix="For each account, confirm the business need. Remove, reduce or time-limit the access "
                             "(time-limited access through Privileged Identity Management needs Entra ID P2).")
    if "DisabledIdentity" in key:
        return dict(Headline="Remove admin rights from disabled accounts",
                    Issue=f"{_plural(count, 'disabled account')} still hold administrator roles.",
                    Why="Admin rights on unused accounts are easy to miss if an account is turned back on.",
                    HowToFix="Remove the admin roles from the disabled accounts.",
                    DoneWhen="No disabled account holds an admin role.")
    if "ServicePurpose" in key:
        return dict(Headline="Record the purpose of service accounts with admin rights",
                    Issue=f"{_plural(count, 'admin account')} look like service or sync accounts, with no purpose "
                          "recorded.",
                    Why="Knowing what each account is for avoids breaking integrations during clean-up.",
                    HowToFix="Record the owner and purpose of each service account.",
                    DoneWhen="Each service account has an owner and a purpose recorded.")
    if "ActivityUnknown" in key:
        return dict(Headline="Check which admin accounts are actually used",
                    Issue=f"Sign-in activity was not available for {_plural(count, 'admin account')}, so unused "
                          "admin rights cannot be spotted.",
                    Why="Unused admin rights are a risk that can simply be removed.",
                    HowToFix="Pull admin sign-in activity and use it to remove rights nobody uses.",
                    DoneWhen="An activity report for admin accounts, with unused rights removed.")
    return None


def _risk(ctx, item, obs, key):
    risk = ctx.facts.get("risk") or {}
    if risk.get("AtRisk") is None:
        return None
    issue = (f"{_plural(risk['AtRisk'], 'account is', 'accounts are')} flagged by Microsoft as possibly "
             f"compromised ({risk.get('High') or 0} high, {risk.get('Medium') or 0} medium, {risk.get('Low') or 0} "
             "low risk).")
    if not risk.get("UserRiskPolicy"):
        issue += " No user-risk policy is set up."
    stated = _first_number(obs)
    change = (f"The original report counted {stated} risky accounts; it left out low-risk accounts."
              if stated is not None and stated != risk["AtRisk"] else "")
    return dict(Headline="Investigate accounts flagged as risky", Issue=issue, Change=change,
                HowToFix="Investigate each flagged account, then reset its password or dismiss the flag. Consider a "
                         "user-risk policy that does this automatically.")


def _device(ctx, item, obs, key):
    intune = ctx.facts.get("intune") or {}
    mde = ctx.facts.get("defender_devices") or {}
    issue = (f"Intune shows {_num(intune.get('ManagedDevices'), 'no')} managed devices, although "
             f"{_plural(intune.get('CompliancePolicies'), 'compliance policy', 'compliance policies')} exist.")
    if mde.get("ReportedManagedByIntune") and intune.get("ManagedDevices") == 0:
        issue += " Defender reports a device as Intune-managed, so the two sources disagree."
    return dict(Headline="Decide how devices used with Copilot are controlled", Issue=issue,
                HowToFix="Confirm how you manage devices (Intune, another tool, or browser-only controls) and check "
                         "that devices used with Copilot are covered.")


def _mde(ctx, item, obs, key):
    mde = ctx.facts.get("defender_devices") or {}
    recon = ctx.facts.get("device_reconciliation") or {}
    issue = (f"Defender for Endpoint lists {_plural(mde.get('Records'), 'device')}, and {mde.get('Active', 0)} "
             "are reporting.")
    if mde.get("NewestLastSeenAgeDays") is not None:
        issue += f" The most recent report was {mde['NewestLastSeenAgeDays']} days ago."
    if recon.get("known_active"):
        issue += f" About {recon['known_active']} devices appear to be in use."
    return dict(Headline="Get devices reporting to Defender", Issue=issue,
                HowToFix="Onboard the devices people use with Copilot, and confirm they show as active in Defender.")


def _sharing(ctx, item, obs, key):
    return dict(Headline="Decide your policy on Anyone sharing links",
                Issue=_sentences(obs, 3) or "Anyone links (no sign-in needed) are allowed.",
                Why="Anyone links let files be opened without signing in. Whether that is acceptable for business "
                    "content is your policy choice.",
                HowToFix="Decide whether Anyone links are needed. If they are, consider making Specific people the "
                         "default link type, keep a short expiry, and review the sites that allow them.")


def _guest(ctx, item, obs, key):
    guests = ctx.facts.get("guests") or {}
    setting = "Every user" if str(guests.get("InviteSetting")).lower() == "everyone" else "Users"
    return dict(Headline="Decide who may invite guests",
                Issue=f"{setting} can invite outside guests (the Microsoft default). "
                      f"{_plural(guests.get('Total'), 'guest account')} exist.",
                HowToFix="Decide whether to limit invitations to certain roles or to require approval.")


def _permissions(ctx, item, obs, key):
    seats = ctx.facts.get("copilot_licenses") or {}
    if "broad_internal" in key:
        everyone = re.search(r"(\d+) permissions? granted to Everyone;", obs or "")
        eeeu = re.search(r"(\d+) permissions? granted to Everyone except external users", obs or "")
        links = re.search(r"(\d+) links? accessible to people in the organization", obs or "")
        sites = re.search(r"(\d+) SharePoint sites?", obs or "")
        bits = []
        if everyone:
            bits.append(f"{everyone.group(1)} permissions granted to Everyone")
        if eeeu:
            bits.append(f"{eeeu.group(1)} granted to Everyone except external users")
        if links:
            bits.append(f"{links.group(1)} company-wide sharing links")
        issue = ("The last oversharing report found " + ", ".join(bits)
                 + (f" across about {sites.group(1)} sites" if sites else "") + ".") if bits else \
            "The last oversharing report found content open to very large groups."
        age = item["Evidence"].get("AgeDays")
        if age:
            issue += f" That report is {age} days old."
        return dict(Headline="Remove company-wide access from SharePoint sites", Issue=issue,
                    HowToFix="Refresh the report, then have site owners remove broad access that is not needed. "
                             "Where a site must stay open, keep it out of Copilot with Restricted Content "
                             "Discovery while it is reviewed.")
    licensed = " Your Copilot licenses include this feature." if seats.get("Enabled") else ""
    return dict(Headline="Run current SharePoint oversharing reports",
                Issue="Current SharePoint oversharing reports are missing or out of date, so the real level of "
                      "oversharing is unknown.",
                HowToFix="In the SharePoint admin center, run the Data access governance reports (site "
                         "permissions, Everyone, and sharing links)." + licensed,
                DoneWhen="Reports dated within the last 35 days.")


def _ownership(ctx, item, obs, key):
    inventory = ctx.facts.get("site_inventory") or {}
    if inventory.get("Count"):
        return dict(Headline="Give every site an owner and clean up inactive sites",
                    Issue=f"The original report said site ownership was not collected, but the site inventory lists "
                          f"owners and activity for {inventory['Count']} sites. Some have no named owner, a disabled "
                          "owner or no recent activity.",
                    HowToFix="Assign at least two owners to each active site. For inactive sites, confirm with the "
                             "owner, then archive (Microsoft 365 Archive) or delete them. SharePoint inactive-site and "
                             "site-ownership policies can keep this up to date.",
                    Change="The original report said site ownership was not collected; it was.")
    return dict(Headline="Confirm every site has an owner",
                Issue="Site ownership and activity were not collected.",
                HowToFix="In the SharePoint admin center, open Active sites, find sites without owners or recent "
                         "activity, and assign an owner or archive the site.")


def _labels(ctx, item, obs, key):
    labels = ctx.facts.get("labels") or {}
    policies = ctx.facts.get("label_policies")
    found = (f"{_plural(labels.get('Count'), 'sensitivity label')} and "
             f"{_plural(policies, 'publishing policy', 'publishing policies')} exist.")
    if "coverage.data.labels" in key:
        if labels.get("Count"):
            return dict(Headline="Confirm sensitivity labels reach Copilot users",
                        Issue=f"The original report said labels were not collected, but {found[0].lower()}{found[1:]} "
                              "How widely they are used on business content is not known.",
                        HowToFix="Confirm which labels users get and how widely they are applied.",
                        Change="The original report said labels were not collected; they were.")
        return dict(Headline="Confirm sensitivity labels are in place",
                    Issue="Sensitivity label information was not collected.",
                    HowToFix="Confirm which labels exist and that they are published to users.")
    if item["ReadinessTreatment"] == "NoAction":
        return dict(Headline="Sensitivity labels are published", Issue=found)
    return dict(Headline="Agree your labeling approach",
                Issue=found + " Whether to use default, mandatory or automatic labeling is not yet decided.",
                HowToFix="Agree the label list, who receives the labels, and whether to use default, mandatory or "
                         "automatic labeling.",
                Change="The original report presented specific labeling rules for Copilot content as required. "
                       "They are a policy choice for you.")


def _encryption(ctx, item, obs, key):
    if item["ReadinessTreatment"] == "NoAction":
        return dict(Headline="Encryption is available", Issue="Rights management (encryption) is turned on.")
    return dict(Headline="Confirm which labels encrypt sensitive files",
                Issue="Rights management (encryption) is turned on. It protects files only where a label or "
                      "permission applies it.",
                HowToFix="Identify which labels encrypt, and confirm they are used on sensitive content.",
                Change="The original report implied encryption protects Copilot content on its own. It protects "
                       "only files where it is applied.")


def _dlp(ctx, item, obs, key):
    if item["ReadinessTreatment"] == "NoAction":
        return dict(Headline="Data loss prevention rules are on", Issue=_sentences(obs, 1) or "DLP rules are on.")
    return dict(Headline="Test your data loss prevention rules",
                Issue="DLP policies exist, but there is no evidence they were tested to work as intended.",
                HowToFix="Run the key policies in simulation mode or a controlled test, and keep the results.")


def _copilot_dlp(ctx, item, obs, key):
    enforced = re.search(r"(\d+) of (\d+) DLP policies are configured in enforcement mode", obs or "")
    copilot = re.search(r"(\d+) enforced polic(?:y|ies) appl(?:y|ies) to Microsoft 365 Copilot", obs or "")
    issue = (f"{enforced.group(1)} of {enforced.group(2)} DLP policies are enforced" if enforced
             else "Data loss prevention rules are configured")
    issue += f", and {copilot.group(1)} of them apply to Copilot." if copilot else "."
    if item["ReadinessTreatment"] == "NoAction":
        return dict(Headline="Data loss prevention covers Copilot", Issue=issue)
    return dict(Headline="Test the Copilot data loss prevention rules", Issue=issue,
                HowToFix="Test that the Copilot rules block the intended prompts or labeled files.")


def _audit(ctx, item, obs, key):
    audit = ctx.facts.get("audit") or {}
    on = "Audit logging is on" if audit.get("UnifiedAuditEnabled") else "Audit logging could not be confirmed"
    state = audit.get("CopilotInteractionSource")
    if item["Topic"] == "AI_INVESTIGATION":
        return dict(Headline="Decide who can investigate Copilot conversations",
                    Issue="No one has decided who may search Copilot prompts and responses, and retrieving them has "
                          "not been tested.",
                    HowToFix="Choose who may investigate, set how long conversations are kept, and run a test "
                             "search.")
    if "coverage.operation" in key:
        check = " did not finish" if state == "Partial" else " was not run" if state != "Complete" else " ran"
        return dict(Headline="Test that Copilot activity shows in the audit log",
                    Issue=f"{on}, but the check for Copilot activity records{check}, so it is not shown that the "
                          "records can be found.",
                    HowToFix="Run an audit search for Copilot activity (record type CopilotInteraction) for a few "
                             "Copilot users, and save the result.")
    if item["ReadinessTreatment"] == "NoAction":
        return dict(Headline="Audit logging is on", Issue=on + ".")
    return dict(Headline="Set audit log retention to your requirement",
                Issue=f"{on}. The original advice to keep logs for 10 years needs an add-on license, and it named "
                      "audit events that do not exist.",
                HowToFix="Decide how long to keep audit records within your licensing (standard audit keeps them "
                         "180 days), and confirm Copilot activity appears in a search.",
                Change="The original advice needed an add-on license it did not mention and cited audit event names "
                       "that Microsoft does not document.")


def _retention(ctx, item, obs, key):
    if item["ReadinessTreatment"] == "NoAction":
        return dict(Headline="Retention policies are in place", Issue="Retention policies are configured.")
    return dict(Headline="Decide how long to keep Copilot conversations",
                Issue="Retention policies exist. The original advice linked retention to Copilot training, but "
                      "Microsoft does not train its AI models on your data.",
                HowToFix="Set retention as a records decision, including a policy for Copilot and AI app "
                         "conversations.",
                Change="The original advice tied retention to Copilot training. Microsoft documents that your data "
                       "is not used to train its models.")


def _lockbox(ctx, item, obs, key):
    return dict(Headline="Decide whether to use Customer Lockbox",
                Issue="Customer Lockbox is off. When on, Microsoft support must ask your approval before accessing "
                      "your content.",
                HowToFix="Decide whether your contracts or regulations require it, and turn it on if they do (it is "
                         "included with E5).")


def _consent(ctx, item, obs, key):
    consent = ctx.facts.get("consent") or {}
    if "high_impact_grants" in key:
        extra = ctx.facts.get("app_permission_apps")
        return dict(Headline="Review apps with high-level access",
                    Issue=f"{_plural(consent.get('HighPrivilegeApps'), 'app')} have high-privilege permissions, and "
                          f"{_num(consent.get('UnverifiedPublishers'))} come from unverified publishers."
                          + (f" The review also found {_plural(extra, 'app')} with tenant-wide application "
                             "permissions (for example, reading all mailboxes or files) that the original count did "
                             "not include." if extra else ""),
                    HowToFix="Check each flagged app's owner, purpose and recent use. Remove unused apps and reduce "
                             "permissions that go beyond what the app needs.",
                    DoneWhen="Each flagged app is kept with a reason, reduced or removed.")
    if consent.get("AllMicrosoftManaged"):
        return dict(Headline="Decide whether to tighten app consent",
                    Issue="Users can approve some apps themselves under Microsoft's recommended default setting, "
                          "which excludes high-risk access such as files, mail and chat.",
                    HowToFix="If your policy needs it, allow user consent only for verified publishers and turn on "
                             "the admin consent request workflow.",
                    Change="The original report treated this as a blocker. The settings are Microsoft's "
                           "recommended defaults, so it is a security improvement, not a blocker.")
    if consent.get("LegacyAssigned"):
        return dict(Headline="Restrict which apps users can approve",
                    Issue="Users can grant any app access to company data without admin review.",
                    HowToFix="Allow user consent only for verified publishers, and turn on the admin consent "
                             "request workflow.")
    return None


def _connectors(ctx, item, obs, key):
    if item["ReadinessTreatment"] == "NoAction":
        return dict(Headline="No outside data sources are connected",
                    Issue="No Graph connectors are connected to Microsoft 365 search and Copilot.")
    return None


def _incidents(ctx, item, obs, key):
    inc = ctx.facts.get("incidents")
    if item["Topic"] == "EXCEPTIONS" and inc:
        return dict(Headline="Set up incident ownership and an exceptions list",
                    Issue=f"{inc['ActiveNoUpdateSinceCreation']} of {inc['Active']} open incidents have not been "
                          "touched since they were created, and there is no list of accepted risks.",
                    HowToFix="Name an owner for the incident queue, set a review routine, and keep a list of "
                             "accepted exceptions, each with an owner and an expiry date.")
    if not inc:
        return None
    issue = (f"{_plural(inc['Active'], 'security incident is', 'security incidents are')} open "
             f"({inc['ActiveHigh']} high severity).")
    if inc["ActiveUnassigned"] == inc["Active"] and inc["Active"]:
        issue += " None has an owner"
    else:
        issue += f" {inc['ActiveUnassigned']} have no owner"
    issue += (f", and the oldest is {inc['OldestActiveAgeDays']} days old." if inc.get("OldestActiveAgeDays")
              is not None else ".")
    return dict(Headline="Work the open security incidents", Issue=issue,
                HowToFix="Assign an owner, handle the high-severity incident first, then close or classify the "
                         "rest.",
                Change="The original report called this a blocker. Open incidents show a handling gap; they "
                       "become a blocker only if triage finds an active compromise.")


def _agents(ctx, item, obs, key):
    pp = ctx.facts.get("power_platform") or {}
    issue = ("Copilot Studio is licensed, but the list of agents and Power Platform apps could not be collected."
             if pp.get("CopilotStudioLicensed") else "The list of AI agents in use was not collected.")
    return dict(Headline="List and approve the AI agents in use", Issue=issue,
                HowToFix="Export the agent list from the Microsoft 365 admin center and Power Platform, and record "
                         "which agents are approved.")


def _shadow(ctx, item, obs, key):
    mde = ctx.facts.get("defender_devices") or {}
    if item["FindingId"] == "RVW-SHADOW-AI-VISIBILITY":
        return dict(Headline="Turn on discovery of unapproved AI tools",
                    Issue="The check for other AI tools found nothing, but that does not prove none are used.",
                    HowToFix="After devices are onboarded, turn on AI app discovery in Purview (it needs the Purview "
                             "browser extension).")
    count = _first_number(obs)
    if not mde.get("Active"):
        return dict(Headline="Onboard devices so endpoint data loss prevention works",
                    Issue=f"{_plural(count, 'endpoint DLP policy', 'endpoint DLP policies')} exist, but no active "
                          "device is onboarded, so they protect nothing yet.",
                    Why="Endpoint DLP stops sensitive data being pasted or uploaded into outside AI sites, but only "
                        "on onboarded devices.",
                    HowToFix="Onboard devices first (see the device items), then confirm the policies apply.",
                    DoneWhen="Devices are onboarded and active, and the endpoint DLP policies apply to them.",
                    Change="The original report marked endpoint DLP as working. No active onboarded device was "
                           "found, so that is not shown.")
    return None


def _ai_data_risk(ctx, item, obs, key):
    return dict(Headline="Run a Purview data risk assessment for AI",
                Issue="No Purview data risk assessment results for AI were available.",
                HowToFix="Review the default weekly assessment in Purview (DSPM for AI). It adds to the SharePoint "
                         "oversharing review; it does not replace it.")


def _adoption(ctx, item, obs, key):
    return dict(Headline="Copilot use cases and success measures",
                Issue="Adoption planning is outside the scope of this security review.")


def _access_review(ctx, item, obs, key):
    return dict(Headline="Consider regular access reviews", Issue="No access reviews are set up.",
                HowToFix="Consider periodic reviews for admin roles and guests (needs Entra ID Governance).")


TEMPLATES = {
    "LICENSING": _licensing, "MFA": _mfa, "CA": _ca, "LEGACY": _legacy, "PRIV": _priv, "RISK": _risk,
    "DEVICE": _device, "MDE": _mde, "SHARING": _sharing, "GUEST": _guest, "PERMISSIONS": _permissions,
    "OWNERSHIP": _ownership, "LABELS": _labels, "ENCRYPTION": _encryption, "DLP": _dlp,
    "COPILOT_DLP": _copilot_dlp, "AUDIT": _audit, "AI_INVESTIGATION": _audit, "RETENTION": _retention,
    "LOCKBOX": _lockbox, "CONSENT": _consent, "CONNECTORS": _connectors, "INCIDENTS": _incidents,
    "EXCEPTIONS": _incidents, "AGENTS": _agents, "SHADOW_AI": _shadow, "AI_DATA_RISK": _ai_data_risk,
    "ADOPTION": _adoption, "ACCESS_REVIEW": _access_review,
}


def _rating_change(item, script_treatment):
    if script_treatment == "BlocksPilot" and item["ReadinessTreatment"] != "BlocksPilot":
        reason = ("the evidence supports checking it before rollout but does not show it stops Copilot on its own"
                  if item["ReadinessTreatment"] == "ConfirmBeforePilot" else "it does not affect Copilot by itself")
        return (f"The original report called this a blocker. We moved it to {GROUP[item['ReadinessTreatment']]} "
                f"because {reason}.")
    return ""


def describe(ctx, item, key="", script_treatment=None):
    """Attach item['Plain'] with customer-readable wording."""
    obs = (item.get("ScriptConclusion") or {}).get("Observation") or ""
    topic = item["Topic"]
    text_value = None
    if item["FindingId"].startswith("RVW-STEP-"):
        text_value = dict(Headline=STEP_HEADLINE.get(topic, item["Title"]),
                          Issue="The assessment did not cover this area.",
                          HowToFix="Send the information listed, or record a decision for this area.")
    elif topic == "SCOPE":
        text_value = _scope(ctx, item, obs)
    elif topic in TEMPLATES:
        text_value = TEMPLATES[topic](ctx, item, obs, key)
    text_value = dict(text_value or {})
    fallback_fix = re.split(r"(?<=[.!?])\s+", item["RecommendedAction"] or "")
    plain = {
        "Headline": text_value.get("Headline") or item["Title"],
        "Issue": text_value.get("Issue") or _sentences(obs) or item["CorrectedInterpretation"],
        "Why": text_value.get("Why") or WHY.get(topic, ""),
        "HowToFix": text_value.get("HowToFix") or " ".join(fallback_fix[:2]),
        "DoneWhen": text_value.get("DoneWhen") or DONE.get(topic, item["ClosureEvidence"].get("RequiredEvidence", "")),
        "Correction": text_value.get("Change") or "",
        "RatingChange": "",
    }
    if "blocker" not in plain["Correction"]:
        plain["RatingChange"] = _rating_change(item, script_treatment)
    if item["ReviewDecision"] == "DuplicateOrOverlapping":
        plain["Correction"], plain["RatingChange"] = f"Covered by {item['OverlapsWith']}.", ""
    plain["Change"] = " ".join(v for v in (plain["Correction"], plain["RatingChange"]) if v)
    if item["ReadinessTreatment"] == "NoAction":
        plain["HowToFix"], plain["DoneWhen"] = "No action needed.", "Not applicable."
    item["Plain"] = {k: ascii_clean(ctx.clean(v)) if v else "" for k, v in plain.items()}
    item["Group"] = GROUP[item["ReadinessTreatment"]]
    item["Area"] = AREA_OF.get(topic, "Other")
