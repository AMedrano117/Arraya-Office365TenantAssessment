"""
Exchange section collector — Graph API accessible data only.
Produces: EmailActivitySummary, EmailActivityTopSenders, EmailActivityTopReceivers,
          PrimaryMailboxStats, ArchiveMailboxStats, AllExchangeGroups, RoomMailboxes

Optional (called from runner.py post-parallel with user IDs from Identity section):
  collect_mailbox_types(client, user_ids) -> SharedMailboxes, EquipmentMailboxes
  collect_inbox_forwarding_rules(client, user_ids, accepted_domains) -> InboxRulesExternalForwarding

Note: Per-mailbox data (AllMailboxes, AllRecipients, MailFlowRules, etc.) requires
the Exchange Online PowerShell module and is not collected here.
"""

from __future__ import annotations

import logging
from typing import Any

from ..graph_client import GraphClient

log = logging.getLogger(__name__)

_PERIOD = "D30"


def collect(client: GraphClient) -> dict[str, Any]:
    log.info("Collecting Exchange section (Graph-available data only)...")
    data: dict[str, Any] = {}

    data["EmailActivitySummary"]      = _collect_email_activity_summary(client)
    sender_rows                        = _collect_email_activity_users(client)
    data["EmailActivityTopSenders"]   = _top_n(sender_rows, "Send Count", 20)
    data["EmailActivityTopReceivers"] = _top_n(sender_rows, "Receive Count", 20)

    mailbox_rows                       = _collect_mailbox_usage(client)
    data["PrimaryMailboxStats"]        = _build_primary_mailbox_stats(mailbox_rows)
    data["ArchiveMailboxStats"]        = _build_archive_stats(mailbox_rows)
    data["MailboxUsageDetails"]        = _build_mailbox_details(mailbox_rows)

    # Distribution groups — available via Graph advanced query (Group.Read.All)
    data["AllExchangeGroups"]         = _collect_distribution_groups(client)

    # Room mailboxes — available via Places API (Place.Read.All)
    data["RoomMailboxes"]             = _collect_room_mailboxes(client)

    # Keys that still require EXO PS module — mark as not collected
    for key in ("AllMailboxes", "AllRecipients", "NonUserMailboxes",
                "LitigationHoldMailboxes",
                "ArchiveMailboxes", "InactiveMailboxes", "MailFlowRules",
                "MailFlowConnectors", "InboxRulesExternalForwarding",
                "ForwardingPolicySummary", "RemoteDomains", "PublicFolderDetails",
                "PublicFolderPerms", "SMTPRelayConfig",
                "SMTPRelaySummary", "SpamFilteringConfig"):
        data.setdefault(key, None)

    # Per-user keys populated by runner.py post-parallel (MailboxSettings.Read)
    for key in ("SharedMailboxes", "EquipmentMailboxes", "MailboxPurposes"):
        data.setdefault(key, None)

    # Index stubs (normally built from AllMailboxes)
    for key in ("AllMailboxes-PrimarySmtpAddress", "AllMailboxes-MailIdentity",
                "AllMailboxes-UserPrincipalName"):
        data.setdefault(key, {})

    log.info("Exchange: core data collected; per-user enrichment deferred to runner")
    return data


# ---------------------------------------------------------------------------
# Raw collectors
# ---------------------------------------------------------------------------

def _collect_email_activity_summary(client: GraphClient) -> dict:
    try:
        rows = client.get_report_csv(f"reports/getEmailActivityCounts(period='{_PERIOD}')")
        if not rows:
            return {}
        latest = rows[-1] if rows else {}
        return {
            "ReportPeriod":  _PERIOD,
            "SendCount":     _int(latest.get("Send", "")),
            "ReceiveCount":  _int(latest.get("Receive", "")),
            "ReadCount":     _int(latest.get("Read", "")),
        }
    except Exception as exc:
        log.debug("Email activity summary unavailable: %s", exc)
        return {}


def _collect_email_activity_users(client: GraphClient) -> list[dict]:
    try:
        return client.get_report_csv(f"reports/getEmailActivityUserDetail(period='{_PERIOD}')")
    except Exception as exc:
        log.debug("Email activity user detail unavailable: %s", exc)
        return []


def _collect_mailbox_usage(client: GraphClient) -> list[dict]:
    try:
        return client.get_report_csv(f"reports/getMailboxUsageDetail(period='{_PERIOD}')")
    except Exception as exc:
        log.debug("Mailbox usage detail unavailable: %s", exc)
        return []


def _collect_room_mailboxes(client: GraphClient) -> list[dict]:
    """GET /places/microsoft.graph.room — requires Place.Read.All (optional)."""
    try:
        raw = client.get(
            "places/microsoft.graph.room",
            params={"$select": "id,displayName,emailAddress,building,floorNumber,capacity"},
        )
        return raw if isinstance(raw, list) else []
    except Exception as exc:
        log.debug("Room mailboxes unavailable (need Place.Read.All): %s", exc)
        return []


def _collect_distribution_groups(client: GraphClient) -> dict:
    select = ("id,displayName,mail,mailNickname,groupTypes,"
              "mailEnabled,securityEnabled,createdDateTime,description")
    filter_str = ("NOT groupTypes/any(c:c eq 'Unified')"
                  " and mailEnabled eq true and securityEnabled eq false")
    try:
        raw = client.get(
            f"groups?$filter={filter_str}&$select={select}&$count=true",
            extra_headers={"ConsistencyLevel": "eventual"},
        )
        return {g["id"]: g for g in (raw if isinstance(raw, list) else [])}
    except Exception as exc:
        log.debug("Distribution groups unavailable: %s", exc)
        return {}


# ---------------------------------------------------------------------------
# Public post-parallel enrichment (called from runner.py with Identity user IDs)
# ---------------------------------------------------------------------------

def collect_mailbox_types(client: GraphClient, user_ids: list[str]) -> dict[str, Any]:
    """
    Batch-fetch mailboxSettings.userPurpose for every user ID.
    Returns {"MailboxPurposes": {uid: purpose}, "SharedMailboxes": [uid,...],
             "EquipmentMailboxes": [uid,...]}
    Returns {} on 403 (MailboxSettings.Read not granted).
    """
    if not user_ids:
        return {}
    reqs = [
        {"id": uid, "method": "GET", "url": f"/users/{uid}/mailboxSettings?$select=userPurpose"}
        for uid in user_ids
    ]
    purposes: dict[str, str] = {}
    try:
        for resp in client.batch(reqs):
            uid     = resp.get("id", "")
            status  = resp.get("status", 0)
            if status == 403:
                log.debug("MailboxSettings.Read not granted — skipping mailbox type detection")
                return {}
            if status == 200:
                body = resp.get("body", {}) or {}
                purposes[uid] = body.get("userPurpose", "user") or "user"
    except Exception as exc:
        log.debug("Mailbox type batch unavailable: %s", exc)
        return {}

    shared    = [uid for uid, p in purposes.items() if p == "shared"]
    equipment = [uid for uid, p in purposes.items() if p == "equipment"]
    return {
        "MailboxPurposes":    purposes,
        "SharedMailboxes":    shared,
        "EquipmentMailboxes": equipment,
    }


def collect_inbox_forwarding_rules(
    client: GraphClient,
    user_ids: list[str],
    accepted_domains: set[str] | None = None,
) -> list[dict]:
    """
    Batch GET /users/{id}/mailFolders/inbox/messageRules for each user.
    Returns rules that have ForwardTo or RedirectTo actions.
    When accepted_domains is provided, each rule includes an ExternalAddresses list.
    Returns [] on 403 (MailboxSettings.Read not granted).
    """
    if not user_ids:
        return []
    reqs = [
        {"id": uid, "method": "GET", "url": f"/users/{uid}/mailFolders/inbox/messageRules"}
        for uid in user_ids
    ]
    results: list[dict] = []
    try:
        for resp in client.batch(reqs):
            uid    = resp.get("id", "")
            status = resp.get("status", 0)
            if status == 403:
                log.debug("MailboxSettings.Read not granted — skipping inbox rule scan")
                return []
            if status != 200:
                continue
            rules = (resp.get("body", {}) or {}).get("value", [])
            for rule in rules:
                actions = rule.get("actions", {}) or {}
                fwd     = [r.get("emailAddress", {}).get("address", "") for r in (actions.get("forwardTo") or [])]
                redir   = [r.get("emailAddress", {}).get("address", "") for r in (actions.get("redirectTo") or [])]
                all_targets = [a for a in fwd + redir if a]
                if not all_targets:
                    continue
                entry: dict[str, Any] = {
                    "UserId":              uid,
                    "RuleName":            rule.get("displayName", ""),
                    "IsEnabled":           rule.get("isEnabled", False),
                    "ForwardToAddresses":  fwd,
                    "RedirectToAddresses": redir,
                }
                if accepted_domains is not None:
                    entry["ExternalAddresses"] = [
                        a for a in all_targets
                        if "@" in a and a.split("@")[1].lower() not in accepted_domains
                    ]
                results.append(entry)
    except Exception as exc:
        log.debug("Inbox rule scan unavailable: %s", exc)
        return []
    return results


# ---------------------------------------------------------------------------
# Builders
# ---------------------------------------------------------------------------

def _build_mailbox_details(rows: list[dict]) -> list[dict]:
    """Per-mailbox records from the usage report CSV, used for storage cross-reference."""
    result = []
    for r in rows:
        if r.get("Is Deleted", "False") == "True":
            continue
        result.append({
            "UserPrincipalName": r.get("User Principal Name", ""),
            "DisplayName":       r.get("Display Name", ""),
            "LastActivityDate":  r.get("Last Activity Date", ""),
            "ItemCount":         _int(r.get("Item Count", "0")),
            "StorageUsedGB":     round(_float(r.get("Storage Used (Byte)", "0")) / (1024 ** 3), 3),
            "StorageQuotaGB":    round(_float(r.get("Prohibit Send/Receive Quota (Byte)", "0")) / (1024 ** 3), 1),
            "HasArchive":        r.get("Has Archive", "False") == "True",
        })
    return result


def _top_n(rows: list[dict], sort_col: str, n: int) -> list[dict]:
    try:
        return sorted(rows, key=lambda r: _int(r.get(sort_col, "")), reverse=True)[:n]
    except Exception:
        return rows[:n]


def _build_primary_mailbox_stats(rows: list[dict]) -> dict:
    if not rows:
        return {}
    total_mb = sum(_float(r.get("Storage Used (Byte)", "0")) for r in rows) / (1024 ** 3)
    return {
        "MailboxCount":     len(rows),
        "TotalStorageGB":   round(total_mb, 2),
        "Source":           f"Graph report (period={_PERIOD})",
    }


def _build_archive_stats(rows: list[dict]) -> dict:
    archive_rows = [r for r in rows if r.get("Has Archive") == "True"]
    total_gb = sum(_float(r.get("Archive Mailbox Size (Byte)", "0")) for r in archive_rows) / (1024 ** 3)
    return {
        "ArchiveMailboxCount": len(archive_rows),
        "TotalArchiveStorageGB": round(total_gb, 2),
    }


def _int(v: str) -> int:
    try:
        return int(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0


def _float(v: str) -> float:
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0
