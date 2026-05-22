"""
Collaboration section collector.
Produces: UnifiedGroups, AllTeams, SharePoint, SharePointSharingSummary, OneDrive,
          Office365GroupsActivityTopGroups, TeamsActivityTopUsers,
          CollaborationActivitySummary, TeamsGroupsCleanupCandidates,
          TeamOwners, TeamMemberCounts, PrivateChannels
"""

from __future__ import annotations

import logging
from datetime import datetime, timezone
from typing import Any

from ..graph_client import GraphClient

log = logging.getLogger(__name__)

_PERIOD      = "D30"
_STALE_DAYS  = 90


def collect(client: GraphClient) -> dict[str, Any]:
    log.info("Collecting Collaboration section...")
    data: dict[str, Any] = {}

    data["UnifiedGroups"]  = _collect_unified_groups(client)
    data["AllTeams"]       = _collect_teams(client)
    data["SharePoint"]     = _collect_sharepoint_settings(client)

    groups_activity        = _collect_groups_activity(client)
    data["Office365GroupsActivityTopGroups"] = groups_activity

    teams_activity         = _collect_teams_activity(client)
    data["TeamsActivityTopUsers"] = teams_activity

    sp_usage               = _collect_sharepoint_usage(client)
    data["SharePointSharingSummary"] = _build_sharepoint_summary(sp_usage)
    data["SharePointSiteUsage"]      = _build_sharepoint_site_details(sp_usage)

    od_usage               = _collect_onedrive_usage(client)
    data["OneDrive"]       = _build_onedrive_summary(od_usage)
    data["OneDriveUsageDetails"] = _build_onedrive_user_details(od_usage)

    data["CollaborationActivitySummary"] = _build_collab_activity_summary(
        data["UnifiedGroups"], groups_activity, teams_activity
    )
    data["TeamsGroupsCleanupCandidates"] = _build_cleanup_candidates(
        data["UnifiedGroups"], data["AllTeams"], groups_activity
    )
    data["TeamsVoice"]        = {}
    data["TeamsVoiceSummary"] = {}

    # Per-team enrichment: owners, member counts, private channels
    team_ids = list(data["AllTeams"].keys())
    data["TeamOwners"]      = _collect_team_owners(client, team_ids)
    data["TeamMemberCounts"] = _collect_team_member_counts(client, team_ids)
    data["PrivateChannels"] = _collect_private_channels(client, team_ids)

    log.info("Collaboration: %d groups, %d teams, %d private-channel teams",
             len(data["UnifiedGroups"]), len(data["AllTeams"]),
             sum(1 for v in data["PrivateChannels"].values() if v))
    return data


# ---------------------------------------------------------------------------
# Raw collectors
# ---------------------------------------------------------------------------

def _collect_unified_groups(client: GraphClient) -> dict:
    select = ("id,displayName,mail,mailNickname,groupTypes,createdDateTime,"
              "expirationDateTime,visibility,onPremisesSyncEnabled,mailEnabled,assignedLabels")
    raw = client.get(
        f"groups?$filter=groupTypes/any(c:c+eq+'Unified')&$select={select}"
        f"&$expand=owners($select=id,displayName,userPrincipalName)"
    )
    return {g["id"]: g for g in (raw if isinstance(raw, list) else [])}


def _collect_teams(client: GraphClient) -> dict:
    select = "id,displayName,visibility,isArchived,description,createdDateTime"
    try:
        raw = client.get(f"teams?$select={select}")
        return {t["id"]: t for t in (raw if isinstance(raw, list) else [])}
    except Exception as exc:
        log.debug("Teams list unavailable (may need Teams.ReadBasic.All): %s", exc)
        return {}


def _collect_team_owners(client: GraphClient, team_ids: list[str]) -> dict[str, list]:
    """
    Returns {team_id: [owner_member_dict, ...]} for all teams via $batch.
    Returns {} if TeamMember.Read.All is not granted (403).
    """
    if not team_ids:
        return {}
    requests_list = [
        {
            "id": tid,
            "method": "GET",
            "url": f"/teams/{tid}/members?$filter=roles/any(r:r eq 'owner')"
                   "&$select=id,displayName,email,roles",
        }
        for tid in team_ids
    ]
    results: dict[str, list] = {}
    try:
        responses = client.batch(requests_list)
        for resp in responses:
            tid    = resp.get("id", "")
            status = resp.get("status", 0)
            if status == 403:
                log.debug("Team owners unavailable: TeamMember.Read.All not granted")
                return {}
            body = resp.get("body", {})
            results[tid] = body.get("value", []) if status == 200 else []
    except Exception as exc:
        log.debug("Team owners batch unavailable: %s", exc)
        results = {tid: [] for tid in team_ids}
    return results


def _collect_team_member_counts(client: GraphClient, team_ids: list[str]) -> dict[str, int]:
    """
    Returns {team_id: member_count} via batch requests with $count=true.
    Returns {} if TeamMember.Read.All is not granted (403).
    """
    if not team_ids:
        return {}
    # /teams/{id}/members does not support $count (conversationMember resource).
    # Teams are backed by M365 groups with the same ID — use the groups endpoint instead.
    requests_list = [
        {
            "id": tid,
            "method": "GET",
            "url": f"/groups/{tid}/members?$count=true&$top=1&$select=id",
            "headers": {"ConsistencyLevel": "eventual"},
        }
        for tid in team_ids
    ]
    counts: dict[str, int] = {}
    try:
        responses = client.batch(requests_list)
        for resp in responses:
            tid    = resp.get("id", "")
            status = resp.get("status", 0)
            if status == 403:
                log.debug("Team member counts unavailable: GroupMember.Read.All not granted")
                return {}
            if status == 200:
                body = resp.get("body", {}) or {}
                # @odata.count is the total count regardless of $top
                odata_count = body.get("@odata.count")
                if odata_count is not None:
                    try:
                        counts[tid] = int(odata_count)
                    except (TypeError, ValueError):
                        counts[tid] = -1
                else:
                    counts[tid] = -1
            else:
                counts[tid] = -1
    except Exception as exc:
        log.debug("Team member counts batch unavailable: %s", exc)
        counts = {tid: -1 for tid in team_ids}
    return counts


def _collect_private_channels(client: GraphClient, team_ids: list[str]) -> dict[str, list]:
    """Returns {team_id: [channel_dict, ...]} — empty list means no private channels."""
    if not team_ids:
        return {}
    requests_list = [
        {
            "id": tid,
            "method": "GET",
            "url": f"/teams/{tid}/channels?$filter=membershipType eq 'private'"
                   "&$select=id,displayName,membershipType,createdDateTime",
        }
        for tid in team_ids
    ]
    results: dict[str, list] = {}
    try:
        responses = client.batch(requests_list)
        for resp in responses:
            tid = resp.get("id", "")
            body = resp.get("body", {})
            results[tid] = body.get("value", []) if resp.get("status", 0) == 200 else []
    except Exception as exc:
        log.debug("Private channels batch unavailable: %s", exc)
        results = {tid: [] for tid in team_ids}
    return results


def _collect_groups_activity(client: GraphClient) -> list[dict]:
    try:
        return client.get_report_csv(f"reports/getOffice365GroupsActivityDetail(period='{_PERIOD}')")
    except Exception as exc:
        log.debug("Groups activity report unavailable: %s", exc)
        return []


def _collect_teams_activity(client: GraphClient) -> list[dict]:
    try:
        return client.get_report_csv(f"reports/getTeamsUserActivityUserDetail(period='{_PERIOD}')")
    except Exception as exc:
        log.debug("Teams user activity report unavailable: %s", exc)
        return []


def _collect_sharepoint_settings(client: GraphClient) -> dict:
    try:
        settings = client.get("admin/sharepoint/settings")
        return {
            "SharingCapability":      settings.get("sharingCapability", ""),
            "DefaultSharingLinkType": settings.get("defaultSharingLinkType", ""),
            "DefaultLinkPermission":  settings.get("defaultLinkPermission", ""),
        }
    except Exception as exc:
        log.debug("SharePoint admin settings unavailable: %s", exc)
        return {}


def _collect_sharepoint_usage(client: GraphClient) -> list[dict]:
    try:
        return client.get_report_csv(f"reports/getSharePointSiteUsageDetail(period='{_PERIOD}')")
    except Exception as exc:
        log.debug("SharePoint site usage report unavailable: %s", exc)
        return []


def _collect_onedrive_usage(client: GraphClient) -> list[dict]:
    try:
        return client.get_report_csv(f"reports/getOneDriveUsageAccountDetail(period='{_PERIOD}')")
    except Exception as exc:
        log.debug("OneDrive usage report unavailable: %s", exc)
        return []


# ---------------------------------------------------------------------------
# Builders
# ---------------------------------------------------------------------------

def _build_sharepoint_summary(rows: list[dict]) -> dict:
    if not rows:
        return {}
    active = [r for r in rows if r.get("Is Deleted", "False") == "False"]
    external_sharing = sum(1 for r in active if r.get("External Sharing", "0") not in ("0", "", None))
    total_gb = sum(_float(r.get("Storage Used (Byte)", "0")) for r in active) / (1024 ** 3)
    return {
        "TotalSites":          len(active),
        "SitesWithExternalSharing": external_sharing,
        "TotalStorageGB":      round(total_gb, 2),
    }


def _build_onedrive_summary(rows: list[dict]) -> dict:
    if not rows:
        return {}
    active = [r for r in rows if r.get("Is Deleted", "False") == "False"]
    total_gb = sum(_float(r.get("Storage Used (Byte)", "0")) for r in active) / (1024 ** 3)
    return {
        "TotalAccounts":   len(active),
        "TotalStorageGB":  round(total_gb, 2),
    }


def _build_sharepoint_site_details(rows: list[dict]) -> list[dict]:
    result = []
    for r in rows:
        if r.get("Is Deleted", "False") == "True":
            continue
        result.append({
            "SiteUrl":           r.get("Site URL", ""),
            "SiteDisplayName":   r.get("Site Display Name", ""),
            "LastActivityDate":  r.get("Last Activity Date", ""),
            "FileCount":         _int(r.get("File Count", "0")),
            "ActiveFileCount":   _int(r.get("Active File Count", "0")),
            "PageViewCount":     _int(r.get("Page View Count", "0")),
            "StorageUsedGB":     round(_float(r.get("Storage Used (Byte)", "0")) / (1024 ** 3), 3),
            "StorageAllocatedGB": round(_float(r.get("Storage Allocated (Byte)", "0")) / (1024 ** 3), 1),
            "ExternalSharing":   r.get("External Sharing", ""),
            "RootWebTemplate":   r.get("Root Web Template", ""),
        })
    return sorted(result, key=lambda x: x.get("StorageUsedGB", 0), reverse=True)


def _build_onedrive_user_details(rows: list[dict]) -> list[dict]:
    result = []
    for r in rows:
        if r.get("Is Deleted", "False") == "True":
            continue
        result.append({
            "OwnerPrincipalName": r.get("Owner Principal Name", ""),
            "OwnerDisplayName":   r.get("Owner Display Name", ""),
            "SiteUrl":            r.get("Site URL", ""),
            "LastActivityDate":   r.get("Last Activity Date", ""),
            "FileCount":          _int(r.get("File Count", "0")),
            "ActiveFileCount":    _int(r.get("Active File Count", "0")),
            "StorageUsedGB":      round(_float(r.get("Storage Used (Byte)", "0")) / (1024 ** 3), 3),
            "StorageAllocatedGB": round(_float(r.get("Storage Allocated (Byte)", "0")) / (1024 ** 3), 1),
        })
    return sorted(result, key=lambda x: x.get("StorageUsedGB", 0), reverse=True)


def _build_collab_activity_summary(groups: dict, groups_activity: list, teams_activity: list) -> dict:
    return {
        "TotalUnifiedGroups": len(groups),
        "GroupsWithActivityInPeriod": sum(
            1 for r in groups_activity
            if _int(r.get("Exchange Emails Received Count", "0")) > 0
            or _int(r.get("SharePoint Active Files Count", "0")) > 0
        ),
        "ActiveTeamsUsersInPeriod": sum(
            1 for r in teams_activity
            if _int(r.get("Team Chat Message Count", "0")) > 0
        ),
    }


def _build_cleanup_candidates(groups: dict, teams: dict, groups_activity: list) -> dict:
    activity_by_id: dict[str, dict] = {}
    for r in groups_activity:
        gid = r.get("Group Id", "")
        if gid:
            activity_by_id[gid] = r

    candidates: dict[str, dict] = {}
    now = datetime.now(timezone.utc)

    for gid, g in groups.items():
        owners = g.get("owners", []) or []
        act    = activity_by_id.get(gid, {})
        last_act = act.get("Report Date") or act.get("Last Activity Date") or ""

        # Member count comes from the activity report CSV; the groups API doesn't return it directly
        member_count    = _int(act.get("Member Count", "-1"))
        ext_member_count = _int(act.get("External Member Count", "0"))

        signals = []
        if not owners:
            signals.append("Ownerless")
        if member_count == 0:
            signals.append("No members")
        if last_act:
            try:
                dt = datetime.fromisoformat(last_act.replace("Z", "+00:00"))
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=timezone.utc)
                if (now - dt).days >= _STALE_DAYS:
                    signals.append("Dormant")
            except (ValueError, AttributeError):
                pass
        elif not act:
            signals.append("No activity in period")

        if signals:
            candidates[gid] = {
                "Name":              g.get("displayName", ""),
                "GroupId":           gid,
                "OwnerCount":        len(owners),
                "MemberCount":       member_count if member_count >= 0 else None,
                "ExternalMemberCount": ext_member_count,
                "RiskSignal":        "; ".join(signals),
                "LastActivityDate":  last_act,
                "IsTeam":            gid in teams,
            }

    return candidates


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
