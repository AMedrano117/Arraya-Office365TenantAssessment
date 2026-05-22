"""
Tenant section collector.
Produces: TenantInfo, Domains, AdConnectConfiguration, FederationConfiguration, ExternalSharingSummary
"""

from __future__ import annotations

import logging
from typing import Any

from ..graph_client import GraphClient

log = logging.getLogger(__name__)


def collect(client: GraphClient) -> dict[str, Any]:
    log.info("Collecting Tenant section...")
    data: dict[str, Any] = {}

    # Organization object
    org_list = client.get("organization?$select=id,displayName,verifiedDomains,onPremisesSyncEnabled,"
                          "onPremisesLastSyncDateTime,onPremisesLastPasswordSyncDateTime,"
                          "countryLetterCode,createdDateTime,tenantType,assignedPlans")
    org = org_list[0] if isinstance(org_list, list) and org_list else (org_list or {})

    data["TenantInfo"] = _build_tenant_info(org)

    # Domains
    domains_raw = client.get("domains?$select=id,isDefault,isVerified,isInitial,authenticationType,"
                             "availabilityStatus,passwordValidityPeriodInDays,passwordNotificationWindowInDays")
    data["Domains"] = {d["id"]: d for d in (domains_raw if isinstance(domains_raw, list) else [])}

    # AdConnect signals from organization
    data["AdConnectConfiguration"] = _build_adconnect(org)

    # Federation configuration for each federated domain
    data["FederationConfiguration"] = _collect_federation(client, data["Domains"])

    # External sharing summary stub (Graph admin/sharepoint settings — beta)
    data["ExternalSharingSummary"] = _collect_external_sharing(client)

    # HybridConfiguration derived from AdConnect
    data["HybridConfiguration"] = _build_hybrid(data["AdConnectConfiguration"])

    # Detailed sync features — requires OnPremDirectorySynchronization.Read.All
    data["HybridSyncDetails"] = _collect_hybrid_sync_details(client)

    # Report display settings — detects if usage reports are anonymized (GUIDs instead of names)
    data["ReportSettings"] = _collect_report_settings(client)

    log.info("Tenant: collected TenantInfo, %d domain(s), AdConnect, Federation", len(data["Domains"]))
    return data


# ---------------------------------------------------------------------------
# Builders
# ---------------------------------------------------------------------------

def _build_tenant_info(org: dict) -> dict:
    verified = [d["name"] for d in (org.get("verifiedDomains") or []) if d.get("isDefault")]
    default_domain = verified[0] if verified else ""
    return {
        "TenantId":          org.get("id", ""),
        "DisplayName":       org.get("displayName", ""),
        "DefaultDomain":     default_domain,
        "CountryCode":       org.get("countryLetterCode", ""),
        "CreatedDateTime":   org.get("createdDateTime", ""),
        "TenantType":        org.get("tenantType", ""),
        "OnPremisesSyncEnabled": org.get("onPremisesSyncEnabled"),
    }


def _build_adconnect(org: dict) -> dict:
    summary = {
        "OnPremisesSyncEnabled":              org.get("onPremisesSyncEnabled"),
        "LastSyncDateTime":                   org.get("onPremisesLastSyncDateTime"),
        "LastPasswordSyncDateTime":           org.get("onPremisesLastPasswordSyncDateTime"),
    }
    return {"Summary": summary, "RecentErrors": None, "SyncServices": [], "ErrorCount": 0}


def _collect_federation(client: GraphClient, domains: dict) -> dict:
    federated = [d_id for d_id, d in domains.items() if d.get("authenticationType") == "Federated"]
    result: dict = {}
    for d_id in federated:
        try:
            fed = client.get(f"domains/{d_id}/federationConfiguration")
            result[d_id] = fed if isinstance(fed, list) else [fed]
        except Exception as exc:
            log.debug("Federation config fetch failed for %s: %s", d_id, exc)
    return result


def _collect_external_sharing(client: GraphClient) -> dict:
    try:
        settings = client.get("admin/sharepoint/settings")
        return {
            "SharingCapability":        settings.get("sharingCapability", ""),
            "DefaultSharingLinkType":   settings.get("defaultSharingLinkType", ""),
            "DefaultLinkPermission":    settings.get("defaultLinkPermission", ""),
            "AllowAnonymousAccess":     settings.get("isExternalUserSelfServiceSignUpEnabled"),
        }
    except Exception as exc:
        log.debug("SharePoint admin settings unavailable: %s", exc)
        return {}


def _collect_report_settings(client: GraphClient) -> dict:
    """
    GET /admin/reportSettings — returns isDisplayConcealedNames.
    When true, all usage report CSVs contain GUIDs instead of user/group names.
    Requires ReportSettings.Read.All. Returns {} on 403.
    """
    try:
        return client.get("admin/reportSettings") or {}
    except Exception as exc:
        log.debug("ReportSettings unavailable (need ReportSettings.Read.All): %s", exc)
        return {}


def _collect_hybrid_sync_details(client: GraphClient) -> dict:
    """
    GET /directory/onPremisesSynchronization — detailed sync config including feature flags.
    Requires OnPremDirectorySynchronization.Read.All. Returns {} on 403/404.
    """
    try:
        raw = client.get("directory/onPremisesSynchronization")
        items = raw if isinstance(raw, list) else ([raw] if raw else [])
        return items[0] if items else {}
    except Exception as exc:
        log.debug("HybridSyncDetails unavailable (need OnPremDirectorySynchronization.Read.All): %s", exc)
        return {}


def _build_hybrid(adconnect: dict) -> dict:
    summary = adconnect.get("Summary", {})
    return {
        "IsHybrid":          bool(summary.get("OnPremisesSyncEnabled")),
        "LastSyncDateTime":  summary.get("LastSyncDateTime"),
    }
