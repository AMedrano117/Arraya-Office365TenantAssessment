"""
Governance section collector.
Produces: PasswordLifecycleSummary, DomainAuthenticationRecords
DLP and Retention policies require the Purview/Compliance PS module and are not collected here.
"""

from __future__ import annotations

import logging
from typing import Any

from ..graph_client import GraphClient

log = logging.getLogger(__name__)

# Exchange Online DKIM selectors and common third-party ones
_DKIM_SELECTORS = ["selector1", "selector2", "google", "k1", "k2", "s1", "s2", "mail"]


def collect(client: GraphClient) -> dict[str, Any]:
    log.info("Collecting Governance section...")
    data: dict[str, Any] = {}

    data["PasswordLifecycleSummary"] = _collect_password_lifecycle(client)

    # Purview/Compliance PS-only datasets
    data["DlpPolicies"]        = None
    data["RetentionPolicies"]  = None

    log.info("Governance: password lifecycle collected; DLP and Retention require Purview module")
    return data


def collect_domain_auth(domain_names: list[str]) -> dict[str, dict]:
    """
    DNS-based SPF/DKIM/DMARC lookup for a list of domain names.
    No GraphClient needed — queries public DNS directly.
    Returns {domain: {SPF, DKIM, DMARC, DKIMSelector, Notes}}.
    """
    try:
        import dns.resolver  # dnspython
    except ImportError:
        log.warning("dnspython not installed; skipping domain auth DNS lookups. Run: pip install dnspython")
        return {}

    results: dict[str, dict] = {}
    for domain in domain_names:
        if not domain or domain.endswith(".onmicrosoft.com"):
            continue
        record = _lookup_domain_auth(dns.resolver, domain)
        results[domain] = record
        log.debug("Domain auth %s: SPF=%s DKIM=%s DMARC=%s", domain, record["SPF"], record["DKIM"], record["DMARC"])

    log.info("Governance: DNS auth records collected for %d domain(s)", len(results))
    return results


def _lookup_domain_auth(resolver, domain: str) -> dict:
    spf   = _check_spf(resolver, domain)
    dmarc, dmarc_policy = _check_dmarc(resolver, domain)
    dkim, dkim_selector = _check_dkim(resolver, domain)
    notes = []
    if dmarc and dmarc_policy:
        notes.append(f"DMARC policy: {dmarc_policy}")
    return {
        "SPF":          "Configured" if spf  else "Not found",
        "DKIM":         "Configured" if dkim else "Not found",
        "DMARC":        "Configured" if dmarc else "Not found",
        "DKIMSelector": dkim_selector or "",
        "DMARCPolicy":  dmarc_policy or "",
        "Notes":        "; ".join(notes) if notes else "",
    }


def _check_spf(resolver, domain: str) -> bool:
    try:
        answers = resolver.resolve(domain, "TXT", raise_on_no_answer=False)
        for rdata in (answers or []):
            txt = "".join(s.decode() if isinstance(s, bytes) else s for s in rdata.strings)
            if txt.lower().startswith("v=spf1"):
                return True
    except Exception:
        pass
    return False


def _check_dmarc(resolver, domain: str) -> tuple[bool, str]:
    try:
        answers = resolver.resolve(f"_dmarc.{domain}", "TXT", raise_on_no_answer=False)
        for rdata in (answers or []):
            txt = "".join(s.decode() if isinstance(s, bytes) else s for s in rdata.strings)
            if "v=dmarc1" in txt.lower():
                policy = ""
                for part in txt.split(";"):
                    part = part.strip()
                    if part.lower().startswith("p="):
                        policy = part[2:].strip()
                        break
                return True, policy
    except Exception:
        pass
    return False, ""


def _check_dkim(resolver, domain: str) -> tuple[bool, str]:
    for selector in _DKIM_SELECTORS:
        host = f"{selector}._domainkey.{domain}"
        try:
            # Try CNAME first (Exchange Online uses CNAME redirect)
            resolver.resolve(host, "CNAME", raise_on_no_answer=False)
            return True, selector
        except Exception:
            pass
        try:
            answers = resolver.resolve(host, "TXT", raise_on_no_answer=False)
            for rdata in (answers or []):
                txt = "".join(s.decode() if isinstance(s, bytes) else s for s in rdata.strings)
                if "v=dkim1" in txt.lower() or "p=" in txt.lower():
                    return True, selector
        except Exception:
            pass
    return False, ""


def _collect_password_lifecycle(client: GraphClient) -> dict:
    try:
        # SSPR policy
        sspr: dict = {}
        try:
            sspr = client.get("policies/authenticationMethodsPolicy") or {}
        except Exception:
            pass

        # Password expiry from domain settings
        domains = client.get("domains?$select=id,isDefault,passwordValidityPeriodInDays,passwordNotificationWindowInDays")
        domains_list = domains if isinstance(domains, list) else []
        default_domain = next((d for d in domains_list if d.get("isDefault")), {})

        sspr_enabled = None
        reg_info = sspr.get("selfServicePasswordReset") or sspr.get("registrationConfiguration") or {}
        if isinstance(reg_info, dict):
            sspr_enabled = reg_info.get("ssprEnabled") or reg_info.get("isEnabled")

        return {
            "PasswordValidityPeriodInDays":      default_domain.get("passwordValidityPeriodInDays"),
            "PasswordNotificationWindowInDays":  default_domain.get("passwordNotificationWindowInDays"),
            "SelfServicePasswordResetEnabled":   sspr_enabled,
            "PasswordNeverExpires":              default_domain.get("passwordValidityPeriodInDays") == 2147483647,
        }
    except Exception as exc:
        log.debug("Password lifecycle collection failed: %s", exc)
        return {}
