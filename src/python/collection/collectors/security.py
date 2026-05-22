"""
Security section collector.
Produces: SecuritySecureScore, SecureScoreActions, SMTPRelaySummary, SpamFilteringSummary
"""

from __future__ import annotations

import logging
from typing import Any

from ..graph_client import GraphClient

log = logging.getLogger(__name__)


def collect(client: GraphClient) -> dict[str, Any]:
    log.info("Collecting Security section...")
    data: dict[str, Any] = {}

    data["SecuritySecureScore"] = _collect_secure_score(client)
    data["SecureScoreActions"]  = _collect_secure_score_controls(client)

    # EXO-only datasets — not available via Graph
    data["SMTPRelayConfig"]   = None
    data["SMTPRelaySummary"]  = {}
    data["SpamFilteringConfig"] = None

    log.info("Security: Secure Score collected; SMTP/spam config requires EXO module")
    return data


def _collect_secure_score(client: GraphClient) -> dict:
    try:
        raw = client.get("security/secureScores?$top=1")
        scores = raw if isinstance(raw, list) else []
        return scores[0] if scores else {}
    except Exception as exc:
        log.debug("SecureScore unavailable: %s", exc)
        return {}


def _collect_secure_score_controls(client: GraphClient) -> list:
    try:
        raw = client.get("security/secureScoreControlProfiles")
        return raw if isinstance(raw, list) else []
    except Exception as exc:
        log.debug("SecureScoreControlProfiles unavailable: %s", exc)
        return []
