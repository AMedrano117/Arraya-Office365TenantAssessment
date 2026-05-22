"""
Microsoft Graph API HTTP client.
Handles: auto-pagination, $batch (20-req chunks), 429 rate-limit retry, CSV report downloads.
"""

from __future__ import annotations

import csv
import io
import logging
import time
from typing import Any

import requests

from .auth import GraphAuthProvider

log = logging.getLogger(__name__)

_V1_BASE   = "https://graph.microsoft.com/v1.0"
_BETA_BASE = "https://graph.microsoft.com/beta"
_TIMEOUT   = 60   # seconds per request
_MAX_RETRY = 5    # max 429 retries


class GraphClient:
    """Thin, stateless Graph REST client with auto-pagination and batching."""

    def __init__(self, auth: GraphAuthProvider) -> None:
        self._auth    = auth
        self._session = requests.Session()

    # ------------------------------------------------------------------
    # Public: single requests
    # ------------------------------------------------------------------

    def get(self, path: str, params: dict | None = None, beta: bool = False,
            extra_headers: dict | None = None) -> Any:
        """
        GET a Graph path and return the full result.
        - If the response has a 'value' array, returns the list (all pages).
        - Otherwise returns the response dict (single object like /organization[0]).
        Raises on HTTP errors.
        Pass extra_headers for advanced query parameters (e.g. ConsistencyLevel: eventual).
        """
        base = _BETA_BASE if beta else _V1_BASE
        url  = f"{base}/{path.lstrip('/')}"
        return self._get_paged(url, params, extra_headers=extra_headers)

    def get_report_csv(self, path: str, beta: bool = False) -> list[dict]:
        """
        Download a Graph usage report CSV (getMailboxUsageDetail, etc.).
        Returns a list of dicts (one per row, keys = CSV headers).
        """
        base = _BETA_BASE if beta else _V1_BASE
        url  = f"{base}/{path.lstrip('/')}"
        resp = self._request("GET", url)
        # Reports redirect to a pre-signed blob URL; requests follows automatically.
        text = resp.text
        reader = csv.DictReader(io.StringIO(text))
        return [dict(row) for row in reader]

    # ------------------------------------------------------------------
    # Public: batch
    # ------------------------------------------------------------------

    def batch(self, requests_list: list[dict]) -> list[dict]:
        """
        Send up to N Graph requests in parallel using the $batch endpoint.
        Each item in requests_list must be a dict with at least {id, method, url}.
        Returns a flat list of response dicts in the same id order.
        """
        results: list[dict] = []
        for chunk in _chunks(requests_list, 20):
            results.extend(self._send_batch(chunk))
        return results

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def _headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self._auth.get_token()}",
            "Accept":        "application/json",
            "Content-Type":  "application/json",
        }

    def _request(self, method: str, url: str, extra_headers: dict | None = None,
                 **kwargs) -> requests.Response:
        for attempt in range(1, _MAX_RETRY + 1):
            headers = {**self._headers(), **(extra_headers or {})}
            resp = self._session.request(
                method, url, headers=headers, timeout=_TIMEOUT, **kwargs
            )
            if resp.status_code == 429:
                retry_after = int(resp.headers.get("Retry-After", "10"))
                log.warning("429 rate-limited; sleeping %ds (attempt %d/%d)", retry_after, attempt, _MAX_RETRY)
                time.sleep(retry_after)
                continue
            if resp.status_code == 401 and attempt == 1:
                # Token may have just expired — force a fresh token next call
                log.debug("401 on attempt 1; token refresh will occur on next header build")
                continue
            resp.raise_for_status()
            return resp
        resp.raise_for_status()  # raise after exhausting retries
        return resp  # unreachable

    def _get_paged(self, url: str, params: dict | None,
                   extra_headers: dict | None = None) -> Any:
        """Fetch all pages of a paged Graph response, merging 'value' arrays."""
        items: list = []
        next_url: str | None = url

        while next_url:
            if next_url == url:
                resp_dict = self._request("GET", url, extra_headers=extra_headers, params=params).json()
            else:
                resp_dict = self._request("GET", next_url, extra_headers=extra_headers).json()

            if "value" in resp_dict:
                items.extend(resp_dict["value"])
                next_url = resp_dict.get("@odata.nextLink")
            else:
                # Single-object response (e.g. /organization returns a dict, not value array)
                return resp_dict

        return items

    def _send_batch(self, reqs: list[dict]) -> list[dict]:
        """POST a single $batch call and return response bodies sorted by id."""
        body = {"requests": reqs}
        resp = self._request("POST", f"{_V1_BASE}/$batch", json=body)
        responses: list[dict] = resp.json().get("responses", [])
        # Sort by original id so callers get predictable ordering
        id_map = {str(r["id"]): r for r in responses}
        return [id_map.get(str(req["id"]), {"id": req["id"], "status": 0, "body": {}}) for req in reqs]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _chunks(lst: list, n: int):
    for i in range(0, len(lst), n):
        yield lst[i : i + n]
