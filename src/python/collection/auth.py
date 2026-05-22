"""
Graph API token acquisition via MSAL.
Supports Interactive (device code), Certificate, and ClientSecret auth modes.
"""

from __future__ import annotations

import logging
import os
from typing import Any

log = logging.getLogger(__name__)

_GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]


class GraphAuthProvider:
    """Acquires and caches Microsoft Graph access tokens via MSAL."""

    def __init__(
        self,
        auth_mode: str,
        tenant_id: str,
        client_id: str,
        cert_thumbprint: str = "",
        client_secret: str = "",
    ) -> None:
        self.auth_mode      = (auth_mode or "Interactive").strip()
        self.tenant_id      = (tenant_id or "").strip()
        self.client_id      = (client_id or "").strip()
        self.cert_thumbprint = (cert_thumbprint or "").strip()
        self.client_secret  = (client_secret or "").strip()
        self._app: Any      = None
        self._token_cache: dict | None = None

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def get_token(self) -> str:
        """Return a valid Bearer token string, refreshing if needed."""
        result = self._acquire()
        if not result or "access_token" not in result:
            error = (result or {}).get("error_description", "unknown error")
            raise RuntimeError(f"Failed to acquire Graph token: {error}")
        return result["access_token"]

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def _acquire(self) -> dict:
        mode = self.auth_mode.lower()
        if mode in ("certificate", "cert"):
            return self._acquire_certificate()
        if mode in ("clientsecret", "secret"):
            return self._acquire_secret()
        return self._acquire_interactive()

    def _confidential_app(self, credential: Any) -> Any:
        import msal  # noqa: PLC0415
        authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        return msal.ConfidentialClientApplication(
            self.client_id, authority=authority, client_credential=credential
        )

    def _acquire_certificate(self) -> dict:
        import msal  # noqa: PLC0415
        cert = _load_certificate(self.cert_thumbprint)
        app  = self._confidential_app(cert)
        result = app.acquire_token_for_client(scopes=_GRAPH_SCOPES)
        if not result or "access_token" not in result:
            # Fallback: try without private key (thumbprint-only, cert in store)
            log.warning("Certificate load may have partially failed; check thumbprint and cert store.")
        return result or {}

    def _acquire_secret(self) -> dict:
        app = self._confidential_app(self.client_secret)
        return app.acquire_token_for_client(scopes=_GRAPH_SCOPES) or {}

    def _acquire_interactive(self) -> dict:
        import msal  # noqa: PLC0415
        authority = f"https://login.microsoftonline.com/{self.tenant_id or 'common'}"
        app = msal.PublicClientApplication(
            self.client_id or "14d82eec-204b-4c2f-b7e8-296a70dab67e",  # well-known Graph Explorer
            authority=authority,
        )
        # Try silent first (cached accounts)
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes=_GRAPH_SCOPES, account=accounts[0])
            if result and "access_token" in result:
                return result

        # Device code flow — works in headless/terminal environments
        flow = app.initiate_device_flow(scopes=_GRAPH_SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to initiate device flow: {flow.get('error_description', flow)}")
        print(f"\n{flow['message']}\n")  # shows the URL and code to the user
        return app.acquire_token_by_device_flow(flow) or {}


# ---------------------------------------------------------------------------
# Certificate loading helpers
# ---------------------------------------------------------------------------

def _load_certificate(thumbprint: str) -> dict:
    """
    Load a certificate from the Windows certificate store (My/CurrentUser) by thumbprint.
    Returns an msal-compatible {thumbprint, private_key} dict.
    Falls back to thumbprint-only if the private key is not exportable.
    """
    if not thumbprint:
        raise ValueError("cert_thumbprint is required for Certificate auth mode")

    # Try Windows certificate store first
    try:
        return _load_from_windows_store(thumbprint)
    except Exception as exc:
        log.debug("Windows cert store lookup failed (%s); trying PEM file fallback", exc)

    # Fallback: look for a PEM file named by thumbprint in the current dir or APPDATA
    pem_path = _find_pem_file(thumbprint)
    if pem_path:
        return _load_from_pem(thumbprint, pem_path)

    # Last resort: pass just the thumbprint and let MSAL try the store itself
    log.warning("Could not load private key for cert %s; using thumbprint-only credential.", thumbprint)
    return {"thumbprint": thumbprint}


def _load_from_windows_store(thumbprint: str) -> dict:
    import subprocess, json as _json  # noqa: PLC0415
    # Use PowerShell to export the cert bytes — avoids adding cryptography dep
    ps_cmd = (
        f"$c = Get-ChildItem Cert:\\CurrentUser\\My\\{thumbprint} -ErrorAction Stop; "
        f"$bytes = $c.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx); "
        f"[Convert]::ToBase64String($bytes)"
    )
    result = subprocess.run(
        ["pwsh", "-NonInteractive", "-NoProfile", "-Command", ps_cmd],
        capture_output=True, text=True, timeout=15,
    )
    if result.returncode != 0 or not result.stdout.strip():
        raise RuntimeError(result.stderr.strip() or "pwsh returned no output")

    import base64  # noqa: PLC0415
    pfx_bytes = base64.b64decode(result.stdout.strip())

    # Try to extract PEM private key via cryptography; fall back to raw PFX bytes (MSAL v1.18+)
    try:
        import warnings  # noqa: PLC0415
        from cryptography.hazmat.primitives.serialization import pkcs12, Encoding, PrivateFormat, NoEncryption  # noqa: PLC0415
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            private_key, _cert, _ = pkcs12.load_key_and_certificates(pfx_bytes, None)
        if private_key is not None:
            pem = private_key.private_bytes(Encoding.PEM, PrivateFormat.PKCS8, NoEncryption()).decode()
            return {"thumbprint": thumbprint, "private_key": pem}
        log.debug("PFX parsed but contained no private key; passing raw PFX bytes to MSAL")
    except Exception as exc:
        log.debug("cryptography PFX parse failed (%s); passing raw PFX bytes to MSAL", exc)

    # MSAL v1.18+ accepts raw PFX bytes directly — no password since Windows export uses none
    return {"thumbprint": thumbprint, "private_key_pfx_bytes": pfx_bytes, "passphrase": None}


def _find_pem_file(thumbprint: str) -> str | None:
    candidates = [
        f"{thumbprint}.pem",
        os.path.join(os.environ.get("APPDATA", ""), f"{thumbprint}.pem"),
        os.path.join(os.path.expanduser("~"), f"{thumbprint}.pem"),
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    return None


def _load_from_pem(thumbprint: str, pem_path: str) -> dict:
    with open(pem_path, "r") as f:
        pem_data = f.read()
    return {"thumbprint": thumbprint, "private_key": pem_data}
