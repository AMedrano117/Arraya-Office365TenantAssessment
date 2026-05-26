"""
Exchange Online data collector via PowerShell subprocess.

Collects EXO-only datasets not available via Graph API:
  AllMailboxes, MailFlowRules, MailFlowConnectors, SpamFilteringConfig,
  RemoteDomains, PublicFolderDetails, InactiveMailboxDetails,
  LitigationHoldMailboxes, ForwardingPolicySummary

Requires: ExchangeOnlineManagement PS module v3+
Auth:
  Certificate mode -- needs Exchange.ManageAsApp app permission
  Interactive mode -- user needs Exchange Admin or Global Reader role
  ClientSecret mode -- not supported by Connect-ExchangeOnline; skipped
"""

from __future__ import annotations

import json
import logging
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

log = logging.getLogger(__name__)

_SCRIPT_REL = Path("src/scripts/assessments/exchange/Get-EXOData.ps1")
_TIMEOUT_S  = 300


def collect(auth: Any, organization: str) -> dict[str, Any]:
    """
    Run the EXO PS subprocess and return the collected datasets.
    Returns {} gracefully when PS unavailable, module missing, or auth unsupported.
    The caller merges the returned dict into the Exchange section of the snapshot.
    """
    if not organization:
        log.debug("No organization domain provided; EXO collection skipped")
        return {}

    auth_mode = (getattr(auth, "auth_mode", "") or "").strip().lower()
    if auth_mode in ("clientsecret", "secret"):
        log.info(
            "EXO collection skipped: ClientSecret auth is not supported by "
            "Connect-ExchangeOnline. Use Certificate or Interactive auth."
        )
        return {}

    pwsh = _find_pwsh()
    if not pwsh:
        log.info("EXO collection skipped: PowerShell not found on PATH")
        return {}

    script = _find_script()
    if not script:
        log.warning("EXO collection skipped: Get-EXOData.ps1 not found")
        return {}

    with tempfile.NamedTemporaryFile(suffix="-exa.json", delete=False) as tmp:
        out_path = tmp.name

    # Map Python auth_mode names to PS param values
    ps_auth_mode = "Certificate" if auth_mode in ("certificate", "cert") else "Interactive"

    cmd = [
        pwsh, "-NonInteractive", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(script),
        "-OutputPath", out_path,
        "-AuthMode", ps_auth_mode,
    ]
    if organization:
        cmd += ["-Organization", organization]

    client_id = getattr(auth, "client_id", "") or ""
    cert_tp    = getattr(auth, "cert_thumbprint", "") or ""
    tenant_id  = getattr(auth, "tenant_id", "") or ""

    if client_id:
        cmd += ["-AppId", client_id]
    if cert_tp and ps_auth_mode == "Certificate":
        cmd += ["-CertificateThumbprint", cert_tp]
    if tenant_id:
        cmd += ["-TenantId", tenant_id]

    log.info("Running EXO PS subprocess (auth=%s org=%s)", ps_auth_mode, organization)
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=_TIMEOUT_S,
            check=False,
        )
    except subprocess.TimeoutExpired:
        log.warning("EXO PS subprocess timed out after %ds", _TIMEOUT_S)
        _rm(out_path)
        return {}
    except Exception as exc:
        log.warning("EXO PS subprocess could not start: %s", exc)
        _rm(out_path)
        return {}

    if proc.returncode != 0:
        stderr_lower = (proc.stderr or "").lower()
        if "exchangeonlinemanagement" in stderr_lower or "module" in stderr_lower:
            log.info("EXO collection skipped: ExchangeOnlineManagement module not installed")
        elif "unauthorized" in stderr_lower or "403" in stderr_lower:
            log.info("EXO collection skipped: Exchange.ManageAsApp permission not granted")
        else:
            log.warning(
                "EXO PS subprocess exited %d. stderr: %s",
                proc.returncode,
                (proc.stderr or "")[:500],
            )
        _rm(out_path)
        return {}

    try:
        raw  = Path(out_path).read_text(encoding="utf-8-sig")
        data = json.loads(raw)
    except FileNotFoundError:
        log.warning("EXO PS subprocess did not produce output file")
        return {}
    except json.JSONDecodeError as exc:
        log.warning("EXO PS subprocess output is not valid JSON: %s", exc)
        return {}
    finally:
        _rm(out_path)

    if not isinstance(data, dict):
        log.warning("EXO PS subprocess returned unexpected type: %s", type(data))
        return {}

    log.info("EXO subprocess collected %d dataset(s)", len(data))
    return data


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _find_pwsh() -> str | None:
    for candidate in ("pwsh", "powershell"):
        found = shutil.which(candidate)
        if found:
            return found
    if sys.platform == "win32":
        for p in (
            r"C:\Program Files\PowerShell\7\pwsh.exe",
            r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
        ):
            if Path(p).exists():
                return p
    return None


def _find_script() -> Path | None:
    # Walk up from this file to find the repo root (contains src/)
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / _SCRIPT_REL
        if candidate.exists():
            return candidate
    return None


def _rm(path: str) -> None:
    try:
        Path(path).unlink(missing_ok=True)
    except Exception:
        pass
