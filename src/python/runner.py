"""
Assessment runner — orchestrates PowerShell collection and Python reporting.
PowerShell is invoked as a subprocess for anything that requires M365 cmdlets.
"""

from __future__ import annotations

import logging
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

from .utils.paths import find_repo_root

log = logging.getLogger(__name__)


def run_collection(
    export_path: Path,
    output_profiles: list[str],
    auth_mode: str = "",
    tenant_id: str = "",
    client_id: str = "",
    certificate_thumbprint: str = "",
    client_secret: str = "",
    skip_auth: bool = False,
    skip_preflight: bool = False,
    use_graph_fallback: bool = False,
    run_improve: bool = False,
    extra_params: dict[str, Any] | None = None,
) -> int:
    """Call the PowerShell data collection script and return its exit code."""
    repo_root = find_repo_root()
    ps_script = repo_root / "src" / "scripts" / "operations" / "Start-M365TenantAssessment.ps1"
    if not ps_script.exists():
        raise FileNotFoundError(f"PowerShell launcher not found: {ps_script}")

    pwsh = _find_pwsh()
    cmd: list[str] = [
        pwsh, "-NonInteractive", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(ps_script),
        "-Action", "Collect",
        "-ExportPath", str(export_path),
        "-OutputProfile", ",".join(output_profiles),
    ]

    if auth_mode:
        cmd += ["-AuthMode", auth_mode]
    if tenant_id:
        cmd += ["-TenantId", tenant_id]
    if client_id:
        cmd += ["-ClientId", client_id]
    if certificate_thumbprint:
        cmd += ["-CertificateThumbprint", certificate_thumbprint]
    if client_secret:
        cmd += ["-ClientSecret", client_secret]
    if skip_auth:
        cmd.append("-SkipAuth")
    if skip_preflight:
        cmd.append("-SkipPermissionPreflight")
    if use_graph_fallback:
        cmd.append("-UseGraphFallback")
    if run_improve:
        cmd.append("-RunImprove")

    log.info("Launching PowerShell collection: %s", " ".join(cmd))
    result = subprocess.run(cmd, check=False)
    return result.returncode


def run_preflight(
    output_profiles: list[str],
    auth_mode: str = "",
    tenant_id: str = "",
    client_id: str = "",
    certificate_thumbprint: str = "",
    client_secret: str = "",
    skip_auth: bool = False,
) -> int:
    repo_root = find_repo_root()
    ps_script = repo_root / "src" / "scripts" / "operations" / "Start-M365TenantAssessment.ps1"
    pwsh = _find_pwsh()
    cmd = [
        pwsh, "-NonInteractive", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(ps_script),
        "-Action", "Preflight",
        "-OutputProfile", ",".join(output_profiles),
    ]
    if auth_mode:
        cmd += ["-AuthMode", auth_mode]
    if tenant_id:
        cmd += ["-TenantId", tenant_id]
    if client_id:
        cmd += ["-ClientId", client_id]
    if certificate_thumbprint:
        cmd += ["-CertificateThumbprint", certificate_thumbprint]
    if client_secret:
        cmd += ["-ClientSecret", client_secret]
    if skip_auth:
        cmd.append("-SkipAuth")

    log.info("Launching PowerShell preflight")
    return subprocess.run(cmd, check=False).returncode


def run_ad_assessment() -> int:
    repo_root = find_repo_root()
    ps_script = repo_root / "src" / "scripts" / "operations" / "Start-M365TenantAssessment.ps1"
    pwsh = _find_pwsh()
    cmd = [
        pwsh, "-NonInteractive", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(ps_script),
        "-Action", "AD",
    ]
    log.info("Launching PowerShell AD assessment")
    return subprocess.run(cmd, check=False).returncode


def _find_pwsh() -> str:
    for candidate in ("pwsh", "powershell"):
        found = shutil.which(candidate)
        if found:
            return found
    if sys.platform == "win32":
        for path in (
            r"C:\Program Files\PowerShell\7\pwsh.exe",
            r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
        ):
            if Path(path).exists():
                return path
    raise RuntimeError("PowerShell (pwsh or powershell) not found on PATH.")
