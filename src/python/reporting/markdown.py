"""
Markdown generation — Python equivalent of Export-TenantToTenantQuestionnaireMarkdown.
Uses stdlib only.
"""

from __future__ import annotations

import logging
from pathlib import Path

from ..snapshot import get_data, get_metadata
from ..utils.converters import ensure_list

log = logging.getLogger(__name__)


def generate_questionnaire(snapshot: dict, output_path: str | Path) -> None:
    output_path = Path(output_path)
    if output_path.suffix.lower() != ".md":
        output_path = output_path.with_suffix(".md")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    lines: list[str] = []
    meta = get_metadata(snapshot)
    tenant = meta.get("TenantDisplayName") or meta.get("TenantDomain") or "Tenant"

    lines += [
        f"# Microsoft 365 Migration Questionnaire — {tenant}",
        "",
        f"> Generated: {meta.get('GeneratedAt', '')}",
        "",
    ]

    data = get_data(snapshot)

    # Tenant overview
    tenant_info = _find_table(data, "TenantInfoSummary", "TenantInfo")
    if tenant_info:
        lines += ["## Tenant Overview", ""]
        lines += _dict_to_table(tenant_info if isinstance(tenant_info, dict) else (ensure_list(tenant_info) or [{}])[0])
        lines.append("")

    # Domain summary
    domains = _find_table(data, "Domains")
    if domains:
        lines += ["## Accepted Domains", ""]
        lines += _list_to_table(ensure_list(domains))
        lines.append("")

    # Mailbox summary
    mailboxes = _find_table(data, "AllMailboxes")
    if mailboxes:
        lines += [f"## Mailboxes ({len(ensure_list(mailboxes))} total)", ""]
        lines += _list_to_table(ensure_list(mailboxes)[:20])
        if len(ensure_list(mailboxes)) > 20:
            lines.append(f"> _(showing first 20 of {len(ensure_list(mailboxes))})_")
        lines.append("")

    # Questions
    lines += [
        "## Migration Questions",
        "",
        "| # | Question | Answer |",
        "|---|----------|--------|",
        "| 1 | What is the target tenant domain? | |",
        "| 2 | What is the planned cutover date? | |",
        "| 3 | Are there any hybrid dependencies? | |",
        "| 4 | Will archive mailboxes be migrated? | |",
        "| 5 | Are there public folders to migrate? | |",
        "| 6 | What third-party apps need reconfiguration? | |",
        "",
    ]

    output_path.write_text("\n".join(lines), encoding="utf-8")
    log.info("Markdown questionnaire saved: %s", output_path)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _find_table(data: dict, *keys: str):
    for key in keys:
        for section in data.values():
            if isinstance(section, dict) and key in section:
                return section[key]
    return None


def _dict_to_table(record: dict) -> list[str]:
    lines = ["| Field | Value |", "|-------|-------|"]
    for k, v in record.items():
        lines.append(f"| {k} | {v if v is not None else ''} |")
    return lines


def _list_to_table(rows: list) -> list[str]:
    if not rows:
        return ["_No data._"]
    if not isinstance(rows[0], dict):
        return [f"- {r}" for r in rows]
    headers = list(rows[0].keys())
    sep = "|".join("---" for _ in headers)
    lines = ["| " + " | ".join(headers) + " |", f"| {sep} |"]
    for r in rows:
        cells = " | ".join(str(r.get(h, "")) for h in headers)
        lines.append(f"| {cells} |")
    return lines
