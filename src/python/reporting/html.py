"""
HTML report generation - tenant assessment dashboard.
Produces a clean overview with stats, workstream cards, and findings.
No raw data table dumps.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from jinja2 import Environment, BaseLoader

from ..snapshot import get_data, get_metadata
from ..utils.converters import to_export_friendly_value

log = logging.getLogger(__name__)

_PHASE_ORDER = {"Immediate": 0, "Near Term": 1, "Planned": 2, "Monitor": 3}
_SEV_ORDER   = {"High": 0, "Medium": 1, "Low": 2}

_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>M365 Assessment &mdash; {{ tenant_name }}</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:#F0F2F7;color:#222;font-size:14px}

/* ---- Brand bar ---- */
.top-bar{height:5px;background:linear-gradient(90deg,#1F3864 60%,#C55A11 100%)}
header{background:#1F3864;color:#fff;padding:28px 48px 22px}
.brand-line{font-size:11px;text-transform:uppercase;letter-spacing:.12em;opacity:.6;margin-bottom:6px}
header h1{font-size:26px;font-weight:700;letter-spacing:-.01em;margin-bottom:4px}
.header-meta{font-size:12px;opacity:.65;margin-top:2px}
.accent-bar{height:4px;background:#C55A11}

/* ---- Layout ---- */
main{max-width:1340px;margin:32px auto;padding:0 40px 60px}
.section-label{font-size:11px;text-transform:uppercase;letter-spacing:.1em;color:#C55A11;font-weight:700;
               margin:32px 0 12px;padding-bottom:6px;border-bottom:2px solid #E8ECF4}

/* ---- Stat cards ---- */
.stats-grid{display:grid;grid-template-columns:repeat(6,1fr);gap:14px;margin-bottom:6px}
.stat-card{background:#fff;border-radius:10px;padding:18px 20px 14px;
           box-shadow:0 1px 5px rgba(0,0,0,.07);border-top:3px solid #1F3864;position:relative}
.stat-card.accent{border-top-color:#C55A11}
.stat-card .val{font-size:28px;font-weight:700;color:#1F3864;line-height:1}
.stat-card.accent .val{color:#C55A11}
.stat-card .lbl{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#888;margin-top:5px}

/* ---- Phase breakdown ---- */
.phase-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:6px}
.phase-card{background:#fff;border-radius:10px;padding:16px 20px;box-shadow:0 1px 5px rgba(0,0,0,.07)}
.phase-card .p-lbl{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#888;margin-bottom:6px}
.phase-card .p-count{font-size:32px;font-weight:700;line-height:1}
.phase-card .p-sub{font-size:11px;color:#999;margin-top:4px}
.ph-immediate{border-left:4px solid #C00000}.ph-immediate .p-count{color:#C00000}
.ph-nearterm{border-left:4px solid #C55A11}.ph-nearterm .p-count{color:#C55A11}
.ph-planned{border-left:4px solid #B8860B}.ph-planned .p-count{color:#B8860B}
.ph-monitor{border-left:4px solid #4A7C3F}.ph-monitor .p-count{color:#4A7C3F}

/* ---- Workstream cards ---- */
.ws-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:6px}
.ws-card{background:#fff;border-radius:10px;padding:18px 20px;box-shadow:0 1px 5px rgba(0,0,0,.07)}
.ws-top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px}
.ws-name{font-weight:700;font-size:15px;color:#1F3864}
.sev-badge{font-size:10px;font-weight:700;padding:3px 9px;border-radius:20px;text-transform:uppercase;letter-spacing:.04em;white-space:nowrap}
.sev-High{background:#FDECEA;color:#C00000;border:1px solid #F5C6C6}
.sev-Medium{background:#FFF0E0;color:#C55A11;border:1px solid #F5CCA0}
.sev-Low{background:#EEF7EB;color:#4A7C3F;border:1px solid #BFE0B5}
.ws-counts{font-size:12px;color:#777;margin-bottom:10px}
.ws-top-issue{font-size:12px;color:#333;background:#F5F7FC;
              border-left:3px solid #C55A11;padding:7px 10px;border-radius:0 6px 6px 0;
              line-height:1.45}

/* ---- Tables ---- */
.card-table{background:#fff;border-radius:10px;box-shadow:0 1px 5px rgba(0,0,0,.07);
            overflow:hidden;margin-bottom:6px}
.card-table-header{background:#1F3864;color:#fff;padding:12px 20px;font-size:12px;
                   font-weight:700;text-transform:uppercase;letter-spacing:.07em;
                   display:flex;justify-content:space-between;align-items:center}
.card-table-header span{opacity:.65;font-weight:400;text-transform:none;letter-spacing:0;font-size:11px}
table{width:100%;border-collapse:collapse;font-size:13px}
thead th{background:#EEF2FA;padding:9px 16px;text-align:left;font-weight:600;
         font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#555;
         border-bottom:2px solid #D9E1F2;white-space:nowrap}
tbody tr:hover{background:#F5F7FD}
tbody tr:nth-child(even){background:#FAFBFE}
td{padding:9px 16px;border-bottom:1px solid #EDF0F8;vertical-align:top;line-height:1.45}
td.finding-text{max-width:360px}
td.remediation-text{max-width:340px;color:#444}

/* ---- Severity dots & badges ---- */
.dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:5px;flex-shrink:0;vertical-align:middle}
.dot-High{background:#C00000}.dot-Medium{background:#C55A11}.dot-Low{background:#4A7C3F}
.sev-cell{display:flex;align-items:center;white-space:nowrap;font-weight:600}

/* ---- Phase pill ---- */
.pill{font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px;text-transform:uppercase;letter-spacing:.04em;white-space:nowrap}
.pill-Immediate{background:#FDECEA;color:#C00000}
.pill-NearTerm{background:#FFF0E0;color:#C55A11}
.pill-Planned{background:#FFF9E0;color:#B8860B}
.pill-Monitor{background:#EEF7EB;color:#4A7C3F}

/* ---- Footer ---- */
footer{text-align:center;color:#aaa;font-size:11px;padding:32px;border-top:1px solid #E8ECF4;margin-top:20px}
</style>
</head>
<body>
<div class="top-bar"></div>
<header>
  <div class="brand-line">Arraya Solutions &bull; Microsoft 365 Assessment</div>
  <h1>{{ tenant_name }}</h1>
  <div class="header-meta">Generated {{ generated_at }} &nbsp;&bull;&nbsp; {{ finding_count }} findings across {{ ws_count }} workstream(s)</div>
</header>
<div class="accent-bar"></div>
<main>

<!-- Tenant Stats -->
<div class="section-label">Tenant Overview</div>
<div class="stats-grid">
  <div class="stat-card"><div class="val">{{ stats.users }}</div><div class="lbl">Users</div></div>
  <div class="stat-card"><div class="val">{{ stats.admins }}</div><div class="lbl">Admins</div></div>
  <div class="stat-card"><div class="val">{{ stats.domains }}</div><div class="lbl">Domains</div></div>
  <div class="stat-card"><div class="val">{{ stats.devices }}</div><div class="lbl">Devices</div></div>
  <div class="stat-card"><div class="val">{{ stats.licenses }}</div><div class="lbl">License SKUs</div></div>
  <div class="stat-card accent"><div class="val">{{ stats.secure_score }}%</div><div class="lbl">Secure Score</div></div>
</div>

<!-- Phase Breakdown -->
<div class="section-label">Findings by Phase</div>
<div class="phase-grid">
  <div class="phase-card ph-immediate">
    <div class="p-lbl">Immediate</div>
    <div class="p-count">{{ phases.Immediate }}</div>
    <div class="p-sub">Action required now</div>
  </div>
  <div class="phase-card ph-nearterm">
    <div class="p-lbl">Near Term</div>
    <div class="p-count">{{ phases.near_term }}</div>
    <div class="p-sub">Within 30&ndash;60 days</div>
  </div>
  <div class="phase-card ph-planned">
    <div class="p-lbl">Planned</div>
    <div class="p-count">{{ phases.Planned }}</div>
    <div class="p-sub">Roadmap items</div>
  </div>
  <div class="phase-card ph-monitor">
    <div class="p-lbl">Monitor</div>
    <div class="p-count">{{ phases.Monitor }}</div>
    <div class="p-sub">Track over time</div>
  </div>
</div>

<!-- Workstream Cards -->
<div class="section-label">Workstream Summary</div>
<div class="ws-grid">
  {% for ws in workstreams %}
  <div class="ws-card">
    <div class="ws-top">
      <span class="ws-name">{{ ws.name }}</span>
      <span class="sev-badge sev-{{ ws.sev }}">{{ ws.sev }}</span>
    </div>
    <div class="ws-counts">{{ ws.total }} finding(s) &nbsp;&middot;&nbsp; <strong>{{ ws.high }}</strong> critical &middot; {{ ws.medium }} medium &middot; {{ ws.low }} low</div>
    {% if ws.top_issue %}
    <div class="ws-top-issue">{{ ws.top_issue }}</div>
    {% endif %}
  </div>
  {% endfor %}
</div>

<!-- Priority Findings -->
<div class="section-label">Priority Findings &mdash; Immediate &amp; Near Term</div>
<div class="card-table">
  <div class="card-table-header">Action Required <span>{{ priority_findings | length }} finding(s)</span></div>
  <table>
    <thead><tr>
      <th>Severity</th><th>Phase</th><th>Workstream</th><th>Area</th>
      <th class="finding-text">Finding</th><th class="remediation-text">Recommended Action</th>
    </tr></thead>
    <tbody>
    {% for f in priority_findings %}
    <tr>
      <td><div class="sev-cell"><span class="dot dot-{{ f.sev }}"></span>{{ f.sev }}</div></td>
      <td><span class="pill pill-{{ f.phase_css }}">{{ f.phase }}</span></td>
      <td>{{ f.workstream }}</td>
      <td>{{ f.area }}</td>
      <td class="finding-text">{{ f.finding }}</td>
      <td class="remediation-text">{{ f.remediation }}</td>
    </tr>
    {% endfor %}
    </tbody>
  </table>
</div>

<!-- All Findings -->
<div class="section-label">All Findings</div>
<div class="card-table">
  <div class="card-table-header">Complete Finding List <span>{{ all_findings | length }} total</span></div>
  <table>
    <thead><tr>
      <th>Severity</th><th>Phase</th><th>Workstream</th><th>Area</th><th>Finding</th>
    </tr></thead>
    <tbody>
    {% for f in all_findings %}
    <tr>
      <td><div class="sev-cell"><span class="dot dot-{{ f.sev }}"></span>{{ f.sev }}</div></td>
      <td><span class="pill pill-{{ f.phase_css }}">{{ f.phase }}</span></td>
      <td>{{ f.workstream }}</td>
      <td>{{ f.area }}</td>
      <td>{{ f.finding }}</td>
    </tr>
    {% endfor %}
    </tbody>
  </table>
</div>

</main>
<footer>Arraya Solutions &bull; Microsoft 365 Tenant Assessment &bull; Confidential &bull; {{ generated_at }}</footer>
</body>
</html>"""


def generate(
    snapshot: dict,
    output_path: str | Path,
    plan: dict | None = None,
) -> None:
    output_path = Path(output_path)
    if output_path.suffix.lower() != ".html":
        output_path = output_path.with_suffix(".html")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    meta = get_metadata(snapshot)
    tenant_obj = meta.get("Tenant") or {}
    tenant_name = (
        (tenant_obj.get("DisplayName") if isinstance(tenant_obj, dict) else None)
        or meta.get("TenantDisplayName")
        or meta.get("TenantDomain")
        or "Unknown Tenant"
    )
    generated_at = meta.get("GeneratedAt", "")
    try:
        from datetime import datetime
        dt = datetime.fromisoformat(str(generated_at).replace("Z", "+00:00"))
        generated_at = dt.strftime("%B %d, %Y")
    except (ValueError, AttributeError):
        pass

    # Tenant stats from snapshot
    data = get_data(snapshot)
    stats = _collect_stats(snapshot, data)

    # Findings from plan
    findings: list[dict] = (plan or {}).get("Findings", [])
    findings_sorted = sorted(
        findings,
        key=lambda f: (
            _SEV_ORDER.get(f.get("Severity", "Low"), 99),
            _PHASE_ORDER.get(f.get("RoadmapPhase", "Monitor"), 99),
        ),
    )

    phase_counts = {p: sum(1 for f in findings if f.get("RoadmapPhase") == p)
                   for p in ("Immediate", "Near Term", "Planned", "Monitor")}

    # Workstream summary cards
    ws_map: dict[str, list[dict]] = {}
    for f in findings:
        ws_map.setdefault(f.get("Workstream", "Governance"), []).append(f)

    workstreams = []
    for ws_name in sorted(ws_map, key=lambda w: min(
        _SEV_ORDER.get(f.get("Severity", "Low"), 99) for f in ws_map[w]
    )):
        group = ws_map[ws_name]
        high_sev = min(group, key=lambda f: _SEV_ORDER.get(f.get("Severity", "Low"), 99))
        top_issue_f = next(
            (f for f in group if f.get("Severity") == "High"),
            group[0],
        )
        top_issue = str(top_issue_f.get("Finding") or "")
        if len(top_issue) > 120:
            top_issue = top_issue[:117] + "..."
        workstreams.append({
            "name":    ws_name,
            "sev":     high_sev.get("Severity", "Low"),
            "total":   len(group),
            "high":    sum(1 for f in group if f.get("Severity") == "High"),
            "medium":  sum(1 for f in group if f.get("Severity") == "Medium"),
            "low":     sum(1 for f in group if f.get("Severity") == "Low"),
            "top_issue": top_issue,
        })

    def _row(f: dict) -> dict:
        phase = f.get("RoadmapPhase", "Monitor")
        finding_text = str(f.get("Finding") or "")
        remed = str(f.get("TechnicalRemediation") or f.get("Recommendation") or "")
        return {
            "sev":        f.get("Severity", "Low"),
            "phase":      phase,
            "phase_css":  phase.replace(" ", ""),
            "workstream": f.get("Workstream", ""),
            "area":       f.get("Area", ""),
            "finding":    finding_text[:200] + ("..." if len(finding_text) > 200 else ""),
            "remediation": remed[:180] + ("..." if len(remed) > 180 else ""),
        }

    priority_findings = [_row(f) for f in findings_sorted
                         if f.get("RoadmapPhase") in ("Immediate", "Near Term")]
    all_findings_rows  = [_row(f) for f in findings_sorted]

    env  = Environment(loader=BaseLoader())
    tmpl = env.from_string(_TEMPLATE)
    html = tmpl.render(
        tenant_name      = tenant_name,
        generated_at     = generated_at,
        finding_count    = len(findings),
        ws_count         = len(ws_map),
        stats            = stats,
        phases           = {
            "Immediate": phase_counts["Immediate"],
            "near_term": phase_counts["Near Term"],
            "Planned":   phase_counts["Planned"],
            "Monitor":   phase_counts["Monitor"],
        },
        workstreams      = workstreams,
        priority_findings= priority_findings,
        all_findings     = all_findings_rows,
    )

    output_path.write_text(html, encoding="utf-8")
    log.info("HTML dashboard saved: %s (%d findings)", output_path, len(findings))


def _collect_stats(snapshot: dict, data: dict) -> dict[str, Any]:
    def _count(section: str, key: str) -> int:
        val = (data.get(section) or {}).get(key)
        if isinstance(val, list):
            return len(val)
        if isinstance(val, dict):
            return len(val)
        return 0 if val is None else 1

    # Secure Score
    ss_raw = ((data.get("Security") or {}).get("SecuritySecureScore") or {})
    if isinstance(ss_raw, dict):
        score = ss_raw.get("CurrentScore") or ss_raw.get("SecureScore") or 0
        max_score = ss_raw.get("MaxScore") or 100
        pct = round(float(score) / float(max_score) * 100, 1) if max_score else 0
    else:
        pct = 0

    return {
        "users":        _count("Identity", "Users"),
        "admins":       _count("Identity", "Admins"),
        "domains":      _count("Tenant", "Domains"),
        "devices":      _count("Identity", "DeviceDetails"),
        "licenses":     _count("Identity", "LicenseSKUs"),
        "secure_score": pct,
    }
