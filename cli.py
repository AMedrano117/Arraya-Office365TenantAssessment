"""
M365 Assessment CLI — Python replacement for Start-M365TenantAssessment.ps1.

Actions:
  full       Full assessment (collect via PowerShell + report via Python)
  preflight  Test M365 connections only (PowerShell)
  collect    Collect data to a JSON snapshot (PowerShell)
  report     Generate reports from an existing snapshot (Python)
  improve    Build improvement plan from a snapshot (Python)
  compare    Compare two snapshots (Python)
  ad         Active Directory assessment (PowerShell)

Usage:
  python cli.py --help
  python cli.py collect --export-path ./output --profile SolutionsEngineer
  python cli.py report --snapshot ./output/snapshot.json
"""

from __future__ import annotations

import sys
from pathlib import Path

import click
from rich.console import Console
from rich.panel import Panel

# Ensure src/python is importable when running cli.py from repo root
sys.path.insert(0, str(Path(__file__).parent))

from src.python import pipeline, runner
from src.python.config import get_output_root, parse_profiles, VALID_PROFILES
from src.python.utils.logging import configure_root

console = Console()

_PROFILES = sorted(VALID_PROFILES)

_AUTH_MODES = ["Interactive", "Certificate", "ClientSecret"]


# ---------------------------------------------------------------------------
# Root group
# ---------------------------------------------------------------------------

@click.group(invoke_without_command=True, context_settings={"help_option_names": ["-h", "--help"]})
@click.option("--verbose", "-v", is_flag=True, default=False, help="Show detailed progress and diagnostic output.")
@click.pass_context
def main(ctx: click.Context, verbose: bool) -> None:
    """Arraya M365 Tenant Assessment - cross-platform Python CLI."""
    ctx.ensure_object(dict)
    ctx.obj["verbose"] = verbose
    configure_root(verbose=verbose)
    if ctx.invoked_subcommand is None:
        _interactive_menu(ctx)


# ---------------------------------------------------------------------------
# Shared options factory
# ---------------------------------------------------------------------------

def _auth_options(f):
    # Applied in reverse order — Click reverses params, so last applied appears first in --help.
    f = click.option("--skip-preflight", is_flag=True, help="Skip the permission preflight check.")(f)
    f = click.option("--skip-auth", is_flag=True, help="Skip authentication; reuse an existing session.")(f)
    f = click.option("--client-secret", default="", help="Client secret (ClientSecret auth mode).")(f)
    f = click.option("--cert-thumbprint", default="", help="Certificate thumbprint (Certificate auth mode).")(f)
    f = click.option("--client-id", default="", help="App registration client ID.")(f)
    f = click.option("--tenant-id", default="", help="Entra tenant ID (GUID).")(f)
    f = click.option(
        "--auth-mode",
        type=click.Choice(_AUTH_MODES, case_sensitive=False),
        default="Interactive",
        show_default=True,
        metavar="[Interactive|Certificate|ClientSecret]",
        help="Authentication mode.",
    )(f)
    return f


def _profile_option(f):
    return click.option(
        "--profile", "-p",
        multiple=True,
        default=("SolutionsEngineer",),
        show_default=True,
        type=click.Choice(_PROFILES, case_sensitive=False),
        metavar="PROFILE",
        help="Output profile(s). Repeat to select multiple.  Choices: " + " | ".join(sorted(_PROFILES)),
    )(f)


def _export_option(f):
    return click.option(
        "--export-path", "-o",
        default=None,
        type=click.Path(),
        help="Output folder or .xlsx path. Defaults to output/ in repo root.",
    )(f)


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

@main.command()
@_auth_options
@_profile_option
@_export_option
@click.option("--skip-improve", is_flag=True, help="Skip improvement plan generation.")
@click.option("--use-graph-fallback", is_flag=True, help="Use Graph API fallback instead of live connections.")
@click.pass_context
def full(ctx, auth_mode, tenant_id, client_id, cert_thumbprint, client_secret,
         skip_auth, skip_preflight, profile, export_path, skip_improve, use_graph_fallback):
    """Full assessment: collect via PowerShell + report via Python."""
    verbose = (ctx.obj or {}).get("verbose", False)
    _print_banner("Full Assessment")
    profiles = parse_profiles(list(profile))
    out = Path(export_path) if export_path else get_output_root()

    console.print("  [dim]Phase 1 of 2[/dim]  Collect  [dim](PowerShell → M365 APIs)[/dim]")
    console.print("  [dim]Phase 2 of 2[/dim]  Report   [dim](Python   → Excel, HTML, Word"
                  + (", Improvement Plan" if not skip_improve else "") + ")[/dim]")
    console.print()
    console.print(f"  [dim]Export path :[/dim] {out}")
    console.print(f"  [dim]Profile(s)  :[/dim] {', '.join(profiles)}")
    console.print(f"  [dim]Auth mode   :[/dim] {auth_mode or 'Interactive'}")
    console.print(f"  [dim]Improve     :[/dim] {'No (skipped)' if skip_improve else 'Yes'}")
    console.print()

    rc = runner.run_collection(
        export_path=out, output_profiles=profiles,
        auth_mode=auth_mode, tenant_id=tenant_id, client_id=client_id,
        certificate_thumbprint=cert_thumbprint, client_secret=client_secret,
        skip_auth=skip_auth, skip_preflight=skip_preflight,
        use_graph_fallback=use_graph_fallback,
    )
    if rc != 0:
        console.print(f"[red]Collection failed (exit {rc}).[/red]")
        sys.exit(rc)

    snapshot_path = _find_latest_snapshot(out)
    if snapshot_path:
        console.print(f"\n[green]Collection complete.[/green] Generating reports...")
        pipeline.run(snapshot_path, output_dir=out, profiles=profiles, skip_improve=skip_improve, verbose=verbose)
        console.print(f"[green]Done.[/green] Outputs: {out}")
    else:
        console.print(f"[yellow]Collection complete.[/yellow] No snapshot was found under {out}; reports were not generated.")
        if not verbose:
            console.print("  [dim]Run with --verbose for diagnostic details.[/dim]")


@main.command()
@_auth_options
@_profile_option
@_export_option
def preflight(auth_mode, tenant_id, client_id, cert_thumbprint, client_secret,
              skip_auth, skip_preflight, profile, export_path):
    """Test M365 connections and permissions (PowerShell)."""
    _print_banner("Preflight")
    profiles = parse_profiles(list(profile))
    rc = runner.run_preflight(
        output_profiles=profiles, auth_mode=auth_mode, tenant_id=tenant_id,
        client_id=client_id, certificate_thumbprint=cert_thumbprint,
        client_secret=client_secret, skip_auth=skip_auth,
    )
    sys.exit(rc)


@main.command()
@_auth_options
@_profile_option
@_export_option
@click.option("--use-graph-fallback", is_flag=True, help="Use Graph API fallback instead of live module connections.")
def collect(auth_mode, tenant_id, client_id, cert_thumbprint, client_secret,
            skip_auth, skip_preflight, profile, export_path, use_graph_fallback):
    """Collect M365 data to a JSON snapshot (PowerShell)."""
    _print_banner("Data Collection")
    profiles = parse_profiles(list(profile))
    out = Path(export_path) if export_path else get_output_root()
    rc = runner.run_collection(
        export_path=out, output_profiles=profiles,
        auth_mode=auth_mode, tenant_id=tenant_id, client_id=client_id,
        certificate_thumbprint=cert_thumbprint, client_secret=client_secret,
        skip_auth=skip_auth, skip_preflight=skip_preflight,
        use_graph_fallback=use_graph_fallback,
    )
    sys.exit(rc)


@main.command()
@click.argument("snapshot", type=click.Path(exists=True), metavar="SNAPSHOT")
@_profile_option
@_export_option
@click.option("--skip-improve", is_flag=True, help="Skip improvement plan generation.")
@click.option("--skip-excel", is_flag=True, help="Skip Excel workbook output.")
@click.option("--skip-html", is_flag=True, help="Skip HTML report output.")
@click.option("--skip-docx", is_flag=True, help="Skip Word document output.")
@click.option("--docx-template", default=None, type=click.Path(), help="Path to a custom .docx template.")
@click.pass_context
def report(ctx, snapshot, profile, export_path, skip_improve, skip_excel, skip_html, skip_docx, docx_template):
    """Generate Excel, HTML, Word reports from an existing snapshot (Python)."""
    verbose = (ctx.obj or {}).get("verbose", False)
    _print_banner("Report Generation")
    snap_path = Path(snapshot)
    out = Path(export_path) if export_path else snap_path.parent
    profiles = parse_profiles(list(profile))
    console.print(f"  [dim]Snapshot:[/dim]  {snap_path.name}")
    console.print(f"  [dim]Output:[/dim]    {out}")
    console.print()
    artifacts = pipeline.run(
        snap_path, output_dir=out, profiles=profiles,
        skip_excel=skip_excel, skip_html=skip_html, skip_docx=skip_docx,
        skip_improve=skip_improve,
        docx_template=Path(docx_template) if docx_template else None,
        verbose=verbose,
    )
    console.print(f"[green]Done.[/green] {len(artifacts)} artifact(s) written to: {out}")


@main.command()
@click.argument("snapshot", type=click.Path(exists=True), metavar="SNAPSHOT")
@click.option("--output-dir", "-o", default=None, type=click.Path(), help="Directory to write improvement plan files. Defaults to <snapshot-dir>/ImprovementPlan.")
def improve(snapshot, output_dir):
    """Build improvement plan from a snapshot (Python)."""
    _print_banner("Improvement Plan")
    from src.python.reporting import improvement_plan
    from src.python import snapshot as snap
    snap_path = Path(snapshot)
    out = Path(output_dir) if output_dir else snap_path.parent / "ImprovementPlan"
    data = snap.load(snap_path)
    improvement_plan.generate(data, out)
    console.print(f"[green]Improvement plan written to:[/green] {out}")


@main.command()
@click.argument("baseline", type=click.Path(exists=True), metavar="BASELINE")
@click.argument("current", type=click.Path(exists=True), metavar="CURRENT")
@click.option("--output-dir", "-o", default=None, type=click.Path(), help="Directory to write the delta report. Defaults to the current snapshot's folder.")
def compare(baseline, current, output_dir):
    """Compare two snapshots and write a delta report (Python)."""
    _print_banner("Snapshot Comparison")
    import json
    from src.python import snapshot as snap
    from src.python.analysis.comparison import compare as do_compare

    base_data = snap.load(Path(baseline))
    curr_data = snap.load(Path(current))
    delta = do_compare(base_data, curr_data)

    out = Path(output_dir) if output_dir else Path(current).parent
    out.mkdir(parents=True, exist_ok=True)
    delta_path = out / "comparison-delta.json"
    delta_path.write_text(json.dumps(delta, indent=2, default=str), encoding="utf-8")
    console.print(f"[green]Delta report:[/green] {delta_path}")
    console.print(
        f"Added: {delta['Summary']['Added']}  "
        f"Removed: {delta['Summary']['Removed']}  "
        f"Changed: {delta['Summary']['Changed']}"
    )


@main.command()
def ad():
    """Active Directory assessment (PowerShell)."""
    _print_banner("Active Directory Assessment")
    rc = runner.run_ad_assessment()
    sys.exit(rc)


# ---------------------------------------------------------------------------
# Interactive menu (when invoked with no subcommand)
# ---------------------------------------------------------------------------

def _interactive_menu(ctx: click.Context) -> None:
    console.print(Panel.fit(
        "[bold cyan]Arraya M365 Tenant Assessment Launcher[/bold cyan]",
        border_style="cyan",
    ))

    menu = {
        "1": ("full",      "Full assessment + improvement plan"),
        "2": ("preflight", "Test M365 connections only"),
        "3": ("collect",   "Collect M365 data to a snapshot"),
        "4": ("report",    "Generate reports from snapshot"),
        "5": ("improve",   "Build improvement plan from snapshot"),
        "6": ("compare",   "Compare two snapshots"),
        "7": ("ad",        "Active Directory assessment"),
    }

    console.print("\n  [dim]── Live ──────────────────────────────────[/dim]")
    for k in ("1", "2", "3"):
        console.print(f"  [cyan]{k}[/cyan]  {menu[k][0]:<12} {menu[k][1]}")
    console.print("\n  [dim]── From Snapshot ──────────────────────────[/dim]")
    for k in ("4", "5", "6"):
        console.print(f"  [cyan]{k}[/cyan]  {menu[k][0]:<12} {menu[k][1]}")
    console.print("\n  [dim]── Other ──────────────────────────────────[/dim]")
    console.print(f"  [cyan]7[/cyan]  {menu['7'][0]:<12} {menu['7'][1]}")
    console.print()

    while True:
        choice = click.prompt("Select option (1-7)", default="1")
        if choice in menu:
            break
        console.print("[red]Invalid choice.[/red]")

    action_name = menu[choice][0]
    console.print(f"\n  [dim]Run:[/dim]  python cli.py {action_name} --help\n")
    ctx.invoke(main.commands[action_name])


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _print_banner(action: str) -> None:
    console.print(Panel.fit(
        f"[bold cyan]M365 Assessment  -  {action}[/bold cyan]",
        border_style="cyan",
    ))


def _find_latest_snapshot(base_dir: Path) -> Path | None:
    # The PS pipeline writes snapshots as *-Snap.json; older patterns kept for compatibility.
    for pattern in ("*-Snap.json", "*AssessmentSnapshot*.json", "*.snapshot.json"):
        candidates = sorted(base_dir.rglob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
        # Exclude manifest, plan, evidence-coverage, and other side-car files.
        candidates = [
            p for p in candidates
            if not any(tok in p.stem for tok in ("manifest", "Plan", "Coverage", "EvidenceCoverage"))
        ]
        if candidates:
            return candidates[0]
    return None


if __name__ == "__main__":
    main()
