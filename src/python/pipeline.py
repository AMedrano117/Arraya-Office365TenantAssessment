"""
Export pipeline — drives Python report generation from a snapshot.
Equivalent of Invoke-M365TenantAssessmentExportPipeline.
"""

from __future__ import annotations

import logging
from pathlib import Path

from rich.console import Console

from . import snapshot as snap
from .analysis import plan as plan_module, evidence_coverage as coverage_module
from .reporting import excel, html, docx, markdown, improvement_plan, engpack, roadmap
from .utils.paths import resolve_snapshot_output_context

log = logging.getLogger(__name__)
_console = Console()


def run(
    snapshot_path: str | Path,
    output_dir: str | Path | None = None,
    profiles: list[str] | None = None,
    skip_excel: bool = False,
    skip_html: bool = False,
    skip_docx: bool = False,
    skip_improve: bool = False,
    docx_template: str | Path | None = None,
    verbose: bool = False,
) -> dict[str, Path]:
    snapshot_path = Path(snapshot_path)
    snapshot_data = snap.load(snapshot_path)
    if snapshot_data is None:
        raise FileNotFoundError(f"Snapshot not found: {snapshot_path}")

    ctx = resolve_snapshot_output_context(
        snapshot_data,
        output_dir or snapshot_path.parent,
    )
    out_dir: Path = ctx["OutputDirectory"]
    stem: str = ctx["FileStem"]
    deliverables = out_dir / "Deliverables"
    support = out_dir / "Support"
    deliverables.mkdir(parents=True, exist_ok=True)
    support.mkdir(parents=True, exist_ok=True)

    artifacts: dict[str, Path] = {}

    # -----------------------------------------------------------------------
    # Analysis: generate enriched Plan.json (always — other steps depend on it)
    # -----------------------------------------------------------------------
    log.info("Generating plan: %s", support)
    generated_plan = plan_module.generate(snapshot_data, support, tenant_name=stem)
    artifacts["plan"] = support / f"{stem}-Plan.json"
    if verbose:
        _console.print(f"  [dim]Plan[/dim]      {artifacts['plan'].name}")

    # -----------------------------------------------------------------------
    # Reports
    # -----------------------------------------------------------------------
    if not skip_excel:
        xlsx_path = deliverables / f"{stem}-Tenant Details.xlsx"
        log.info("Generating Excel workbook: %s", xlsx_path)
        _flatten_and_export_excel(snapshot_data, xlsx_path)
        artifacts["excel"] = xlsx_path
        if verbose:
            _console.print(f"  [dim]Excel[/dim]     {xlsx_path.name}")

    if not skip_html:
        html_path = deliverables / f"{stem}-Report.html"
        log.info("Generating HTML report: %s", html_path)
        html.generate(snapshot_data, html_path, plan=generated_plan)
        artifacts["html"] = html_path
        if verbose:
            _console.print(f"  [dim]HTML[/dim]      {html_path.name}")

    if not skip_docx:
        docx_path = deliverables / f"{stem}-Assessment.docx"
        log.info("Generating Word document: %s", docx_path)
        docx.generate(snapshot_data, docx_path, template_path=docx_template, plan=generated_plan)
        artifacts["docx"] = docx_path
        if verbose:
            _console.print(f"  [dim]Word[/dim]      {docx_path.name}")

    if not skip_improve:
        # Evidence coverage JSON (support artifact for SE review)
        coverage_data = coverage_module.generate(
            snapshot_data,
            support,
            tenant_name=stem,
            snapshot_path=str(snapshot_path),
        )
        coverage_path = support / f"{stem}-SolutionsEngineerEvidenceCoverage.json"
        artifacts["coverage"] = coverage_path
        if verbose:
            _console.print(f"  [dim]Coverage[/dim]  {coverage_path.name}")

        # EngPack.docx (engineer-facing deliverable)
        engpack_path = deliverables / f"{stem}-EngPack.docx"
        log.info("Generating EngPack: %s", engpack_path)
        engpack.generate(
            generated_plan,
            snapshot_data,
            engpack_path,
            snapshot_path=str(snapshot_path),
            support_dir=support,
            coverage=coverage_data,
        )
        artifacts["engpack"] = engpack_path
        if verbose:
            _console.print(f"  [dim]EngPack[/dim]   {engpack_path.name}")

        # Remediation Roadmap Word doc
        from datetime import date as _date
        roadmap_date = _date.today().strftime("%Y-%m-%d")
        roadmap_path = deliverables / f"{stem}-Microsoft 365 Remediation Roadmap-{roadmap_date}.docx"
        log.info("Generating roadmap: %s", roadmap_path)
        roadmap.generate(generated_plan, snapshot_data, roadmap_path)
        artifacts["roadmap"] = roadmap_path
        if verbose:
            _console.print(f"  [dim]Roadmap[/dim]   {roadmap_path.name}")

        # Basic improvement plan markdown (lightweight summary)
        improve_dir = out_dir / "ImprovementPlan"
        log.info("Generating improvement plan: %s", improve_dir)
        improvement_plan.generate(snapshot_data, improve_dir)
        artifacts["improve_dir"] = improve_dir
        if verbose:
            _console.print(f"  [dim]Improve[/dim]   {improve_dir.name}/")

    log.info("Pipeline complete. Artifacts: %s", list(artifacts.values()))
    return artifacts


def _flatten_and_export_excel(snapshot_data: dict, xlsx_path: Path) -> None:
    flat: dict = {}
    data = snapshot_data.get("Data") or {}
    for section in data.values():
        if isinstance(section, dict):
            flat.update(section)
    derived = snapshot_data.get("Derived") or {}
    flat.update(derived)

    excel.export(flat, xlsx_path)
