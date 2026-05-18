"""
Export pipeline — drives Python report generation from a snapshot.
Equivalent of Invoke-M365TenantAssessmentExportPipeline.
"""

from __future__ import annotations

import logging
from pathlib import Path

from . import snapshot as snap
from .reporting import excel, html, docx, markdown, improvement_plan
from .utils.paths import resolve_snapshot_output_context

log = logging.getLogger(__name__)


def run(
    snapshot_path: str | Path,
    output_dir: str | Path | None = None,
    profiles: list[str] | None = None,
    skip_excel: bool = False,
    skip_html: bool = False,
    skip_docx: bool = False,
    skip_improve: bool = False,
    docx_template: str | Path | None = None,
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
    deliverables.mkdir(parents=True, exist_ok=True)

    artifacts: dict[str, Path] = {}

    if not skip_excel:
        xlsx_path = deliverables / f"{stem}-Tenant Details.xlsx"
        log.info("Generating Excel workbook: %s", xlsx_path)
        _flatten_and_export_excel(snapshot_data, xlsx_path)
        artifacts["excel"] = xlsx_path

    if not skip_html:
        html_path = deliverables / f"{stem}-Report.html"
        log.info("Generating HTML report: %s", html_path)
        html.generate(snapshot_data, html_path)
        artifacts["html"] = html_path

    if not skip_docx:
        docx_path = deliverables / f"{stem}-Assessment.docx"
        log.info("Generating Word document: %s", docx_path)
        docx.generate(snapshot_data, docx_path, template_path=docx_template)
        artifacts["docx"] = docx_path

    if not skip_improve:
        improve_dir = out_dir / "ImprovementPlan"
        log.info("Generating improvement plan: %s", improve_dir)
        improvement_plan.generate(snapshot_data, improve_dir)
        artifacts["improve_dir"] = improve_dir

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
