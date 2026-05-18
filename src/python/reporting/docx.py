"""
Word document generation — Python equivalent of CustomerAssessmentDocx.ps1.
Uses python-docx to fill a template or generate from scratch.
"""

from __future__ import annotations

import logging
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from ..snapshot import get_data, get_metadata
from ..utils.converters import ensure_list

log = logging.getLogger(__name__)

_NAVY = RGBColor(0x1F, 0x38, 0x64)


def generate(
    snapshot: dict,
    output_path: str | Path,
    template_path: str | Path | None = None,
) -> None:
    output_path = Path(output_path)
    if output_path.suffix.lower() != ".docx":
        output_path = output_path.with_suffix(".docx")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if template_path and Path(template_path).exists():
        doc = Document(str(template_path))
    else:
        doc = Document()
        _setup_styles(doc)

    _write_cover(doc, snapshot)
    _write_data_sections(doc, snapshot)

    doc.save(str(output_path))
    log.info("Word document saved: %s", output_path)


# ---------------------------------------------------------------------------
# Document builders
# ---------------------------------------------------------------------------

def _setup_styles(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)


def _write_cover(doc: Document, snapshot: dict) -> None:
    meta = get_metadata(snapshot)
    tenant = meta.get("TenantDisplayName") or meta.get("TenantDomain") or "Tenant"
    generated_at = meta.get("GeneratedAt", "")

    p = doc.add_heading("Microsoft 365 Tenant Assessment", level=0)
    p.runs[0].font.color.rgb = _NAVY

    info = doc.add_paragraph()
    info.add_run(f"Tenant: {tenant}").bold = True
    doc.add_paragraph(f"Generated: {generated_at}")
    doc.add_page_break()


def _write_data_sections(doc: Document, snapshot: dict) -> None:
    data = get_data(snapshot)
    for section_name, tables in data.items():
        if not isinstance(tables, dict) or not tables:
            continue

        doc.add_heading(section_name, level=1)

        for table_name, rows in tables.items():
            rows_list = ensure_list(rows)
            if not rows_list:
                continue

            doc.add_heading(table_name, level=2)

            if isinstance(rows_list[0], dict):
                headers = list(rows_list[0].keys())
                table = doc.add_table(rows=1, cols=len(headers))
                table.style = "Table Grid"

                hdr_row = table.rows[0]
                for idx, header in enumerate(headers):
                    cell = hdr_row.cells[idx]
                    cell.text = header
                    run = cell.paragraphs[0].runs[0]
                    run.bold = True
                    run.font.color.rgb = _NAVY

                for record in rows_list[:500]:
                    if not isinstance(record, dict):
                        continue
                    row_cells = table.add_row().cells
                    for idx, header in enumerate(headers):
                        val = record.get(header)
                        row_cells[idx].text = str(val) if val is not None else ""
            else:
                for item in rows_list[:500]:
                    doc.add_paragraph(str(item), style="List Bullet")
