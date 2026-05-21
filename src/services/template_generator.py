from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from config import APP_CONFIG


def _style_header_cell(cell) -> None:
    side = Side(style="thin", color="999999")
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", fgColor="D9EAD3")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(left=side, right=side, top=side, bottom=side)


def _style_body_cell(cell) -> None:
    side = Side(style="thin", color="DDDDDD")
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    cell.border = Border(left=side, right=side, top=side, bottom=side)


def _column_width(header: str) -> int:
    wide_headers = {
        "Nama peserta didik",
        "Tempat dan tanggal lahir",
        "Alamat",
        "Nama ayah",
        "Nama ibu",
        "Nama wali",
        "Alamat wali",
        "Tanggal pembagian rapor",
        "Tanggal masuk",
    }

    medium_headers = {
        "Nomor induk_NISN",
        "Diterima di kelompok",
        "Diterima tanggal",
        "Pekerjaan ayah",
        "Pekerjaan ibu",
        "Pekerjaan wali",
        "Nama panggilan",
    }

    if header in wide_headers:
        return 28

    if header in medium_headers:
        return 20

    if header in APP_CONFIG.excel.score_headers:
        return 14

    if header in {"S", "I", "A"}:
        return 8

    return 16


def generate_student_data_template(
    output_dir: Path | None = None,
    max_students: int | None = None,
) -> Path:
    """
    Generate file Excel kosong untuk input data murid dan nilai.

    Catatan:
    - Tidak memasukkan P1-P17.
    - P1-P17 akan dibuat otomatis saat generate database.
    - Header mengikuti APP_CONFIG.excel.student_data_headers.
    """
    config = APP_CONFIG.excel
    output_dir = output_dir or APP_CONFIG.output_dir
    max_students = max_students or config.default_max_students

    output_dir.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = config.student_data_sheet

    headers = config.student_data_headers

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        _style_header_cell(cell)
        ws.column_dimensions[get_column_letter(col_idx)].width = _column_width(header)

    for row_idx in range(2, max_students + 2):
        for col_idx in range(1, len(headers) + 1):
            _style_body_cell(ws.cell(row=row_idx, column=col_idx))

        # Default supaya user tidak perlu isi berulang.
        header_to_col = {header: idx for idx, header in enumerate(headers, start=1)}

        if "Semester" in header_to_col:
            ws.cell(row=row_idx, column=header_to_col["Semester"], value=config.semester_value)

        if "T.P." in header_to_col:
            ws.cell(row=row_idx, column=header_to_col["T.P."], value="2025/2026")

        if "Kelompok" in header_to_col:
            ws.cell(row=row_idx, column=header_to_col["Kelompok"], value="B4")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"TEMPLATE_DATA_MURID_NILAI_{max_students}_MURID_{timestamp}.xlsx"

    wb.save(output_path)
    return output_path