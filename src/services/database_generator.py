from __future__ import annotations

from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any
import re

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from config import ExcelLayoutConfig
from models import GeneratedWorkbookResult
from utils.excel_reader import (
    normalize,
    normalize_header,
    nickname_from_name,
    read_rows_by_header,
)


def header_map(ws: Worksheet) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        header = normalize(ws.cell(1, col).value)
        if header:
            mapping[normalize_header(header)] = col
    return mapping


def get_col(headers: dict[str, int], header_name: str) -> int | None:
    return headers.get(normalize_header(header_name))


def capture_row_styles(ws: Worksheet, source_row: int) -> list[dict[str, Any]]:
    styles: list[dict[str, Any]] = []
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(source_row, col)
        styles.append(
            {
                "font": copy(cell.font),
                "fill": copy(cell.fill),
                "border": copy(cell.border),
                "alignment": copy(cell.alignment),
                "number_format": cell.number_format,
                "protection": copy(cell.protection),
            }
        )
    return styles


def apply_row_styles(ws: Worksheet, row: int, styles: list[dict[str, Any]]) -> None:
    for col, style in enumerate(styles, start=1):
        cell = ws.cell(row, col)
        cell.font = copy(style["font"])
        cell.fill = copy(style["fill"])
        cell.border = copy(style["border"])
        cell.alignment = copy(style["alignment"])
        cell.number_format = style["number_format"]
        cell.protection = copy(style["protection"])


def clear_existing_data(ws: Worksheet) -> None:
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)


def write_if_header_exists(ws: Worksheet, headers: dict[str, int], row_idx: int, header_name: str, value: Any) -> bool:
    col = get_col(headers, header_name)
    if not col:
        return False
    ws.cell(row=row_idx, column=col).value = value
    return True


def canonical_header(header: str, config: ExcelLayoutConfig) -> str:
    cleaned = normalize(header)
    if not cleaned:
        return ""

    for alias, target in config.header_aliases.items():
        if normalize_header(alias) == normalize_header(cleaned):
            return target

    return cleaned


def normalize_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = normalize(value)
    if not text or text == "-":
        return None

    # Indonesia sering pakai koma decimal.
    text = text.replace(".", "") if re.fullmatch(r"\d{1,3}(\.\d{3})+(,\d+)?", text) else text
    text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return None


def format_number_for_excel(value: float | None) -> float | None:
    if value is None:
        return None
    rounded = round(value, 2)
    if rounded.is_integer():
        return int(rounded)
    return rounded


def predicate_from_score(value: Any, config: ExcelLayoutConfig) -> str:
    score = normalize_number(value)
    if score is None:
        return ""

    rule = config.predicate
    if score < rule.kurang_below:
        return "D"
    if score <= rule.cukup_max:
        return "C"
    if score <= rule.baik_max:
        return "B"
    if score <= rule.sangat_baik_max:
        return "A"
    return ""


def build_source_row_with_aliases(source_row: dict[str, Any], config: ExcelLayoutConfig) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for header, value in source_row.items():
        canonical = canonical_header(header, config)
        if canonical:
            result[canonical] = value
    return result


def compute_score_stats(row: dict[str, Any], config: ExcelLayoutConfig) -> tuple[float | None, float | None]:
    scores: list[float] = []
    for score_header, _predicate_header in config.score_to_predicate_headers:
        value = normalize_number(row.get(score_header))
        if value is not None:
            scores.append(value)

    if not scores:
        return None, None

    total = sum(scores)
    average = total / len(scores)
    return total, average


def write_row_to_database(
    db_ws: Worksheet,
    headers: dict[str, int],
    output_row: int,
    index: int,
    source_row: dict[str, Any],
    config: ExcelLayoutConfig,
) -> None:
    normalized_source = build_source_row_with_aliases(source_row, config)

    name = normalize(normalized_source.get(config.name_header))
    nickname = normalize(normalized_source.get("Nama panggilan")) or nickname_from_name(name)

    # Wajib/standar.
    write_if_header_exists(db_ws, headers, output_row, "No.", f"{index}.")
    write_if_header_exists(db_ws, headers, output_row, "Nama peserta didik", name)
    write_if_header_exists(db_ws, headers, output_row, "Nama panggilan", nickname)
    write_if_header_exists(db_ws, headers, output_row, "Semester", normalized_source.get("Semester") or config.semester_value)

    # Copy semua kolom input yang memang ada di template database.
    for source_header, value in normalized_source.items():
        write_if_header_exists(db_ws, headers, output_row, source_header, value)

    # Auto predikat P1-P17 berdasarkan nilai angka.
    for score_header, predicate_header in config.score_to_predicate_headers:
        score_value = normalized_source.get(score_header)
        if score_value in (None, ""):
            continue

        write_if_header_exists(db_ws, headers, output_row, score_header, score_value)
        write_if_header_exists(db_ws, headers, output_row, predicate_header, predicate_from_score(score_value, config))

    # Auto total dan rata-rata kalau tidak diisi di input.
    total, average = compute_score_stats(normalized_source, config)
    if total is not None:
        if normalized_source.get("Jumlah semua") in (None, ""):
            write_if_header_exists(db_ws, headers, output_row, "Jumlah semua", format_number_for_excel(total))
        if normalized_source.get("Rata-rata") in (None, ""):
            write_if_header_exists(db_ws, headers, output_row, "Rata-rata", format_number_for_excel(average))


def generate_database(
    student_data_path: Path,
    database_template_path: Path,
    output_dir: Path,
    config: ExcelLayoutConfig,
    max_students: int | None = None,
) -> GeneratedWorkbookResult:
    """
    Generate database dari file data murid/nilai terpisah.

    Perilaku utama:
    - Data murid/nilai dibaca dari Sheet1 file baru, bukan Sheet3.
    - Header input harus mengikuti template ringkas: Pend_Agama, PMP, dst.
    - P1-P17 dibuat otomatis dari nilai angka.
    - D/C/B/A dibuat dari aturan predikat.
    - Header yang sama/alias akan langsung diinject ke DATABASE SEMESTER 1.xlsx.
    """
    if not student_data_path.exists():
        raise FileNotFoundError(f"File data murid/nilai tidak ditemukan: {student_data_path}")
    if not database_template_path.exists():
        raise FileNotFoundError(f"Template database tidak ditemukan: {database_template_path}")

    rows = read_rows_by_header(
        workbook_path=student_data_path,
        sheet_name=config.student_data_sheet,
        required_header=config.name_header,
        max_rows=max_students,
    )
    if not rows:
        raise ValueError("Tidak ada data murid yang bisa dibaca dari file data murid/nilai.")

    db_wb = load_workbook(database_template_path)
    db_ws = db_wb[config.database_sheet] if config.database_sheet in db_wb.sheetnames else db_wb.active
    headers = header_map(db_ws)

    if not headers:
        raise ValueError("Header database tidak terbaca. Pastikan baris 1 berisi nama kolom.")

    style_source_row = 2 if db_ws.max_row >= 2 else 1
    row_styles = capture_row_styles(db_ws, style_source_row)

    clear_existing_data(db_ws)

    for index, source_row in enumerate(rows, start=1):
        output_row = index + 1
        apply_row_styles(db_ws, output_row, row_styles)
        write_row_to_database(
            db_ws=db_ws,
            headers=headers,
            output_row=output_row,
            index=index,
            source_row=source_row,
            config=config,
        )

    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"DATABASE_SEMESTER_1_{len(rows)}_MURID_{timestamp}.xlsx"
    db_wb.save(output_path)

    return GeneratedWorkbookResult(
        path=output_path,
        student_count=len(rows),
        message=f"Database semester 1 berhasil dibuat untuk {len(rows)} murid. Nilai dan P1-P17 sudah diisi otomatis.",
    )
