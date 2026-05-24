from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from src.config import APP_CONFIG
from src.services.runtime_files import read_bytes


EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DEFAULT_INDICATOR_FILENAME = "RAPOR HIJAU_INDIKATOR.xlsx"


def render_predicate_table() -> None:
    st.markdown(
        """
| Nilai | Predikat |
|---:|---|
| x < 7 | D = Kurang |
| 7 ≤ x ≤ 8 | C = Cukup |
| 8 < x ≤ 9 | B = Baik |
| 9 < x ≤ 10 | A = Sangat Baik |
"""
    )


def template_tsv(headers: tuple[str, ...]) -> str:
    return "\t".join(headers)


def format_date_id(value: date) -> str:
    return value.strftime("%d/%m/%Y")


def semester_header_from_value(value: str) -> str:
    text = str(value).strip().lower()

    if "2" in text or "dua" in text or "ii" in text:
        return "SEMESTER II"

    return "SEMESTER I"


def render_download_button(
    *,
    label: str,
    path: Path,
    mime: str = EXCEL_MIME,
) -> None:
    if not path.exists():
        st.warning(f"File hasil tidak ditemukan: {path}")
        return

    st.download_button(
        label=label,
        data=read_bytes(path),
        file_name=path.name,
        mime=mime,
        use_container_width=True,
    )


def default_indicator_path() -> Path:
    return APP_CONFIG.input_dir / DEFAULT_INDICATOR_FILENAME


def save_student_editor_dataframe(
    *,
    dataframe: pd.DataFrame,
    output_dir: Path,
    row_count: int,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"DATA_MURID_NILAI_EDITED_{row_count}_MURID_{timestamp}.xlsx"

    dataframe.to_excel(output_path, index=False, sheet_name="Sheet1")

    return output_path