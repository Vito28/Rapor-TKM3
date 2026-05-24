from __future__ import annotations

from pathlib import Path
from typing import Any

import streamlit as st

from src.utils.excel_reader import extract_student_names, read_rows_by_header


def render_feature_4_preview(
    *,
    saved_student_path: Path | None,
    runtime_config: Any,
    max_students: int,
) -> None:
    st.subheader("4. Preview data murid/nilai")

    if saved_student_path:
        try:
            names = extract_student_names(
                workbook_path=saved_student_path,
                sheet_name=runtime_config.student_data_sheet,
                name_header=runtime_config.name_header,
                max_students=int(max_students),
            )

            rows = read_rows_by_header(
                workbook_path=saved_student_path,
                sheet_name=runtime_config.student_data_sheet,
                required_header=runtime_config.name_header,
                max_rows=int(max_students),
            )

            st.success(f"Terbaca {len(names)} murid dari file data murid/nilai.")
            st.dataframe(rows, use_container_width=True, hide_index=True)

        except Exception as exc:
            st.error(f"Preview gagal: {exc}")
    else:
        st.info("Belum ada data murid/nilai untuk preview.")

    st.divider()