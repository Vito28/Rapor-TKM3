from __future__ import annotations

from pathlib import Path
from typing import Any

import streamlit as st

from src.services.database_generator import generate_database
from src.ui_parts.common import render_download_button, render_predicate_table, template_tsv


def render_feature_5_generate_database(
    *,
    runtime_dir: Path,
    runtime_config: Any,
    max_students: int,
    common_values: dict[str, object],
    target_semester_header: str,
    saved_student_path: Path | None,
    saved_database_template_path: Path | None,
    indicator_path_for_database: Path | None,
) -> None:
    st.subheader("5. Generate database")

    st.write(
        "Nilai angka akan masuk ke database. Jika nilai kosong, database tetap kosong. "
        "P1-P17 dibuat otomatis hanya untuk nilai yang terisi. "
        f"Narasi indikator hanya memakai BSB untuk sangat baik dan BB untuk bimbingan."
    )

    manual_narratives = st.session_state.get("manual_indicator_narratives_by_name")

    if manual_narratives:
        st.success(f"Narasi indikator siap masuk database: {len(manual_narratives)} murid.")
    else:
        st.warning(
            "Narasi indikator belum tersedia. Klik Apply Kategori lalu Submit Pilihan Indikator di bagian 3C."
        )

    can_generate_database = saved_student_path is not None and saved_database_template_path is not None

    if st.button(
        "Generate Database Semester 1",
        disabled=not can_generate_database,
        type="primary",
        use_container_width=True,
    ):
        try:
            result = generate_database(
                student_data_path=saved_student_path,
                database_template_path=saved_database_template_path,
                output_dir=runtime_dir,
                config=runtime_config,
                max_students=int(max_students),
                default_values=common_values,
                indicator_path=indicator_path_for_database,
                target_semester_header=target_semester_header,
                manual_indicator_narratives_by_name=manual_narratives,
            )

            st.session_state["database_result_path"] = str(result.path)
            st.session_state["database_result_message"] = result.message

        except Exception as exc:
            st.error(f"Generate database gagal: {exc}")

    if "database_result_message" in st.session_state:
        st.success(st.session_state["database_result_message"])

    if "database_result_path" in st.session_state:
        database_path = Path(st.session_state["database_result_path"])
        render_download_button(
            label="Download Database Semester 1",
            path=database_path,
        )

    with st.expander("Struktur file data murid/nilai yang benar"):
        st.write("Header yang disarankan. Copy baris ini ke Excel kalau mau buat manual:")
        st.code(template_tsv(runtime_config.student_data_headers), language="tsv")

        st.write("Predikat tidak perlu diinput manual. Sistem menghitung P1 sampai P17 dari nilai angka:")
        render_predicate_table()

        st.write(
            "Kolom biodata memang banyak, tetapi file input bisa di-freeze pada kolom Nama. "
            "Kolom P1/P2/P3 tidak perlu dibuat di input karena akan dibuat otomatis di output database."
        )