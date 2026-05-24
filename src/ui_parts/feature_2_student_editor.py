from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from src.services.horizontal_generator import generate_horizontal_format
from src.services.runtime_files import save_uploaded_file
from src.services.template_generator import generate_student_data_template
from src.ui_parts.common import (
    DEFAULT_INDICATOR_FILENAME,
    default_indicator_path,
    render_download_button,
    save_student_editor_dataframe,
)


def build_student_editor_dataframe(
    *,
    headers: tuple[str, ...],
    row_count: int,
    config: Any,
    common_values: dict[str, object],
    with_dummy_data: bool,
) -> pd.DataFrame:
    rows: list[dict[str, object]] = []

    for index in range(1, row_count + 1):
        row: dict[str, object] = {header: "" for header in headers}

        for header, value in common_values.items():
            if header in row and value not in ("", None):
                row[header] = value

        if with_dummy_data:
            row.update(
                {
                    "Nama peserta didik": f"Murid {index:02d}",
                    "Nama panggilan": f"Murid{index}",
                    "Nomor induk_NISN": f"2025{index:04d}",
                    "Jenis kelamin": "L",
                    "Tempat dan tanggal lahir": "Medan, 1 Januari 2020",
                    "Agama": "Kristen",
                    "Anak ke": 1,
                    "Alamat": "Medan",
                    "Telepon": "-",
                    "Diterima di kelompok": row.get("Kelompok", "B4") or "B4",
                    "Diterima tanggal": row.get("Tanggal masuk", "01/07/2025") or "01/07/2025",
                    "Nama ayah": f"Ayah Murid {index:02d}",
                    "Nama ibu": f"Ibu Murid {index:02d}",
                    "Pekerjaan ayah": "-",
                    "Pekerjaan ibu": "-",
                    "Nama wali": "-",
                    "Alamat wali": "-",
                    "Telepon wali": "-",
                    "Pekerjaan wali": "-",
                    "Kelakuan": "A",
                    "S": "-",
                    "I": "-",
                    "A": "-",
                    "Naik ke ": "-",
                }
            )

            for score_header in config.score_headers:
                row[score_header] = round(8 + ((index % 7) * 0.2), 1)

        rows.append(row)

    return pd.DataFrame(rows, columns=list(headers))


def student_editor_column_config(config: Any) -> dict:
    column_config = {
        "Nama peserta didik": st.column_config.Column(
            "Nama peserta didik",
            pinned=True,
            width="medium",
            required=True,
        ),
        "Nama panggilan": st.column_config.Column(
            "Nama panggilan",
            pinned=True,
            width="small",
        ),
    }

    for score_header in config.score_headers:
        column_config[score_header] = st.column_config.NumberColumn(
            score_header,
            min_value=0.0,
            max_value=10.0,
            step=0.1,
            format="%.1f",
            width="small",
        )

    return column_config


def render_feature_2_student_editor(
    *,
    runtime_dir: Path,
    runtime_config: Any,
    max_students: int,
    common_values: dict[str, object],
) -> None:
    st.subheader("2. Buat dan edit data murid/nilai")

    st.write(
        "Buat template data murid/nilai, lihat preview, edit langsung di browser, "
        "lalu generate hasil edit atau tabel indikator horizontal."
    )

    temp_col1, temp_col2, temp_col3 = st.columns([1, 1, 2])

    with temp_col1:
        template_rows = st.number_input(
            "Jumlah baris template",
            min_value=1,
            max_value=100,
            value=int(max_students),
            step=1,
        )

    with temp_col2:
        with_dummy_data = st.checkbox("Isi dummy data", value=True)

    with temp_col3:
        if st.button("Generate Template Excel", use_container_width=True):
            try:
                result = generate_student_data_template(
                    output_dir=runtime_dir,
                    config=runtime_config,
                    row_count=int(template_rows),
                    with_dummy_data=with_dummy_data,
                    default_values=common_values,
                )

                st.session_state["template_result_path"] = str(result.path)
                st.session_state["template_result_message"] = result.message

            except Exception as exc:
                st.error(f"Generate template gagal: {exc}")

        if "template_result_message" in st.session_state:
            st.success(st.session_state["template_result_message"])

        if "template_result_path" in st.session_state:
            template_path = Path(st.session_state["template_result_path"])
            render_download_button(
                label="Download Template Excel",
                path=template_path,
            )

    st.markdown("#### Preview dan edit langsung")

    preview_col1, preview_col2 = st.columns([1, 2])

    with preview_col1:
        if st.button("See Preview / Edit", use_container_width=True):
            st.session_state["student_editor_df"] = build_student_editor_dataframe(
                headers=runtime_config.student_data_headers,
                row_count=int(template_rows),
                config=runtime_config,
                common_values=common_values,
                with_dummy_data=with_dummy_data,
            )

    with preview_col2:
        if "student_editor_df" in st.session_state:
            st.info(
                "Kolom Nama peserta didik dibuat freeze/pinned. "
                "Untuk layar HP, gunakan scroll horizontal pada tabel."
            )

    if "student_editor_df" in st.session_state:
        editor_height = min(620, 120 + int(template_rows) * 38)

        edited_df = st.data_editor(
            st.session_state["student_editor_df"],
            key="student_data_editor",
            use_container_width=True,
            hide_index=True,
            height=editor_height,
            num_rows="dynamic",
            column_config=student_editor_column_config(runtime_config),
        )

        st.session_state["student_editor_df"] = edited_df

        edit_col1, edit_col2 = st.columns([1, 2])

        with edit_col1:
            if st.button("Generate Hasil Edit", type="primary", use_container_width=True):
                try:
                    edited_output_path = save_student_editor_dataframe(
                        dataframe=edited_df,
                        output_dir=runtime_dir,
                        row_count=len(edited_df),
                    )

                    st.session_state["edited_student_file_path"] = str(edited_output_path)
                    st.success(f"Hasil edit berhasil dibuat: {edited_output_path.name}")

                except Exception as exc:
                    st.error(f"Generate hasil edit gagal: {exc}")

        with edit_col2:
            if "edited_student_file_path" in st.session_state:
                edited_path = Path(st.session_state["edited_student_file_path"])
                render_download_button(
                    label="Download Hasil Edit Data Murid/Nilai",
                    path=edited_path,
                )

    st.markdown("#### Generate tabel indikator dari hasil edit")

    rubric_path_to_use: Path | None = None
    default_rubric_path = default_indicator_path()

    if default_rubric_path.exists():
        rubric_path_to_use = default_rubric_path
        st.success(f"File indikator ditemukan otomatis: {default_rubric_path.name}")
    else:
        st.warning(
            "File indikator default tidak ditemukan di folder data/input. "
            "Silakan upload file indikator/rubrik secara manual."
        )

        uploaded_rubric_fallback = st.file_uploader(
            "Upload file indikator/rubrik",
            type=["xlsx"],
            key="fallback_rubric_upload_for_indicator",
            help=f"Upload {DEFAULT_INDICATOR_FILENAME} jika file default tidak ditemukan.",
        )

        if uploaded_rubric_fallback is not None:
            rubric_path_to_use = save_uploaded_file(
                uploaded_rubric_fallback,
                runtime_dir,
                "rubric_indicator_fallback.xlsx",
            )
            st.success("File indikator berhasil diupload dan siap dipakai.")

    has_editor_data = "student_editor_df" in st.session_state

    if not has_editor_data:
        st.warning("Klik See Preview / Edit dulu agar data murid tersedia.")

    can_generate_indicator_from_editor = has_editor_data and rubric_path_to_use is not None

    if st.button(
        "Generate Tabel Indikator dari Hasil Edit",
        disabled=not can_generate_indicator_from_editor,
        use_container_width=True,
    ):
        try:
            edited_df = st.session_state["student_editor_df"]

            edited_student_path = save_student_editor_dataframe(
                dataframe=edited_df,
                output_dir=runtime_dir,
                row_count=len(edited_df),
            )

            st.session_state["edited_student_file_path"] = str(edited_student_path)

            result = generate_horizontal_format(
                rubric_path=rubric_path_to_use,
                student_data_path=edited_student_path,
                output_dir=runtime_dir,
                config=runtime_config,
                max_students=int(max_students),
            )

            st.session_state["indicator_result_path"] = str(result.path)
            st.session_state["indicator_result_message"] = result.message

        except Exception as exc:
            st.error(f"Generate tabel indikator gagal: {exc}")

    if "indicator_result_message" in st.session_state:
        st.success(st.session_state["indicator_result_message"])

    if "indicator_result_path" in st.session_state:
        indicator_path = Path(st.session_state["indicator_result_path"])
        render_download_button(
            label="Download Tabel Indikator",
            path=indicator_path,
        )

    st.divider()