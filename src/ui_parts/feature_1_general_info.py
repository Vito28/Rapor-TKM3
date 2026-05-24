from __future__ import annotations

from datetime import date
from typing import Any

import streamlit as st

from src.ui_parts.common import format_date_id, semester_header_from_value


def render_feature_1_general_info(runtime_config: Any) -> tuple[dict[str, object], str]:
    st.subheader("1. Informasi umum rapor")

    info_col1, info_col2, info_col3 = st.columns(3)

    with info_col1:
        tahun_pelajaran = st.text_input(
            "Tahun Pelajaran",
            value="2025/2026",
        )

        kelas = st.text_input(
            "Kelompok / Kelas",
            value="B4",
        )

    with info_col2:
        nama_guru = st.text_input(
            "Nama Guru",
            value="",
            placeholder="Contoh: Ibu Maria",
        )

        semester_value = st.text_input(
            "Semester",
            value=runtime_config.semester_value,
        )

    with info_col3:
        tanggal_masuk = st.date_input(
            "Tanggal Masuk",
            value=date(2025, 7, 1),
            format="DD/MM/YYYY",
        )

        tanggal_pembagian_rapor = st.date_input(
            "Tanggal Pembagian Rapor",
            value=date(2025, 12, 20),
            format="DD/MM/YYYY",
        )

    common_values: dict[str, object] = {
        "T.P.": tahun_pelajaran,
        "Kelompok": kelas,
        "Semester": semester_value,
        "Guru": nama_guru,
        "Tanggal masuk": format_date_id(tanggal_masuk),
        "Tanggal pembagian rapor": format_date_id(tanggal_pembagian_rapor),
    }

    target_semester_header = semester_header_from_value(semester_value)

    st.info(f"Semester indikator yang akan dibaca untuk narasi: **{target_semester_header}**")
    st.divider()

    return common_values, target_semester_header