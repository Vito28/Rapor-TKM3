from __future__ import annotations

from dataclasses import replace
from typing import Any

import streamlit as st


def render_sidebar(base_config: Any) -> tuple[Any, int]:
    with st.sidebar:
        st.header("Pengaturan")

        max_students = st.number_input(
            "Maksimal murid diproses",
            min_value=1,
            max_value=100,
            value=base_config.default_max_students,
            step=1,
        )

        with st.expander("Advanced sheet config"):
            template_sheet = st.text_input("Sheet indikator", value=base_config.template_sheet)
            student_sheet = st.text_input("Sheet file data murid/nilai", value=base_config.student_data_sheet)
            database_sheet = st.text_input("Sheet database template", value=base_config.database_sheet)
            name_header = st.text_input("Header nama murid", value=base_config.name_header)

        runtime_config = replace(
            base_config,
            template_sheet=template_sheet,
            student_data_sheet=student_sheet,
            database_sheet=database_sheet,
            name_header=name_header,
            default_max_students=int(max_students),
        )

    return runtime_config, int(max_students)