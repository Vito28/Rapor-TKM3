from __future__ import annotations

import streamlit as st

from src.config import APP_CONFIG
from src.services.runtime_files import get_session_runtime_dir
from src.ui_parts.sidebar import render_sidebar
from src.ui_parts.feature_1_general_info import render_feature_1_general_info
from src.ui_parts.feature_2_student_editor import render_feature_2_student_editor
from src.ui_parts.feature_3_indicator_mapping import render_feature_3_indicator_mapping
from src.ui_parts.feature_4_preview import render_feature_4_preview
from src.ui_parts.feature_5_generate_database import render_feature_5_generate_database


def render_app() -> None:
    st.set_page_config(
        page_title="Rapor Mail Merge Generator",
        page_icon="📘",
        layout="wide",
    )

    st.title("📘 Rapor Mail Merge Generator")
    st.caption(
        "Generate format indikator horizontal dan database semester 1 "
        "dari file data murid/nilai terpisah."
    )

    base_config = APP_CONFIG.excel
    runtime_dir = get_session_runtime_dir(APP_CONFIG.runtime_dir)

    runtime_config, max_students = render_sidebar(base_config)

    common_values, target_semester_header = render_feature_1_general_info(
        runtime_config=runtime_config,
    )

    render_feature_2_student_editor(
        runtime_dir=runtime_dir,
        runtime_config=runtime_config,
        max_students=int(max_students),
        common_values=common_values,
    )

    feature3_state = render_feature_3_indicator_mapping(
        runtime_dir=runtime_dir,
        runtime_config=runtime_config,
        max_students=int(max_students),
        target_semester_header=target_semester_header,
    )

    render_feature_4_preview(
        saved_student_path=feature3_state["saved_student_path"],
        runtime_config=runtime_config,
        max_students=int(max_students),
    )

    render_feature_5_generate_database(
        runtime_dir=runtime_dir,
        runtime_config=runtime_config,
        max_students=int(max_students),
        common_values=common_values,
        target_semester_header=target_semester_header,
        saved_student_path=feature3_state["saved_student_path"],
        saved_database_template_path=feature3_state["saved_database_template_path"],
        indicator_path_for_database=feature3_state["indicator_path_for_database"],
    )