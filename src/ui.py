from __future__ import annotations

from anyio import Path
import streamlit as st

from dataclasses import replace

from src.config import APP_CONFIG, ExcelLayoutConfig
from src.services.database_generator import generate_database
from src.services.horizontal_generator import generate_horizontal_format
from src.services.runtime_files import get_session_runtime_dir, read_bytes, save_uploaded_file
from src.services.template_generator import generate_student_data_template
from src.utils.excel_reader import extract_student_names, read_rows_by_header


def _render_predicate_table() -> None:
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


def _template_tsv(headers: tuple[str, ...]) -> str:
    return "\t".join(headers)


def render_app() -> None:
    st.set_page_config(
        page_title="Rapor Mail Merge Generator",
        page_icon="📘",
        layout="wide",
    )

    st.title("📘 Rapor Mail Merge Generator")
    st.caption("Generate format indikator horizontal dan database semester 1 dari file data murid/nilai terpisah.")

    config = APP_CONFIG.excel
    runtime_dir = get_session_runtime_dir(APP_CONFIG.runtime_dir)

    with st.sidebar:
        st.header("Pengaturan")
        max_students = st.number_input(
            "Maksimal murid diproses",
            min_value=1,
            max_value=100,
            value=config.default_max_students,
            step=1,
        )

        with st.expander("Advanced sheet config"):
            template_sheet = st.text_input("Sheet indikator", value=config.template_sheet)
            student_sheet = st.text_input("Sheet file data murid/nilai", value=config.student_data_sheet)
            database_sheet = st.text_input("Sheet database template", value=config.database_sheet)
            name_header = st.text_input("Header nama murid", value=config.name_header)

        runtime_config = replace(
            config,
            template_sheet=template_sheet,
            student_data_sheet=student_sheet,
            database_sheet=database_sheet,
            name_header=name_header,
            default_max_students=int(max_students),
        )

    st.subheader("0. Buat template input murid/nilai")
    st.write(
        "Gunakan ini kalau belum punya file data murid/nilai. "
        "Template ini memakai Sheet1 dan header ringkas tanpa A_ dan tanpa P1-P17 manual."
    )
    temp_col1, temp_col2, temp_col3 = st.columns([1, 1, 2])
    with temp_col1:
        template_rows = st.number_input("Jumlah baris template", min_value=1, max_value=100, value=int(max_students), step=1)
    with temp_col2:
        with_dummy_data = st.checkbox("Isi dummy data", value=True)
    with temp_col3:
        if st.button("Generate Template Data Murid/Nilai", use_container_width=True):
            try:
                result = generate_student_data_template(
                    output_dir=runtime_dir,
                    config=runtime_config,
                    row_count=int(template_rows),
                    with_dummy_data=with_dummy_data,
                )
                st.session_state["template_result_path"] = str(result.path)
                st.session_state["template_result_message"] = result.message
                st.success(result.message)
            except Exception as exc:
                st.error(f"Generate template gagal: {exc}")

        if "template_result_path" in st.session_state:
            template_path = Path(st.session_state["template_result_path"])
        
            if template_path.exists():
                st.download_button(
                    label="Download Template Data Murid/Nilai",
                    data=read_bytes(template_path),
                    file_name=template_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    st.subheader("1. Upload file")
    col1, col2, col3 = st.columns(3)

    with col1:
        rubric_file = st.file_uploader(
            "File indikator/rubrik",
            type=["xlsx"],
            help="Contoh: RAPOR HIJAU_INDIKATOR.xlsx. Yang dipakai Sheet1.",
        )

    with col2:
        student_file = st.file_uploader(
            "File data murid & nilai",
            type=["xlsx"],
            help="Workbook terpisah. Pakai Sheet1. Header nilai: Pend_Agama, PMP, Penguasaan Aktif, dst.",
        )

    with col3:
        database_template_file = st.file_uploader(
            "Template database semester 1",
            type=["xlsx"],
            help="Contoh: DATABASE SEMESTER 1.xlsx.",
        )

    saved_rubric_path = None
    saved_student_path = None
    saved_database_template_path = None

    if rubric_file:
        saved_rubric_path = save_uploaded_file(rubric_file, runtime_dir, "rubric_indicator.xlsx")

    if student_file:
        saved_student_path = save_uploaded_file(student_file, runtime_dir, "student_data.xlsx")

    if database_template_file:
        saved_database_template_path = save_uploaded_file(database_template_file, runtime_dir, "database_template.xlsx")

    st.subheader("2. Preview data murid/nilai")
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
        st.info("Upload file data murid/nilai untuk melihat preview.")

    st.subheader("3. Generate output")
    action_col1, action_col2 = st.columns(2)

    with action_col1:
        st.markdown("**Generate Format Indikator Horizontal**")
        st.write("Mengambil A:E dari Sheet1 indikator, lalu membuat blok nama murid ke samping.")

        can_generate_format = saved_rubric_path is not None and saved_student_path is not None
        if st.button("Generate Format", disabled=not can_generate_format, use_container_width=True):
            try:
                result = generate_horizontal_format(
                    rubric_path=saved_rubric_path,
                    student_data_path=saved_student_path,
                    output_dir=runtime_dir,
                    config=runtime_config,
                    max_students=int(max_students),
                )
                st.success(result.message)
                st.download_button(
                    label="Download Format Indikator",
                    data=read_bytes(result.path),
                    file_name=result.path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as exc:
                st.error(f"Generate format gagal: {exc}")

    with action_col2:
        st.markdown("**Generate Database Semester 1**")
        st.write("Inject biodata/nilai ke template database. P1-P17 otomatis dari nilai angka.")

        can_generate_database = saved_student_path is not None and saved_database_template_path is not None
        if st.button("Generate Database", disabled=not can_generate_database, use_container_width=True):
            try:
                result = generate_database(
                    student_data_path=saved_student_path,
                    database_template_path=saved_database_template_path,
                    output_dir=runtime_dir,
                    config=runtime_config,
                    max_students=int(max_students),
                )
                st.success(result.message)
                st.download_button(
                    label="Download Database Semester 1",
                    data=read_bytes(result.path),
                    file_name=result.path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as exc:
                st.error(f"Generate database gagal: {exc}")

    with st.expander("Struktur file data murid/nilai yang benar"):
        st.write("Header yang disarankan. Copy baris ini ke Excel kalau mau buat manual:")
        st.code(_template_tsv(runtime_config.student_data_headers), language="tsv")
        st.write("Predikat tidak perlu diinput manual. Sistem menghitung P1 sampai P17 dari nilai angka:")
        _render_predicate_table()
        st.write(
            "Kolom biodata memang banyak, tetapi file input bisa di-freeze pada kolom Nama. "
            "Kolom P1/P2/P3 tidak perlu dibuat di input karena akan dibuat otomatis di output database."
        )
