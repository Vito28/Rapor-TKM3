from __future__ import annotations

from pathlib import Path
from uuid import uuid4

import streamlit as st


def get_session_runtime_dir(base_runtime_dir: Path) -> Path:
    if "runtime_id" not in st.session_state:
        st.session_state["runtime_id"] = uuid4().hex

    runtime_dir = base_runtime_dir / st.session_state["runtime_id"]
    runtime_dir.mkdir(parents=True, exist_ok=True)
    return runtime_dir


def save_uploaded_file(uploaded_file, output_dir: Path, filename: str | None = None) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)

    safe_filename = filename or uploaded_file.name
    output_path = output_dir / safe_filename

    output_path.write_bytes(uploaded_file.getbuffer())
    return output_path


def read_bytes(path: Path | str) -> bytes:
    file_path = Path(path)

    if not file_path.exists():
        raise FileNotFoundError(f"File tidak ditemukan: {file_path}")

    return file_path.read_bytes()