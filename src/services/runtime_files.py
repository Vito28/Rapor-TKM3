from __future__ import annotations

from pathlib import Path
from uuid import uuid4

import streamlit as st


def get_session_runtime_dir(base_dir: Path) -> Path:
    if "runtime_id" not in st.session_state:
        st.session_state.runtime_id = uuid4().hex

    runtime_dir = base_dir / st.session_state.runtime_id
    runtime_dir.mkdir(parents=True, exist_ok=True)
    return runtime_dir


def save_uploaded_file(uploaded_file, target_dir: Path, filename: str | None = None) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    safe_name = filename or uploaded_file.name
    target_path = target_dir / safe_name
    target_path.write_bytes(uploaded_file.getbuffer())
    return target_path


def read_bytes(path: Path) -> bytes:
    return path.read_bytes()
