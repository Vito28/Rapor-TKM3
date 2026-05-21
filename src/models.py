from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class GeneratedWorkbookResult:
    path: Path
    student_count: int
    message: str
