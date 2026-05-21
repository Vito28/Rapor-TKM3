from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class ScorePredicateConfig:
    """Aturan predikat dari nilai angka."""

    # Sesuai tabel predikat:
    # D = Kurang      : x < 7
    # C = Cukup       : 7 <= x <= 8
    # B = Baik        : 8 < x <= 9
    # A = Sangat Baik : 9 < x <= 10
    # Praktisnya dibuat tanpa gap decimal agar nilai 7, 8, 9 tetap terbaca.
    kurang_below: float = 7.0
    cukup_max: float = 8.0
    baik_max: float = 9.0
    sangat_baik_max: float = 10.0


@dataclass(frozen=True)
class ExcelLayoutConfig:
    # File indikator tetap memakai Sheet1 dari RAPOR HIJAU_INDIKATOR.xlsx.
    template_sheet: str = "Sheet1"

    # File data murid/nilai sekarang workbook terpisah, sheet default Sheet1.
    student_data_sheet: str = "Sheet1"

    # Template database biasanya punya sheet bernama "Sheet 1".
    database_sheet: str = "Sheet 1"

    name_header: str = "Nama peserta didik"
    semester_value: str = "1 (satu)"

    # Layout format indikator horizontal.
    left_copy_until_col: int = 5          # A:E dari sheet indikator asli
    first_student_col: int = 6            # F
    columns_per_student: int = 8          # SEM I 4 kolom + SEM II 4 kolom

    name_row: int = 3
    semester_row: int = 4
    score_row: int = 5
    data_start_row: int = 6

    semester_labels: tuple[str, str] = ("SEMESTER I", "SEMESTER II")
    score_labels: tuple[str, str, str, str] = ("BB", "MB", "BSH", "BSB")

    default_max_students: int = 21

    # Header input nilai ringkas. Tidak perlu P1/P2/P3 diisi manual.
    score_to_predicate_headers: tuple[tuple[str, str], ...] = (
        ("Pend_Agama", "P1"),
        ("PMP", "P2"),
        ("Penguasaan Aktif", "P3"),
        ("Penguasaan Pasif", "P4"),
        ("Syair", "P5"),
        ("Dramatisasi", "P6"),
        ("Menggambar", "P7"),
        ("Prakarya", "P8"),
        ("Menyanyi", "P9"),
        ("Bermain", "P10"),
        ("Penjas", "P11"),
        ("PLH", "P12"),
        ("Matematika", "P13"),
        ("Mengenal Bentuk", "P14"),
        ("Menulis", "P15"),
        ("B. Ingg", "P16"),
        ("B. Mand.", "P17"),
    )

    # Header input template data murid/nilai yang disarankan.
    # A_ prefix dihapus. P1-P17 tidak dimasukkan karena dibuat otomatis.
    student_data_headers: tuple[str, ...] = (
        "Nama peserta didik",
        "Nama panggilan",
        "Nomor induk_NISN",
        "Jenis kelamin",
        "Tempat dan tanggal lahir",
        "Agama",
        "Anak ke",
        "Alamat",
        "Telepon",
        "Diterima di kelompok",
        "Diterima tanggal",
        "Nama ayah",
        "Nama ibu",
        "Pekerjaan ayah",
        "Pekerjaan ibu",
        "Nama wali",
        "Alamat wali",
        "Telepon wali",
        "Pekerjaan wali",
        "Kelompok",
        "Semester",
        "T.P.",
        "Pend_Agama",
        "PMP",
        "Penguasaan Aktif",
        "Penguasaan Pasif",
        "Syair",
        "Dramatisasi",
        "Menggambar",
        "Prakarya",
        "Menyanyi",
        "Bermain",
        "Penjas",
        "PLH",
        "Matematika",
        "Mengenal Bentuk",
        "Menulis",
        "B. Ingg",
        "B. Mand.",
        "Kelakuan",
        "S",
        "I",
        "A",
        "Tanggal pembagian rapor",
        "Guru",
        "Naik ke ",
        "Tanggal masuk",
    )

    # Alias supaya input yang lebih pendek tetap bisa masuk ke template database.
    header_aliases: dict[str, str] = field(
        default_factory=lambda: {
            "NISN": "Nomor induk_NISN",
            "Nomor Induk/NISN": "Nomor induk_NISN",
            "Nomor Induk NISN": "Nomor induk_NISN",
            "Nama Murid": "Nama peserta didik",
            "Nama siswa": "Nama peserta didik",
            "Nama peserta didik": "Nama peserta didik",
            "Tahun Pelajaran": "T.P.",
            "TP": "T.P.",
            "B Inggris": "B. Ingg",
            "Bahasa Inggris": "B. Ingg",
            "B Mandarin": "B. Mand.",
            "Bahasa Mandarin": "B. Mand.",
            "Sakit": "S",
            "Izin": "I",
            "Alpa": "A",
        }
    )

    predicate: ScorePredicateConfig = ScorePredicateConfig()


@dataclass(frozen=True)
class AppConfig:
    root_dir: Path = Path(__file__).resolve().parents[2]
    input_dir: Path = root_dir / "data" / "input"
    output_dir: Path = root_dir / "data" / "output"
    runtime_dir: Path = root_dir / "data" / "runtime"
    excel: ExcelLayoutConfig = ExcelLayoutConfig()


APP_CONFIG = AppConfig()
