from __future__ import annotations

from pathlib import Path


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def data_path(filename: str) -> Path:
    return project_root() / "data" / filename


def excel_path(filename: str) -> Path:
    return project_root() / "excel" / filename


def output_path(filename: str) -> Path:
    path = project_root() / "output" / filename
    path.parent.mkdir(parents=True, exist_ok=True)
    return path
