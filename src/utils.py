from __future__ import annotations

import logging
import sys
from pathlib import Path

from .config import LOG_FILE


def setup_logging() -> None:
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
        encoding="utf-8",
    )


def clean_sheet_title(title: str) -> str:
    invalid = set(r'[]:*?/\\')
    cleaned = "".join("_" if char in invalid else char for char in title).strip()
    return (cleaned or "Sheet")[:31]


def ensure_xlsx_suffix(filename: str) -> str:
    name = filename.strip() or "result.xlsx"
    return name if name.lower().endswith(".xlsx") else f"{name}.xlsx"


def display_path(path: Path) -> str:
    return str(path.expanduser().resolve())


def resource_path(relative_path: str) -> Path:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base_path / relative_path
