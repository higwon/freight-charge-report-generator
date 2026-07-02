from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


APP_NAME = "Freight Charge Report Generator"
APP_VERSION = "0.1.0"
SOURCE_SHEET_NAME = "LV1DOC"
RAW_DATA_SHEET_NAME = "RawData"
LOG_DIR = Path("logs")
LOG_FILE = LOG_DIR / "app.log"


@dataclass(frozen=True)
class BusinessRules:
    required_columns: tuple[str, ...] = (
        "Job Date",
        "Func Code",
        "Arrival Port",
        "Depart Port",
        "Customer Name",
        "Loc Amt",
    )
    port_columns: dict[str, str] = field(
        default_factory=lambda: {
            "AE": "Depart Port",
            "OA": "Arrival Port",
            "OE": "Depart Port",
            "AI": "Arrival Port",
            "OI": "Arrival Port",
        }
    )
    default_port_column: str = "Arrival Port"
    amount_format: str = "#,##0.##"
    default_month_format: str = "%Y/%m"
