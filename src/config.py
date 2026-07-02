from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


APP_NAME = "Freight Charge Report Generator"
APP_VERSION = "0.4.0"
APP_AUTHOR = "DA_GYEONG"
SOURCE_DATA_SHEET_NAME = "Source"
ALL_SHEET_NAME = "ALL"
HEADER_SCAN_ROW_LIMIT = 20
LOG_DIR = Path("logs")
LOG_FILE = LOG_DIR / "app.log"


@dataclass(frozen=True)
class BusinessRules:
    func_codes: tuple[str, ...] = ("AE", "AI", "OA", "OE", "OI")
    export_func_codes: tuple[str, ...] = ("AE", "OE", "OA")
    import_func_codes: tuple[str, ...] = ("AI", "OI")
    required_columns: tuple[str, ...] = (
        "Job Date",
        "Func Code",
        "Arrival Port",
        "Depart Port",
        "AR / AP Type",
        "Customer Name",
        "Loc Amt",
    )
    port_columns: dict[str, str] = field(
        default_factory=lambda: {
            "AE": "Arrival Port",
            "OA": "Arrival Port",
            "OE": "Arrival Port",
            "AI": "Depart Port",
            "OI": "Depart Port",
        }
    )
    default_port_column: str = "Arrival Port"
    amount_format: str = "#,##0.##"
    default_month_format: str = "%Y/%m"
