from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


APP_NAME = "Freight Charge Report Generator"
APP_VERSION = "0.8.0"
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
    header_aliases: dict[str, tuple[str, ...]] = field(
        default_factory=lambda: {
            "Job Date": ("Job Date", "JOB일자", "JOB 일자", "작업일자", "일자"),
            "Func Code": ("Func Code", "기능코드", "Func. Code"),
            "Arrival Port": ("Arrival Port", "도착지포트", "도착항", "도착 Port"),
            "Depart Port": ("Depart Port", "출발지포트", "출발항", "출발 Port"),
            "AR / AP Type": ("AR / AP Type", "AR/AP Type", "매출매입구분", "매출/매입 구분"),
            "Customer Name": ("Customer Name", "거래처명", "거래처 이름", "Customer"),
            "Customer Code": ("Customer Code", "거래처코드", "거래처 코드"),
            "Loc Amt": ("Loc Amt", "현지금액", "현지 금액", "Local Amount"),
        }
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

    def canonical_column_name(self, value: Any) -> str | None:
        normalized = self.normalize_header(value)
        for canonical, aliases in self.header_aliases.items():
            if normalized in {self.normalize_header(alias) for alias in aliases}:
                return canonical
        return None

    @staticmethod
    def normalize_header(value: Any) -> str:
        return " ".join(str(value).replace("\ufeff", "").split()).casefold()
