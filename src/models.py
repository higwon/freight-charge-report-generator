from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum
from pathlib import Path
from typing import Any


Row = list[Any]


class ReportFormat(str, Enum):
    CLASSIC = "classic"
    ANALYTIC = "analytic"
    AR_AP_MONTHLY = "ar_ap_monthly"


@dataclass(frozen=True)
class WorkbookData:
    headers: list[str]
    rows: list[Row]
    source_sheet_name: str = ""

    @property
    def record_count(self) -> int:
        return len(self.rows)


@dataclass(frozen=True)
class CustomerSummary:
    name: str
    amount: Decimal


@dataclass(frozen=True)
class PortSummary:
    name: str
    amount: Decimal
    customers: list[CustomerSummary] = field(default_factory=list)


@dataclass(frozen=True)
class MonthSummary:
    label: str
    amount: Decimal
    ports: list[PortSummary] = field(default_factory=list)


@dataclass(frozen=True)
class FuncCodeSummary:
    func_code: str
    amount: Decimal
    months: list[MonthSummary] = field(default_factory=list)


@dataclass(frozen=True)
class ArApMonthlyAmount:
    ar: Decimal = Decimal("0")
    ap: Decimal = Decimal("0")

    @property
    def difference(self) -> Decimal:
        return self.ar - self.ap


@dataclass(frozen=True)
class ArApSummaryRow:
    port: str | None
    customer_code: str
    customer_name: str
    monthly: dict[str, ArApMonthlyAmount] = field(default_factory=dict)


@dataclass(frozen=True)
class ArApFuncCodeSummary:
    func_code: str
    category: str
    port_label: str | None
    months: list[str]
    rows: list[ArApSummaryRow] = field(default_factory=list)


@dataclass(frozen=True)
class GenerationRequest:
    source_path: Path
    output_path: Path
    report_format: ReportFormat = ReportFormat.AR_AP_MONTHLY


@dataclass(frozen=True)
class GenerationResult:
    output_path: Path
    record_count: int
    source_sheet_name: str
    func_codes: list[str]
    summary_sheet_count: int
    report_format: ReportFormat
