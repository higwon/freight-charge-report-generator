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
class GenerationRequest:
    source_path: Path
    output_path: Path
    report_format: ReportFormat = ReportFormat.CLASSIC


@dataclass(frozen=True)
class GenerationResult:
    output_path: Path
    record_count: int
    source_sheet_name: str
    func_codes: list[str]
    summary_sheet_count: int
    report_format: ReportFormat
