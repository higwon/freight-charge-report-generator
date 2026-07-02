from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any


Row = list[Any]


@dataclass(frozen=True)
class WorkbookData:
    headers: list[str]
    rows: list[Row]

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


@dataclass(frozen=True)
class GenerationResult:
    output_path: Path
    record_count: int
    func_codes: list[str]
    summary_sheet_count: int
