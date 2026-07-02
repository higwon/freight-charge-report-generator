from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Any

from .config import BusinessRules
from .exceptions import InvalidAmountError, InvalidDateError, MissingRequiredColumnsError
from .models import CustomerSummary, FuncCodeSummary, MonthSummary, PortSummary, WorkbookData


class ReportTransformer:
    def __init__(self, rules: BusinessRules | None = None) -> None:
        self.rules = rules or BusinessRules()

    def transform(self, data: WorkbookData, month_format: str | None = None) -> list[FuncCodeSummary]:
        index = self._header_index(data.headers)
        self._validate_columns(index)
        date_format = month_format or self.rules.default_month_format

        grouped: dict[str, dict[str, dict[str, dict[str, Decimal]]]] = defaultdict(
            lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(Decimal)))
        )

        for row_number, row in enumerate(data.rows, start=2):
            func_code = str(self._value(row, index["Func Code"]) or "").strip()
            if not func_code:
                continue

            month_label = self._month_label(self._value(row, index["Job Date"]), date_format, row_number)
            port_column = self.rules.port_columns.get(func_code, self.rules.default_port_column)
            port_value = self._value(row, index[port_column])
            customer_value = self._value(row, index["Customer Name"])
            port_name = str(port_value or "").strip() or "(Blank)"
            customer_name = str(customer_value if customer_value is not None else "")
            if not customer_name.strip():
                customer_name = "(Blank)"
            amount = self._amount(self._value(row, index["Loc Amt"]), row_number)
            grouped[func_code][month_label][port_name][customer_name] += amount

        summaries: list[FuncCodeSummary] = []
        for func_code in sorted(grouped, key=self._sort_key):
            month_summaries: list[MonthSummary] = []
            func_total = Decimal("0")
            for month in sorted(grouped[func_code]):
                port_summaries: list[PortSummary] = []
                month_total = Decimal("0")
                for port in sorted(grouped[func_code][month], key=self._sort_key):
                    customers = [
                        CustomerSummary(name=customer, amount=amount)
                        for customer, amount in sorted(
                            grouped[func_code][month][port].items(), key=lambda item: self._sort_key(item[0])
                        )
                    ]
                    port_total = sum((item.amount for item in customers), Decimal("0"))
                    month_total += port_total
                    port_summaries.append(PortSummary(name=port, amount=port_total, customers=customers))
                func_total += month_total
                month_summaries.append(MonthSummary(label=month, amount=month_total, ports=port_summaries))
            summaries.append(FuncCodeSummary(func_code=func_code, amount=func_total, months=month_summaries))
        return summaries

    def _validate_columns(self, index: dict[str, int]) -> None:
        missing = [column for column in self.rules.required_columns if column not in index]
        if missing:
            raise MissingRequiredColumnsError(f"필수 열이 없습니다: {', '.join(missing)}")

    @staticmethod
    def _header_index(headers: list[str]) -> dict[str, int]:
        return {header.strip(): position for position, header in enumerate(headers) if header}

    @staticmethod
    def _value(row: list[Any], index: int) -> Any:
        return row[index] if index < len(row) else None

    @staticmethod
    def _month_label(value: Any, output_format: str, row_number: int) -> str:
        if isinstance(value, datetime):
            return value.strftime(output_format)
        text = str(value or "").strip()
        for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y/%m", "%Y-%m", "%Y%m%d"):
            try:
                return datetime.strptime(text, fmt).strftime(output_format)
            except ValueError:
                continue
        raise InvalidDateError(f"{row_number}행의 Job Date가 올바르지 않습니다: {value!r}")

    @staticmethod
    def _amount(value: Any, row_number: int) -> Decimal:
        if value in (None, ""):
            return Decimal("0")
        try:
            return Decimal(str(value).replace(",", "").strip())
        except (InvalidOperation, AttributeError) as exc:
            raise InvalidAmountError(f"{row_number}행의 Loc Amt가 올바르지 않습니다: {value!r}") from exc

    @staticmethod
    def _sort_key(value: str) -> str:
        return value.casefold()
