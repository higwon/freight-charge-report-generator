from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Any

from .config import ALL_SHEET_NAME, BusinessRules
from .exceptions import InvalidAmountError, InvalidDateError, MissingRequiredColumnsError
from .models import (
    ArApFuncCodeSummary,
    ArApMonthlyAmount,
    ArApSummaryRow,
    CustomerSummary,
    FuncCodeSummary,
    MonthSummary,
    PortSummary,
    WorkbookData,
)


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
        all_grouped: dict[str, dict[str, dict[str, Decimal]]] = defaultdict(
            lambda: defaultdict(lambda: defaultdict(Decimal))
        )

        for row_number, row in enumerate(data.rows, start=2):
            func_code = str(self._value(row, index["Func Code"]) or "").strip()
            if func_code not in self.rules.func_codes:
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
            all_grouped[month_label][port_name][customer_name] += amount

        summaries = [self._build_summary(ALL_SHEET_NAME, all_grouped)]
        for func_code in self.rules.func_codes:
            if func_code not in grouped:
                continue
            summaries.append(self._build_summary(func_code, grouped[func_code]))
        return summaries

    def transform_ar_ap_monthly(
        self, data: WorkbookData, month_format: str | None = None
    ) -> list[ArApFuncCodeSummary]:
        index = self._header_index(data.headers)
        self._validate_columns(index)
        date_format = month_format or self.rules.default_month_format
        customer_code_index = index.get("Customer Code")

        grouped: dict[
            str, dict[tuple[str | None, str, str], dict[str, dict[str, Decimal]]]
        ] = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(Decimal))))
        months_by_code: dict[str, set[str]] = defaultdict(set)

        for row_number, row in enumerate(data.rows, start=2):
            func_code = str(self._value(row, index["Func Code"]) or "").strip()
            if not func_code:
                continue
            ar_ap_type = str(self._value(row, index["AR / AP Type"]) or "").strip().upper()
            if ar_ap_type not in {"AR", "AP"}:
                continue

            month_label = self._month_label(self._value(row, index["Job Date"]), date_format, row_number)
            port_column = self._ar_ap_port_column(func_code)
            port_name: str | None = None
            if port_column:
                port_name = str(self._value(row, index[port_column]) or "").strip() or "(Blank)"
            customer_code = ""
            if customer_code_index is not None:
                customer_code = str(self._value(row, customer_code_index) or "").strip()
            customer_name = str(self._value(row, index["Customer Name"]) or "").strip() or "(Blank)"
            amount = self._amount(self._value(row, index["Loc Amt"]), row_number)

            grouped[func_code][(port_name, customer_code, customer_name)][month_label][ar_ap_type] += amount
            months_by_code[func_code].add(month_label)

        summaries: list[ArApFuncCodeSummary] = []
        for func_code in sorted(grouped, key=self._sort_key):
            months = sorted(months_by_code[func_code])
            rows = []
            for key, monthly_values in sorted(grouped[func_code].items(), key=lambda item: self._summary_row_sort_key(item[0])):
                month_map = {
                    month: ArApMonthlyAmount(
                        ar=monthly_values[month].get("AR", Decimal("0")),
                        ap=monthly_values[month].get("AP", Decimal("0")),
                    )
                    for month in months
                }
                rows.append(
                    ArApSummaryRow(
                        port=key[0],
                        customer_code=key[1],
                        customer_name=key[2],
                        monthly=month_map,
                    )
                )
            port_label = self._ar_ap_port_label(func_code)
            summaries.append(
                ArApFuncCodeSummary(
                    func_code=func_code,
                    category=self._ar_ap_category(func_code),
                    port_label=port_label,
                    months=months,
                    rows=rows,
                )
            )
        return summaries

    def _build_summary(
        self, func_code: str, grouped: dict[str, dict[str, dict[str, Decimal]]]
    ) -> FuncCodeSummary:
        month_summaries: list[MonthSummary] = []
        func_total = Decimal("0")
        for month in sorted(grouped):
            port_summaries: list[PortSummary] = []
            month_total = Decimal("0")
            for port in sorted(grouped[month], key=self._sort_key):
                customers = [
                    CustomerSummary(name=customer, amount=amount)
                    for customer, amount in sorted(
                        grouped[month][port].items(), key=lambda item: self._sort_key(item[0])
                    )
                ]
                port_total = sum((item.amount for item in customers), Decimal("0"))
                month_total += port_total
                port_summaries.append(PortSummary(name=port, amount=port_total, customers=customers))
            func_total += month_total
            month_summaries.append(MonthSummary(label=month, amount=month_total, ports=port_summaries))
        return FuncCodeSummary(func_code=func_code, amount=func_total, months=month_summaries)

    def _validate_columns(self, index: dict[str, int]) -> None:
        missing = [column for column in self.rules.required_columns if column not in index]
        if missing:
            raise MissingRequiredColumnsError(f"필수 열이 없습니다: {', '.join(missing)}")

    @staticmethod
    def _header_index(headers: list[str]) -> dict[str, int]:
        expected_names = set(BusinessRules().required_columns) | {"Customer Code"}
        expected = {" ".join(name.split()).casefold(): name for name in expected_names}
        result: dict[str, int] = {}
        for position, header in enumerate(headers):
            normalized = " ".join(str(header).replace("\ufeff", "").split()).casefold()
            if normalized in expected:
                result[expected[normalized]] = position
        return result

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

    def _ar_ap_port_column(self, func_code: str) -> str | None:
        if func_code in self.rules.export_func_codes:
            return "Arrival Port"
        if func_code in self.rules.import_func_codes:
            return "Depart Port"
        return None

    def _ar_ap_port_label(self, func_code: str) -> str | None:
        if func_code in self.rules.export_func_codes:
            return "도착항"
        if func_code in self.rules.import_func_codes:
            return "출발항"
        return None

    def _ar_ap_category(self, func_code: str) -> str:
        if func_code in self.rules.export_func_codes:
            return "수출"
        if func_code in self.rules.import_func_codes:
            return "수입"
        return "기타"

    @staticmethod
    def _summary_row_sort_key(key: tuple[str | None, str, str]) -> tuple[str, str, str]:
        port, customer_code, customer_name = key
        return ((port or "").casefold(), customer_name.casefold(), customer_code.casefold())
