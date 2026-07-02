from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.utils.exceptions import IllegalCharacterError

from .config import SOURCE_DATA_SHEET_NAME
from .exceptions import OutputPermissionError
from .formatter import Formatter, ReportStyle
from .models import FuncCodeSummary, WorkbookData
from .utils import clean_sheet_title


class ReportWriter:
    def __init__(self, formatter: Formatter | None = None) -> None:
        self.formatter = formatter or Formatter()

    def write(
        self,
        output_path: Path,
        source_data: WorkbookData,
        summaries: list[FuncCodeSummary],
        style: ReportStyle,
    ) -> None:
        workbook = Workbook()
        source_sheet = workbook.active
        source_sheet.title = SOURCE_DATA_SHEET_NAME
        self._write_raw_data(source_sheet, source_data)
        self.formatter.style_raw_sheet(source_sheet)

        for summary in summaries:
            sheet = workbook.create_sheet(clean_sheet_title(summary.func_code))
            self._write_summary_sheet(sheet, summary, style)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            workbook.save(output_path)
        except PermissionError as exc:
            raise OutputPermissionError(
                "출력 파일을 저장할 수 없습니다. 결과 파일이 열려 있다면 닫은 뒤 다시 시도하세요."
            ) from exc

    def _write_raw_data(self, sheet, source_data: WorkbookData) -> None:
        sheet.append(source_data.headers)
        for row in source_data.rows:
            padded = list(row) + [None] * max(0, len(source_data.headers) - len(row))
            try:
                sheet.append(padded[: len(source_data.headers)])
            except IllegalCharacterError:
                sheet.append([self._safe_cell_value(value) for value in padded[: len(source_data.headers)]])

    def _write_summary_sheet(self, sheet, summary: FuncCodeSummary, style: ReportStyle) -> None:
        self.formatter.apply_sheet_layout(sheet, style)
        sheet["A1"] = "Func Code"
        sheet["B1"] = summary.func_code
        self.formatter.style_row(sheet, 1, "title", style)
        self.formatter.style_row(sheet, 2, "blank", style)
        sheet["A3"] = "행 레이블"
        sheet["B3"] = "합계 : Loc Amt"
        self.formatter.style_row(sheet, 3, "header", style)

        row_index = 4
        for month in summary.months:
            month_row = row_index
            sheet.cell(row=row_index, column=1, value=month.label)
            sheet.cell(row=row_index, column=2, value=self._numeric(month.amount))
            self.formatter.style_row(sheet, row_index, "month", style)
            row_index += 1
            for port in month.ports:
                port_row = row_index
                sheet.cell(row=row_index, column=1, value=port.name)
                sheet.cell(row=row_index, column=2, value=self._numeric(port.amount))
                self.formatter.style_row(sheet, row_index, "port", style)
                row_index += 1
                for customer in port.customers:
                    sheet.cell(row=row_index, column=1, value=customer.name)
                    sheet.cell(row=row_index, column=2, value=self._numeric(customer.amount))
                    self.formatter.style_row(sheet, row_index, "customer", style)
                    sheet.row_dimensions[row_index].outlineLevel = 2
                    row_index += 1
                sheet.row_dimensions[port_row].outlineLevel = 1

        sheet.cell(row=row_index, column=1, value="총합계")
        sheet.cell(row=row_index, column=2, value=self._numeric(summary.amount))
        self.formatter.style_row(sheet, row_index, "grand_total", style)

    @staticmethod
    def _numeric(amount: Decimal) -> int | float:
        if amount == amount.to_integral_value():
            return int(amount)
        return float(amount)

    @staticmethod
    def _safe_cell_value(value: Any) -> Any:
        if isinstance(value, str):
            return "".join(char for char in value if ord(char) >= 32 or char in "\t\n\r")
        return value
