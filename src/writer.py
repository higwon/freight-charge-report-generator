from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.exceptions import IllegalCharacterError

from .config import ALL_SHEET_NAME, SOURCE_DATA_SHEET_NAME
from .exceptions import OutputPermissionError
from .formatter import Formatter, ReportStyle
from .models import FuncCodeSummary, ReportFormat, WorkbookData
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
        report_format: ReportFormat = ReportFormat.CLASSIC,
    ) -> None:
        workbook = Workbook()
        if report_format == ReportFormat.ANALYTIC:
            self._write_analytic_workbook(workbook, source_data, summaries, style)
        else:
            self._write_classic_workbook(workbook, source_data, summaries, style)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            workbook.save(output_path)
        except PermissionError as exc:
            raise OutputPermissionError(
                "출력 파일을 저장할 수 없습니다. 결과 파일이 열려 있다면 닫은 뒤 다시 시도하세요."
            ) from exc

    def _write_classic_workbook(
        self, workbook: Workbook, source_data: WorkbookData, summaries: list[FuncCodeSummary], style: ReportStyle
    ) -> None:
        source_sheet = workbook.active
        source_sheet.title = SOURCE_DATA_SHEET_NAME
        self._write_raw_data(source_sheet, source_data)
        self.formatter.style_raw_sheet(source_sheet)
        for summary in summaries:
            sheet = workbook.create_sheet(clean_sheet_title(summary.func_code))
            self._write_summary_sheet(sheet, summary, style)

    def _write_analytic_workbook(
        self, workbook: Workbook, source_data: WorkbookData, summaries: list[FuncCodeSummary], style: ReportStyle
    ) -> None:
        func_summaries = [summary for summary in summaries if summary.func_code != ALL_SHEET_NAME]
        overview = workbook.active
        overview.title = "Overview"
        self._write_overview(overview, source_data, func_summaries, style)

        all_sheet = workbook.create_sheet(ALL_SHEET_NAME)
        self._write_analytic_all(all_sheet, func_summaries, style)
        for summary in func_summaries:
            sheet = workbook.create_sheet(clean_sheet_title(summary.func_code))
            self._write_summary_sheet(sheet, summary, style)

        source_sheet = workbook.create_sheet(SOURCE_DATA_SHEET_NAME)
        self._write_raw_data(source_sheet, source_data)
        self.formatter.style_raw_sheet(source_sheet, enable_filter=True)

    def _write_overview(
        self, sheet, source_data: WorkbookData, summaries: list[FuncCodeSummary], style: ReportStyle
    ) -> None:
        sheet.sheet_view.showGridLines = False
        sheet.freeze_panes = "A5"
        sheet.merge_cells("A1:K1")
        sheet["A1"] = "Freight Charge Overview"
        sheet["A1"].font = Font(name="맑은 고딕", size=18, bold=True, color="FFFFFF")
        sheet["A1"].fill = PatternFill("solid", fgColor="1F4E78")
        sheet["A1"].alignment = Alignment(horizontal="left", vertical="center")
        sheet.row_dimensions[1].height = 30

        total_amount = sum((summary.amount for summary in summaries), Decimal("0"))
        metadata = (
            ("A3", "원본 시트", "B3", source_data.source_sheet_name),
            ("D3", "Source 행 수", "E3", source_data.record_count),
            ("G3", "총 운임", "H3", self._numeric(total_amount)),
            ("J3", "생성 일시", "K3", datetime.now().strftime("%Y-%m-%d %H:%M")),
        )
        for label_cell, label, value_cell, value in metadata:
            sheet[label_cell] = label
            sheet[label_cell].font = Font(name="맑은 고딕", bold=True, color="44546A")
            sheet[value_cell] = value
        sheet["H3"].number_format = style.amount_format

        monthly: dict[str, Decimal] = defaultdict(Decimal)
        ports: dict[str, Decimal] = defaultdict(Decimal)
        customers: dict[str, Decimal] = defaultdict(Decimal)
        for summary in summaries:
            for month in summary.months:
                monthly[month.label] += month.amount
                for port in month.ports:
                    ports[port.name] += port.amount
                    for customer in port.customers:
                        customers[customer.name] += customer.amount

        func_end = self._write_overview_table(
            sheet, 5, 1, "Func Code", [(s.func_code, s.amount) for s in summaries], style
        )
        month_end = self._write_overview_table(sheet, 5, 4, "월별 합계", sorted(monthly.items()), style)
        self._write_overview_table(
            sheet, 5, 7, "상위 Port", sorted(ports.items(), key=lambda item: item[1], reverse=True)[:10], style
        )
        self._write_overview_table(
            sheet, 5, 10, "상위 Customer", sorted(customers.items(), key=lambda item: item[1], reverse=True)[:10], style
        )

        if func_end >= 6:
            chart = BarChart()
            chart.title = "Func Code별 운임"
            chart.y_axis.title = "Loc Amt"
            chart.add_data(Reference(sheet, min_col=2, min_row=5, max_row=func_end), titles_from_data=True)
            chart.set_categories(Reference(sheet, min_col=1, min_row=6, max_row=func_end))
            chart.height = 7
            chart.width = 12
            sheet.add_chart(chart, "A17")
        if month_end >= 6:
            chart = LineChart()
            chart.title = "월별 운임 추이"
            chart.y_axis.title = "Loc Amt"
            chart.add_data(Reference(sheet, min_col=5, min_row=5, max_row=month_end), titles_from_data=True)
            chart.set_categories(Reference(sheet, min_col=4, min_row=6, max_row=month_end))
            chart.height = 7
            chart.width = 12
            sheet.add_chart(chart, "G17")

        for column, width in {"A": 16, "B": 16, "D": 16, "E": 16, "G": 24, "H": 16, "J": 42, "K": 16}.items():
            sheet.column_dimensions[column].width = width

    @staticmethod
    def _write_overview_table(sheet, start_row: int, start_col: int, title: str, rows, style: ReportStyle) -> int:
        header_fill = PatternFill("solid", fgColor="D9EAF7")
        border = Border(bottom=Side(style="thin", color="9EB6CE"))
        sheet.cell(start_row, start_col, title)
        sheet.cell(start_row, start_col + 1, "Loc Amt")
        for cell in (sheet.cell(start_row, start_col), sheet.cell(start_row, start_col + 1)):
            cell.font = Font(name="맑은 고딕", bold=True, color="1F2937")
            cell.fill = header_fill
            cell.border = border
        for offset, (label, amount) in enumerate(rows, start=1):
            row = start_row + offset
            sheet.cell(row, start_col, label)
            amount_cell = sheet.cell(row, start_col + 1, ReportWriter._numeric(amount))
            amount_cell.number_format = style.amount_format
        return start_row + len(rows)

    def _write_analytic_all(self, sheet, summaries: list[FuncCodeSummary], style: ReportStyle) -> None:
        self.formatter.apply_sheet_layout(sheet, style)
        sheet["A1"] = "통합 보고서"
        sheet["B1"] = ALL_SHEET_NAME
        self.formatter.style_row(sheet, 1, "title", style)
        self.formatter.style_row(sheet, 2, "blank", style)
        sheet["A3"] = "행 레이블"
        sheet["B3"] = "합계 : Loc Amt"
        self.formatter.style_row(sheet, 3, "header", style)

        month_lookup = {
            summary.func_code: {month.label: month for month in summary.months} for summary in summaries
        }
        months = sorted({month.label for summary in summaries for month in summary.months})
        row_index = 4
        for month_label in months:
            month_total = sum(
                (month_lookup[s.func_code][month_label].amount for s in summaries if month_label in month_lookup[s.func_code]),
                Decimal("0"),
            )
            sheet.cell(row_index, 1, month_label)
            sheet.cell(row_index, 2, self._numeric(month_total))
            self.formatter.style_row(sheet, row_index, "month", style)
            row_index += 1
            for summary in summaries:
                month = month_lookup[summary.func_code].get(month_label)
                if month is None:
                    continue
                sheet.cell(row_index, 1, summary.func_code)
                sheet.cell(row_index, 2, self._numeric(month.amount))
                self.formatter.style_row(sheet, row_index, "func", style)
                sheet.row_dimensions[row_index].outlineLevel = 1
                row_index += 1
                for port in month.ports:
                    sheet.cell(row_index, 1, port.name)
                    sheet.cell(row_index, 2, self._numeric(port.amount))
                    self.formatter.style_row(sheet, row_index, "analytic_port", style)
                    sheet.row_dimensions[row_index].outlineLevel = 2
                    row_index += 1
                    for customer in port.customers:
                        sheet.cell(row_index, 1, customer.name)
                        sheet.cell(row_index, 2, self._numeric(customer.amount))
                        self.formatter.style_row(sheet, row_index, "analytic_customer", style)
                        sheet.row_dimensions[row_index].outlineLevel = 3
                        row_index += 1

        sheet.cell(row_index, 1, "총합계")
        sheet.cell(row_index, 2, self._numeric(sum((s.amount for s in summaries), Decimal("0"))))
        self.formatter.style_row(sheet, row_index, "grand_total", style)

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
