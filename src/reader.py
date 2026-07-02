from __future__ import annotations

from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
import zipfile

from openpyxl import load_workbook

from .config import SOURCE_SHEET_NAME
from .exceptions import MissingInputFileError, MissingWorksheetError
from .models import WorkbookData

SS_NS = "{urn:schemas-microsoft-com:office:spreadsheet}"


class SourceWorkbookReader:
    def read(self, path: Path, sheet_name: str = SOURCE_SHEET_NAME) -> WorkbookData:
        if not path.exists():
            raise MissingInputFileError(f"입력 파일을 찾을 수 없습니다: {path}")

        if self._is_excel_xml(path):
            return self._read_excel_xml(path, sheet_name)
        if zipfile.is_zipfile(path):
            return self._read_xlsx(path, sheet_name)
        return self._read_xls(path, sheet_name)

    @staticmethod
    def _is_excel_xml(path: Path) -> bool:
        with path.open("rb") as file:
            return file.read(64).lstrip().startswith(b"<?xml")

    def _read_excel_xml(self, path: Path, sheet_name: str) -> WorkbookData:
        headers: list[str] | None = None
        rows: list[list[Any]] = []
        found_sheet = False
        in_target = False

        for event, elem in ET.iterparse(path, events=("start", "end")):
            if event == "start" and elem.tag.endswith("Worksheet"):
                current = elem.attrib.get(f"{SS_NS}Name")
                in_target = current == sheet_name
                found_sheet = found_sheet or in_target
            elif in_target and event == "end" and elem.tag.endswith("Row"):
                values = self._xml_row_values(elem)
                if headers is None:
                    headers = ["" if value is None else str(value) for value in values]
                else:
                    rows.append(values)
                elem.clear()

        if not found_sheet or headers is None:
            raise MissingWorksheetError(f"원본 파일에 '{sheet_name}' 시트가 없습니다.")
        return WorkbookData(headers=headers, rows=rows)

    @staticmethod
    def _xml_row_values(row: ET.Element) -> list[Any]:
        values: list[Any] = []
        column_index = 1
        for cell in row:
            if not cell.tag.endswith("Cell"):
                continue
            explicit_index = cell.attrib.get(f"{SS_NS}Index")
            if explicit_index:
                while column_index < int(explicit_index):
                    values.append(None)
                    column_index += 1
            values.append(SourceWorkbookReader._xml_cell_value(cell))
            column_index += 1
        return values

    @staticmethod
    def _xml_cell_value(cell: ET.Element) -> Any:
        for child in cell:
            if child.tag.endswith("Data"):
                return child.text
        return None

    @staticmethod
    def _read_xlsx(path: Path, sheet_name: str) -> WorkbookData:
        workbook = load_workbook(path, data_only=True, read_only=True)
        if sheet_name not in workbook.sheetnames:
            raise MissingWorksheetError(f"원본 파일에 '{sheet_name}' 시트가 없습니다.")
        sheet = workbook[sheet_name]
        iterator = sheet.iter_rows(values_only=True)
        try:
            headers = ["" if value is None else str(value) for value in next(iterator)]
        except StopIteration as exc:
            raise MissingWorksheetError(f"'{sheet_name}' 시트가 비어 있습니다.") from exc
        return WorkbookData(headers=headers, rows=[list(row) for row in iterator])

    @staticmethod
    def _read_xls(path: Path, sheet_name: str) -> WorkbookData:
        try:
            import xlrd
        except ImportError as exc:
            raise MissingInputFileError("구형 .xls 파일을 읽으려면 xlrd가 필요합니다.") from exc

        workbook = xlrd.open_workbook(str(path))
        if sheet_name not in workbook.sheet_names():
            raise MissingWorksheetError(f"원본 파일에 '{sheet_name}' 시트가 없습니다.")
        sheet = workbook.sheet_by_name(sheet_name)
        if sheet.nrows == 0:
            raise MissingWorksheetError(f"'{sheet_name}' 시트가 비어 있습니다.")
        headers = ["" if value is None else str(value) for value in sheet.row_values(0)]
        rows = [sheet.row_values(index) for index in range(1, sheet.nrows)]
        return WorkbookData(headers=headers, rows=rows)
