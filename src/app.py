from __future__ import annotations

import logging
import sys
from pathlib import Path
from collections.abc import Callable

from .config import ALL_SHEET_NAME, APP_NAME
from .exceptions import ReportGeneratorError
from .formatter import Formatter
from .models import GenerationRequest, GenerationResult
from .reader import SourceWorkbookReader
from .transformer import ReportTransformer
from .utils import resource_path, setup_logging
from .writer import ReportWriter

LOGGER = logging.getLogger(__name__)


class ReportGeneratorService:
    def __init__(
        self,
        reader: SourceWorkbookReader | None = None,
        transformer: ReportTransformer | None = None,
        formatter: Formatter | None = None,
        writer: ReportWriter | None = None,
    ) -> None:
        self.reader = reader or SourceWorkbookReader()
        self.transformer = transformer or ReportTransformer()
        self.formatter = formatter or Formatter()
        self.writer = writer or ReportWriter(self.formatter)

    def generate(self, request: GenerationRequest, progress: Callable[[int, str], None] | None = None) -> GenerationResult:
        def emit(percent: int, message: str) -> None:
            if progress:
                progress(percent, message)

        LOGGER.info("Conversion start")
        LOGGER.info("Input file: %s", request.source_path)
        LOGGER.info("Output file: %s", request.output_path)

        emit(10, "원본 파일을 여는 중입니다.")
        source_data = self.reader.read(request.source_path)
        emit(35, f"'{source_data.source_sheet_name}' 시트에서 Source {source_data.record_count:,}행을 읽었습니다.")

        style = self.formatter.default_style()
        emit(50, "월, 포트, 고객 기준으로 금액을 집계하는 중입니다.")
        summaries = self.transformer.transform(source_data, month_format=style.month_format)
        func_codes = [summary.func_code for summary in summaries if summary.func_code != ALL_SHEET_NAME]
        emit(70, f"ALL 및 Func Code {len(func_codes)}개를 집계했습니다: {', '.join(func_codes)}")

        LOGGER.info("Number of records: %s", source_data.record_count)
        LOGGER.info("Func Codes found: %s", ", ".join(func_codes))

        emit(85, "Source와 요약 시트를 작성하고 있습니다.")
        self.writer.write(request.output_path, source_data, summaries, style)
        emit(100, f"저장 완료: {request.output_path.name}")
        LOGGER.info("Success")
        return GenerationResult(
            output_path=request.output_path,
            record_count=source_data.record_count,
            source_sheet_name=source_data.source_sheet_name,
            func_codes=func_codes,
            summary_sheet_count=len(summaries),
        )


def main() -> int:
    setup_logging()
    from PySide6.QtGui import QIcon
    from PySide6.QtWidgets import QApplication

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setWindowIcon(QIcon(str(resource_path("assets/app_icon.ico"))))

    try:
        from ui.main_window import MainWindow

        window = MainWindow()
        window.show()
        return app.exec()
    except ReportGeneratorError:
        LOGGER.exception("Application error")
        raise
    except Exception:
        LOGGER.exception("Unexpected application error")
        raise
