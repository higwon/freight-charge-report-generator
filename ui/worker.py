from __future__ import annotations

import logging

from PySide6.QtCore import QObject, Signal, Slot

from src.app import ReportGeneratorService
from src.exceptions import ReportGeneratorError
from src.models import GenerationRequest, GenerationResult

LOGGER = logging.getLogger(__name__)


class GenerateWorker(QObject):
    progress = Signal(int)
    status = Signal(str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, request: GenerationRequest) -> None:
        super().__init__()
        self.request = request

    @Slot()
    def run(self) -> None:
        try:
            service = ReportGeneratorService()

            def report(percent: int, message: str) -> None:
                self.progress.emit(percent)
                self.status.emit(message)

            result: GenerationResult = service.generate(self.request, progress=report)
            self.finished.emit(result)
        except ReportGeneratorError as exc:
            LOGGER.exception("Generation failed")
            self.failed.emit(str(exc))
        except Exception as exc:
            LOGGER.exception("Unexpected generation failure")
            self.failed.emit(f"예상하지 못한 오류가 발생했습니다: {exc}")
