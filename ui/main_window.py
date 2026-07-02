from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QVBoxLayout,
    QWidget,
)

from src.config import APP_NAME
from src.models import GenerationRequest, GenerationResult
from src.utils import ensure_xlsx_suffix
from ui.worker import GenerateWorker


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(820, 430)
        self.worker_thread: QThread | None = None
        self.worker: GenerateWorker | None = None
        self.last_output_path: Path | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        central = QWidget(self)
        root = QVBoxLayout(central)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(16)

        title = QLabel(APP_NAME)
        title.setObjectName("Title")
        title.setAlignment(Qt.AlignLeft)
        root.addWidget(title)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.setFormAlignment(Qt.AlignTop)
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(12)

        self.source_edit = QLineEdit()
        self.output_folder_edit = QLineEdit(str(Path.cwd() / "output"))
        self.output_name_edit = QLineEdit("result.xlsx")

        form.addRow("Source Excel", self._file_row(self.source_edit, self._choose_source))
        form.addRow("Output Folder", self._file_row(self.output_folder_edit, self._choose_output_folder))
        form.addRow("Output Filename", self.output_name_edit)
        root.addLayout(form)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFormat("대기 중")
        root.addWidget(self.progress)

        status_panel = QFrame()
        status_panel.setObjectName("StatusPanel")
        status_layout = QVBoxLayout(status_panel)
        status_layout.setContentsMargins(14, 12, 14, 12)
        status_layout.setSpacing(6)

        self.status_label = QLabel("작업 대기 중입니다.")
        self.status_label.setObjectName("StatusLabel")
        self.status_label.setWordWrap(True)
        self.summary_label = QLabel("아직 생성된 결과가 없습니다.")
        self.summary_label.setObjectName("SummaryLabel")
        self.summary_label.setWordWrap(True)
        status_layout.addWidget(self.status_label)
        status_layout.addWidget(self.summary_label)
        root.addWidget(status_panel)

        actions = QHBoxLayout()
        actions.addStretch(1)
        self.open_file_button = QPushButton("Open Result")
        self.open_file_button.setEnabled(False)
        self.open_file_button.clicked.connect(self._open_result_file)
        self.open_folder_button = QPushButton("Open Output Folder")
        self.open_folder_button.setEnabled(False)
        self.open_folder_button.clicked.connect(self._open_output_folder)
        self.generate_button = QPushButton("Generate")
        self.generate_button.clicked.connect(self._generate)
        actions.addWidget(self.open_file_button)
        actions.addWidget(self.open_folder_button)
        actions.addWidget(self.generate_button)
        root.addLayout(actions)

        self.setCentralWidget(central)
        self.setStyleSheet(
            """
            QWidget { font-family: 'Segoe UI', 'Malgun Gothic'; font-size: 10.5pt; color: #20242a; }
            QMainWindow { background: #f6f7f9; }
            QLabel#Title { font-size: 18pt; font-weight: 700; color: #20242a; }
            QLabel#StatusLabel { font-weight: 600; color: #1f2937; }
            QLabel#SummaryLabel { color: #4b5563; }
            QFrame#StatusPanel { border: 1px solid #d5dbe3; border-radius: 6px; background: #ffffff; }
            QLineEdit { padding: 7px 9px; border: 1px solid #c7ccd4; border-radius: 4px; background: white; }
            QPushButton { padding: 8px 14px; border-radius: 4px; border: 1px solid #9aa5b1; background: #ffffff; }
            QPushButton:hover { background: #eef3f8; }
            QPushButton:disabled { color: #87909a; background: #eceff3; }
            QProgressBar { height: 20px; border: 1px solid #c7ccd4; border-radius: 4px; background: white; text-align: center; }
            QProgressBar::chunk { background: #2b6cb0; border-radius: 3px; }
            """
        )

    def _file_row(self, edit: QLineEdit, callback) -> QWidget:
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        button = QPushButton("Browse")
        button.clicked.connect(callback)
        layout.addWidget(edit, 1)
        layout.addWidget(button)
        return container

    def _choose_source(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Source Excel 선택",
            str(Path.home()),
            "Excel Files (*.xls *.xlsx *.xls.xls *.xlsx.xlsx);;All Files (*)",
        )
        if path:
            self.source_edit.setText(path)

    def _choose_output_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Output Folder 선택", self.output_folder_edit.text())
        if path:
            self.output_folder_edit.setText(path)

    def _generate(self) -> None:
        source = Path(self.source_edit.text().strip())
        output_folder = Path(self.output_folder_edit.text().strip())
        filename = ensure_xlsx_suffix(self.output_name_edit.text())
        request = GenerationRequest(source_path=source, output_path=output_folder / filename)

        if not source.exists():
            QMessageBox.warning(self, "입력 오류", "Source Excel 파일을 선택하세요.")
            return

        self.generate_button.setEnabled(False)
        self.open_file_button.setEnabled(False)
        self.open_folder_button.setEnabled(False)
        self.progress.setValue(0)
        self.progress.setFormat("0%")
        self.status_label.setText("작업을 시작합니다.")
        self.summary_label.setText(f"입력 파일: {source.name}")

        self.worker_thread = QThread(self)
        self.worker = GenerateWorker(request)
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._update_progress)
        self.worker.status.connect(self.status_label.setText)
        self.worker.finished.connect(self._generation_finished)
        self.worker.failed.connect(self._generation_failed)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker_thread.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.finished.connect(lambda: setattr(self, "worker", None))
        self.worker_thread.finished.connect(lambda: setattr(self, "worker_thread", None))
        self.worker_thread.start()

    def _update_progress(self, value: int) -> None:
        self.progress.setValue(value)
        self.progress.setFormat(f"{value}%")

    def _generation_finished(self, result: GenerationResult) -> None:
        self.last_output_path = result.output_path
        self.generate_button.setEnabled(True)
        self.open_file_button.setEnabled(True)
        self.open_folder_button.setEnabled(True)
        self.status_label.setText("성공: 결과 파일 생성이 완료되었습니다.")
        self.summary_label.setText(
            "결과 요약: "
            f"RawData {result.record_count:,}행, "
            f"요약 시트 {result.summary_sheet_count}개({', '.join(result.func_codes)}), "
            f"저장 위치 {result.output_path}"
        )
        QMessageBox.information(self, "완료", f"결과 파일을 생성했습니다.\n{result.output_path}")

    def _generation_failed(self, message: str) -> None:
        self.generate_button.setEnabled(True)
        self.open_file_button.setEnabled(bool(self.last_output_path))
        self.open_folder_button.setEnabled(bool(self.last_output_path))
        self.progress.setValue(0)
        self.progress.setFormat("실패")
        self.status_label.setText("실패: 결과 파일을 생성하지 못했습니다.")
        self.summary_label.setText(message)
        QMessageBox.critical(self, "생성 실패", message)

    def _open_output_folder(self) -> None:
        folder = self.last_output_path.parent if self.last_output_path else Path(self.output_folder_edit.text())
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))

    def _open_result_file(self) -> None:
        if self.last_output_path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.last_output_path)))
