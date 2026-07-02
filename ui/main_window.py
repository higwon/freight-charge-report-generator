from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, Qt, QUrl
from PySide6.QtGui import QDesktopServices, QIcon, QPixmap
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

from src.config import APP_AUTHOR, APP_NAME, APP_VERSION
from src.models import GenerationRequest, GenerationResult
from src.utils import ensure_xlsx_suffix, resource_path
from ui.worker import GenerateWorker


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(QIcon(str(resource_path("assets/app_icon.ico"))))
        self.resize(900, 540)
        self.setMinimumSize(760, 500)
        self.worker_thread: QThread | None = None
        self.worker: GenerateWorker | None = None
        self.last_output_path: Path | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        central = QWidget(self)
        root = QVBoxLayout(central)
        root.setContentsMargins(24, 18, 24, 18)
        root.setSpacing(12)

        header = QHBoxLayout()
        header.setSpacing(10)
        mascot = QLabel()
        mascot.setFixedSize(48, 48)
        mascot.setPixmap(self._pixmap("assets/header_mascot.png", 48, 48))
        mascot.setAlignment(Qt.AlignCenter)
        title = QLabel(APP_NAME)
        title.setObjectName("Title")
        header.addWidget(mascot)
        header.addWidget(title)
        header.addStretch(1)
        metadata = QLabel(f"v{APP_VERSION}\n{APP_AUTHOR}")
        metadata.setObjectName("Metadata")
        metadata.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        header.addWidget(metadata)
        root.addLayout(header)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(10)
        self.source_edit = QLineEdit()
        self.output_folder_edit = QLineEdit(str(Path.cwd() / "output"))
        self.output_name_edit = QLineEdit("result.xlsx")
        form.addRow("Source Excel", self._file_row(self.source_edit, self._choose_source))
        form.addRow("Output Folder", self._file_row(self.output_folder_edit, self._choose_output_folder))
        form.addRow("Output Filename", self.output_name_edit)
        root.addLayout(form)

        progress_panel = QFrame()
        progress_panel.setObjectName("Panel")
        progress_layout = QVBoxLayout(progress_panel)
        progress_layout.setContentsMargins(14, 10, 14, 10)
        progress_layout.setSpacing(6)
        progress_header = QHBoxLayout()
        progress_title = QLabel("진행 상태")
        progress_title.setObjectName("PanelTitle")
        self.status_label = QLabel("작업 대기 중")
        self.status_label.setObjectName("StatusLabel")
        self.progress_percent = QLabel("0%")
        self.progress_percent.setObjectName("ProgressPercent")
        progress_header.addWidget(progress_title)
        progress_header.addSpacing(18)
        progress_header.addWidget(self.status_label, 1)
        progress_header.addWidget(self.progress_percent)
        progress_layout.addLayout(progress_header)
        self.progress_detail = QLabel("Source Excel을 선택한 뒤 생성을 시작하세요.")
        self.progress_detail.setObjectName("DetailLabel")
        progress_layout.addWidget(self.progress_detail)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        progress_layout.addWidget(self.progress)
        root.addWidget(progress_panel)

        result_panel = QFrame()
        result_panel.setObjectName("Panel")
        result_layout = QHBoxLayout(result_panel)
        result_layout.setContentsMargins(14, 10, 14, 10)
        result_layout.setSpacing(14)
        self.result_image = QLabel()
        self.result_image.setFixedSize(76, 76)
        self.result_image.setAlignment(Qt.AlignCenter)
        self.result_image.setVisible(False)
        result_layout.addWidget(self.result_image)
        result_text = QVBoxLayout()
        result_title = QLabel("결과 요약")
        result_title.setObjectName("PanelTitle")
        self.summary_label = QLabel("아직 생성된 결과가 없습니다.")
        self.summary_label.setObjectName("SummaryLabel")
        self.summary_label.setWordWrap(True)
        result_text.addWidget(result_title)
        result_text.addWidget(self.summary_label, 1)
        result_layout.addLayout(result_text, 1)
        root.addWidget(result_panel, 1)

        actions = QHBoxLayout()
        actions.addStretch(1)
        self.open_file_button = QPushButton("결과 열기")
        self.open_file_button.setEnabled(False)
        self.open_file_button.clicked.connect(self._open_result_file)
        self.open_folder_button = QPushButton("폴더 열기")
        self.open_folder_button.setEnabled(False)
        self.open_folder_button.clicked.connect(self._open_output_folder)
        self.generate_button = QPushButton("생성")
        self.generate_button.setObjectName("PrimaryButton")
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
            QLabel#Title { font-size: 20pt; font-weight: 700; color: #111827; }
            QLabel#Metadata { color: #4b5563; line-height: 1.4; }
            QLabel#PanelTitle { font-size: 11pt; font-weight: 700; color: #111827; }
            QLabel#StatusLabel { font-weight: 600; color: #1f2937; }
            QLabel#ProgressPercent { font-size: 13pt; font-weight: 700; color: #1d4ed8; }
            QLabel#DetailLabel, QLabel#SummaryLabel { color: #4b5563; }
            QFrame#Panel { border: 1px solid #d5dbe3; border-radius: 6px; background: #ffffff; }
            QLineEdit { padding: 7px 9px; border: 1px solid #c7ccd4; border-radius: 4px; background: white; }
            QPushButton { min-width: 82px; padding: 8px 14px; border-radius: 4px; border: 1px solid #9aa5b1; background: #ffffff; }
            QPushButton:hover { background: #eef3f8; }
            QPushButton:disabled { color: #87909a; background: #eceff3; }
            QPushButton#PrimaryButton { color: white; font-weight: 700; border-color: #1d4ed8; background: #1d4ed8; }
            QPushButton#PrimaryButton:hover { background: #1e40af; }
            QProgressBar { height: 12px; border: 1px solid #c7ccd4; border-radius: 4px; background: white; }
            QProgressBar::chunk { background: #2563eb; border-radius: 3px; }
            """
        )

    @staticmethod
    def _pixmap(relative_path: str, width: int, height: int) -> QPixmap:
        return QPixmap(str(resource_path(relative_path))).scaled(
            width, height, Qt.KeepAspectRatio, Qt.SmoothTransformation
        )

    def _file_row(self, edit: QLineEdit, callback) -> QWidget:
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        button = QPushButton("찾아보기")
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
        self.result_image.setVisible(False)
        self.progress.setValue(0)
        self.progress_percent.setText("0%")
        self.status_label.setText("파일 확인 중")
        self.progress_detail.setText(f"입력 파일: {source.name}")
        self.summary_label.setText("결과 파일을 생성하고 있습니다.")

        self.worker_thread = QThread(self)
        self.worker = GenerateWorker(request)
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._update_progress)
        self.worker.status.connect(self.progress_detail.setText)
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
        self.progress_percent.setText(f"{value}%")
        stages = {10: "Source 탐색 중", 35: "데이터 읽기 완료", 50: "금액 집계 중", 70: "요약 생성 중", 85: "Excel 작성 중", 100: "완료"}
        self.status_label.setText(stages.get(value, "처리 중"))

    def _generation_finished(self, result: GenerationResult) -> None:
        self.last_output_path = result.output_path
        self.generate_button.setEnabled(True)
        self.open_file_button.setEnabled(True)
        self.open_folder_button.setEnabled(True)
        self.result_image.setPixmap(self._pixmap("assets/status_success.png", 76, 76))
        self.result_image.setVisible(True)
        self.status_label.setText("완료 · 결과 파일 저장")
        self.progress_detail.setText(f"{result.summary_sheet_count}/{result.summary_sheet_count} 요약 시트 생성 완료")
        self.summary_label.setText(
            f"생성 완료\nSource: {result.source_sheet_name} · {result.record_count:,}행\n"
            f"요약 시트: ALL, {', '.join(result.func_codes)}\n출력 파일: {result.output_path}"
        )
        QMessageBox.information(self, "완료", f"결과 파일을 생성했습니다.\n{result.output_path}")

    def _generation_failed(self, message: str) -> None:
        self.generate_button.setEnabled(True)
        self.open_file_button.setEnabled(bool(self.last_output_path))
        self.open_folder_button.setEnabled(bool(self.last_output_path))
        self.progress.setValue(0)
        self.progress_percent.setText("실패")
        self.status_label.setText("생성 실패")
        self.progress_detail.setText("입력 파일과 오류 내용을 확인하세요.")
        self.result_image.setPixmap(self._pixmap("assets/status_failure.png", 76, 76))
        self.result_image.setVisible(True)
        self.summary_label.setText(message)
        QMessageBox.critical(self, "생성 실패", message)

    def _open_output_folder(self) -> None:
        folder = self.last_output_path.parent if self.last_output_path else Path(self.output_folder_edit.text())
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))

    def _open_result_file(self) -> None:
        if self.last_output_path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.last_output_path)))
