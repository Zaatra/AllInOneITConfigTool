"""System configuration status dashboard."""
from __future__ import annotations

from pathlib import Path
from typing import Callable, Dict

from PySide6.QtCore import Qt, QThreadPool
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QGridLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from services.system_config import ApplyStepResult, ConfigCheckResult, SystemConfigService
from allinone_it_config.constants import FixedSystemConfig
from ui.workers import ServiceWorker

LogCallback = Callable[[str], None]


class SystemTab(QWidget):
    def __init__(
        self,
        config: FixedSystemConfig,
        log_callback: LogCallback,
        thread_pool: QThreadPool,
    ) -> None:
        super().__init__()
        self._config = config
        self._log = log_callback
        self._thread_pool = thread_pool
        self._service = SystemConfigService(config)
        self._status_labels: Dict[str, QLabel] = {}
        self._workers: set[ServiceWorker] = set()
        self._busy = False
        self._build_ui()
        self._start_check()

    def _track_worker(self, worker: ServiceWorker) -> None:
        self._workers.add(worker)
        worker.signals.finished.connect(lambda *_: self._workers.discard(worker))
        worker.signals.error.connect(lambda *_: self._workers.discard(worker))

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        description = QLabel("System status compared against All-In-One IT Configuration Tool policy values")
        layout.addWidget(description)

        grid = QGridLayout()
        layout.addLayout(grid)

        entries = [
            ("Timezone", "System timezone"),
            ("Power Plan", "Active power profile"),
            ("Fast Boot", "Fast startup registry"),
            ("Desktop Icons", "Desktop icon visibility"),
            ("Locale", "System locale & date format"),
            ("Default User Profile", "Default user profile (new users)"),
        ]
        for row, (key, caption) in enumerate(entries):
            label = QLabel(caption)
            value_label = QLabel("Checking...")
            value_label.setAlignment(Qt.AlignLeft)
            value_label.setProperty("statusKey", key)
            grid.addWidget(label, row, 0)
            grid.addWidget(value_label, row, 1)
            self._status_labels[key] = value_label

        self._btn_apply = QPushButton("Apply All")
        self._btn_apply.clicked.connect(self._start_apply)
        layout.addWidget(self._btn_apply)
        layout.addStretch()

    def _start_check(self) -> None:
        if self._busy:
            return
        self._busy = True
        self._btn_apply.setEnabled(False)
        for label in self._status_labels.values():
            label.setText("Checking...")
            label.setStyleSheet("")
        worker = ServiceWorker(self._service.check)
        worker.signals.finished.connect(self._handle_check_results)
        worker.signals.error.connect(self._handle_error)
        self._track_worker(worker)
        self._thread_pool.start(worker)

    def _handle_check_results(self, results: list[ConfigCheckResult]) -> None:
        failures = 0
        for result in results:
            label = self._status_labels.get(result.name)
            if not label:
                continue
            label.setText(self._format_result_text(result))
            label.setStyleSheet(self._format_result_style(result))
            if not result.in_desired_state:
                failures += 1
        self._busy = False
        self._btn_apply.setEnabled(True)
        summary = "All settings compliant." if failures == 0 else f"{failures} setting(s) require attention."
        self._log(summary)

    def _start_apply(self) -> None:
        if self._busy:
            QMessageBox.information(self, "In Progress", "Please wait for current operation to finish.")
            return
        self._busy = True
        self._btn_apply.setEnabled(False)
        self._log("Applying system configuration values...")
        worker = ServiceWorker(self._run_apply)
        worker.signals.finished.connect(self._handle_apply_finished)
        worker.signals.error.connect(self._handle_error)
        self._track_worker(worker)
        self._thread_pool.start(worker)

    def _run_apply(self) -> list[ApplyStepResult]:
        return self._service.apply_with_results()

    def _handle_apply_finished(self, results: list[ApplyStepResult] | None) -> None:
        if results:
            failures = 0
            for result in results:
                status = "OK" if result.success else "FAILED"
                detail = f" - {result.detail}" if result.detail else ""
                self._log(f"{result.name}: {status}{detail}")
                if not result.success:
                    failures += 1
            if failures:
                self._log(f"{failures} apply step(s) failed.")
        self._log("System configuration applied. Refreshing status...")
        self._busy = False
        self._start_check()

    def _format_result_text(self, result: ConfigCheckResult) -> str:
        status_icon = "✓" if result.in_desired_state else "✗"
        return f"{status_icon} {result.actual} (target: {result.expected})"

    def _format_result_style(self, result: ConfigCheckResult) -> str:
        color = "#4caf50" if result.in_desired_state else "#f44336"
        return f"color: {color}; font-weight: bold;"

    def _handle_error(self, message: str) -> None:
        self._log(f"[ERROR] {message}")
        self._busy = False
        self._btn_apply.setEnabled(True)
