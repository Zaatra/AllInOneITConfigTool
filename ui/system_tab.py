"""System configuration status dashboard."""
from __future__ import annotations

from typing import Callable, Dict

from PySide6.QtCore import Qt, QThreadPool
from PySide6.QtWidgets import (
    QCheckBox,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from allinone_it_config.constants import FixedSystemConfig
from services.system_config import (
    ApplyStepResult,
    ConfigCheckResult,
    DiagnosticStepResult,
    SystemConfigService,
)
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
        self._setting_checks: Dict[str, QCheckBox] = {}
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

        entries = {
            "Timezone": "System timezone",
            "Power Plan": "Active power profile",
            "Fast Boot": "Fast startup registry",
            "Desktop Icons": "Desktop icon visibility",
            "Locale": "Locale, date/time format, language features, and region",
            "Default User Profile": "Default user profile (new users)",
            "Default Apps": "Chrome/Outlook default associations",
        }
        available_steps = self._service.available_apply_steps()

        grid.addWidget(QLabel("Apply"), 0, 0)
        grid.addWidget(QLabel("Setting"), 0, 1)
        grid.addWidget(QLabel("Status"), 0, 2)

        for row, key in enumerate(available_steps, start=1):
            caption = entries.get(key, key)
            checkbox = QCheckBox()
            checkbox.setChecked(True)
            label = QLabel(caption)
            value_label = QLabel("Checking...")
            value_label.setAlignment(Qt.AlignLeft)
            value_label.setProperty("statusKey", key)
            grid.addWidget(checkbox, row, 0, alignment=Qt.AlignCenter)
            grid.addWidget(label, row, 1)
            grid.addWidget(value_label, row, 2)
            self._setting_checks[key] = checkbox
            self._status_labels[key] = value_label

        button_row = QHBoxLayout()
        self._btn_select_all = QPushButton("Select All")
        self._btn_select_all.clicked.connect(lambda: self._set_all_selection(True))
        button_row.addWidget(self._btn_select_all)

        self._btn_deselect_all = QPushButton("Deselect All")
        self._btn_deselect_all.clicked.connect(lambda: self._set_all_selection(False))
        button_row.addWidget(self._btn_deselect_all)

        self._btn_apply = QPushButton("Apply Selected")
        self._btn_apply.clicked.connect(self._start_apply)
        button_row.addWidget(self._btn_apply)

        self._btn_diagnostics = QPushButton("Run Diagnostics")
        self._btn_diagnostics.clicked.connect(self._start_diagnostics)
        button_row.addWidget(self._btn_diagnostics)

        layout.addLayout(button_row)

        diagnostics_label = QLabel("Diagnostics")
        layout.addWidget(diagnostics_label)

        self._diagnostics_view = QPlainTextEdit()
        self._diagnostics_view.setReadOnly(True)
        self._diagnostics_view.setPlaceholderText("Run diagnostics to collect command and registry details.")
        self._diagnostics_view.setMinimumHeight(180)
        layout.addWidget(self._diagnostics_view)
        layout.addStretch()

    def _start_check(self) -> None:
        if self._busy:
            return
        self._busy = True
        self._set_controls_enabled(False)
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
        self._set_controls_enabled(True)
        summary = "All settings compliant." if failures == 0 else f"{failures} setting(s) require attention."
        self._log(summary)

    def _start_apply(self) -> None:
        if self._busy:
            QMessageBox.information(self, "In Progress", "Please wait for current operation to finish.")
            return
        selected_steps = [name for name, checkbox in self._setting_checks.items() if checkbox.isChecked()]
        if not selected_steps:
            QMessageBox.information(self, "No Selection", "Select at least one setting to apply.")
            return
        self._busy = True
        self._set_controls_enabled(False)
        self._log(f"Applying selected system settings: {', '.join(selected_steps)}")
        worker = ServiceWorker(self._run_apply, selected_steps)
        worker.signals.finished.connect(self._handle_apply_finished)
        worker.signals.error.connect(self._handle_error)
        self._track_worker(worker)
        self._thread_pool.start(worker)

    def _run_apply(self, selected_steps: list[str]) -> list[ApplyStepResult]:
        return self._service.apply_with_results(selected_steps)

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

    def _start_diagnostics(self) -> None:
        if self._busy:
            QMessageBox.information(self, "In Progress", "Please wait for current operation to finish.")
            return
        self._busy = True
        self._set_controls_enabled(False)
        self._diagnostics_view.setPlainText("Collecting diagnostics...")
        worker = ServiceWorker(self._service.diagnostics)
        worker.signals.finished.connect(self._handle_diagnostics_finished)
        worker.signals.error.connect(self._handle_error)
        self._track_worker(worker)
        self._thread_pool.start(worker)

    def _handle_diagnostics_finished(self, results: list[DiagnosticStepResult] | None) -> None:
        lines: list[str] = []
        if results:
            for result in results:
                lines.append(self._format_diagnostic_line(result))
        self._diagnostics_view.setPlainText("\n".join(lines))
        self._log("Diagnostics completed.")
        self._busy = False
        self._set_controls_enabled(True)

    def _set_all_selection(self, selected: bool) -> None:
        for checkbox in self._setting_checks.values():
            checkbox.setChecked(selected)

    def _set_controls_enabled(self, enabled: bool) -> None:
        self._btn_apply.setEnabled(enabled)
        self._btn_diagnostics.setEnabled(enabled)
        self._btn_select_all.setEnabled(enabled)
        self._btn_deselect_all.setEnabled(enabled)
        for checkbox in self._setting_checks.values():
            checkbox.setEnabled(enabled)

    def _format_result_text(self, result: ConfigCheckResult) -> str:
        status_icon = "✓" if result.in_desired_state else "✗"
        return f"{status_icon} {result.actual} (target: {result.expected})"

    def _format_diagnostic_line(self, result: DiagnosticStepResult) -> str:
        status = "OK" if result.success else "FAIL"
        detail = f" :: {result.detail}" if result.detail else ""
        return f"[{status}] {result.name}{detail}"

    def _format_result_style(self, result: ConfigCheckResult) -> str:
        color = "#4caf50" if result.in_desired_state else "#f44336"
        return f"color: {color}; font-weight: bold;"

    def _handle_error(self, message: str) -> None:
        self._log(f"[ERROR] {message}")
        self._diagnostics_view.appendPlainText(f"[ERROR] {message}")
        self._busy = False
        self._set_controls_enabled(True)
