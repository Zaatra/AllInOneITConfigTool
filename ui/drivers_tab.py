"""Drivers tab UI for scanning/downloading/installing HP drivers."""
from __future__ import annotations

from pathlib import Path
from typing import Callable, Iterable, List

from PySide6.QtCore import Qt, QThreadPool
from PySide6.QtWidgets import (
    QAbstractItemView,
    QHBoxLayout,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from services.drivers import DriverOperationResult, DriverRecord, DriverService
from ui.workers import ServiceWorker

LogCallback = Callable[[str], None]


class DriversTab(QWidget):
    def __init__(
        self,
        log_callback: LogCallback,
        thread_pool: QThreadPool,
        *,
        working_dir: Path | None = None,
    ) -> None:
        super().__init__()
        self._log = log_callback
        self._thread_pool = thread_pool
        self._service = DriverService(working_dir=working_dir or Path.cwd())
        self._records: list[DriverRecord] = []
        self._busy = False
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)

        button_row = QHBoxLayout()
        self._btn_scan = QPushButton("Scan Drivers")
        self._btn_download = QPushButton("Download Selected")
        self._btn_install = QPushButton("Install Selected")
        self._btn_select_all = QPushButton("Select All")
        self._btn_select_none = QPushButton("Select None")
        for btn in (self._btn_scan, self._btn_download, self._btn_install):
            btn.setMinimumWidth(150)
        button_row.addWidget(self._btn_scan)
        button_row.addWidget(self._btn_download)
        button_row.addWidget(self._btn_install)
        button_row.addStretch()
        button_row.addWidget(self._btn_select_all)
        button_row.addWidget(self._btn_select_none)
        layout.addLayout(button_row)

        self._table = QTableWidget(0, 6, self)
        self._table.setHorizontalHeaderLabels(["Select", "Source", "Name", "Installed", "Latest", "Status"])
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.verticalHeader().setVisible(False)
        header = self._table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, header.ResizeToContents)
        header.setSectionResizeMode(1, header.ResizeToContents)
        header.setSectionResizeMode(2, header.Stretch)
        header.setSectionResizeMode(3, header.ResizeToContents)
        header.setSectionResizeMode(4, header.ResizeToContents)
        header.setSectionResizeMode(5, header.ResizeToContents)
        layout.addWidget(self._table)

        self._btn_scan.clicked.connect(self._start_scan)
        self._btn_select_all.clicked.connect(lambda: self._set_all(Qt.Checked))
        self._btn_select_none.clicked.connect(lambda: self._set_all(Qt.Unchecked))
        self._btn_download.clicked.connect(lambda: self._start_operation("download"))
        self._btn_install.clicked.connect(lambda: self._start_operation("install"))

    def _start_scan(self) -> None:
        if self._busy:
            return
        self._busy = True
        self._set_buttons_enabled(False)
        self._log("Scanning for HP driver updates...")
        worker = ServiceWorker(self._service.scan)
        worker.signals.finished.connect(self._handle_scan_results)
        worker.signals.error.connect(self._handle_error)
        self._thread_pool.start(worker)

    def _handle_scan_results(self, records: Iterable[DriverRecord]) -> None:
        self._records = list(records)
        self._populate_table()
        self._log(f"Driver scan complete. Found {len(self._records)} entries.")
        self._busy = False
        self._set_buttons_enabled(True)

    def _start_operation(self, op: str) -> None:
        if self._busy:
            QMessageBox.information(self, "In Progress", "Wait for the current operation to finish.")
            return
        selected = self._selected_records()
        if not selected:
            QMessageBox.information(self, "No Selection", "Select at least one driver entry.")
            return
        self._busy = True
        self._set_buttons_enabled(False)
        action = self._service.download if op == "download" else self._service.install
        self._log(f"Running {op} for {len(selected)} driver(s)...")
        worker = ServiceWorker(action, selected)
        worker.signals.finished.connect(lambda result, op=op: self._handle_driver_results(op, result))
        worker.signals.error.connect(self._handle_error)
        self._thread_pool.start(worker)

    def _handle_driver_results(self, op: str, results: Iterable[DriverOperationResult]) -> None:
        for result in results:
            status = "OK" if result.success else "FAIL"
            self._log(f"[{status}] {op} :: {result.driver.name} -> {result.message}")
        if op == "download":
            self._populate_table()
        self._busy = False
        self._set_buttons_enabled(True)

    def _populate_table(self) -> None:
        self._table.setRowCount(len(self._records))
        for row, record in enumerate(self._records):
            checkbox = QTableWidgetItem()
            checkbox.setFlags(Qt.ItemIsSelectable | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            checkbox.setCheckState(Qt.Unchecked)
            checkbox.setData(Qt.UserRole, row)
            self._table.setItem(row, 0, checkbox)

            self._table.setItem(row, 1, QTableWidgetItem(record.source))
            self._table.setItem(row, 2, QTableWidgetItem(record.name))
            installed = record.installed_version or "Unknown"
            latest = record.latest_version or "Unknown"
            self._table.setItem(row, 3, QTableWidgetItem(installed))
            self._table.setItem(row, 4, QTableWidgetItem(latest))
            status_text = record.status
            if record.output_path:
                status_text += " (cached)"
            self._table.setItem(row, 5, QTableWidgetItem(status_text))

    def _selected_records(self) -> List[DriverRecord]:
        selections: list[DriverRecord] = []
        for row in range(self._table.rowCount()):
            item = self._table.item(row, 0)
            if item and item.checkState() == Qt.Checked:
                idx = item.data(Qt.UserRole)
                if isinstance(idx, int) and 0 <= idx < len(self._records):
                    selections.append(self._records[idx])
        return selections

    def _set_all(self, state: Qt.CheckState) -> None:
        for row in range(self._table.rowCount()):
            item = self._table.item(row, 0)
            if item:
                item.setCheckState(state)

    def _set_buttons_enabled(self, enabled: bool) -> None:
        for button in (self._btn_scan, self._btn_download, self._btn_install, self._btn_select_all, self._btn_select_none):
            button.setEnabled(enabled)

    def _handle_error(self, message: str) -> None:
        self._log(f"[ERROR] {message}")
        self._busy = False
        self._set_buttons_enabled(True)
*** End File
