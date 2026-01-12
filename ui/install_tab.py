"""Installation tab with selectable applications and async actions."""
from __future__ import annotations

from pathlib import Path
from typing import Callable, Iterable

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

from services.installer import InstallerService, OperationResult
from trustpal.app_registry import AppEntry, AppRegistry
from ui.workers import ServiceWorker

LogCallback = Callable[[str], None]


class InstallTab(QWidget):
    def __init__(
        self,
        registry: AppRegistry,
        log_callback: LogCallback,
        thread_pool: QThreadPool,
        *,
        working_dir: Path | None = None,
    ) -> None:
        super().__init__()
        self._registry = registry
        self._log = log_callback
        self._thread_pool = thread_pool
        self._service = InstallerService(registry.entries, working_dir=working_dir or Path.cwd())
        self._busy = False
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)

        button_row = QHBoxLayout()
        self._btn_download = QPushButton("Download Selected")
        self._btn_install = QPushButton("Install Selected")
        self._btn_select_all = QPushButton("Select All")
        self._btn_select_none = QPushButton("Select None")
        button_row.addWidget(self._btn_download)
        button_row.addWidget(self._btn_install)
        button_row.addStretch()
        button_row.addWidget(self._btn_select_all)
        button_row.addWidget(self._btn_select_none)
        layout.addLayout(button_row)

        self._table = QTableWidget(len(self._registry.entries), 4, self)
        self._table.setHorizontalHeaderLabels(["Select", "Category", "Application", "Mode"])
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.verticalHeader().setVisible(False)
        header = self._table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, header.ResizeToContents)
        header.setSectionResizeMode(1, header.ResizeToContents)
        header.setSectionResizeMode(2, header.Stretch)
        header.setSectionResizeMode(3, header.ResizeToContents)
        layout.addWidget(self._table)

        for row, app in enumerate(self._registry.entries):
            checkbox = QTableWidgetItem()
            checkbox.setFlags(Qt.ItemIsSelectable | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            checkbox.setCheckState(Qt.Unchecked)
            checkbox.setData(Qt.UserRole, app.name)
            self._table.setItem(row, 0, checkbox)

            self._table.setItem(row, 1, QTableWidgetItem(app.category))
            self._table.setItem(row, 2, QTableWidgetItem(app.name))
            self._table.setItem(row, 3, QTableWidgetItem(app.download_mode))

        self._btn_download.clicked.connect(lambda: self._start_action("download_selected"))
        self._btn_install.clicked.connect(lambda: self._start_action("install_selected"))
        self._btn_select_all.clicked.connect(self._select_all)
        self._btn_select_none.clicked.connect(self._select_none)

    def _select_all(self) -> None:
        for row in range(self._table.rowCount()):
            item = self._table.item(row, 0)
            if item:
                item.setCheckState(Qt.Checked)

    def _select_none(self) -> None:
        for row in range(self._table.rowCount()):
            item = self._table.item(row, 0)
            if item:
                item.setCheckState(Qt.Unchecked)

    def _start_action(self, action: str) -> None:
        if self._busy:
            QMessageBox.information(self, "In Progress", "Please wait for the current operation to complete.")
            return
        selection = self._selected_apps()
        if not selection:
            QMessageBox.information(self, "No Selection", "Select at least one application to continue.")
            return
        if not hasattr(self._service, action):
            QMessageBox.warning(self, "Unsupported", f"Unknown installer action: {action}")
            return
        self._busy = True
        self._set_buttons_enabled(False)
        self._log(f"Starting {action.replace('_', ' ')} for {len(selection)} app(s) ...")
        worker = ServiceWorker(getattr(self._service, action), selection)
        worker.signals.finished.connect(lambda result, action=action: self._handle_results(action, result))
        worker.signals.error.connect(self._handle_error)
        self._thread_pool.start(worker)

    def _handle_results(self, action: str, results: Iterable[OperationResult]) -> None:
        for result in results:
            status = "OK" if result.success else "FAIL"
            self._log(f"[{status}] {action} :: {result.app.name} -> {result.message}")
        self._busy = False
        self._set_buttons_enabled(True)

    def _handle_error(self, message: str) -> None:
        self._log(f"[ERROR] {message}")
        self._busy = False
        self._set_buttons_enabled(True)

    def _selected_apps(self) -> list[str]:
        selection: list[str] = []
        for row in range(self._table.rowCount()):
            item = self._table.item(row, 0)
            if item and item.checkState() == Qt.Checked:
                app_name = item.data(Qt.UserRole)
                if isinstance(app_name, str):
                    selection.append(app_name)
        return selection

    def _set_buttons_enabled(self, enabled: bool) -> None:
        for button in (self._btn_download, self._btn_install, self._btn_select_all, self._btn_select_none):
            button.setEnabled(enabled)
