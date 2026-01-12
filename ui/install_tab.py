"""Installation tab with selectable applications and async actions."""
from __future__ import annotations

from pathlib import Path
from typing import Callable, Iterable

from PySide6.QtCore import Qt, QThreadPool
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QHeaderView,
    QHBoxLayout,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from services.app_status import AppStatusService, AppUpdateResult, InstalledInfo
from services.installer import InstallerService, OperationResult
from allinone_it_config.app_registry import AppRegistry
from ui.workers import ServiceWorker

LogCallback = Callable[[str], None]


class InstallTab(QWidget):
    COL_SELECT = 0
    COL_CATEGORY = 1
    COL_APP = 2
    COL_INSTALLED = 3
    COL_LATEST = 4
    COL_STATUS = 5
    COL_OFFLINE = 6
    COL_MODE = 7

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
        self._working_dir = working_dir or Path.cwd()
        self._service = InstallerService(registry.entries, working_dir=self._working_dir)
        self._status_service = AppStatusService(registry.entries, working_dir=self._working_dir)
        self._installed_map: dict[str, InstalledInfo] = {}
        self._row_by_name: dict[str, int] = {}
        self._busy = False
        self._build_ui()
        self._start_installed_scan()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)

        button_row = QHBoxLayout()
        self._btn_download = QPushButton("Download Offline")
        self._btn_install = QPushButton("Install Selected")
        self._btn_check_updates = QPushButton("Check for Updates")
        self._btn_select_all = QPushButton("Select All")
        self._btn_select_none = QPushButton("Select None")
        button_row.addWidget(self._btn_download)
        button_row.addWidget(self._btn_install)
        button_row.addWidget(self._btn_check_updates)
        button_row.addStretch()
        button_row.addWidget(self._btn_select_all)
        button_row.addWidget(self._btn_select_none)
        layout.addLayout(button_row)

        self._update_progress = QProgressBar(self)
        self._update_progress.setVisible(False)
        self._update_progress.setTextVisible(True)
        self._update_progress.setFormat("Checking updates... %p%")
        layout.addWidget(self._update_progress)

        self._table = QTableWidget(len(self._registry.entries), 8, self)
        self._table.setHorizontalHeaderLabels(
            ["Select", "Category", "Application", "Installed", "Latest", "Status", "Offline", "Mode"]
        )
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.verticalHeader().setVisible(False)
        header = self._table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(self.COL_SELECT, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(self.COL_CATEGORY, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(self.COL_APP, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.COL_INSTALLED, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.COL_LATEST, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.COL_STATUS, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(self.COL_OFFLINE, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(self.COL_MODE, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self._table)

        for row, app in enumerate(self._registry.entries):
            checkbox = QTableWidgetItem()
            checkbox.setFlags(Qt.ItemIsSelectable | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            checkbox.setCheckState(Qt.Unchecked)
            checkbox.setData(Qt.UserRole, app.name)
            self._table.setItem(row, self.COL_SELECT, checkbox)

            self._table.setItem(row, self.COL_CATEGORY, QTableWidgetItem(app.category))
            self._table.setItem(row, self.COL_APP, QTableWidgetItem(app.name))
            self._table.setItem(row, self.COL_INSTALLED, QTableWidgetItem("Scanning..."))
            self._table.setItem(row, self.COL_LATEST, QTableWidgetItem(""))
            self._table.setItem(row, self.COL_STATUS, QTableWidgetItem("Pending"))
            self._table.setItem(row, self.COL_OFFLINE, QTableWidgetItem("Checking..."))
            self._table.setItem(row, self.COL_MODE, QTableWidgetItem(app.download_mode))
            self._row_by_name[app.name] = row

        self._btn_download.clicked.connect(lambda: self._start_action("download_selected"))
        self._btn_install.clicked.connect(lambda: self._start_action("install_selected"))
        self._btn_check_updates.clicked.connect(self._start_update_check)
        self._btn_select_all.clicked.connect(self._select_all)
        self._btn_select_none.clicked.connect(self._select_none)
        self._refresh_offline_status()

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
        self._update_progress.setVisible(False)
        if action == "download_selected":
            self._refresh_offline_status()
        if action == "install_selected":
            self._start_installed_scan()

    def _handle_error(self, message: str) -> None:
        self._log(f"[ERROR] {message}")
        self._busy = False
        self._set_buttons_enabled(True)
        self._update_progress.setVisible(False)

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
        for button in (
            self._btn_download,
            self._btn_install,
            self._btn_check_updates,
            self._btn_select_all,
            self._btn_select_none,
        ):
            button.setEnabled(enabled)

    def _start_installed_scan(self) -> None:
        if self._busy:
            return
        self._busy = True
        self._set_buttons_enabled(False)
        self._log("Scanning installed applications...")
        worker = ServiceWorker(self._status_service.scan_installed)
        worker.signals.finished.connect(self._handle_installed_results)
        worker.signals.error.connect(self._handle_error)
        self._thread_pool.start(worker)

    def _handle_installed_results(self, results: Iterable[InstalledInfo]) -> None:
        self._installed_map = {info.app.name: info for info in results}
        for info in results:
            row = self._row_by_name.get(info.app.name)
            if row is None:
                continue
            self._set_item_text(row, self.COL_INSTALLED, info.installed_text)
            self._set_item_text(row, self.COL_LATEST, "")
            status_text, level = self._status_for_installed(info)
            self._set_item_text(row, self.COL_STATUS, status_text)
            self._apply_status_color(row, level)
        self._busy = False
        self._set_buttons_enabled(True)

    def _refresh_offline_status(self) -> None:
        for app in self._registry.entries:
            row = self._row_by_name.get(app.name)
            if row is None:
                continue
            local_info = self._service.get_local_installer_info(app, include_downloads=True)
            if local_info.exists:
                text = "Ready"
            elif self._service.is_downloadable(app):
                text = "Downloadable"
            else:
                text = "Not Downloadable"
            self._set_item_text(row, self.COL_OFFLINE, text)

    def _start_update_check(self) -> None:
        if self._busy:
            QMessageBox.information(self, "In Progress", "Please wait for the current operation to complete.")
            return
        self._busy = True
        self._set_buttons_enabled(False)
        self._log("Checking latest application versions...")
        total = self._table.rowCount()
        self._update_progress.setRange(0, total)
        self._update_progress.setValue(0)
        self._update_progress.setFormat("Checking updates... %v/%m (%p%)")
        self._update_progress.setVisible(True)
        for row in range(self._table.rowCount()):
            self._set_item_text(row, self.COL_LATEST, "Checking...")
            self._set_item_text(row, self.COL_STATUS, "Checking")
            self._apply_status_color(row, "checking")
        installed_map = dict(self._installed_map) or None
        worker = ServiceWorker(self._status_service.check_updates, installed_map)
        worker.kwargs["progress_callback"] = worker.signals.progress.emit
        worker.signals.finished.connect(self._handle_update_results)
        worker.signals.error.connect(self._handle_error)
        worker.signals.progress.connect(self._handle_update_progress)
        self._thread_pool.start(worker)

    def _handle_update_results(self, results: Iterable[AppUpdateResult]) -> None:
        for result in results:
            row = self._row_by_name.get(result.app.name)
            if row is None:
                continue
            self._set_item_text(row, self.COL_INSTALLED, result.installed_text)
            self._set_item_text(row, self.COL_LATEST, result.latest_text)
            self._set_item_text(row, self.COL_STATUS, result.status)
            self._apply_status_color(row, result.status_level)
        self._busy = False
        self._set_buttons_enabled(True)
        self._update_progress.setVisible(False)

    def _handle_update_progress(self, current: int, total: int, app_name: str) -> None:
        if total > 0:
            self._update_progress.setRange(0, total)
        self._update_progress.setValue(current)

    def _set_item_text(self, row: int, column: int, text: str) -> None:
        item = self._table.item(row, column)
        if item is None:
            item = QTableWidgetItem(text)
            self._table.setItem(row, column, item)
        else:
            item.setText(text)

    def _status_for_installed(self, info: InstalledInfo) -> tuple[str, str]:
        if not info.is_known:
            return "Unknown", "unknown"
        if info.is_installed:
            return "Installed", "installed"
        return "Not Installed", "not_installed"

    def _apply_status_color(self, row: int, level: str) -> None:
        color = self._status_color(level)
        if color is None:
            return
        for col in range(self.COL_CATEGORY, self.COL_MODE + 1):
            item = self._table.item(row, col)
            if item:
                item.setForeground(color)

    def _status_color(self, level: str) -> QColor | None:
        palette = {
            "up_to_date": QColor("#2ecc71"),
            "update_available": QColor("#f39c12"),
            "not_installed": QColor("#e74c3c"),
            "installed": QColor("#27ae60"),
            "unknown": QColor("#9aa7b2"),
            "checking": QColor("#f1c40f"),
        }
        return palette.get(level)
