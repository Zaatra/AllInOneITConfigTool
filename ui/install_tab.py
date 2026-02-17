"""Installation tab with selectable applications and async actions."""
from __future__ import annotations

import time
from pathlib import Path
from typing import Callable, Iterable
from urllib.parse import urlparse

from PySide6.QtCore import Qt, QThreadPool, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
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
from services.installer import InstallerService, LocalInstallerInfo, LocalInstallerVersionInfo, OperationResult
from allinone_it_config.app_registry import AppEntry, AppRegistry, build_registry
from allinone_it_config.paths import get_application_directory
from allinone_it_config.user_settings import SettingsStore, UserSettings
from ui.settings_dialog import SettingsDialog
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
        settings: UserSettings | None = None,
        settings_store: SettingsStore | None = None,
    ) -> None:
        super().__init__()
        self._registry = registry
        self._log = log_callback
        self._thread_pool = thread_pool
        self._working_dir = working_dir or get_application_directory()
        self._settings_store = settings_store or SettingsStore()
        self._settings = settings or self._settings_store.load()
        self._service = InstallerService(registry.entries, working_dir=self._working_dir, settings=self._settings)
        self._status_service = AppStatusService(
            registry.entries,
            working_dir=self._working_dir,
            settings=self._settings,
        )
        self._installed_map: dict[str, InstalledInfo] = {}
        self._latest_versions: dict[str, str] = {}
        self._row_by_name: dict[str, int] = {}
        self._workers: set[ServiceWorker] = set()
        self._busy = False
        self._action_label = ""
        self._action_current = 0
        self._action_total = 0
        self._action_app = ""
        self._action_started_at: float | None = None
        self._last_update_check_app = ""
        self._last_speed_message = ""
        self._last_speed_log_at: float = 0.0
        self._build_ui()
        self._start_installed_scan()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)

        button_row = QHBoxLayout()
        self._btn_download = QPushButton("Download Offline")
        self._btn_install = QPushButton("Install Selected")
        self._btn_check_updates = QPushButton("Check for Updates")
        self._btn_refresh_offline = QPushButton("Refresh Offline")
        self._btn_settings = QPushButton("Settings")
        self._btn_select_all = QPushButton("Select All")
        self._btn_select_none = QPushButton("Select None")
        button_row.addWidget(self._btn_download)
        button_row.addWidget(self._btn_install)
        button_row.addWidget(self._btn_check_updates)
        button_row.addWidget(self._btn_refresh_offline)
        button_row.addWidget(self._btn_settings)
        button_row.addStretch()
        button_row.addWidget(self._btn_select_all)
        button_row.addWidget(self._btn_select_none)
        layout.addLayout(button_row)

        self._action_progress = QProgressBar(self)
        self._action_progress.setVisible(False)
        self._action_progress.setTextVisible(True)
        self._action_progress.setFormat("Working...")
        layout.addWidget(self._action_progress)

        self._update_progress = QProgressBar(self)
        self._update_progress.setVisible(False)
        self._update_progress.setTextVisible(True)
        self._update_progress.setFormat("Checking updates... %p%")
        layout.addWidget(self._update_progress)
        self._action_timer = QTimer(self)
        self._action_timer.setInterval(1000)
        self._action_timer.timeout.connect(self._tick_action_timer)

        self._table = QTableWidget(0, 8, self)
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
        header.setSectionResizeMode(self.COL_STATUS, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(self.COL_OFFLINE, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(self.COL_MODE, QHeaderView.ResizeMode.Interactive)
        self._table.setColumnWidth(self.COL_STATUS, 150)
        self._table.setColumnWidth(self.COL_OFFLINE, 150)
        self._table.setColumnWidth(self.COL_MODE, 130)
        layout.addWidget(self._table)

        self._populate_table()

        self._btn_download.clicked.connect(lambda: self._start_action("download_selected"))
        self._btn_install.clicked.connect(lambda: self._start_action("install_selected"))
        self._btn_check_updates.clicked.connect(self._start_update_check)
        self._btn_refresh_offline.clicked.connect(self._refresh_offline_status_clicked)
        self._btn_select_all.clicked.connect(self._select_all)
        self._btn_select_none.clicked.connect(self._select_none)
        self._btn_settings.clicked.connect(self._open_settings_dialog)
        self._refresh_offline_status()
        if not self._settings_store.exists():
            QTimer.singleShot(0, lambda: self._open_settings_dialog(force=True))

    def _track_worker(self, worker: ServiceWorker) -> None:
        self._workers.add(worker)
        worker.signals.finished.connect(lambda *_: self._workers.discard(worker))
        worker.signals.error.connect(lambda *_: self._workers.discard(worker))

    def _populate_table(self) -> None:
        self._table.setRowCount(len(self._registry.entries))
        self._table.clearContents()
        self._row_by_name.clear()
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

    def _open_settings_dialog(self, *, force: bool = False) -> bool:
        if self._busy and not force:
            QMessageBox.information(self, "In Progress", "Please wait for the current operation to complete.")
            return False
        dialog = SettingsDialog(self._settings, self._settings_store, self)
        if dialog.exec() == QDialog.Accepted:
            self._apply_registry(build_registry(self._settings))
            return True
        return False

    def _apply_registry(self, registry: AppRegistry) -> None:
        self._registry = registry
        self._service = InstallerService(registry.entries, working_dir=self._working_dir, settings=self._settings)
        self._status_service = AppStatusService(
            registry.entries,
            working_dir=self._working_dir,
            settings=self._settings,
        )
        self._installed_map.clear()
        self._latest_versions.clear()
        self._populate_table()
        self._refresh_offline_status()
        if not self._busy:
            self._start_installed_scan()

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
        if not self._ensure_settings_for(action, selection):
            return
        if action == "install_selected":
            if not self._confirm_local_version_overrides(selection):
                return
        force_updates: set[str] = set()
        if action == "download_selected" and self._latest_versions:
            selection_set = {name.lower() for name in selection}
            for app in self._registry.entries:
                if app.name.lower() not in selection_set:
                    continue
                local_info = self._service.get_local_installer_info(app, include_downloads=True)
                if not local_info.exists:
                    continue
                if self._offline_version_status(app, local_info) == "Outdated":
                    force_updates.add(app.name)
        if not hasattr(self._service, action):
            QMessageBox.warning(self, "Unsupported", f"Unknown installer action: {action}")
            return
        self._busy = True
        self._set_buttons_enabled(False)
        action_label = "Downloading" if action == "download_selected" else "Installing"
        self._begin_action_progress(action_label, len(selection))
        self._log(f"Starting {action.replace('_', ' ')} for {len(selection)} app(s) ...")
        worker = ServiceWorker(getattr(self._service, action), selection)
        worker.kwargs["progress_callback"] = worker.signals.progress.emit
        worker.kwargs["status_callback"] = worker.signals.message.emit
        if action == "download_selected":
            worker.kwargs["force_updates"] = force_updates
        worker.signals.finished.connect(lambda result, action=action: self._handle_results(action, result))
        worker.signals.error.connect(self._handle_error)
        worker.signals.progress.connect(self._handle_action_progress)
        worker.signals.message.connect(self._handle_action_message)
        self._track_worker(worker)
        self._thread_pool.start(worker)

    def _handle_results(self, action: str, results: Iterable[OperationResult]) -> None:
        try:
            for result in results:
                status = "OK" if result.success else "FAIL"
                self._log(f"[{status}] {action} :: {result.app.name} -> {result.message}")
        except Exception as exc:
            self._log(f"[ERROR] Failed to process results: {exc}")
        self._busy = False
        self._set_buttons_enabled(True)
        self._complete_action_progress()
        self._update_progress.setVisible(False)
        if action == "download_selected":
            self._refresh_offline_status()
        if action == "install_selected":
            self._start_installed_scan()

    def _handle_error(self, message: str) -> None:
        self._log(f"[ERROR] {message}")
        self._busy = False
        self._set_buttons_enabled(True)
        self._complete_action_progress()
        self._update_progress.setVisible(False)

    def _ensure_settings_for(self, action: str, selection: list[str]) -> bool:
        missing = self._missing_settings(action, selection)
        if not missing:
            return True
        message = "Settings required for selected apps:\n" + "\n".join(f"- {item}" for item in missing)
        reply = QMessageBox.question(
            self,
            "Settings Required",
            f"{message}\n\nOpen settings now?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return False
        if not self._open_settings_dialog():
            return False
        missing = self._missing_settings(action, selection)
        if missing:
            message = "Still missing:\n" + "\n".join(f"- {item}" for item in missing)
            QMessageBox.warning(self, "Settings Incomplete", message)
            return False
        return True

    def _missing_settings(self, action: str, selection: list[str]) -> list[str]:
        missing: list[str] = []
        selected = {name.lower() for name in selection}
        if "crowdstrike falcon sensor".lower() in selected:
            if not self._settings.crowdstrike_cid.strip():
                missing.append("CrowdStrike CID")
            if action == "download_selected":
                if not _is_sharepoint_url(self._settings.crowdstrike_download_url):
                    missing.append("CrowdStrike SharePoint URL")
        if "forticlient vpn".lower() in selected and action == "download_selected":
            if not _is_sharepoint_url(self._settings.forticlient_download_url):
                missing.append("FortiClient VPN SharePoint URL")
        if action == "install_selected":
            if "office 2024 ltsc".lower() in selected:
                if not _file_exists(self._settings.office_2024_xml_path):
                    missing.append("Office 2024 XML path")
            if "office 365 ent".lower() in selected:
                if not _file_exists(self._settings.office_365_xml_path):
                    missing.append("Office 365 XML path")
            if "winrar".lower() in selected:
                if not _file_exists(self._settings.winrar_license_path):
                    missing.append("WinRAR license file path")
            if "teamviewer".lower() in selected:
                if self._settings.teamviewer_install_mode.strip().lower() == "msi":
                    if not _file_exists(self._settings.teamviewer_msi_path):
                        missing.append("TeamViewer MSI path")
                    if not self._settings.teamviewer_customconfig_id.strip():
                        missing.append("TeamViewer CUSTOMCONFIGID")
                    if not self._settings.teamviewer_assignment_id.strip():
                        missing.append("TeamViewer ASSIGNMENTID")
                    settings_file = self._settings.teamviewer_settings_file.strip()
                    if not settings_file:
                        missing.append("TeamViewer SETTINGSFILE path")
                    elif not settings_file.lower().endswith(".tvopt"):
                        missing.append("TeamViewer SETTINGSFILE must end with .tvopt")
                    elif not _file_exists(settings_file):
                        missing.append("TeamViewer SETTINGSFILE path")
        return missing

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
            self._btn_refresh_offline,
            self._btn_settings,
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
        self._track_worker(worker)
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
        self._refresh_offline_status()
        self._busy = False
        self._set_buttons_enabled(True)

    def _refresh_offline_status_clicked(self) -> None:
        if self._busy:
            QMessageBox.information(self, "In Progress", "Please wait for the current operation to complete.")
            return
        self._start_installed_scan()

    def _refresh_offline_status(self) -> None:
        for app in self._registry.entries:
            row = self._row_by_name.get(app.name)
            if row is None:
                continue
            if app.name == "Office Deployment Tool":
                odt_path = self._settings.odt_setup_path.strip()
                odt_version = self._status_service.get_local_odt_version()
                if _file_exists(odt_path):
                    text = "Ready"
                elif odt_version:
                    text = "Ready"
                else:
                    self._set_item_text(row, self.COL_OFFLINE, "Managed")
                    continue
                latest_text = self._latest_versions.get(app.name)
                if odt_version and latest_text:
                    suffix = self._status_service.offline_installer_status(
                        app,
                        LocalInstallerVersionInfo(version=odt_version),
                        latest_text,
                    )
                    if suffix:
                        text = f"{text} - {suffix}"
                self._set_item_text(row, self.COL_OFFLINE, text)
                continue
            local_info = self._service.get_local_installer_info(app, include_downloads=True)
            if local_info.exists:
                text = "Ready"
                suffix = self._offline_version_status(app, local_info)
                if suffix:
                    text = f"{text} - {suffix}"
            elif self._service.is_downloadable(app):
                text = "Downloadable"
            else:
                text = "Not Downloadable"
            self._set_item_text(row, self.COL_OFFLINE, text)

    def _offline_version_status(self, app: AppEntry, local_info: LocalInstallerInfo) -> str | None:
        latest_text = self._latest_versions.get(app.name)
        if not latest_text:
            return None
        local_versions = self._service.get_local_installer_versions(app, local_info)
        return self._status_service.offline_installer_status(app, local_versions, latest_text)

    def _start_update_check(self) -> None:
        if self._busy:
            QMessageBox.information(self, "In Progress", "Please wait for the current operation to complete.")
            return
        self._busy = True
        self._set_buttons_enabled(False)
        self._log("Checking latest application versions...")
        self._last_update_check_app = ""
        total = self._table.rowCount()
        self._update_progress.setRange(0, total)
        self._update_progress.setValue(0)
        self._update_progress.setFormat("Checking updates... %v/%m (%p%)")
        self._update_progress.setVisible(True)
        for row in range(self._table.rowCount()):
            self._set_item_text(row, self.COL_LATEST, "Checking...")
            self._set_item_text(row, self.COL_STATUS, "Checking")
            self._apply_status_color(row, "checking")
        worker = ServiceWorker(self._status_service.check_updates)
        worker.kwargs["progress_callback"] = worker.signals.progress.emit
        worker.signals.finished.connect(self._handle_update_results)
        worker.signals.error.connect(self._handle_error)
        worker.signals.progress.connect(self._handle_update_progress)
        self._track_worker(worker)
        self._thread_pool.start(worker)

    def _handle_update_results(self, results: Iterable[AppUpdateResult]) -> None:
        results_list = list(results)
        self._last_update_check_app = ""
        self._latest_versions = {result.app.name: result.latest_text for result in results_list}
        for result in results_list:
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
        self._refresh_offline_status()

    def _handle_update_progress(self, current: int, total: int, app_name: str) -> None:
        if app_name.startswith("START:"):
            current_app = app_name.split(":", 1)[1].strip()
            if current_app and current_app != self._last_update_check_app:
                self._last_update_check_app = current_app
                self._log(f"Checking for update {current_app}...")
            if current_app:
                self._update_progress.setFormat(f"Checking updates... %v/%m (%p%) | {current_app}")
            if total > 0:
                self._update_progress.setRange(0, total)
            self._update_progress.setValue(current)
            return
        if total > 0:
            self._update_progress.setRange(0, total)
        self._update_progress.setValue(current)
        if app_name:
            self._update_progress.setFormat(f"Checking updates... %v/%m (%p%) | {app_name}")

    def _handle_action_progress(self, current: int, total: int, app_name: str) -> None:
        self._action_current = current
        self._action_total = total if total > 0 else self._action_total
        self._action_progress.setRange(0, max(self._action_total, 1))
        self._action_progress.setValue(current)
        if not self._action_app:
            self._action_app = app_name
        self._update_action_progress_text()

    def _handle_action_message(self, message: str) -> None:
        if not self._busy:
            return
        self._action_app = message
        self._update_action_progress_text()
        if self._action_label != "Downloading":
            return
        if not _looks_like_speed_message(message):
            return
        now = time.monotonic()
        if message != self._last_speed_message or (now - self._last_speed_log_at) >= 1.5:
            self._last_speed_message = message
            self._last_speed_log_at = now
            self._log(f"[SPEED] {message}")

    def _tick_action_timer(self) -> None:
        if not self._action_progress.isVisible():
            self._action_timer.stop()
            return
        self._update_action_progress_text()

    def _begin_action_progress(self, label: str, total: int) -> None:
        self._action_label = label
        self._action_total = max(total, 1)
        self._action_current = 0
        self._action_app = ""
        self._action_started_at = time.monotonic()
        self._last_speed_message = ""
        self._last_speed_log_at = 0.0
        self._action_progress.setRange(0, self._action_total)
        self._action_progress.setValue(0)
        self._action_progress.setVisible(True)
        self._update_action_progress_text()
        self._action_timer.start()

    def _end_action_progress(self) -> None:
        self._action_timer.stop()
        self._action_progress.setVisible(False)
        self._action_started_at = None

    def _complete_action_progress(self) -> None:
        if self._action_total:
            self._action_current = self._action_total
            self._action_progress.setValue(self._action_total)
            self._update_action_progress_text()
        self._end_action_progress()

    def _update_action_progress_text(self) -> None:
        elapsed = 0
        if self._action_started_at is not None:
            elapsed = int(time.monotonic() - self._action_started_at)
        app_part = f" | {self._action_app}" if self._action_app else ""
        text = f"{self._action_label} {self._action_current}/{self._action_total}{app_part} | {_format_elapsed(elapsed)}"
        self._action_progress.setFormat(text)

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

    def _confirm_local_version_overrides(self, selection: list[str]) -> bool:
        selected = {name.lower() for name in selection}
        warnings: list[str] = []
        for app in self._registry.entries:
            if app.name.lower() in selected:
                warnings.extend(self._service.local_version_override_warnings(app))
        if not warnings:
            return True
        message = (
            "Local installers found that will override selected versions:\n"
            + "\n".join(f"- {item}" for item in warnings)
            + "\n\nContinue anyway?"
        )
        reply = QMessageBox.warning(
            self,
            "Local Installer Override",
            message,
            QMessageBox.Yes | QMessageBox.No,
        )
        return reply == QMessageBox.Yes


def _file_exists(path: str) -> bool:
    if not path:
        return False
    candidate = Path(path)
    try:
        return candidate.exists() and candidate.is_file()
    except OSError:
        return False


def _format_elapsed(total_seconds: int) -> str:
    minutes, seconds = divmod(max(total_seconds, 0), 60)
    hours, minutes = divmod(minutes, 60)
    if hours:
        return f"{hours}:{minutes:02d}:{seconds:02d}"
    return f"{minutes:02d}:{seconds:02d}"


def _looks_like_speed_message(message: str) -> bool:
    text = message.lower()
    return "downloading " in text and "/s" in text


def _is_sharepoint_url(value: str) -> bool:
    cleaned = value.strip()
    if not cleaned:
        return False
    parsed = urlparse(cleaned)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        return False
    host = parsed.netloc.lower()
    return host.endswith("sharepoint.com") or ".sharepoint." in host or host.startswith("sharepoint.")
