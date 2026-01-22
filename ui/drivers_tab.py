"""Drivers tab UI for scanning/downloading/installing HP drivers."""
from __future__ import annotations

import os
import re
import subprocess
from pathlib import Path
from pathlib import PureWindowsPath
from typing import Callable, Iterable, List

from PySide6.QtCore import Qt, QThreadPool
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from allinone_it_config.paths import get_application_directory
from allinone_it_config.user_settings import SettingsStore, UserSettings
from services.drivers import DriverOperationResult, DriverRecord, DriverService
from ui.driver_settings_dialog import DriverSettingsDialog
from ui.legacy_credentials_dialog import LegacyCredentialsDialog
from ui.workers import ServiceWorker

LogCallback = Callable[[str], None]


class DriversTab(QWidget):
    def __init__(
        self,
        log_callback: LogCallback,
        thread_pool: QThreadPool,
        *,
        working_dir: Path | None = None,
        settings: UserSettings | None = None,
        settings_store: SettingsStore | None = None,
    ) -> None:
        super().__init__()
        self._log = log_callback
        self._thread_pool = thread_pool
        self._working_dir = working_dir or get_application_directory()
        self._settings_store = settings_store or SettingsStore()
        self._settings = settings or self._settings_store.load()
        self._refresh_service()
        self._records_by_source: dict[str, list[DriverRecord]] = {"HPIA": [], "CMSL": [], "LEGACY": []}
        self._view_by_source: dict[str, list[DriverRecord]] = {"HPIA": [], "CMSL": [], "LEGACY": []}
        self._workers: set[ServiceWorker] = set()
        self._busy = False
        self._last_action: tuple[str, str] | None = None
        self._build_ui()

    def _track_worker(self, worker: ServiceWorker) -> None:
        self._workers.add(worker)
        worker.signals.finished.connect(lambda *_: self._workers.discard(worker))
        worker.signals.error.connect(lambda *_: self._workers.discard(worker))

    def _refresh_service(self) -> None:
        legacy_root = self._settings.hp_legacy_repo_root.strip()
        if legacy_root:
            legacy_root = self._normalize_unc_path(legacy_root)
        self._service = DriverService(
            working_dir=self._working_dir,
            legacy_repo_root=legacy_root or None,
        )

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        self._tabs = QTabWidget(self)
        self._panels: dict[str, dict[str, QWidget]] = {}
        self._panels["HPIA"] = self._create_panel("HPIA")
        self._panels["CMSL"] = self._create_panel("CMSL", show_filter=True)
        self._panels["LEGACY"] = self._create_panel("Legacy", show_settings=True)
        self._tabs.addTab(self._panels["HPIA"]["widget"], "HPIA")
        self._tabs.addTab(self._panels["CMSL"]["widget"], "CMSL")
        self._tabs.addTab(self._panels["LEGACY"]["widget"], "Legacy")
        layout.addWidget(self._tabs)

    def _create_panel(
        self,
        source: str,
        *,
        show_settings: bool = False,
        show_filter: bool = False,
    ) -> dict[str, QWidget]:
        panel = QWidget()
        layout = QVBoxLayout(panel)

        button_row = QHBoxLayout()
        btn_scan = QPushButton(f"Scan {source}")
        btn_download = QPushButton("Download Selected")
        btn_install = QPushButton("Install Selected")
        btn_select_all = QPushButton("Select All")
        btn_select_none = QPushButton("Select None")
        for btn in (btn_scan, btn_download, btn_install):
            btn.setMinimumWidth(150)
        button_row.addWidget(btn_scan)
        button_row.addWidget(btn_download)
        button_row.addWidget(btn_install)
        if show_settings:
            btn_settings = QPushButton("Settings")
            button_row.addWidget(btn_settings)
            btn_settings.clicked.connect(self._open_driver_settings)
        else:
            btn_settings = None
        if show_filter:
            filter_label = QLabel("Category")
            filter_combo = QComboBox()
            filter_combo.addItem("All")
            button_row.addWidget(filter_label)
            button_row.addWidget(filter_combo)
            filter_combo.currentTextChanged.connect(lambda _: self._populate_table(source))
        else:
            filter_combo = None
        button_row.addStretch()
        button_row.addWidget(btn_select_all)
        button_row.addWidget(btn_select_none)
        layout.addLayout(button_row)

        table = QTableWidget(0, 7, panel)
        table.setHorizontalHeaderLabels(["Select", "Source", "Name", "Category", "Installed", "Latest", "Status"])
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.verticalHeader().setVisible(False)
        header = table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(table)

        btn_scan.clicked.connect(lambda: self._start_scan(source))
        btn_select_all.clicked.connect(lambda: self._set_all(table, Qt.Checked))
        btn_select_none.clicked.connect(lambda: self._set_all(table, Qt.Unchecked))
        btn_download.clicked.connect(lambda: self._start_operation(source, "download"))
        btn_install.clicked.connect(lambda: self._start_operation(source, "install"))

        return {
            "widget": panel,
            "table": table,
            "btn_scan": btn_scan,
            "btn_download": btn_download,
            "btn_install": btn_install,
            "btn_select_all": btn_select_all,
            "btn_select_none": btn_select_none,
            "btn_settings": btn_settings,
            "category_filter": filter_combo,
        }

    def _start_scan(self, source: str) -> None:
        if self._busy:
            return
        self._refresh_service()
        self._busy = True
        self._last_action = ("scan", source)
        self._set_buttons_enabled(False)
        self._log(f"Scanning {source} drivers...")
        if source == "HPIA":
            action = self._service.scan_hpia
        elif source == "CMSL":
            action = self._service.scan_cmsl_catalog
        else:
            action = self._service.scan_legacy
        worker = ServiceWorker(action)
        worker.signals.finished.connect(lambda records, src=source: self._handle_scan_results(src, records))
        worker.signals.error.connect(self._handle_error)
        self._track_worker(worker)
        self._thread_pool.start(worker)

    def _handle_scan_results(self, source: str, records: Iterable[DriverRecord]) -> None:
        self._records_by_source[source.upper()] = list(records)
        if source.upper() == "CMSL":
            self._update_category_filter(source, self._records_by_source[source.upper()])
        self._populate_table(source)
        self._log(f"{source} scan complete. Found {len(self._records_by_source[source.upper()])} entries.")
        for warning in self._service.last_scan_warnings:
            self._log(f"[WARN] {warning}")
        self._busy = False
        self._set_buttons_enabled(True)

    def _start_operation(self, source: str, op: str) -> None:
        if self._busy:
            QMessageBox.information(self, "In Progress", "Wait for the current operation to finish.")
            return
        selected = self._selected_records(source)
        if not selected:
            QMessageBox.information(self, "No Selection", "Select at least one driver entry.")
            return
        self._busy = True
        self._last_action = (op, source)
        self._set_buttons_enabled(False)
        action = self._service.download if op == "download" else self._service.install
        self._log(f"Running {op} for {len(selected)} driver(s) from {source}...")
        worker = ServiceWorker(action, selected)
        worker.signals.finished.connect(lambda result, op=op, src=source: self._handle_driver_results(src, op, result))
        worker.signals.error.connect(self._handle_error)
        self._track_worker(worker)
        self._thread_pool.start(worker)

    def _handle_driver_results(self, source: str, op: str, results: Iterable[DriverOperationResult]) -> None:
        for result in results:
            status = "OK" if result.success else "FAIL"
            self._log(f"[{status}] {op} :: {result.driver.name} -> {result.message}")
        if op == "download":
            self._populate_table(source)
        self._busy = False
        self._set_buttons_enabled(True)

    def _populate_table(self, source: str) -> None:
        key = source.upper()
        table = self._panel_table(key)
        records = self._filtered_records(source)
        self._view_by_source[key] = records
        table.setRowCount(len(records))
        for row, record in enumerate(records):
            table.setRowHeight(row, 28)
            checkbox = QTableWidgetItem()
            checkbox.setFlags(Qt.ItemIsSelectable | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            checkbox.setCheckState(Qt.Unchecked)
            checkbox.setData(Qt.UserRole, row)
            table.setItem(row, 0, checkbox)

            self._set_badge_cell(table, row, 1, record.source, self._source_badge_style(record.source))
            table.setItem(row, 2, QTableWidgetItem(record.name))
            category = record.category or "Other"
            table.setItem(row, 3, QTableWidgetItem(category))
            installed = record.installed_version or ("N/A" if record.status.lower() == "catalog" else "Unknown")
            latest = record.latest_version or "Unknown"
            installed_item = QTableWidgetItem(installed)
            installed_item.setTextAlignment(Qt.AlignCenter)
            table.setItem(row, 4, installed_item)
            latest_item = QTableWidgetItem(latest)
            latest_item.setTextAlignment(Qt.AlignCenter)
            table.setItem(row, 5, latest_item)
            status_text = record.status
            if record.output_path:
                status_text += " (cached)"
            self._set_badge_cell(table, row, 6, status_text, self._status_badge_style(record.status))
            self._apply_version_colors(table, row, record.status)

    def _selected_records(self, source: str) -> List[DriverRecord]:
        key = source.upper()
        table = self._panel_table(key)
        records = self._view_by_source.get(key, [])
        selections: list[DriverRecord] = []
        for row in range(table.rowCount()):
            item = table.item(row, 0)
            if item and item.checkState() == Qt.Checked:
                idx = item.data(Qt.UserRole)
                if isinstance(idx, int) and 0 <= idx < len(records):
                    selections.append(records[idx])
        return selections

    def _set_all(self, table: QTableWidget, state: Qt.CheckState) -> None:
        for row in range(table.rowCount()):
            item = table.item(row, 0)
            if item:
                item.setCheckState(state)

    def _set_buttons_enabled(self, enabled: bool) -> None:
        for panel in self._panels.values():
            for key in ("btn_scan", "btn_download", "btn_install", "btn_select_all", "btn_select_none", "btn_settings"):
                button = panel.get(key)
                if isinstance(button, QPushButton):
                    button.setEnabled(enabled)

    def _handle_error(self, message: str) -> None:
        self._busy = False
        self._set_buttons_enabled(True)
        if self._maybe_prompt_legacy_credentials(message):
            return
        self._log(f"[ERROR] {message}")

    def _maybe_prompt_legacy_credentials(self, message: str) -> bool:
        if not self._looks_like_legacy_auth_error(message):
            return False
        legacy_root = self._settings.hp_legacy_repo_root.strip()
        if not legacy_root:
            return False
        share_path = self._normalize_unc_path(self._extract_unc_path(message) or legacy_root)
        share_root = self._unc_share_root(share_path) or share_path
        if not share_root.startswith("\\\\"):
            return False
        self._log("[WARN] Legacy repository requires credentials.")
        dialog = LegacyCredentialsDialog(share_root, self)
        if not dialog.exec():
            self._log("[WARN] Legacy repository credentials not provided.")
            return True
        username, password = dialog.credentials()
        if not username or not password:
            QMessageBox.warning(self, "Missing Credentials", "Enter both a username and password.")
            return True
        success, detail = self._connect_to_share(share_root, username, password)
        if not success:
            QMessageBox.critical(self, "Legacy Repo Authentication Failed", detail)
            self._log(f"[ERROR] Legacy repo authentication failed: {detail}")
            return True
        if self._last_action == ("scan", "Legacy"):
            self._log("Legacy repo credentials accepted. Retrying scan...")
            self._start_scan("Legacy")
        else:
            self._log("Legacy repo credentials accepted. Retry the legacy operation.")
        return True

    def _looks_like_legacy_auth_error(self, message: str) -> bool:
        lowered = message.lower()
        patterns = (
            "winerror 1326",
            "user name or password is incorrect",
            "logon failure",
        )
        return any(pattern in lowered for pattern in patterns)

    def _extract_unc_path(self, message: str) -> str | None:
        quoted = re.search(r"'(\\\\[^']+)'", message)
        if quoted:
            return quoted.group(1)
        raw = re.search(r"(\\\\[^\\r\\n]+)", message)
        if raw:
            return raw.group(1)
        return None

    def _normalize_unc_path(self, path: str) -> str:
        cleaned = path.strip()
        if len(cleaned) >= 2 and cleaned[0] == cleaned[-1] and cleaned[0] in ('"', "'"):
            cleaned = cleaned[1:-1].strip()
        if cleaned.startswith("//"):
            cleaned = "\\\\" + cleaned.lstrip("/")
        cleaned = cleaned.replace("/", "\\")
        return cleaned

    def _unc_share_root(self, path: str) -> str | None:
        normalized = self._normalize_unc_path(path)
        if not normalized.startswith("\\\\"):
            return None
        parts = PureWindowsPath(normalized).parts
        if not parts:
            return None
        return parts[0].rstrip("\\")

    def _connect_to_share(self, share: str, username: str, password: str) -> tuple[bool, str]:
        if os.name != "nt":
            return (False, "Network credentials can only be applied on Windows.")
        normalized = self._normalize_unc_path(share).rstrip("\\")
        share_root = self._unc_share_root(normalized) or normalized
        cmd = ["net", "use", share_root, password, f"/user:{username}", "/persistent:no"]
        result = subprocess.run(cmd, capture_output=True, text=True, check=False)
        if result.returncode == 0:
            return (True, "Connected")
        detail = (result.stderr or result.stdout or "Unknown error").strip()
        return (False, detail)

    def _set_badge_cell(self, table: QTableWidget, row: int, column: int, text: str, palette: tuple[str, str, str]) -> None:
        label = QLabel(text)
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet(
            "QLabel {"
            f"color: {palette[0]};"
            f"background-color: {palette[1]};"
            f"border: 1px solid {palette[2]};"
            "border-radius: 8px;"
            "padding: 2px 6px;"
            "font-weight: 600;"
            "}"
        )
        table.setCellWidget(row, column, label)

    def _source_badge_style(self, source: str) -> tuple[str, str, str]:
        palette = {
            "HPIA": ("#dbeafe", "#1e3a8a", "#3b82f6"),
            "CMSL": ("#ccfbf1", "#0f766e", "#14b8a6"),
            "LEGACY": ("#e5e7eb", "#374151", "#6b7280"),
        }
        return palette.get(source.upper(), ("#e5e7eb", "#4b5563", "#9ca3af"))

    def _status_badge_style(self, status: str) -> tuple[str, str, str]:
        palette = {
            "critical": ("#fee2e2", "#7f1d1d", "#ef4444"),
            "update available": ("#fee2e2", "#7f1d1d", "#ef4444"),
            "recommended": ("#fef3c7", "#78350f", "#f59e0b"),
            "optional": ("#e0f2fe", "#075985", "#38bdf8"),
            "up to date": ("#dcfce7", "#14532d", "#22c55e"),
            "installed": ("#dcfce7", "#14532d", "#22c55e"),
            "not installed": ("#e5e7eb", "#4b5563", "#9ca3af"),
            "legacy": ("#e5e7eb", "#374151", "#9ca3af"),
            "catalog": ("#e0f2fe", "#075985", "#38bdf8"),
            "unknown": ("#e5e7eb", "#4b5563", "#9ca3af"),
        }
        return palette.get(status.lower(), ("#e5e7eb", "#4b5563", "#9ca3af"))

    def _apply_version_colors(self, table: QTableWidget, row: int, status: str) -> None:
        installed_item = table.item(row, 4)
        latest_item = table.item(row, 5)
        if not installed_item or not latest_item:
            return
        status_key = status.lower()
        if status_key in {"up to date", "installed"}:
            installed_item.setForeground(QColor("#22c55e"))
        elif status_key in {"update available", "critical"}:
            latest_item.setForeground(QColor("#ef4444"))
        elif status_key == "recommended":
            latest_item.setForeground(QColor("#f59e0b"))
        elif status_key == "optional":
            latest_item.setForeground(QColor("#38bdf8"))
        elif status_key in {"not installed", "unknown"}:
            installed_item.setForeground(QColor("#9ca3af"))

    def _filtered_records(self, source: str) -> list[DriverRecord]:
        key = source.upper()
        records = list(self._records_by_source.get(key, []))
        if key != "CMSL":
            return records
        panel = self._panels.get(key, {})
        combo = panel.get("category_filter")
        if not isinstance(combo, QComboBox):
            return records
        selection = combo.currentText().strip()
        if not selection or selection == "All":
            return records
        return [record for record in records if (record.category or "Other") == selection]

    def _update_category_filter(self, source: str, records: list[DriverRecord]) -> None:
        key = source.upper()
        panel = self._panels.get(key, {})
        combo = panel.get("category_filter")
        if not isinstance(combo, QComboBox):
            return
        categories = sorted({(record.category or "Other") for record in records})
        current = combo.currentText()
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("All")
        for category in categories:
            combo.addItem(category)
        combo.blockSignals(False)
        if current in categories:
            combo.setCurrentText(current)
        else:
            combo.setCurrentIndex(0)

    def _panel_table(self, source: str) -> QTableWidget:
        key = source.upper()
        table = self._panels[key]["table"]
        return table  # type: ignore[return-value]

    def _open_driver_settings(self) -> None:
        dialog = DriverSettingsDialog(self._settings, self._settings_store, self)
        if dialog.exec():
            self._refresh_service()
            self._log("Driver settings saved.")
