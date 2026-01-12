"""Main window for the All-In-One IT Configuration Tool."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QThreadPool
from PySide6.QtWidgets import (
    QMainWindow,
    QSplitter,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from allinone_it_config.app_registry import build_registry
from allinone_it_config.constants import IMMUTABLE_CONFIG
from allinone_it_config.user_settings import SettingsStore
from ui.install_tab import InstallTab
from ui.drivers_tab import DriversTab
from ui.system_tab import SystemTab
from ui.theme import apply_dark_theme


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("All-In-One IT Configuration Tool")
        self.resize(1200, 800)
        apply_dark_theme()
        self._thread_pool = QThreadPool.globalInstance()
        self._settings_store = SettingsStore()
        self._settings = self._settings_store.load()
        self._log_view = QTextEdit()
        self._log_view.setReadOnly(True)
        self._log_view.setMinimumHeight(120)
        self._tabs = QTabWidget()
        self._tabs.addTab(self._create_install_tab(), "Applications")
        self._tabs.addTab(self._create_drivers_tab(), "Drivers")
        self._tabs.addTab(self._create_system_tab(), "System Config")

        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(self._tabs)
        splitter.addWidget(self._log_view)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 1)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.addWidget(splitter)
        self.setCentralWidget(container)

    def log_message(self, message: str) -> None:
        self._log_view.append(message)

    def _create_install_tab(self) -> QWidget:
        return InstallTab(
            build_registry(self._settings),
            log_callback=self.log_message,
            thread_pool=self._thread_pool,
            working_dir=Path.cwd(),
            settings=self._settings,
            settings_store=self._settings_store,
        )

    def _create_drivers_tab(self) -> QWidget:
        return DriversTab(
            log_callback=self.log_message,
            thread_pool=self._thread_pool,
            working_dir=Path.cwd(),
        )

    def _create_system_tab(self) -> QWidget:
        return SystemTab(
            IMMUTABLE_CONFIG.system,
            log_callback=self.log_message,
            thread_pool=self._thread_pool,
        )
