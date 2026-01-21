"""Settings dialog for driver/legacy repository configuration."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from allinone_it_config.user_settings import SettingsStore, UserSettings


class DriverSettingsDialog(QDialog):
    def __init__(self, settings: UserSettings, store: SettingsStore, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._settings = settings
        self._store = store
        self.setWindowTitle("Driver Repo Settings")
        self.setMinimumWidth(520)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        form = QFormLayout()

        self._legacy_repo = QLineEdit(self._settings.hp_legacy_repo_root)
        form.addRow("HP Legacy Repo Root", self._make_dir_picker(self._legacy_repo, "Select HP Legacy Repo Root"))
        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _make_dir_picker(self, field: QLineEdit, title: str) -> QWidget:
        container = QWidget()
        row = QHBoxLayout(container)
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(field)
        browse = QPushButton("Browse")
        browse.clicked.connect(lambda: self._browse_for_dir(field, title))
        row.addWidget(browse)
        return container

    def _browse_for_dir(self, field: QLineEdit, title: str) -> None:
        current = field.text().strip()
        start_dir = current or str(Path.home())
        path = QFileDialog.getExistingDirectory(self, title, start_dir)
        if path:
            field.setText(path)

    def _save(self) -> None:
        self._settings.hp_legacy_repo_root = self._legacy_repo.text().strip()
        self._store.save(self._settings)
        self.accept()
