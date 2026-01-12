"""Settings dialog for user-provided installer values."""
from __future__ import annotations

import re
import shutil
import subprocess
from pathlib import Path

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from allinone_it_config.user_settings import SettingsStore, UserSettings


class SettingsDialog(QDialog):
    def __init__(self, settings: UserSettings, store: SettingsStore, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._settings = settings
        self._store = store
        self.setWindowTitle("Installer Settings")
        self.setMinimumWidth(560)
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        form = QFormLayout()

        self._crowdstrike_cid = QLineEdit(self._settings.crowdstrike_cid)
        self._crowdstrike_cid.setPlaceholderText("Example: 753B2C62DBF...-27")
        form.addRow("CrowdStrike CID", self._crowdstrike_cid)

        self._crowdstrike_url = QLineEdit(self._settings.crowdstrike_download_url)
        self._crowdstrike_url.setPlaceholderText("https://sharepoint/...")
        form.addRow("CrowdStrike Download URL", self._crowdstrike_url)

        self._office_2024_path = QLineEdit(self._settings.office_2024_xml_path)
        form.addRow("Office 2024 XML", self._make_path_picker(self._office_2024_path, "Select Office 2024 XML", "XML Files (*.xml);;All Files (*)"))

        self._office_365_path = QLineEdit(self._settings.office_365_xml_path)
        form.addRow("Office 365 XML", self._make_path_picker(self._office_365_path, "Select Office 365 XML", "XML Files (*.xml);;All Files (*)"))

        self._winrar_license = QLineEdit(self._settings.winrar_license_path)
        form.addRow("WinRAR License File", self._make_path_picker(self._winrar_license, "Select WinRAR License", "Key Files (*.key);;All Files (*)"))

        self._java_version = QComboBox()
        self._java_version.setEditable(True)
        self._java_version.addItems(["", "8", "11", "17", "21"])
        self._java_version.setCurrentText(self._settings.java_version)
        if self._java_version.lineEdit():
            self._java_version.lineEdit().setPlaceholderText("8.0.391.13")
        java_row = QWidget()
        java_layout = QHBoxLayout(java_row)
        java_layout.setContentsMargins(0, 0, 0, 0)
        java_layout.addWidget(self._java_version)
        self._btn_java_versions = QPushButton("List Versions")
        self._btn_java_versions.clicked.connect(self._list_java_versions)
        java_layout.addWidget(self._btn_java_versions)
        form.addRow("Java Version", java_row)

        self._java_hint = QLabel("Leave blank for latest; enter a full winget version if needed.")
        form.addRow("", self._java_hint)

        self._teamviewer_args = QLineEdit(self._settings.teamviewer_args)
        self._teamviewer_args.setPlaceholderText("/qn")
        form.addRow("TeamViewer Host Args", self._teamviewer_args)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _make_path_picker(self, field: QLineEdit, title: str, filter_text: str) -> QWidget:
        container = QWidget()
        row = QHBoxLayout(container)
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(field)
        browse = QPushButton("Browse")
        browse.clicked.connect(lambda: self._browse_for_path(field, title, filter_text))
        row.addWidget(browse)
        return container

    def _browse_for_path(self, field: QLineEdit, title: str, filter_text: str) -> None:
        current = field.text().strip()
        start_dir = str(Path(current).parent) if current else str(Path.home())
        path, _ = QFileDialog.getOpenFileName(self, title, start_dir, filter_text)
        if path:
            field.setText(path)

    def _save(self) -> None:
        cid_value = self._crowdstrike_cid.text().strip()
        if cid_value.upper().startswith("CID="):
            cid_value = cid_value[4:].strip()
        self._settings.crowdstrike_cid = cid_value
        self._settings.crowdstrike_download_url = self._crowdstrike_url.text().strip()
        self._settings.office_2024_xml_path = self._office_2024_path.text().strip()
        self._settings.office_365_xml_path = self._office_365_path.text().strip()
        self._settings.winrar_license_path = self._winrar_license.text().strip()
        self._settings.java_version = self._java_version.currentText().strip()
        self._settings.teamviewer_args = self._teamviewer_args.text().strip()
        self._store.save(self._settings)
        self.accept()

    def _list_java_versions(self) -> None:
        exe = shutil.which("winget")
        if not exe:
            QMessageBox.warning(self, "Winget Missing", "winget executable not found in PATH.")
            return
        cmd = [
            exe,
            "show",
            "--id",
            "Oracle.JavaRuntimeEnvironment",
            "--exact",
            "--versions",
            "--accept-source-agreements",
        ]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
        except Exception as exc:
            QMessageBox.warning(self, "Winget Error", f"Unable to query winget: {exc}")
            return
        if result.returncode != 0:
            message = result.stderr.strip() or result.stdout.strip() or "Unknown winget error"
            QMessageBox.warning(self, "Winget Error", message)
            return
        versions = _extract_versions(result.stdout)
        text = "Available Java versions (winget):"
        if versions:
            text += "\n" + "\n".join(versions)
        else:
            text += "\nNo versions parsed from winget output."
        box = QMessageBox(self)
        box.setWindowTitle("Java Versions")
        box.setText(text)
        if result.stdout.strip():
            box.setDetailedText(result.stdout.strip())
        box.exec()


def _extract_versions(output: str) -> list[str]:
    lines = output.splitlines()
    versions: list[str] = []
    for line in lines:
        match = re.match(r"\s*([0-9]+(?:\.[0-9]+){1,3})\s*$", line)
        if match:
            versions.append(match.group(1))
    if versions:
        return versions
    for line in lines:
        match = re.search(r"\b([0-9]+(?:\.[0-9]+){1,3})\b", line)
        if match:
            versions.append(match.group(1))
    return sorted(set(versions))
