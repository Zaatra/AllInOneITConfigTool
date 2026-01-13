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
    QInputDialog,
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
        self._crowdstrike_cid.setPlaceholderText("Example: CID=00000000000000000000000000000000-00")
        form.addRow("CrowdStrike CID", self._crowdstrike_cid)

        self._crowdstrike_url = QLineEdit(self._settings.crowdstrike_download_url)
        form.addRow("CrowdStrike Download URL", self._crowdstrike_url)

        self._office_2024_path = QLineEdit(self._settings.office_2024_xml_path)
        form.addRow("Office 2024 XML", self._make_path_picker(self._office_2024_path, "Select Office 2024 XML", "XML Files (*.xml);;All Files (*)"))

        self._office_365_path = QLineEdit(self._settings.office_365_xml_path)
        form.addRow("Office 365 XML", self._make_path_picker(self._office_365_path, "Select Office 365 XML", "XML Files (*.xml);;All Files (*)"))

        self._odt_setup_path = QLineEdit(self._settings.odt_setup_path)
        form.addRow(
            "Office Deployment Tool EXE",
            self._make_path_picker(
                self._odt_setup_path,
                "Select Office Deployment Tool",
                "Executable Files (*.exe);;All Files (*)",
            ),
        )

        self._winrar_license = QLineEdit(self._settings.winrar_license_path)
        form.addRow("WinRAR License File", self._make_path_picker(self._winrar_license, "Select WinRAR License", "Key Files (*.key);;All Files (*)"))

        self._hp_legacy_repo = QLineEdit(self._settings.hp_legacy_repo_root)
        form.addRow("HP Legacy Repo Root", self._make_dir_picker(self._hp_legacy_repo, "Select HP Legacy Repo Root"))

        self._java_version = QComboBox()
        self._java_version.setEditable(True)
        self._java_version.setCurrentText(self._settings.java_version)
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

        self._teamviewer_mode = QComboBox()
        self._teamviewer_mode.addItem("Standard (winget)", "winget")
        self._teamviewer_mode.addItem("Custom MSI", "msi")
        mode = self._settings.teamviewer_install_mode.strip().lower()
        if mode != "msi":
            mode = "winget"
        index = self._teamviewer_mode.findData(mode)
        if index >= 0:
            self._teamviewer_mode.setCurrentIndex(index)
        form.addRow("TeamViewer Install Mode", self._teamviewer_mode)

        self._teamviewer_args = QLineEdit(self._settings.teamviewer_args)
        form.addRow("TeamViewer Host Args", self._teamviewer_args)

        self._teamviewer_msi_path = QLineEdit(self._settings.teamviewer_msi_path)
        self._teamviewer_msi_row = self._make_path_picker(
            self._teamviewer_msi_path,
            "Select TeamViewer MSI",
            "MSI Files (*.msi);;All Files (*)",
        )
        form.addRow("TeamViewer MSI", self._teamviewer_msi_row)

        self._teamviewer_customconfig = QLineEdit(self._settings.teamviewer_customconfig_id)
        form.addRow("TeamViewer CUSTOMCONFIGID", self._teamviewer_customconfig)

        self._teamviewer_assignment = QLineEdit(self._settings.teamviewer_assignment_id)
        form.addRow("TeamViewer ASSIGNMENTID", self._teamviewer_assignment)

        self._teamviewer_settings_file = QLineEdit(self._settings.teamviewer_settings_file)
        self._teamviewer_settings_row = self._make_path_picker(
            self._teamviewer_settings_file,
            "Select TeamViewer Settings File",
            "TVOPT Files (*.tvopt);;All Files (*)",
        )
        form.addRow("TeamViewer SETTINGSFILE", self._teamviewer_settings_row)

        self._teamviewer_msi_args = QLineEdit()
        self._teamviewer_msi_args.setReadOnly(True)
        form.addRow("TeamViewer MSI Args", self._teamviewer_msi_args)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._teamviewer_mode.currentIndexChanged.connect(self._update_teamviewer_mode_ui)
        self._teamviewer_customconfig.textChanged.connect(self._update_teamviewer_msi_args)
        self._teamviewer_assignment.textChanged.connect(self._update_teamviewer_msi_args)
        self._teamviewer_settings_file.textChanged.connect(self._update_teamviewer_msi_args)
        self._update_teamviewer_msi_args()
        self._update_teamviewer_mode_ui()

    def _make_path_picker(self, field: QLineEdit, title: str, filter_text: str) -> QWidget:
        container = QWidget()
        row = QHBoxLayout(container)
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(field)
        browse = QPushButton("Browse")
        browse.clicked.connect(lambda: self._browse_for_path(field, title, filter_text))
        row.addWidget(browse)
        return container

    def _make_dir_picker(self, field: QLineEdit, title: str) -> QWidget:
        container = QWidget()
        row = QHBoxLayout(container)
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(field)
        browse = QPushButton("Browse")
        browse.clicked.connect(lambda: self._browse_for_dir(field, title))
        row.addWidget(browse)
        return container

    def _browse_for_path(self, field: QLineEdit, title: str, filter_text: str) -> None:
        current = field.text().strip()
        start_dir = str(Path(current).parent) if current else str(Path.home())
        path, _ = QFileDialog.getOpenFileName(self, title, start_dir, filter_text)
        if path:
            field.setText(path)

    def _browse_for_dir(self, field: QLineEdit, title: str) -> None:
        current = field.text().strip()
        start_dir = current or str(Path.home())
        path = QFileDialog.getExistingDirectory(self, title, start_dir)
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
        self._settings.odt_setup_path = self._odt_setup_path.text().strip()
        self._settings.winrar_license_path = self._winrar_license.text().strip()
        self._settings.hp_legacy_repo_root = self._hp_legacy_repo.text().strip()
        self._settings.java_version = self._java_version.currentText().strip()
        teamviewer_mode = self._teamviewer_mode.currentData() or "winget"
        teamviewer_msi_path = self._teamviewer_msi_path.text().strip()
        teamviewer_customconfig = self._teamviewer_customconfig.text().strip()
        teamviewer_assignment = self._teamviewer_assignment.text().strip()
        teamviewer_settings_file = self._teamviewer_settings_file.text().strip()
        if teamviewer_mode == "msi":
            missing = []
            if not teamviewer_msi_path:
                missing.append("TeamViewer MSI path")
            elif not Path(teamviewer_msi_path).is_file():
                missing.append("TeamViewer MSI file not found")
            if not teamviewer_customconfig:
                missing.append("TeamViewer CUSTOMCONFIGID")
            if not teamviewer_assignment:
                missing.append("TeamViewer ASSIGNMENTID")
            if not teamviewer_settings_file:
                missing.append("TeamViewer SETTINGSFILE")
            elif not teamviewer_settings_file.lower().endswith(".tvopt"):
                missing.append("TeamViewer SETTINGSFILE must end with .tvopt")
            elif not Path(teamviewer_settings_file).is_file():
                missing.append("TeamViewer SETTINGSFILE not found")
            if missing:
                message = "TeamViewer MSI settings missing:\n" + "\n".join(f"- {item}" for item in missing)
                QMessageBox.warning(self, "Settings Required", message)
                return
        self._settings.teamviewer_install_mode = str(teamviewer_mode)
        self._settings.teamviewer_args = self._teamviewer_args.text().strip()
        self._settings.teamviewer_msi_path = teamviewer_msi_path
        self._settings.teamviewer_customconfig_id = teamviewer_customconfig
        self._settings.teamviewer_assignment_id = teamviewer_assignment
        self._settings.teamviewer_settings_file = teamviewer_settings_file
        self._store.save(self._settings)
        self.accept()

    def _update_teamviewer_mode_ui(self) -> None:
        use_msi = self._teamviewer_mode.currentData() == "msi"
        self._teamviewer_args.setEnabled(not use_msi)
        for widget in (
            self._teamviewer_msi_row,
            self._teamviewer_customconfig,
            self._teamviewer_assignment,
            self._teamviewer_settings_row,
            self._teamviewer_msi_args,
        ):
            widget.setEnabled(use_msi)

    def _update_teamviewer_msi_args(self) -> None:
        parts = ["/qn", "/norestart"]
        custom_config = self._teamviewer_customconfig.text().strip()
        if custom_config:
            parts.append(f"CUSTOMCONFIGID={custom_config}")
        assignment = self._teamviewer_assignment.text().strip()
        if assignment:
            parts.append(f"ASSIGNMENTID={assignment}")
        settings_file = self._teamviewer_settings_file.text().strip()
        if settings_file:
            if not (settings_file.startswith('\"') and settings_file.endswith('\"')):
                settings_file = f'\"{settings_file}\"'
            parts.append(f"SETTINGSFILE={settings_file}")
        self._teamviewer_msi_args.setText(" ".join(parts))

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
        if not versions:
            text = "No versions parsed from winget output."
            box = QMessageBox(self)
            box.setWindowTitle("Java Versions")
            box.setText(text)
            if result.stdout.strip():
                box.setDetailedText(result.stdout.strip())
            box.exec()
            return

        choice, ok = QInputDialog.getItem(
            self,
            "Java Versions",
            "Select a Java version:",
            versions,
            0,
            False,
        )
        if ok and choice:
            self._java_version.setCurrentText(choice)


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
