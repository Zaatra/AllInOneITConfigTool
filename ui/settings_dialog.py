"""Settings dialog for user-provided installer values."""
from __future__ import annotations

import json
import re
import shutil
import subprocess
from pathlib import Path
from urllib.parse import urlparse

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
        self._crowdstrike_cid_label = QLabel("CrowdStrike CID")
        form.addRow(self._crowdstrike_cid_label, self._crowdstrike_cid)

        self._crowdstrike_url = QLineEdit(self._settings.crowdstrike_download_url)
        self._crowdstrike_url.setPlaceholderText("SharePoint link (..sharepoint...)")
        self._crowdstrike_url_label = QLabel("CrowdStrike SharePoint URL")
        form.addRow(self._crowdstrike_url_label, self._crowdstrike_url)

        self._forticlient_url = QLineEdit(self._settings.forticlient_download_url)
        self._forticlient_url.setPlaceholderText("SharePoint link (..sharepoint...)")
        self._forticlient_url_label = QLabel("FortiClient VPN SharePoint URL")
        form.addRow(self._forticlient_url_label, self._forticlient_url)

        self._office_2024_path = QLineEdit(self._settings.office_2024_xml_path)
        self._office_2024_label = QLabel("Office 2024 XML")
        form.addRow(
            self._office_2024_label,
            self._make_path_picker(self._office_2024_path, "Select Office 2024 XML", "XML Files (*.xml);;All Files (*)"),
        )

        self._office_365_path = QLineEdit(self._settings.office_365_xml_path)
        self._office_365_label = QLabel("Office 365 XML")
        form.addRow(
            self._office_365_label,
            self._make_path_picker(self._office_365_path, "Select Office 365 XML", "XML Files (*.xml);;All Files (*)"),
        )

        self._odt_setup_path = QLineEdit(self._settings.odt_setup_path)
        self._odt_setup_label = QLabel("Office Deployment Tool EXE")
        form.addRow(
            self._odt_setup_label,
            self._make_path_picker(
                self._odt_setup_path,
                "Select Office Deployment Tool",
                "Executable Files (*.exe);;All Files (*)",
            ),
        )

        self._winrar_license = QLineEdit(self._settings.winrar_license_path)
        self._winrar_license_label = QLabel("WinRAR License File")
        form.addRow(
            self._winrar_license_label,
            self._make_path_picker(self._winrar_license, "Select WinRAR License", "Key Files (*.key);;All Files (*)"),
        )

        self._hp_legacy_repo = QLineEdit(self._settings.hp_legacy_repo_root)
        self._hp_legacy_repo_label = QLabel("HP Legacy Repo Root")
        form.addRow(self._hp_legacy_repo_label, self._make_dir_picker(self._hp_legacy_repo, "Select HP Legacy Repo Root"))

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
        self._java_version_label = QLabel("Java Version")
        form.addRow(self._java_version_label, java_row)

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
        self._teamviewer_mode_label = QLabel("TeamViewer Install Mode")
        form.addRow(self._teamviewer_mode_label, self._teamviewer_mode)

        self._teamviewer_args = QLineEdit(self._settings.teamviewer_args)
        self._teamviewer_args_label = QLabel("TeamViewer Host Args")
        form.addRow(self._teamviewer_args_label, self._teamviewer_args)

        self._teamviewer_msi_path = QLineEdit(self._settings.teamviewer_msi_path)
        self._teamviewer_msi_row = self._make_path_picker(
            self._teamviewer_msi_path,
            "Select TeamViewer MSI",
            "MSI Files (*.msi);;All Files (*)",
        )
        self._teamviewer_msi_label = QLabel("TeamViewer MSI")
        form.addRow(self._teamviewer_msi_label, self._teamviewer_msi_row)

        self._teamviewer_customconfig = QLineEdit(self._settings.teamviewer_customconfig_id)
        self._teamviewer_customconfig_label = QLabel("TeamViewer CUSTOMCONFIGID")
        form.addRow(self._teamviewer_customconfig_label, self._teamviewer_customconfig)

        self._teamviewer_assignment = QLineEdit(self._settings.teamviewer_assignment_id)
        self._teamviewer_assignment_label = QLabel("TeamViewer ASSIGNMENTID")
        form.addRow(self._teamviewer_assignment_label, self._teamviewer_assignment)

        self._teamviewer_settings_file = QLineEdit(self._settings.teamviewer_settings_file)
        self._teamviewer_settings_row = self._make_path_picker(
            self._teamviewer_settings_file,
            "Select TeamViewer Settings File",
            "TVOPT Files (*.tvopt);;All Files (*)",
        )
        self._teamviewer_settings_label = QLabel("TeamViewer SETTINGSFILE")
        form.addRow(self._teamviewer_settings_label, self._teamviewer_settings_row)

        self._teamviewer_msi_args = QLineEdit()
        self._teamviewer_msi_args.setReadOnly(True)
        self._teamviewer_msi_args_label = QLabel("TeamViewer MSI Args")
        form.addRow(self._teamviewer_msi_args_label, self._teamviewer_msi_args)

        layout.addLayout(form)

        import_export_row = QHBoxLayout()
        self._btn_import = QPushButton("Import...")
        self._btn_export = QPushButton("Export...")
        import_export_row.addWidget(self._btn_import)
        import_export_row.addWidget(self._btn_export)
        import_export_row.addStretch()
        layout.addLayout(import_export_row)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._teamviewer_mode.currentIndexChanged.connect(self._update_teamviewer_mode_ui)
        self._teamviewer_customconfig.textChanged.connect(self._update_teamviewer_msi_args)
        self._teamviewer_assignment.textChanged.connect(self._update_teamviewer_msi_args)
        self._teamviewer_settings_file.textChanged.connect(self._update_teamviewer_msi_args)
        self._teamviewer_msi_path.textChanged.connect(self._update_validation)
        self._teamviewer_args.textChanged.connect(self._update_validation)
        self._crowdstrike_cid.textChanged.connect(self._update_validation)
        self._crowdstrike_url.textChanged.connect(self._update_validation)
        self._forticlient_url.textChanged.connect(self._update_validation)
        self._office_2024_path.textChanged.connect(self._update_validation)
        self._office_365_path.textChanged.connect(self._update_validation)
        self._odt_setup_path.textChanged.connect(self._update_validation)
        self._winrar_license.textChanged.connect(self._update_validation)
        self._hp_legacy_repo.textChanged.connect(self._update_validation)
        self._java_version.currentTextChanged.connect(self._update_validation)
        self._btn_import.clicked.connect(self._import_settings)
        self._btn_export.clicked.connect(self._export_settings)
        self._update_teamviewer_msi_args()
        self._update_teamviewer_mode_ui()
        self._update_validation()

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
        next_settings = self._settings_from_fields()
        teamviewer_mode = next_settings.teamviewer_install_mode
        if teamviewer_mode == "msi":
            missing = self._teamviewer_msi_issues()
            if missing:
                message = "TeamViewer MSI settings missing:\n" + "\n".join(f"- {item}" for item in missing)
                QMessageBox.warning(self, "Settings Required", message)
                return
        for key, value in next_settings.to_dict().items():
            setattr(self._settings, key, value)
        self._store.save(self._settings)
        self.accept()

    def _import_settings(self) -> None:
        start_dir = str(self._store.path.parent if self._store.path.parent.exists() else Path.home())
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Import Settings",
            start_dir,
            "JSON Files (*.json);;All Files (*)",
        )
        if not path:
            return
        try:
            payload = json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception as exc:
            QMessageBox.warning(self, "Import Failed", f"Unable to read settings file:\n{exc}")
            return
        if not isinstance(payload, dict):
            QMessageBox.warning(self, "Import Failed", "Settings file must contain a JSON object.")
            return
        loaded = UserSettings.from_dict(payload)
        self._set_fields_from_settings(loaded)
        self._update_validation()
        QMessageBox.information(self, "Import Complete", "Settings loaded. Click Save to apply them.")

    def _export_settings(self) -> None:
        start_dir = str(self._store.path.parent if self._store.path.parent.exists() else Path.home())
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Settings",
            str(Path(start_dir) / "settings-export.json"),
            "JSON Files (*.json);;All Files (*)",
        )
        if not path:
            return
        export_settings = self._settings_from_fields()
        try:
            Path(path).write_text(
                json.dumps(export_settings.to_dict(), indent=2, sort_keys=True),
                encoding="utf-8",
            )
        except Exception as exc:
            QMessageBox.warning(self, "Export Failed", f"Unable to write settings file:\n{exc}")
            return
        QMessageBox.information(self, "Export Complete", f"Settings exported to:\n{path}")

    def _settings_from_fields(self) -> UserSettings:
        cid_value = self._crowdstrike_cid.text().strip()
        if cid_value.upper().startswith("CID="):
            cid_value = cid_value[4:].strip()
        teamviewer_mode = self._teamviewer_mode.currentData() or "winget"
        return UserSettings(
            crowdstrike_cid=cid_value,
            crowdstrike_download_url=self._crowdstrike_url.text().strip(),
            forticlient_download_url=self._forticlient_url.text().strip(),
            office_2024_xml_path=self._office_2024_path.text().strip(),
            office_365_xml_path=self._office_365_path.text().strip(),
            odt_setup_path=self._odt_setup_path.text().strip(),
            winrar_license_path=self._winrar_license.text().strip(),
            java_version=self._java_version.currentText().strip(),
            teamviewer_args=self._teamviewer_args.text().strip(),
            hp_legacy_repo_root=self._hp_legacy_repo.text().strip(),
            teamviewer_install_mode=str(teamviewer_mode),
            teamviewer_msi_path=self._teamviewer_msi_path.text().strip(),
            teamviewer_customconfig_id=self._teamviewer_customconfig.text().strip(),
            teamviewer_assignment_id=self._teamviewer_assignment.text().strip(),
            teamviewer_settings_file=self._teamviewer_settings_file.text().strip(),
        )

    def _set_fields_from_settings(self, settings: UserSettings) -> None:
        self._crowdstrike_cid.setText(settings.crowdstrike_cid)
        self._crowdstrike_url.setText(settings.crowdstrike_download_url)
        self._forticlient_url.setText(settings.forticlient_download_url)
        self._office_2024_path.setText(settings.office_2024_xml_path)
        self._office_365_path.setText(settings.office_365_xml_path)
        self._odt_setup_path.setText(settings.odt_setup_path)
        self._winrar_license.setText(settings.winrar_license_path)
        self._hp_legacy_repo.setText(settings.hp_legacy_repo_root)
        self._java_version.setCurrentText(settings.java_version)
        mode = settings.teamviewer_install_mode.strip().lower()
        if mode != "msi":
            mode = "winget"
        index = self._teamviewer_mode.findData(mode)
        if index >= 0:
            self._teamviewer_mode.setCurrentIndex(index)
        self._teamviewer_args.setText(settings.teamviewer_args)
        self._teamviewer_msi_path.setText(settings.teamviewer_msi_path)
        self._teamviewer_customconfig.setText(settings.teamviewer_customconfig_id)
        self._teamviewer_assignment.setText(settings.teamviewer_assignment_id)
        self._teamviewer_settings_file.setText(settings.teamviewer_settings_file)
        self._update_teamviewer_msi_args()
        self._update_teamviewer_mode_ui()

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
        self._update_validation()

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
        self._update_validation()

    def _teamviewer_msi_issues(self) -> list[str]:
        missing: list[str] = []
        teamviewer_msi_path = self._clean_path_value(self._teamviewer_msi_path.text())
        teamviewer_customconfig = self._teamviewer_customconfig.text().strip()
        teamviewer_assignment = self._teamviewer_assignment.text().strip()
        teamviewer_settings_file = self._clean_path_value(self._teamviewer_settings_file.text())
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
        return missing

    def _update_validation(self) -> None:
        use_msi = self._teamviewer_mode.currentData() == "msi"
        self._set_label_valid(self._crowdstrike_cid_label, self._is_crowdstrike_cid_valid())
        self._set_label_valid(
            self._crowdstrike_url_label,
            self._is_sharepoint_url_valid(self._crowdstrike_url.text(), allow_empty=True),
        )
        self._set_label_valid(
            self._forticlient_url_label,
            self._is_sharepoint_url_valid(self._forticlient_url.text(), allow_empty=True),
        )
        self._set_label_valid(self._office_2024_label, self._is_file_valid(self._office_2024_path.text(), suffixes=(".xml",)))
        self._set_label_valid(self._office_365_label, self._is_file_valid(self._office_365_path.text(), suffixes=(".xml",)))
        self._set_label_valid(
            self._odt_setup_label,
            self._is_file_valid(self._odt_setup_path.text(), suffixes=(".exe",), allow_empty=True),
        )
        self._set_label_valid(self._winrar_license_label, self._is_file_valid(self._winrar_license.text(), suffixes=(".key",)))
        self._set_label_valid(self._hp_legacy_repo_label, self._is_dir_valid(self._hp_legacy_repo.text(), allow_empty=True))
        self._set_label_valid(self._java_version_label, self._is_java_version_valid())
        self._set_label_valid(self._teamviewer_mode_label, True)
        self._set_label_valid(self._teamviewer_args_label, (not use_msi) or not self._teamviewer_msi_issues())
        self._set_label_valid(
            self._teamviewer_msi_label,
            (not use_msi) or self._is_file_valid(self._teamviewer_msi_path.text(), suffixes=(".msi",)),
        )
        self._set_label_valid(self._teamviewer_customconfig_label, (not use_msi) or bool(self._teamviewer_customconfig.text().strip()))
        self._set_label_valid(self._teamviewer_assignment_label, (not use_msi) or bool(self._teamviewer_assignment.text().strip()))
        self._set_label_valid(
            self._teamviewer_settings_label,
            (not use_msi) or self._is_teamviewer_settings_file_valid(),
        )
        self._set_label_valid(self._teamviewer_msi_args_label, (not use_msi) or not self._teamviewer_msi_issues())

    def _set_label_valid(self, label: QLabel, valid: bool) -> None:
        color = "#4caf50" if valid else "#f44336"
        label.setStyleSheet(f"color: {color}; font-weight: bold;")

    def _clean_path_value(self, value: str) -> str:
        cleaned = value.strip()
        if len(cleaned) >= 2 and cleaned[0] == cleaned[-1] and cleaned[0] in ('"', "'"):
            cleaned = cleaned[1:-1].strip()
        return cleaned

    def _is_file_valid(
        self,
        value: str,
        *,
        suffixes: tuple[str, ...] | None = None,
        allow_empty: bool = False,
    ) -> bool:
        cleaned = self._clean_path_value(value)
        if not cleaned:
            return allow_empty
        path = Path(cleaned)
        if suffixes and path.suffix.lower() not in suffixes:
            return False
        return path.exists() and path.is_file()

    def _is_dir_valid(self, value: str, *, allow_empty: bool = False) -> bool:
        cleaned = self._clean_path_value(value)
        if not cleaned:
            return allow_empty
        path = Path(cleaned)
        return path.exists() and path.is_dir()

    def _is_url_valid(self, value: str, *, allow_empty: bool = False) -> bool:
        cleaned = value.strip()
        if not cleaned:
            return allow_empty
        parsed = urlparse(cleaned)
        return parsed.scheme in {"http", "https"} and bool(parsed.netloc)

    def _is_sharepoint_url_valid(self, value: str, *, allow_empty: bool = False) -> bool:
        cleaned = value.strip()
        if not cleaned:
            return allow_empty
        parsed = urlparse(cleaned)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            return False
        host = parsed.netloc.lower()
        return host.endswith("sharepoint.com") or ".sharepoint." in host or host.startswith("sharepoint.")

    def _is_crowdstrike_cid_valid(self) -> bool:
        cid_value = self._crowdstrike_cid.text().strip()
        if cid_value.upper().startswith("CID="):
            cid_value = cid_value[4:].strip()
        return bool(re.fullmatch(r"[0-9A-Fa-f]{32}(?:-[0-9A-Fa-f]{2})?", cid_value))

    def _is_java_version_valid(self) -> bool:
        version = self._java_version.currentText().strip()
        if not version:
            return True
        return bool(re.fullmatch(r"\d+(?:\.\d+){1,3}", version))

    def _is_teamviewer_settings_file_valid(self) -> bool:
        settings_file = self._clean_path_value(self._teamviewer_settings_file.text())
        if not settings_file:
            return False
        if not settings_file.lower().endswith(".tvopt"):
            return False
        return Path(settings_file).is_file()

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
