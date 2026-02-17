"""User-configurable settings persisted locally."""
from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any


SETTINGS_DIRNAME = ".allinone_it_config"
SETTINGS_FILENAME = "settings.json"


def default_settings_path() -> Path:
    return Path.home() / SETTINGS_DIRNAME / SETTINGS_FILENAME


@dataclass
class UserSettings:
    crowdstrike_cid: str = ""
    crowdstrike_download_url: str = ""
    forticlient_download_url: str = ""
    office_2024_xml_path: str = ""
    office_365_xml_path: str = ""
    odt_setup_path: str = ""
    winrar_license_path: str = ""
    java_version: str = ""
    teamviewer_args: str = ""
    teamviewer_install_mode: str = ""
    teamviewer_msi_path: str = ""
    teamviewer_customconfig_id: str = ""
    teamviewer_assignment_id: str = ""
    teamviewer_settings_file: str = ""

    def to_dict(self) -> dict[str, str]:
        return {
            "crowdstrike_cid": self.crowdstrike_cid,
            "crowdstrike_download_url": self.crowdstrike_download_url,
            "forticlient_download_url": self.forticlient_download_url,
            "office_2024_xml_path": self.office_2024_xml_path,
            "office_365_xml_path": self.office_365_xml_path,
            "odt_setup_path": self.odt_setup_path,
            "winrar_license_path": self.winrar_license_path,
            "java_version": self.java_version,
            "teamviewer_args": self.teamviewer_args,
            "teamviewer_install_mode": self.teamviewer_install_mode,
            "teamviewer_msi_path": self.teamviewer_msi_path,
            "teamviewer_customconfig_id": self.teamviewer_customconfig_id,
            "teamviewer_assignment_id": self.teamviewer_assignment_id,
            "teamviewer_settings_file": self.teamviewer_settings_file,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "UserSettings":
        def _get(key: str) -> str:
            value = data.get(key, "")
            return str(value) if value is not None else ""

        return cls(
            crowdstrike_cid=_get("crowdstrike_cid"),
            crowdstrike_download_url=_get("crowdstrike_download_url"),
            forticlient_download_url=_get("forticlient_download_url"),
            office_2024_xml_path=_get("office_2024_xml_path"),
            office_365_xml_path=_get("office_365_xml_path"),
            odt_setup_path=_get("odt_setup_path"),
            winrar_license_path=_get("winrar_license_path"),
            java_version=_get("java_version"),
            teamviewer_args=_get("teamviewer_args"),
            teamviewer_install_mode=_get("teamviewer_install_mode"),
            teamviewer_msi_path=_get("teamviewer_msi_path"),
            teamviewer_customconfig_id=_get("teamviewer_customconfig_id"),
            teamviewer_assignment_id=_get("teamviewer_assignment_id"),
            teamviewer_settings_file=_get("teamviewer_settings_file"),
        )

    def load_office_xml(self, app_name: str) -> str:
        if app_name == "Office 2024 LTSC":
            path_str = self.office_2024_xml_path.strip()
        elif app_name == "Office 365 Ent":
            path_str = self.office_365_xml_path.strip()
        else:
            raise ValueError(f"No Office XML mapping for {app_name}")
        if not path_str:
            raise ValueError(f"Office XML path not configured for {app_name}")
        path = Path(path_str)
        if not path.exists() or not path.is_file():
            raise FileNotFoundError(f"Office XML file not found: {path}")
        try:
            return path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            return path.read_text(encoding="utf-8-sig")


class SettingsStore:
    def __init__(self, path: Path | None = None) -> None:
        self._path = path or default_settings_path()

    @property
    def path(self) -> Path:
        return self._path

    def exists(self) -> bool:
        return self._path.exists()

    def load(self) -> UserSettings:
        if not self._path.exists():
            return UserSettings()
        try:
            data = json.loads(self._path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return UserSettings()
        if not isinstance(data, dict):
            return UserSettings()
        return UserSettings.from_dict(data)

    def save(self, settings: UserSettings) -> None:
        self._path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(settings.to_dict(), indent=2, sort_keys=True)
        self._path.write_text(payload, encoding="utf-8")
