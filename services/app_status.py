"""Application status/version scanning and update checks."""
from __future__ import annotations

import re
import shutil
import subprocess
import sys
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Mapping

from services.installer import (
    HPSADownloader,
    IVMSDownloader,
    LocalInstallerVersionInfo,
    WingetClient,
    WingetError,
    _download_sharepoint_exe,
    _is_sharepoint_url,
)
from allinone_it_config.app_registry import AppEntry
from allinone_it_config.paths import get_application_directory
from allinone_it_config.user_settings import UserSettings

try:  # Windows-only dependency, optional for non-Windows hosts
    import winreg  # type: ignore
except ImportError:  # pragma: no cover - not available on Linux runners
    winreg = None  # type: ignore


_STATUS_NOT_INSTALLED = "Not Installed"
_STATUS_INSTALLED = "Installed"
_STATUS_UP_TO_DATE = "Up to Date"
_STATUS_UPDATE_AVAILABLE = "Update Available"
_STATUS_UNKNOWN = "Unknown"

_LEVEL_NOT_INSTALLED = "not_installed"
_LEVEL_INSTALLED = "installed"
_LEVEL_UP_TO_DATE = "up_to_date"
_LEVEL_UPDATE_AVAILABLE = "update_available"
_LEVEL_UNKNOWN = "unknown"

_LATEST_UNKNOWN = {"", "N/A", "Error", "Winget Missing", "Unknown"}


@dataclass(frozen=True)
class InstalledInfo:
    app: AppEntry
    installed_text: str
    installed_version: str | None
    installed_x86: str | None
    installed_x64: str | None
    is_installed: bool
    is_known: bool


@dataclass(frozen=True)
class AppUpdateResult:
    app: AppEntry
    installed_text: str
    latest_text: str
    status: str
    status_level: str


@dataclass(frozen=True)
class _UninstallEntry:
    display_name: str
    display_version: str


class AppStatusService:
    def __init__(
        self,
        apps: Iterable[AppEntry],
        *,
        working_dir: Path | str | None = None,
        winget_client: WingetClient | None = None,
        settings: UserSettings | None = None,
    ) -> None:
        self._apps = list(apps)
        self._working_dir = Path(working_dir) if working_dir is not None else get_application_directory()
        self._winget = winget_client or WingetClient()
        self._settings = settings or UserSettings()
        self._direct_downloaders = {"iVMS-4200": IVMSDownloader(), "HP Support Asst": HPSADownloader()}

    def scan_installed(self) -> list[InstalledInfo]:
        entries = self._read_uninstall_entries()
        results: list[InstalledInfo] = []
        for app in self._apps:
            if app.name == "Office Deployment Tool":
                version = self._get_local_odt_version()
                if version:
                    results.append(
                        InstalledInfo(
                            app=app,
                            installed_text=version,
                            installed_version=version,
                            installed_x86=None,
                            installed_x64=None,
                            is_installed=True,
                            is_known=True,
                        )
                    )
                else:
                    results.append(
                        InstalledInfo(
                            app=app,
                            installed_text=_STATUS_NOT_INSTALLED,
                            installed_version=None,
                            installed_x86=None,
                            installed_x64=None,
                            is_installed=False,
                            is_known=True,
                        )
                    )
                continue
            if app.name == "HP Support Asst":
                version = self._get_hp_support_appx_version()
                if version:
                    results.append(
                        InstalledInfo(
                            app=app,
                            installed_text=version,
                            installed_version=version,
                            installed_x86=None,
                            installed_x64=None,
                            is_installed=True,
                            is_known=True,
                        )
                    )
                else:
                    results.append(
                        InstalledInfo(
                            app=app,
                            installed_text=_STATUS_NOT_INSTALLED,
                            installed_version=None,
                            installed_x86=None,
                            installed_x64=None,
                            is_installed=False,
                            is_known=True,
                        )
                    )
                continue
            if app.dual_arch and app.vc_key:
                vc_map = self._get_vc_installed_map(app.vc_key, entries)
                installed_text = self._format_vc_versions(vc_map)
                is_installed = bool(vc_map["x86"] or vc_map["x64"])
                results.append(
                    InstalledInfo(
                        app=app,
                        installed_text=installed_text,
                        installed_version=None,
                        installed_x86=vc_map["x86"],
                        installed_x64=vc_map["x64"],
                        is_installed=is_installed,
                        is_known=True,
                    )
                )
                continue

            if app.detection_pattern:
                version = self._get_best_version(app.detection_pattern, entries)
                if version:
                    results.append(
                        InstalledInfo(
                            app=app,
                            installed_text=version,
                            installed_version=version,
                            installed_x86=None,
                            installed_x64=None,
                            is_installed=True,
                            is_known=True,
                        )
                    )
                else:
                    results.append(
                        InstalledInfo(
                            app=app,
                            installed_text=_STATUS_NOT_INSTALLED,
                            installed_version=None,
                            installed_x86=None,
                            installed_x64=None,
                            is_installed=False,
                            is_known=True,
                        )
                    )
            else:
                results.append(
                    InstalledInfo(
                        app=app,
                        installed_text=_STATUS_UNKNOWN,
                        installed_version=None,
                        installed_x86=None,
                        installed_x64=None,
                        is_installed=False,
                        is_known=False,
                    )
                )
        return results

    def check_updates(
        self,
        installed_map: Mapping[str, InstalledInfo] | None = None,
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> list[AppUpdateResult]:
        if installed_map is None:
            installed_map = {info.app.name: info for info in self.scan_installed()}
        if self._winget.is_available():
            try:
                self._winget.update_sources("winget")
                if any(app.source == "msstore" for app in self._apps):
                    self._winget.update_sources("msstore")
            except Exception:
                pass
        results: list[AppUpdateResult] = []
        total = len(self._apps)
        for index, app in enumerate(self._apps, start=1):
            info = installed_map.get(app.name)
            if info is None:
                info = InstalledInfo(
                    app=app,
                    installed_text=_STATUS_UNKNOWN,
                    installed_version=None,
                    installed_x86=None,
                    installed_x64=None,
                    is_installed=False,
                    is_known=False,
                )
            latest = self._get_latest_version(app)
            status, level = self._evaluate_status(app, info, latest)
            results.append(
                AppUpdateResult(
                    app=app,
                    installed_text=info.installed_text,
                    latest_text=latest,
                    status=status,
                    status_level=level,
                )
            )
            if progress_callback:
                progress_callback(index, total, app.name)
        return results

    def offline_installer_status(
        self,
        app: AppEntry,
        local_versions: LocalInstallerVersionInfo,
        latest_text: str,
    ) -> str | None:
        if not local_versions.has_any():
            return None
        if self._latest_unknown(latest_text):
            return None
        if app.name in {"Office 2024 LTSC", "Office 365 Ent"}:
            if not local_versions.version:
                return None
            ok = self._office_versions_ok(local_versions.version, latest_text)
            return "Latest" if ok else "Outdated"
        if app.dual_arch:
            if not local_versions.version_x86:
                return None
            if _is_64bit() and not local_versions.version_x64:
                return None
            latest_x86, latest_x64 = _parse_dual_text(latest_text)
            if not latest_x86 or latest_x86 in _LATEST_UNKNOWN:
                return None
            if _is_64bit() and (not latest_x64 or latest_x64 in _LATEST_UNKNOWN):
                return None
            ok86 = _version_ge(local_versions.version_x86, latest_x86)
            ok64 = True
            if _is_64bit():
                ok64 = _version_ge(local_versions.version_x64 or "", latest_x64)
            return "Latest" if ok86 and ok64 else "Outdated"
        if not local_versions.version:
            return None
        return "Latest" if _version_ge(local_versions.version, latest_text) else "Outdated"

    def _read_uninstall_entries(self) -> list[_UninstallEntry]:
        if winreg is None:
            return []
        base_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        wow_path = r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        views = [
            (getattr(winreg, "KEY_WOW64_64KEY", 0), base_path),
            (getattr(winreg, "KEY_WOW64_64KEY", 0), wow_path),
            (getattr(winreg, "KEY_WOW64_32KEY", 0), base_path),
        ]
        entries: dict[tuple[str, str], _UninstallEntry] = {}
        for view, path in views:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ | view) as root:  # type: ignore[arg-type]
                    index = 0
                    while True:
                        try:
                            sub_name = winreg.EnumKey(root, index)
                        except OSError:
                            break
                        index += 1
                        try:
                            with winreg.OpenKey(root, sub_name) as subkey:
                                name = self._reg_value(subkey, "DisplayName")
                                if not name:
                                    continue
                                version = self._reg_value(subkey, "DisplayVersion")
                                entry = _UninstallEntry(display_name=name, display_version=version or "")
                                entries[(entry.display_name, entry.display_version)] = entry
                        except OSError:
                            continue
            except OSError:
                continue
        return list(entries.values())

    def _reg_value(self, key, value_name: str) -> str | None:
        try:
            value, _ = winreg.QueryValueEx(key, value_name)
        except OSError:
            return None
        if value is None:
            return None
        return str(value)

    def _get_best_version(self, pattern: str, entries: Iterable[_UninstallEntry]) -> str | None:
        regex = re.compile(pattern, re.IGNORECASE)
        best_value: tuple[tuple[int, ...], str] | None = None
        fallback: str | None = None
        for entry in entries:
            if not regex.search(entry.display_name):
                continue
            if not entry.display_version:
                continue
            normalized = _normalize_version(entry.display_version)
            if normalized:
                version_key = _version_tuple(normalized)
                if version_key:
                    if best_value is None or version_key > best_value[0]:
                        best_value = (version_key, normalized)
                if fallback is None:
                    fallback = normalized
            elif fallback is None:
                fallback = entry.display_version
        if best_value:
            return best_value[1]
        return fallback

    def _get_hp_support_appx_version(self) -> str | None:
        if sys.platform != "win32":
            return None
        powershell = shutil.which("powershell") or shutil.which("pwsh")
        if not powershell:
            return None
        script = (
            "$pkg = $null; "
            "try { $pkg = Get-AppxPackage -AllUsers -ErrorAction Stop | "
            "Where-Object { $_.Name -match 'HPSupportAssistant|HP.*SupportAssistant|HPInc.*Support' } | "
            "Sort-Object Version -Descending | Select-Object -First 1 } catch { "
            "try { $pkg = Get-AppxPackage -ErrorAction Stop | "
            "Where-Object { $_.Name -match 'HPSupportAssistant|HP.*SupportAssistant|HPInc.*Support' } | "
            "Sort-Object Version -Descending | Select-Object -First 1 } catch {} } "
            "if ($pkg) { $pkg.Version.ToString() }"
        )
        try:
            completed = subprocess.run(
                [powershell, "-NoProfile", "-Command", script],
                capture_output=True,
                text=True,
                check=False,
            )
        except OSError:
            return None
        if completed.returncode != 0:
            return None
        output = completed.stdout.strip()
        if not output:
            return None
        return _normalize_version(output) or output

    def _get_vc_installed_map(self, vc_key: str, entries: Iterable[_UninstallEntry]) -> dict[str, str | None]:
        patterns = {
            "2005": r"^Microsoft Visual C\+\+ 2005 Redistributable",
            "2008": r"^Microsoft Visual C\+\+ 2008 Redistributable",
            "2010": r"^Microsoft Visual C\+\+ 2010",
            "2012": r"^Microsoft Visual C\+\+ 2012 Redistributable",
            "2013": r"^Microsoft Visual C\+\+ 2013 Redistributable",
            "2015+": r"^Microsoft Visual C\+\+.*((2015-2022)|(2015|2017|2019|2022)|v14).*(Redistributable)",
        }
        pattern = patterns.get(vc_key)
        if not pattern:
            return {"x86": None, "x64": None}
        regex = re.compile(pattern, re.IGNORECASE)
        best86: tuple[tuple[int, ...], str] | None = None
        best64: tuple[tuple[int, ...], str] | None = None
        for entry in entries:
            if not regex.search(entry.display_name):
                continue
            normalized = _normalize_version(entry.display_version)
            if not normalized:
                continue
            version_key = _version_tuple(normalized)
            if not version_key:
                continue
            name = entry.display_name
            arch = None
            if re.search(r"\b(x64|64-bit|amd64)\b", name, re.IGNORECASE):
                arch = "x64"
            elif re.search(r"\b(x86|32-bit|x32)\b", name, re.IGNORECASE):
                arch = "x86"
            elif vc_key == "2005":
                arch = "x86"
            if arch == "x86":
                if best86 is None or version_key > best86[0]:
                    best86 = (version_key, normalized)
            elif arch == "x64":
                if best64 is None or version_key > best64[0]:
                    best64 = (version_key, normalized)
        return {
            "x86": best86[1] if best86 else None,
            "x64": best64[1] if best64 else None,
        }

    def _format_vc_versions(self, vc_map: Mapping[str, str | None]) -> str:
        x86 = vc_map.get("x86") or "Missing"
        x64 = vc_map.get("x64") or "Missing"
        return f"x86: {x86} | x64: {x64}"

    def _get_local_odt_version(self) -> str | None:
        settings_path = self._settings.odt_setup_path.strip()
        if settings_path:
            candidate = Path(settings_path)
            version = _get_file_version(candidate) if candidate.suffix.lower() == ".exe" else None
            if version:
                return _normalize_version(version) or version
        version_dir = self._working_dir / "odt_versions"
        if not version_dir.exists():
            version_dir = None
        best_value: tuple[tuple[int, ...], str] | None = None
        fallback: str | None = None
        if version_dir:
            for path in version_dir.glob("*.txt"):
                try:
                    value = path.read_text(encoding="utf-8").strip()
                except OSError:
                    continue
                if not value:
                    continue
                normalized = _normalize_version(value)
                if normalized:
                    version_key = _version_tuple(normalized)
                    if version_key:
                        if best_value is None or version_key > best_value[0]:
                            best_value = (version_key, normalized)
                        continue
                if fallback is None:
                    fallback = value
        if not best_value and not fallback:
            legacy_root = self._working_dir / "Office"
            legacy_paths = [
                legacy_root / "office_2024_ltsc" / "setup.version.txt",
                legacy_root / "office_365_ent" / "setup.version.txt",
            ]
            for path in legacy_paths:
                if not path.exists():
                    continue
                try:
                    value = path.read_text(encoding="utf-8").strip()
                except OSError:
                    continue
                if value:
                    return _normalize_version(value) or value
        if best_value:
            return best_value[1]
        if fallback:
            return fallback
        for candidate in self._odt_setup_candidates():
            version = _get_file_version(candidate)
            if version:
                return _normalize_version(version) or version
        return None

    def _odt_setup_candidates(self) -> Iterable[Path]:
        candidates = [
            self._working_dir / "setup.exe",
            self._working_dir / "OfficeSetup" / "setup.exe",
        ]
        for name in ("office_2024_ltsc", "office_365_ent"):
            candidates.append(self._working_dir / name / "setup.exe")
            candidates.append(self._working_dir / "Office" / name / "setup.exe")
        for candidate in candidates:
            if candidate.exists() and candidate.is_file():
                yield candidate

    def _get_latest_version(self, app: AppEntry) -> str:
        if app.name == "Office 2024 LTSC":
            return self._get_office_latest("https://learn.microsoft.com/en-us/officeupdates/update-history-office-2024")
        if app.name == "Office 365 Ent":
            return self._get_office_latest(
                "https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date"
            )
        if app.name == "CrowdStrike Falcon Sensor":
            return self._get_crowdstrike_latest(app)
        if app.name == "FortiClient VPN":
            return self._get_forticlient_latest(app)
        if app.name == "HP Support Asst":
            if app.name in self._direct_downloaders:
                return self._get_direct_latest(app.name)
            if app.winget_id:
                winget_version = self._get_winget_latest(app.winget_id, app.source)
                if not self._latest_unknown(winget_version) and winget_version != "N/A":
                    return winget_version
            return self._get_hp_support_latest()
        if app.name in self._direct_downloaders:
            return self._get_direct_latest(app.name)
        if app.dual_arch:
            return self._get_dual_winget_latest(app)
        if app.winget_id:
            return self._get_winget_latest(app.winget_id, app.source)
        return "N/A"

    def _get_crowdstrike_latest(self, app: AppEntry) -> str:
        share_url = self._settings.crowdstrike_download_url.strip()
        if not share_url:
            return "N/A"
        if not _is_sharepoint_url(share_url):
            return "Error"
        stem = app.file_stem or _safe_name(app.name)
        filename = f"{stem}.exe"
        destination_dir = self._working_dir / "downloads" / _safe_name(app.name)
        try:
            dest_path = _download_sharepoint_exe(share_url, destination_dir, filename, force=True)
        except Exception:
            return "Error"
        version = _get_file_version(dest_path)
        if not version:
            return "Unknown"
        return _normalize_version(version) or version

    def _get_forticlient_latest(self, app: AppEntry) -> str:
        share_url = self._settings.forticlient_download_url.strip()
        if not share_url:
            return "N/A"
        if not _is_sharepoint_url(share_url):
            return "Error"
        stem = app.file_stem or _safe_name(app.name)
        filename = f"{stem}.exe"
        destination_dir = self._working_dir / "downloads" / _safe_name(app.name)
        try:
            dest_path = _download_sharepoint_exe(share_url, destination_dir, filename, force=True)
        except Exception:
            return "Error"
        version = _get_file_version(dest_path)
        if not version:
            return "Unknown"
        return _normalize_version(version) or version

    def _get_winget_latest(self, package_id: str, source: str | None) -> str:
        if not self._winget.is_available():
            return "Winget Missing"
        source = source or "winget"
        try:
            version = self._winget.show_package_version(package_id, source=source)
        except WingetError:
            return "Winget Missing"
        except Exception:
            return "Error"
        if not version:
            return "N/A"
        normalized = _normalize_version(version) or version
        return normalized

    def _get_dual_winget_latest(self, app: AppEntry) -> str:
        x86 = self._get_winget_latest(app.winget_id_x86 or "", app.source) if app.winget_id_x86 else "N/A"
        x64 = self._get_winget_latest(app.winget_id_x64 or "", app.source) if app.winget_id_x64 else "N/A"
        if x86 == x64 and x86 not in _LATEST_UNKNOWN:
            return x86
        if x86 in _LATEST_UNKNOWN and x64 in _LATEST_UNKNOWN:
            if x86 == x64:
                return x86
            return "N/A"
        return f"x86: {x86} | x64: {x64}"

    def _get_direct_latest(self, app_name: str) -> str:
        downloader = self._direct_downloaders.get(app_name)
        if not downloader:
            return "N/A"
        try:
            info = downloader.fetch()
        except Exception:
            return "Error"
        if not info or not info.version:
            return "N/A"
        return info.version

    def _get_hp_support_latest(self) -> str:
        url = "https://hpsa-redirectors.hpcloud.hp.com/common/hpsaredirector.js"
        content = self._fetch_text(url)
        if not content:
            return "N/A"
        match = re.search(r'(?m)^\s*return\s+"([0-9.]+)"', content)
        if not match:
            return "N/A"
        return _normalize_version(match.group(1)) or match.group(1)

    def _get_office_latest(self, url: str) -> str:
        content = self._fetch_text(url)
        if not content:
            return "N/A"
        match = re.search(r"Version\s+(\d+)\s+\(Build\s+([\d.]+)\)", content)
        if not match:
            return "N/A"
        version = match.group(1)
        build = match.group(2)
        return f"{version} (Build {build})"

    def _fetch_text(self, url: str) -> str | None:
        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120",
                "Accept-Language": "en-US,en;q=0.9",
            },
        )
        try:
            with urllib.request.urlopen(request, timeout=20) as response:
                return response.read().decode("utf-8", errors="ignore")
        except Exception:
            return None

    def _evaluate_status(self, app: AppEntry, info: InstalledInfo, latest_text: str) -> tuple[str, str]:
        if not info.is_known:
            return _STATUS_UNKNOWN, _LEVEL_UNKNOWN
        if not info.is_installed:
            return _STATUS_NOT_INSTALLED, _LEVEL_NOT_INSTALLED
        if self._latest_unknown(latest_text):
            return _STATUS_INSTALLED, _LEVEL_INSTALLED
        if app.dual_arch:
            ok = self._vc_versions_ok(info, latest_text)
            if ok:
                return _STATUS_UP_TO_DATE, _LEVEL_UP_TO_DATE
            return _STATUS_UPDATE_AVAILABLE, _LEVEL_UPDATE_AVAILABLE
        if app.name in {"Office 2024 LTSC", "Office 365 Ent"}:
            ok = self._office_versions_ok(info.installed_text, latest_text)
            if ok:
                return _STATUS_UP_TO_DATE, _LEVEL_UP_TO_DATE
            return _STATUS_UPDATE_AVAILABLE, _LEVEL_UPDATE_AVAILABLE
        ok = _version_ge(info.installed_text, latest_text)
        if ok:
            return _STATUS_UP_TO_DATE, _LEVEL_UP_TO_DATE
        return _STATUS_UPDATE_AVAILABLE, _LEVEL_UPDATE_AVAILABLE

    def _latest_unknown(self, latest_text: str) -> bool:
        if latest_text in _LATEST_UNKNOWN:
            return True
        if "x86:" in latest_text and "x64:" in latest_text:
            x86, x64 = _parse_dual_text(latest_text)
            if (x86 in _LATEST_UNKNOWN) and (x64 in _LATEST_UNKNOWN):
                return True
        return False

    def _vc_versions_ok(self, info: InstalledInfo, latest_text: str) -> bool:
        latest_x86, latest_x64 = _parse_dual_text(latest_text)
        if not latest_x86 and not latest_x64:
            return False
        ok86 = False
        ok64 = False
        if info.installed_x86 and latest_x86 and latest_x86 not in _LATEST_UNKNOWN:
            ok86 = _version_ge(info.installed_x86, latest_x86)
        if _is_64bit():
            if info.installed_x64 and latest_x64 and latest_x64 not in _LATEST_UNKNOWN:
                ok64 = _version_ge(info.installed_x64, latest_x64)
        else:
            ok64 = True
        return ok86 and ok64

    def _office_versions_ok(self, installed_text: str, latest_text: str) -> bool:
        installed_build = _office_installed_build(installed_text)
        latest_build = _office_latest_build(latest_text)
        if installed_build and latest_build:
            return _version_ge(installed_build, latest_build)
        return installed_text == latest_text


def _normalize_version(value: str | None) -> str | None:
    if not value or not value.strip():
        return None
    cleaned = value.strip()
    if re.match(r"^\d+\.\d+$", cleaned):
        return f"{cleaned}.0.0"
    if re.match(r"^\d+\.\d+\.\d+$", cleaned):
        return f"{cleaned}.0"
    if re.match(r"^\d+\.\d+\.\d+\.\d+$", cleaned):
        return cleaned
    match = re.search(r"(\d+\.\d+(?:\.\d+){0,2})", cleaned)
    if match:
        return _normalize_version(match.group(1))
    return cleaned


def _version_tuple(value: str | None) -> tuple[int, ...] | None:
    if not value:
        return None
    parts = value.split(".")
    if not all(part.isdigit() for part in parts):
        return None
    return tuple(int(part) for part in parts)


def _version_ge(installed: str, latest: str) -> bool:
    inst_norm = _normalize_version(installed)
    latest_norm = _normalize_version(latest)
    inst_tuple = _version_tuple(inst_norm)
    latest_tuple = _version_tuple(latest_norm)
    if inst_tuple and latest_tuple:
        return inst_tuple >= latest_tuple
    return installed.strip() == latest.strip()


def _parse_dual_text(text: str) -> tuple[str | None, str | None]:
    match = re.search(r"x86:\s*([0-9.]+|Missing|N/A|Winget Missing|Error)\s*\|\s*x64:\s*([0-9.]+|Missing|N/A|Winget Missing|Error)", text)
    if match:
        return match.group(1), match.group(2)
    if text and text not in _LATEST_UNKNOWN:
        return text, text
    return None, None


def _safe_name(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9_-]+", "_", name).lower()


def _office_installed_build(installed_text: str) -> str | None:
    match = re.search(r"^16\.0\.(.+)$", installed_text)
    if match:
        return match.group(1)
    return None


def _office_latest_build(latest_text: str) -> str | None:
    match = re.search(r"\(Build\s+([\d.]+)\)", latest_text)
    if match:
        return match.group(1)
    match = re.search(r"Build\s+([\d.]+)", latest_text)
    if match:
        return match.group(1)
    return None


def _is_64bit() -> bool:
    return sys.maxsize > 2**32


def _get_file_version(path: Path) -> str | None:
    if sys.platform != "win32":
        return None
    if not path.exists() or not path.is_file():
        return None
    try:
        import ctypes
        from ctypes import wintypes
    except Exception:
        return None

    class VS_FIXEDFILEINFO(ctypes.Structure):
        _fields_ = [
            ("dwSignature", wintypes.DWORD),
            ("dwStrucVersion", wintypes.DWORD),
            ("dwFileVersionMS", wintypes.DWORD),
            ("dwFileVersionLS", wintypes.DWORD),
            ("dwProductVersionMS", wintypes.DWORD),
            ("dwProductVersionLS", wintypes.DWORD),
            ("dwFileFlagsMask", wintypes.DWORD),
            ("dwFileFlags", wintypes.DWORD),
            ("dwFileOS", wintypes.DWORD),
            ("dwFileType", wintypes.DWORD),
            ("dwFileSubtype", wintypes.DWORD),
            ("dwFileDateMS", wintypes.DWORD),
            ("dwFileDateLS", wintypes.DWORD),
        ]

    size = ctypes.windll.version.GetFileVersionInfoSizeW(str(path), None)
    if not size:
        return None
    data = ctypes.create_string_buffer(size)
    if not ctypes.windll.version.GetFileVersionInfoW(str(path), 0, size, data):
        return None
    value = ctypes.c_void_p()
    length = wintypes.UINT()
    if not ctypes.windll.version.VerQueryValueW(data, "\\", ctypes.byref(value), ctypes.byref(length)):
        return None
    if length.value < ctypes.sizeof(VS_FIXEDFILEINFO):
        return None
    fixed = ctypes.cast(value.value, ctypes.POINTER(VS_FIXEDFILEINFO)).contents
    if fixed.dwSignature != 0xFEEF04BD:
        return None
    major = fixed.dwFileVersionMS >> 16
    minor = fixed.dwFileVersionMS & 0xFFFF
    build = fixed.dwFileVersionLS >> 16
    revision = fixed.dwFileVersionLS & 0xFFFF
    return f"{major}.{minor}.{build}.{revision}"
