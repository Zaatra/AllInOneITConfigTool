"""Application installation orchestration."""
from __future__ import annotations

import http.cookiejar
import os
import re
import shlex
import shutil
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Mapping, Protocol, Sequence

from allinone_it_config.app_registry import AppEntry
from allinone_it_config.paths import get_application_directory
from allinone_it_config.user_settings import UserSettings


@dataclass
class CommandExecutionResult:
    command: Sequence[str]
    returncode: int
    stdout: str
    stderr: str

    @property
    def succeeded(self) -> bool:
        return self.returncode == 0


class WingetError(RuntimeError):
    pass


class WingetClient:
    """Thin wrapper around the winget CLI."""

    VERSION_PATTERN = re.compile(r"Version\s*:\s*(.+)", re.IGNORECASE)

    def __init__(self, executable: str | None = None):
        exe_path = executable or shutil.which("winget")
        if not exe_path:
            fallback = self._find_winget_fallback()
            exe_path = str(fallback) if fallback else None
        self._executable = Path(exe_path) if exe_path else None

    def is_available(self) -> bool:
        return self._executable is not None

    def install_package(
        self,
        package_id: str,
        *,
        source: str | None = None,
        override: str | None = None,
        version: str | None = None,
        silent: bool = True,
        force: bool = True,
    ) -> CommandExecutionResult:
        cmd = self._build_base_command("install", package_id, source, force)
        if version:
            cmd.extend(["--version", version])
        if silent:
            cmd.append("--silent")
        if override:
            cmd.extend(["--override", override])
        return self._run(cmd)

    def download_package(
        self,
        package_id: str,
        destination: Path,
        *,
        source: str | None = None,
        version: str | None = None,
        locale: str | None = None,
        force: bool = True,
    ) -> CommandExecutionResult:
        destination.mkdir(parents=True, exist_ok=True)
        cmd = self._build_base_command("download", package_id, source, force)
        if version:
            cmd.extend(["--version", version])
        if locale:
            cmd.extend(["--locale", locale])
        cmd.extend(["--disable-interactivity", "--verbose", "-d", str(destination)])
        return self._run(cmd)

    def show_package_version(self, package_id: str, *, source: str | None = None, locale: str | None = None) -> str | None:
        if not self._executable:
            raise WingetError("winget executable not found in PATH")
        cmd = [str(self._executable), "show", "--id", package_id, "--exact", "--accept-source-agreements"]
        if locale:
            cmd.extend(["--locale", locale])
        if source:
            cmd.extend(["--source", source])
        result = self._run(cmd)
        if result.returncode != 0:
            return None
        for line in result.stdout.splitlines():
            match = self.VERSION_PATTERN.search(line)
            if match:
                return match.group(1).strip()
        return None

    def update_sources(self, name: str | None = None) -> CommandExecutionResult | None:
        if not self._executable:
            return None
        cmd = [str(self._executable), "source", "update"]
        if name:
            cmd.extend(["--name", name])
        return self._run(cmd)

    def _build_base_command(
        self,
        verb: str,
        package_id: str,
        source: str | None,
        force: bool,
    ) -> list[str]:
        if not self._executable:
            raise WingetError("winget executable not found in PATH")
        cmd = [str(self._executable), verb, "--id", package_id, "--exact"]
        if force and verb != "download":
            cmd.append("--force")
        cmd.extend(["--accept-package-agreements", "--accept-source-agreements"])
        if source:
            cmd.extend(["--source", source])
        return cmd

    def _run(self, cmd: list[str]) -> CommandExecutionResult:
        completed = subprocess.run(cmd, capture_output=True, text=True, check=False)
        return CommandExecutionResult(cmd, completed.returncode, completed.stdout, completed.stderr)

    def _find_winget_fallback(self) -> Path | None:
        local_appdata = os.environ.get("LOCALAPPDATA")
        if local_appdata:
            candidate = Path(local_appdata) / "Microsoft" / "WindowsApps" / "winget.exe"
            if candidate.exists():
                return candidate
        program_files = os.environ.get("ProgramFiles")
        if program_files:
            base = Path(program_files) / "WindowsApps"
            try:
                candidates = base.glob("Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe/winget.exe")
            except OSError:
                candidates = []
            for candidate in candidates:
                if candidate.exists():
                    return candidate
        return None


class OfficeInstaller:
    """Handles Office Deployment Tool download and execution."""

    def __init__(
        self,
        working_dir: Path,
        *,
        winget_client: WingetClient | None = None,
        template_loader: Callable[[str], str] | None = None,
        odt_setup_path: str | None = None,
        office_root: Path | None = None,
    ) -> None:
        self._working_dir = Path(working_dir)
        self._winget = winget_client or WingetClient()
        self._template_loader = template_loader
        self._office_root = Path(office_root) if office_root else self._working_dir
        self._odt_version_dir = self._working_dir / "odt_versions"
        self._odt_setup_path = Path(odt_setup_path) if odt_setup_path else None

    def office_dir(self, app_name: str) -> Path:
        return self._office_root / _safe_name(app_name)

    def setup_path(self, app_name: str) -> Path:
        return self.office_dir(app_name) / "setup.exe"

    def config_path(self, app_name: str) -> Path:
        return self.office_dir(app_name) / "config.xml"

    def _version_path(self, app_name: str) -> Path:
        return self._odt_version_dir / f"{_safe_name(app_name)}.txt"

    def _legacy_office_dir(self, app_name: str) -> Path:
        return self._working_dir / "Office" / _safe_name(app_name)

    def _legacy_office_dirs(self, app_name: str) -> tuple[Path, ...]:
        safe_name = _safe_name(app_name)
        return (
            self._working_dir / safe_name,
            self._legacy_office_dir(app_name),
        )

    def odt_override_path(self) -> Path | None:
        if not self._odt_setup_path:
            return None
        if self._odt_setup_path.suffix.lower() != ".exe":
            return None
        if self._odt_setup_path.exists() and self._odt_setup_path.is_file():
            return self._odt_setup_path
        return None

    def setup_executable(self, app_name: str) -> Path:
        override = self.odt_override_path()
        if override:
            return override
        return self.setup_path(app_name)

    def has_payload(self, app_name: str) -> bool:
        if _office_payload_present(self.office_dir(app_name)):
            return True
        for legacy_dir in self._legacy_office_dirs(app_name):
            if _office_payload_present(legacy_dir):
                return True
        return False

    def payload_version(self, app_name: str) -> str | None:
        versions: list[str] = []
        for office_dir in (self.office_dir(app_name), *self._legacy_office_dirs(app_name)):
            versions.extend(_office_payload_versions(office_dir))
        if not versions:
            return None
        return max(versions, key=_version_tuple)

    def _clean_office_dir(self, office_dir: Path) -> None:
        if not office_dir.exists():
            return
        keep = {"setup.exe", "config.xml"}
        for entry in office_dir.iterdir():
            if entry.name.lower() in keep:
                continue
            if entry.is_dir():
                shutil.rmtree(entry, ignore_errors=True)
            else:
                try:
                    entry.unlink()
                except OSError:
                    continue

    def _recover_setup_exe(self, app_name: str, office_dir: Path, setup_path: Path) -> bool:
        try:
            candidates = sorted(
                {path for path in office_dir.rglob("setup.exe")},
                key=lambda path: (len(path.parts), str(path).lower()),
            )
        except OSError:
            candidates = []
        for candidate in candidates:
            if candidate == setup_path:
                return True
            try:
                candidate.replace(setup_path)
                return True
            except OSError:
                continue
        for legacy_dir in self._legacy_office_dirs(app_name):
            legacy_setup = legacy_dir / "setup.exe"
            if not legacy_setup.exists():
                continue
            try:
                shutil.copy2(legacy_setup, setup_path)
                return True
            except OSError:
                return False
        return False

    def ensure_setup(self, app_name: str) -> CommandExecutionResult | None:
        if self.odt_override_path():
            return None
        if not self._winget.is_available():
            raise WingetError("winget unavailable to fetch Office Deployment Tool")
        office_dir = self.office_dir(app_name)
        office_dir.mkdir(parents=True, exist_ok=True)
        setup_path = self.setup_path(app_name)
        version_path = self._version_path(app_name)
        latest_version = None
        try:
            latest_version = self._winget.show_package_version("Microsoft.OfficeDeploymentTool")
        except WingetError:
            latest_version = None
        latest_version = latest_version.strip() if latest_version else ""
        if setup_path.exists() and latest_version and version_path.exists():
            current_version = version_path.read_text(encoding="utf-8").strip()
            if current_version == latest_version:
                return None
        override = f'/quiet /extract:"{office_dir}"'
        result = self._winget.install_package(
            "Microsoft.OfficeDeploymentTool",
            override=override,
        )
        if not setup_path.exists():
            self._recover_setup_exe(app_name, office_dir, setup_path)
        if not setup_path.exists():
            raise FileNotFoundError("setup.exe missing after Office Deployment Tool extraction")
        version_path.parent.mkdir(parents=True, exist_ok=True)
        version_path.write_text(latest_version or "unknown", encoding="utf-8")
        return result

    def install(self, app_name: str) -> CommandExecutionResult:
        if not self._template_loader:
            raise ValueError("Office template loader not configured")
        template_xml = self._template_loader(app_name)
        if not template_xml.strip():
            raise ValueError(f"Office XML template empty for {app_name}")
        office_dir = self.office_dir(app_name)
        office_dir.mkdir(parents=True, exist_ok=True)
        config_path = self.config_path(app_name)
        config_path.write_text(template_xml, encoding="utf-8")
        self.ensure_setup(app_name)
        cmd = [str(self.setup_executable(app_name)), "/configure", str(config_path)]
        completed = subprocess.run(cmd, capture_output=True, text=True, check=False, cwd=office_dir)
        return CommandExecutionResult(cmd, completed.returncode, completed.stdout, completed.stderr)

    def download(self, app_name: str, *, status_callback: Callable[[str], None] | None = None) -> CommandExecutionResult:
        if not self._template_loader:
            raise ValueError("Office template loader not configured")
        template_xml = self._template_loader(app_name)
        if not template_xml.strip():
            raise ValueError(f"Office XML template empty for {app_name}")
        office_dir = self.office_dir(app_name)
        office_dir.mkdir(parents=True, exist_ok=True)
        config_path = self.config_path(app_name)
        config_path.write_text(template_xml, encoding="utf-8")
        self.ensure_setup(app_name)
        self._clean_office_dir(office_dir)
        cmd = [str(self.setup_executable(app_name)), "/download", str(config_path)]
        completed = _run_office_download(cmd, office_dir, label=app_name, status_callback=status_callback)
        return completed


@dataclass(frozen=True)
class DirectDownloadInfo:
    version: str
    url: str
    filename: str | None = None


@dataclass(frozen=True)
class LocalInstallerInfo:
    exists: bool
    path: Path | None = None
    path_x86: Path | None = None
    path_x64: Path | None = None


@dataclass(frozen=True)
class LocalInstallerVersionInfo:
    version: str | None = None
    version_x86: str | None = None
    version_x64: str | None = None
    path: Path | None = None
    path_x86: Path | None = None
    path_x64: Path | None = None

    def has_any(self) -> bool:
        return bool(
            self.version
            or self.version_x86
            or self.version_x64
            or self.path
            or self.path_x86
            or self.path_x64
        )


class DirectDownloader(Protocol):
    def fetch(self) -> DirectDownloadInfo:
        ...


@dataclass(frozen=True)
class LicenseApplyResult:
    success: bool
    message: str


class ConfiguredUrlDownloader:
    def __init__(self, url: str, *, default_filename: str, version: str | None = None) -> None:
        self._url = url
        self._default_filename = default_filename
        self._version = version or "custom"

    def fetch(self) -> DirectDownloadInfo:
        url = self._url.strip()
        if not url:
            raise RuntimeError("Download URL not configured")
        filename = _filename_from_url(url) or self._default_filename
        return DirectDownloadInfo(version=self._version, url=url, filename=filename)


class IVMSDownloader:
    """Scrapes Hikvision's site for the latest iVMS-4200 build."""

    SERIES_URL = "https://www.hikvision.com/en/support/download/software/ivms4200-series/"
    FALLBACK_URL = "https://www.hikvision.com/us-en/support/download/software/"
    VERSION_PATTERNS = (
        re.compile(r"iVMS-4200V(\d+\.\d+\.\d+\.\d+)_E", re.IGNORECASE),
        re.compile(r"iVMS-4200\s+V(\d+\.\d+\.\d+\.\d+)", re.IGNORECASE),
    )
    LINK_PATTERN = re.compile(
        r"(https?://[^\s\"']+iVMS-4200[^\s\"']+?\.exe|/content/dam/[^\s\"']+iVMS-4200[^\s\"']+?\.exe)",
        re.IGNORECASE,
    )

    def fetch(self) -> DirectDownloadInfo:
        html = ""
        try:
            html = self._fetch_html(self.SERIES_URL)
        except Exception:
            html = ""
        links = self._extract_links(html)
        if not links:
            try:
                html = self._fetch_html(self.FALLBACK_URL)
                links = self._extract_links(html)
            except Exception:
                links = []
        if links:
            best_url, version = self._pick_latest_link(links)
            if best_url and version:
                return DirectDownloadInfo(version=version, url=best_url)
        versions = self._extract_versions(html)
        if not versions:
            raise RuntimeError("Unable to determine latest iVMS-4200 version")
        version = max(versions, key=_version_tuple)
        url = self._build_url(version)
        return DirectDownloadInfo(version=version, url=url)

    def _fetch_html(self, url: str) -> str:
        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120",
                "Accept-Language": "en-US,en;q=0.9",
            },
        )
        with urllib.request.urlopen(request, timeout=20) as response:
            return response.read().decode("utf-8", errors="ignore")

    def _build_url(self, version: str) -> str:
        dash = "v" + version.replace(".", "-")
        return (
            "https://www.hikvision.com/content/dam/hikvision/en/support/download/vms/"
            "ivms4200-series/software-download/{dash}/iVMS-4200V{version}_E.exe".format(dash=dash, version=version)
        )

    def _extract_links(self, html: str) -> list[str]:
        links = []
        for match in self.LINK_PATTERN.finditer(html):
            link = match.group(1)
            if not link:
                continue
            if link.startswith("/"):
                link = "https://www.hikvision.com" + link
            links.append(link)
        return list(dict.fromkeys(links))

    def _pick_latest_link(self, links: list[str]) -> tuple[str | None, str | None]:
        best_url = None
        best_version = None
        for link in links:
            match = re.search(r"iVMS-4200V(\d+(?:\.\d+){1,3})_E\.exe", link, re.IGNORECASE)
            if not match:
                continue
            version = match.group(1)
            if not best_version or _version_tuple(version) > _version_tuple(best_version):
                best_version = version
                best_url = link
        return best_url, best_version

    def _extract_versions(self, html: str) -> set[str]:
        found: set[str] = set()
        for pattern in self.VERSION_PATTERNS:
            for match in pattern.finditer(html):
                ver = match.group(1)
                if ver:
                    found.add(ver)
        return found


class HPSADownloader:
    """Parses HP's redirector JS to find the active Support Assistant installer."""

    SOURCE_URL = "https://hpsa-redirectors.hpcloud.hp.com/common/hpsaredirector.js"
    EXE_PATTERN = re.compile(r'return\s+.*?"(ftp\.hp\.com.*?\.exe)"', re.IGNORECASE)
    VERSION_PATTERN = re.compile(r"//\s*([0-9.]+)")

    def fetch(self) -> DirectDownloadInfo:
        request = urllib.request.Request(self.SOURCE_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(request, timeout=20) as response:
            content = response.read().decode("utf-8", errors="ignore")
        for line in content.splitlines():
            trimmed = line.strip()
            if not trimmed or trimmed.startswith("//"):
                continue
            match = self.EXE_PATTERN.search(trimmed)
            if not match:
                continue
            relative = match.group(1)
            url = "https://" + relative
            version = "Unknown"
            version_match = self.VERSION_PATTERN.search(trimmed)
            if version_match:
                version = version_match.group(1)
            filename = _filename_from_url(url) or "hp_support_assistant.exe"
            return DirectDownloadInfo(version=version, url=url, filename=filename)
        raise RuntimeError("Active HP Support Assistant link not found")


@dataclass
class OperationResult:
    app: AppEntry
    operation: str
    success: bool
    message: str
    stdout: str = ""
    stderr: str = ""


class InstallerService:
    def __init__(
        self,
        apps: Iterable[AppEntry],
        *,
        working_dir: Path | str | None = None,
        winget_client: WingetClient | None = None,
        office_installer: OfficeInstaller | None = None,
        direct_downloaders: Mapping[str, DirectDownloader] | None = None,
        settings: UserSettings | None = None,
    ) -> None:
        self._apps = list(apps)
        self._working_dir = Path(working_dir) if working_dir is not None else get_application_directory()
        self._downloads_dir = self._working_dir / "downloads"
        self._winget = winget_client or WingetClient()
        self._settings = settings or UserSettings()
        self._office = office_installer or OfficeInstaller(
            self._working_dir,
            winget_client=self._winget,
            template_loader=self._settings.load_office_xml,
            odt_setup_path=self._settings.odt_setup_path,
            office_root=self._downloads_dir,
        )
        default_direct = {"iVMS-4200": IVMSDownloader(), "HP Support Asst": HPSADownloader()}
        self._direct_downloaders = dict(default_direct)
        if self._settings.crowdstrike_download_url.strip():
            self._direct_downloaders["CrowdStrike Falcon Sensor"] = ConfiguredUrlDownloader(
                self._settings.crowdstrike_download_url.strip(),
                default_filename="crowdstrike_falcon_sensor.exe",
            )
        if self._settings.forticlient_download_url.strip():
            self._direct_downloaders["FortiClient VPN"] = ConfiguredUrlDownloader(
                self._settings.forticlient_download_url.strip(),
                default_filename="forticlient_vpn.exe",
            )
        if direct_downloaders:
            self._direct_downloaders.update(direct_downloaders)

    def is_downloadable(self, app: AppEntry) -> bool:
        if app.download_mode in {"localonly", "onlineonly"}:
            return False
        if app.download_mode == "direct":
            return app.name in self._direct_downloaders
        if app.download_mode == "winget" and app.source and app.source != "winget":
            return False
        return True

    def get_local_installer_info(self, app: AppEntry, *, include_downloads: bool = True) -> LocalInstallerInfo:
        if app.installer_path:
            candidate = Path(app.installer_path)
            if candidate.exists() and candidate.is_file():
                return LocalInstallerInfo(True, path=candidate)
            return LocalInstallerInfo(False)
        if app.download_mode == "onlineonly":
            return LocalInstallerInfo(False)
        if app.download_mode == "office":
            setup_path = self._office.setup_path(app.name)
            if setup_path.exists():
                return LocalInstallerInfo(True, path=setup_path)
            if self._office.has_payload(app.name):
                override = self._office.odt_override_path()
                return LocalInstallerInfo(True, path=override if override else None)
            return LocalInstallerInfo(False)
        search_dirs = [self._working_dir]
        if include_downloads:
            search_dirs.append(self._downloads_dir)
            search_dirs.append(self._downloads_dir / _safe_name(app.name))
        if app.dual_arch:
            path_x86 = None
            path_x64 = None
            if app.file_stem_x86:
                patterns = [f"{app.file_stem_x86}_*.exe", f"{app.file_stem_x86}_*.msi"]
                path_x86 = self._best_local_by_patterns(search_dirs, patterns)
            if app.file_stem_x64:
                patterns = [f"{app.file_stem_x64}_*.exe", f"{app.file_stem_x64}_*.msi"]
                path_x64 = self._best_local_by_patterns(search_dirs, patterns)
            return LocalInstallerInfo(bool(path_x86 or path_x64), path_x86=path_x86, path_x64=path_x64)
        patterns: list[str] = []
        if app.file_stem:
            patterns.extend([f"{app.file_stem}_*.exe", f"{app.file_stem}_*.msi"])
        exact_names = tuple(app.local_alt_names) if app.download_mode in {"localonly", "direct"} else ()
        if app.file_stem and app.winget_version:
            normalized = _normalize_version_string(app.winget_version) or app.winget_version
            safe_version = _safe_file_part(normalized)
            version_pattern = f"{app.file_stem}_{safe_version}.*"
            for directory in search_dirs:
                for candidate in directory.glob(version_pattern):
                    if candidate.suffix.lower() in {".exe", ".msi"}:
                        return LocalInstallerInfo(True, path=candidate)
        path = self._best_local_by_patterns(search_dirs, patterns, exact_names)
        if not path and include_downloads:
            fallback_dir = self._downloads_dir / _safe_name(app.name)
            candidates = list(fallback_dir.glob("*.exe")) + list(fallback_dir.glob("*.msi"))
            best_path, _ = _best_versioned_installer(candidates)
            path = best_path or _pick_best_candidate(candidates)
        return LocalInstallerInfo(bool(path), path=path)

    def get_local_installer_versions(self, app: AppEntry, info: LocalInstallerInfo) -> LocalInstallerVersionInfo:
        if not info.exists:
            return LocalInstallerVersionInfo()
        if app.download_mode == "office":
            version = self._office.payload_version(app.name)
            return LocalInstallerVersionInfo(version=version)
        if app.dual_arch:
            return LocalInstallerVersionInfo(
                version_x86=_local_version_from_path(info.path_x86),
                version_x64=_local_version_from_path(info.path_x64),
                path_x86=info.path_x86,
                path_x64=info.path_x64,
            )
        return LocalInstallerVersionInfo(
            version=_local_version_from_path(info.path),
            path=info.path,
        )

    def install_selected(
        self,
        selection: Iterable[str],
        *,
        progress_callback: Callable[[int, int, str], None] | None = None,
        status_callback: Callable[[str], None] | None = None,
    ) -> list[OperationResult]:
        results: list[OperationResult] = []
        apps = list(self._selected_apps(selection))
        total = len(apps)
        for index, app in enumerate(apps, start=1):
            if status_callback:
                status_callback(app.name)
            results.append(self._install_app(app))
            if progress_callback:
                progress_callback(index, total, app.name)
        return results

    def download_selected(
        self,
        selection: Iterable[str],
        *,
        progress_callback: Callable[[int, int, str], None] | None = None,
        status_callback: Callable[[str], None] | None = None,
        force_updates: Iterable[str] | None = None,
    ) -> list[OperationResult]:
        results: list[OperationResult] = []
        apps = list(self._selected_apps(selection))
        total = len(apps)
        force_set = {name.lower() for name in force_updates or []}
        for index, app in enumerate(apps, start=1):
            if status_callback:
                status_callback(app.name)
            force_update = app.name.lower() in force_set
            results.append(self._download_app(app, status_callback=status_callback, force_update=force_update))
            if progress_callback:
                progress_callback(index, total, app.name)
        return results

    def local_version_override_warnings(self, app: AppEntry) -> list[str]:
        if not app.winget_version or not app.winget_version.strip():
            return []
        info = self.get_local_installer_info(app, include_downloads=True)
        paths: list[Path] = []
        if app.dual_arch:
            if info.path_x86:
                paths.append(info.path_x86)
            if info.path_x64:
                paths.append(info.path_x64)
        else:
            if info.path:
                paths.append(info.path)
        if not paths:
            return []
        expected = _normalize_version_string(app.winget_version) or app.winget_version.strip()
        warnings: list[str] = []
        for path in paths:
            actual = _extract_version_from_filename(path.name)
            actual_normalized = _normalize_version_string(actual) if actual else None
            if actual_normalized == expected:
                continue
            actual_display = actual_normalized or "unknown"
            warnings.append(
                f"{app.name}: local installer {path.name} (version {actual_display}) will be used instead of {expected}"
            )
        return warnings

    def _selected_apps(self, selection: Iterable[str]) -> Iterable[AppEntry]:
        wanted = {name.lower() for name in selection}
        for app in self._apps:
            if app.name.lower() in wanted:
                yield app

    def _install_from_local(self, app: AppEntry, info: LocalInstallerInfo) -> OperationResult:
        if app.dual_arch:
            results: list[OperationResult] = []
            success = True
            if info.path_x86:
                results.append(self._run_local_installer(app, info.path_x86))
            if _is_64bit() and info.path_x64:
                results.append(self._run_local_installer(app, info.path_x64))
            if not results:
                return OperationResult(app, "install", False, "Local installer not found")
            success = all(result.success for result in results)
            stdout = "\n".join(result.stdout for result in results if result.stdout)
            stderr = "\n".join(result.stderr for result in results if result.stderr)
            message = "Installed from local files" if success else "Local install failed"
            return OperationResult(app, "install", success, message, stdout, stderr)
        if info.path:
            return self._run_local_installer(app, info.path)
        return OperationResult(app, "install", False, "Local installer not found")

    def _install_app(self, app: AppEntry) -> OperationResult:
        if app.name == "Office Deployment Tool":
            return OperationResult(app, "install", True, "ODT is managed automatically with Office downloads")
        if app.name == "CrowdStrike Falcon Sensor" and not _has_crowdstrike_cid(app, self._settings):
            return OperationResult(app, "install", False, "CrowdStrike CID not configured in settings")
        if app.download_mode in {"winget", "onlineonly"}:
            local_info = self.get_local_installer_info(app, include_downloads=True)
            if local_info.exists:
                result = self._install_from_local(app, local_info)
            else:
                result = self._install_via_winget(app)
            return self._apply_post_install_steps(app, result)
        if app.download_mode == "office":
            try:
                result = self._office.install(app.name)
                return OperationResult(
                    app,
                    "install",
                    result.succeeded,
                    "Office deployment finished" if result.succeeded else "Office deployment failed",
                    result.stdout,
                    result.stderr,
                )
            except Exception as exc:
                return OperationResult(app, "install", False, str(exc))
        if app.download_mode == "direct":
            local_info = self.get_local_installer_info(app, include_downloads=True)
            if local_info.exists:
                result = self._install_from_local(app, local_info)
                return self._apply_post_install_steps(app, result)
            download_result = self._download_direct(app)
            if not download_result.success:
                return OperationResult(app, "install", False, download_result.message, download_result.stdout, download_result.stderr)
            local_info = self.get_local_installer_info(app, include_downloads=True)
            if local_info.exists:
                result = self._install_from_local(app, local_info)
                return self._apply_post_install_steps(app, result)
            return OperationResult(app, "install", False, "Downloaded installer missing after download")
        if app.download_mode == "localonly":
            local_info = self.get_local_installer_info(app, include_downloads=False)
            if not local_info.exists:
                return OperationResult(app, "install", False, "Local installer not found in working directory")
            result = self._install_from_local(app, local_info)
            return self._apply_post_install_steps(app, result)
        return OperationResult(app, "install", False, f"Download mode {app.download_mode} not implemented")

    def _download_app(
        self,
        app: AppEntry,
        *,
        status_callback: Callable[[str], None] | None = None,
        force_update: bool = False,
    ) -> OperationResult:
        if app.name == "Office Deployment Tool":
            return OperationResult(app, "download", True, "ODT is managed automatically with Office downloads")
        if app.download_mode == "onlineonly":
            return OperationResult(app, "download", True, "Online-only package; offline download not available")
        if app.download_mode == "winget":
            return self._download_via_winget(app, status_callback=status_callback)
        if app.download_mode == "office":
            try:
                result = self._office.download(app.name, status_callback=status_callback)
                message = "Office download completed" if result.succeeded else "Office download failed"
                if result.succeeded and "ODT_DOWNLOAD_COMPLETE" in result.stderr:
                    message = "Office download completed (setup.exe stopped after download)"
                if result.succeeded and "ODT_IDLE_TIMEOUT" in result.stderr:
                    message = "Office download completed (setup.exe stopped after inactivity)"
                return OperationResult(app, "download", result.succeeded, message, result.stdout, result.stderr)
            except Exception as exc:
                return OperationResult(app, "download", False, str(exc))
        if app.download_mode == "direct":
            return self._download_direct(app, status_callback=status_callback, force_update=force_update)
        if app.download_mode == "localonly":
            return OperationResult(app, "download", True, "Local-only package; place installer manually")
        return OperationResult(app, "download", False, f"Download mode {app.download_mode} not implemented")

    def _install_via_winget(self, app: AppEntry) -> OperationResult:
        package_ids = self._package_ids_for(app)
        if not package_ids:
            return OperationResult(app, "install", False, "No Winget package id configured")
        if not self._winget.is_available():
            return OperationResult(app, "install", False, "winget executable not found")
        stdout_parts: list[str] = []
        stderr_parts: list[str] = []
        success = True
        version = app.winget_version.strip() if app.winget_version else None
        for package_id in package_ids:
            try:
                result = self._winget.install_package(
                    package_id,
                    source=app.source,
                    override=app.args or None,
                    version=version,
                )
            except WingetError as exc:
                return OperationResult(app, "install", False, str(exc))
            stdout_parts.append(result.stdout)
            stderr_parts.append(result.stderr)
            success = success and result.succeeded
        message = "Installed via winget" if success else "winget install failed"
        return OperationResult(app, "install", success, message, "\n".join(stdout_parts), "\n".join(stderr_parts))

    def _download_via_winget(
        self,
        app: AppEntry,
        *,
        status_callback: Callable[[str], None] | None = None,
    ) -> OperationResult:
        if app.source and app.source != "winget":
            return OperationResult(app, "download", True, f"Download not supported for source {app.source}")
        package_ids = self._package_ids_for(app)
        if not package_ids:
            return OperationResult(app, "download", False, "No Winget package id configured")
        if not self._winget.is_available():
            return OperationResult(app, "download", False, "winget executable not found")
        stdout_parts: list[str] = []
        stderr_parts: list[str] = []
        messages: list[str] = []
        success = True
        version_override = app.winget_version.strip() if app.winget_version else None
        target_root = self._downloads_dir / _safe_name(app.name)
        target_root.mkdir(parents=True, exist_ok=True)
        packages: list[tuple[str, str]] = []
        if app.dual_arch:
            if app.winget_id_x86:
                stem = app.file_stem_x86 or app.file_stem or _safe_name(app.name)
                packages.append((app.winget_id_x86, stem))
            if _is_64bit() and app.winget_id_x64:
                stem = app.file_stem_x64 or app.file_stem or _safe_name(app.name)
                packages.append((app.winget_id_x64, stem))
        elif app.winget_id:
            stem = app.file_stem or _safe_name(app.name)
            packages.append((app.winget_id, stem))
        for package_id, stem in packages:
            if version_override:
                version_raw = version_override
            else:
                try:
                    version_raw = self._winget.show_package_version(package_id, source=app.source)
                except WingetError as exc:
                    success = False
                    messages.append(f"{stem}: {exc}")
                    continue
            version = _normalize_version_string(version_raw) or version_raw or "unknown"
            safe_version = _safe_file_part(version)
            existing = self._find_existing_versioned_file(target_root, stem, safe_version)
            if existing:
                messages.append(f"{stem}: already have {existing.name}")
                continue
            temp_dir = target_root / f"temp_{_safe_name(stem)}"
            shutil.rmtree(temp_dir, ignore_errors=True)
            monitor = _start_speed_monitor(temp_dir, app.name, status_callback)
            try:
                result = self._winget.download_package(
                    package_id,
                    destination=temp_dir,
                    source=app.source,
                    version=version_override,
                )
            except WingetError as exc:
                success = False
                messages.append(f"{stem}: {exc}")
                continue
            finally:
                _stop_speed_monitor(monitor)
            stdout_parts.append(result.stdout)
            stderr_parts.append(result.stderr)
            if not result.succeeded:
                shutil.rmtree(temp_dir, ignore_errors=True)
                error_msg = result.stderr.strip() or result.stdout.strip() or f"download failed (exit code {result.returncode})"
                fallback_path = None
                if app.name == "Chrome" and _winget_hash_mismatch(error_msg):
                    try:
                        fallback_path = self._download_chrome_direct(
                            target_root,
                            stem,
                            safe_version,
                            status_callback=status_callback,
                        )
                    except Exception as exc:
                        success = False
                        messages.append(f"{stem}: {error_msg}; direct download failed ({exc})")
                        continue
                if fallback_path:
                    messages.append(f"{stem}: winget hash mismatch; downloaded {fallback_path.name} direct")
                else:
                    success = False
                    messages.append(f"{stem}: {error_msg}")
                continue
            installer = self._find_downloaded_installer(temp_dir)
            if not installer:
                shutil.rmtree(temp_dir, ignore_errors=True)
                success = False
                messages.append(f"{stem}: installer not found after download")
                continue
            dest_path = target_root / f"{stem}_{safe_version}{installer.suffix.lower()}"
            try:
                shutil.move(str(installer), dest_path)
            except OSError as exc:
                shutil.rmtree(temp_dir, ignore_errors=True)
                success = False
                messages.append(f"{stem}: rename failed ({exc})")
                continue
            shutil.rmtree(temp_dir, ignore_errors=True)
            _remove_versioned_files(target_root, stem, keep_name=dest_path.name)
            messages.append(f"{stem}: downloaded {dest_path.name}")
        message = "; ".join(messages) if messages else "No packages downloaded"
        return OperationResult(app, "download", success, message, "\n".join(stdout_parts), "\n".join(stderr_parts))

    def _download_direct(
        self,
        app: AppEntry,
        *,
        status_callback: Callable[[str], None] | None = None,
        force_update: bool = False,
    ) -> OperationResult:
        downloader = self._direct_downloaders.get(app.name)
        if not downloader:
            return OperationResult(app, "download", False, f"No direct downloader registered for {app.name}")
        try:
            info = downloader.fetch()
        except Exception as exc:
            return OperationResult(app, "download", False, f"Direct download failed: {exc}")
        stem = app.file_stem or _safe_name(app.name)
        version = _safe_file_part(info.version)
        filename = info.filename or f"{stem}_{version}.exe"
        destination_dir = self._downloads_dir / _safe_name(app.name)
        destination_dir.mkdir(parents=True, exist_ok=True)
        dest_path = destination_dir / filename
        latest_version = _normalize_version_string(info.version) or info.version
        latest_tuple = _version_tuple(latest_version)
        existing_installers = _list_installers(destination_dir)
        best_path, best_tuple = _best_versioned_installer(existing_installers)
        local_up_to_date = False
        update_needed = False
        if latest_tuple and best_tuple:
            local_up_to_date = best_tuple >= latest_tuple
            update_needed = best_tuple < latest_tuple
        should_update = force_update or update_needed
        if existing_installers and local_up_to_date and not force_update:
            best_name = best_path.name if best_path else dest_path.name
            return OperationResult(app, "download", True, f"Installer already present: {best_name}")
        if should_update:
            _remove_installers_in_dir(destination_dir)
        if app.name in {"CrowdStrike Falcon Sensor", "FortiClient VPN"}:
            if not _is_sharepoint_url(info.url):
                return OperationResult(app, "download", False, f"{app.name} download URL must be a SharePoint link")
            already_present = dest_path.exists() and _file_has_exe_header(dest_path)
            try:
                dest_path = _download_sharepoint_exe(
                    info.url,
                    destination_dir,
                    filename,
                    status_callback=status_callback,
                    label=app.name,
                    force=should_update,
                )
            except Exception as exc:
                return OperationResult(app, "download", False, f"Download error: {exc}")
            if already_present and not should_update:
                return OperationResult(app, "download", True, f"Installer already present: {dest_path.name}")
            return OperationResult(app, "download", True, f"Downloaded {dest_path.name}")
        if dest_path.exists() and not should_update:
            return OperationResult(app, "download", True, f"Installer already present: {dest_path.name}")
        try:
            self._download_file(info.url, dest_path, status_callback=status_callback, label=app.name)
        except Exception as exc:
            return OperationResult(app, "download", False, f"Download error: {exc}")
        return OperationResult(app, "download", True, f"Downloaded {dest_path.name}")

    def _download_chrome_direct(
        self,
        target_root: Path,
        stem: str,
        safe_version: str,
        *,
        status_callback: Callable[[str], None] | None = None,
    ) -> Path:
        url = _chrome_download_url()
        dest_path = target_root / f"{stem}_{safe_version}.msi"
        self._download_file(url, dest_path, status_callback=status_callback, label="Chrome")
        _remove_versioned_files(target_root, stem, keep_name=dest_path.name)
        return dest_path

    def _best_local_by_patterns(
        self,
        search_dirs: Sequence[Path],
        patterns: Sequence[str],
        exact_names: Sequence[str] = (),
    ) -> Path | None:
        candidates: list[Path] = []
        for directory in search_dirs:
            if not directory.exists():
                continue
            for pattern in patterns:
                candidates.extend(directory.glob(pattern))
            for name in exact_names:
                candidate = directory / name
                if candidate.exists():
                    candidates.append(candidate)
        return _pick_best_candidate(candidates)

    def _find_existing_versioned_file(self, target_root: Path, stem: str, version: str) -> Path | None:
        pattern = f"{stem}_{version}.*"
        for directory in (target_root, self._working_dir):
            for candidate in directory.glob(pattern):
                if candidate.suffix.lower() in {".exe", ".msi"}:
                    return candidate
        return None

    def _find_downloaded_installer(self, temp_dir: Path) -> Path | None:
        if not temp_dir.exists():
            return None
        for candidate in sorted(temp_dir.rglob("*")):
            if candidate.is_file() and candidate.suffix.lower() in {".exe", ".msi"}:
                return candidate
        return None

    def _package_ids_for(self, app: AppEntry) -> list[str]:
        ids: list[str] = []
        if app.dual_arch:
            if app.winget_id_x86:
                ids.append(app.winget_id_x86)
            if _is_64bit() and app.winget_id_x64:
                ids.append(app.winget_id_x64)
            return ids
        if app.winget_id:
            ids.append(app.winget_id)
        return ids

    def _run_local_installer(self, app: AppEntry, path: Path) -> OperationResult:
        if path.suffix.lower() == ".msi":
            cmd = ["msiexec", "/i", str(path)]
        else:
            cmd = [str(path)]
        if app.args:
            cmd.extend(shlex.split(app.args))
        completed = subprocess.run(cmd, capture_output=True, text=True, check=False)
        success = completed.returncode == 0
        message = "Local install completed" if success else "Local install failed"
        return OperationResult(app, "install", success, message, completed.stdout, completed.stderr)

    def _download_file(
        self,
        url: str,
        destination: Path,
        *,
        status_callback: Callable[[str], None] | None = None,
        label: str | None = None,
    ) -> None:
        _download_file_with_final_url(url, destination, status_callback=status_callback, label=label)

    def _apply_post_install_steps(self, app: AppEntry, result: OperationResult) -> OperationResult:
        if not result.success:
            return result
        if app.name != "WinRAR":
            return result
        license_result = _apply_winrar_license(self._settings)
        if license_result.success:
            message = f"{result.message}; {license_result.message}" if license_result.message else result.message
            return OperationResult(app, result.operation, True, message, result.stdout, result.stderr)
        message = f"{result.message}; {license_result.message}" if license_result.message else "WinRAR license copy failed"
        return OperationResult(app, result.operation, False, message, result.stdout, result.stderr)


def _download_file_with_final_url(
    url: str,
    destination: Path,
    *,
    status_callback: Callable[[str], None] | None = None,
    label: str | None = None,
) -> str:
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120",
        "Accept": "*/*",
        "Accept-Language": "en-US,en;q=0.9",
    }
    parsed = urllib.parse.urlparse(url)
    referer = None
    if "hikvision.com" in parsed.netloc.lower():
        referer = IVMSDownloader.SERIES_URL
        headers["Referer"] = referer
        headers["Origin"] = "https://www.hikvision.com"
    request = urllib.request.Request(url, headers=headers)
    destination.parent.mkdir(parents=True, exist_ok=True)
    cookie_jar = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cookie_jar))
    if referer:
        try:
            opener.open(urllib.request.Request(referer, headers=headers), timeout=20).read(1)
        except Exception:
            pass
    with opener.open(request, timeout=60) as response, destination.open("wb") as handle:
        final_url = response.geturl()
        last_time = time.monotonic() - 0.5
        last_bytes = 0
        downloaded = 0
        while True:
            chunk = response.read(256 * 1024)
            if not chunk:
                break
            handle.write(chunk)
            downloaded += len(chunk)
            now = time.monotonic()
            if now - last_time >= 0.5:
                delta_bytes = downloaded - last_bytes
                _emit_speed_update(
                    status_callback,
                    label or "Downloading",
                    delta_bytes=delta_bytes,
                    elapsed=now - last_time,
                )
                last_time = now
                last_bytes = downloaded
        if downloaded > last_bytes:
            now = time.monotonic()
            _emit_speed_update(
                status_callback,
                label or "Downloading",
                delta_bytes=downloaded - last_bytes,
                elapsed=now - last_time,
            )
    return final_url


def _is_sharepoint_url(url: str) -> bool:
    parsed = urllib.parse.urlparse(url.strip())
    if parsed.scheme not in {"http", "https"}:
        return False
    host = parsed.netloc.lower()
    if not host:
        return False
    return host.endswith("sharepoint.com") or ".sharepoint." in host or host.startswith("sharepoint.")


def _sharepoint_download_url(url: str) -> str:
    cleaned = url.strip()
    parsed = urllib.parse.urlsplit(cleaned)
    if not parsed.scheme or not parsed.netloc:
        return f"{cleaned}&download=1" if "?" in cleaned else f"{cleaned}?download=1"
    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    query = [(key, value) for key, value in query if key.lower() != "download"]
    query.append(("download", "1"))
    new_query = urllib.parse.urlencode(query)
    return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, new_query, parsed.fragment))


def _download_file_with_requests(
    url: str,
    destination: Path,
    *,
    status_callback: Callable[[str], None] | None = None,
    label: str | None = None,
    timeout: float = 120.0,
) -> None:
    try:
        import requests
    except ImportError as exc:
        raise RuntimeError("SharePoint downloads require the requests package (pip install requests).") from exc
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120",
        "Accept": "*/*",
        "Accept-Language": "en-US,en;q=0.9",
    }
    destination.parent.mkdir(parents=True, exist_ok=True)
    with requests.get(
        url,
        headers=headers,
        stream=True,
        allow_redirects=True,
        timeout=timeout,
    ) as response:
        response.raise_for_status()
        last_time = time.monotonic() - 0.5
        last_bytes = 0
        downloaded = 0
        with destination.open("wb") as handle:
            for chunk in response.iter_content(chunk_size=1024 * 1024):
                if not chunk:
                    continue
                handle.write(chunk)
                downloaded += len(chunk)
                now = time.monotonic()
                if now - last_time >= 0.5:
                    delta_bytes = downloaded - last_bytes
                    _emit_speed_update(
                        status_callback,
                        label or "Downloading",
                        delta_bytes=delta_bytes,
                        elapsed=now - last_time,
                    )
                    last_time = now
                    last_bytes = downloaded
        if downloaded > last_bytes:
            now = time.monotonic()
            _emit_speed_update(
                status_callback,
                label or "Downloading",
                delta_bytes=downloaded - last_bytes,
                elapsed=now - last_time,
            )


def _file_has_exe_header(path: Path) -> bool:
    try:
        with path.open("rb") as handle:
            header = handle.read(2)
    except OSError:
        return False
    return len(header) == 2 and header[0] == 0x4D and header[1] == 0x5A


def _contains_permission_error(exc: BaseException) -> bool:
    if isinstance(exc, PermissionError):
        return True
    for arg in getattr(exc, "args", ()):
        if isinstance(arg, PermissionError):
            return True
    cause = exc.__cause__ or exc.__context__
    if cause and cause is not exc:
        return _contains_permission_error(cause)
    return False


def _pick_sharepoint_temp_path(
    destination_dir: Path,
    filename: str,
) -> Path:
    temp_path = destination_dir / f".{filename}.download"
    if not temp_path.exists():
        return temp_path
    try:
        temp_path.unlink()
        time.sleep(0.25)
        return temp_path
    except OSError:
        suffix = f"{int(time.time())}-{os.getpid()}"
        fallback = destination_dir / f".{filename}.{suffix}.download"
        return fallback


def _download_sharepoint_exe(
    share_url: str,
    destination_dir: Path,
    filename: str,
    *,
    status_callback: Callable[[str], None] | None = None,
    label: str | None = None,
    force: bool = False,
) -> Path:
    share_url = share_url.strip()
    if not share_url:
        raise RuntimeError("SharePoint download URL not configured")
    if not _is_sharepoint_url(share_url):
        raise RuntimeError("Download URL is not a SharePoint link")
    destination_dir.mkdir(parents=True, exist_ok=True)
    dest_path = destination_dir / filename
    if not force and dest_path.exists() and _file_has_exe_header(dest_path):
        return dest_path
    temp_path = _pick_sharepoint_temp_path(destination_dir, filename)
    download_url = _sharepoint_download_url(share_url)
    for attempt in range(1, 4):
        try:
            _download_file_with_requests(
                download_url,
                temp_path,
                status_callback=status_callback,
                label=label,
            )
            break
        except Exception as exc:
            if not _contains_permission_error(exc):
                raise
            if attempt >= 3:
                raise
            try:
                if temp_path.exists():
                    temp_path.unlink()
            except OSError:
                pass
            time.sleep(0.5 * attempt)
            temp_path = _pick_sharepoint_temp_path(destination_dir, filename)
    if _file_has_exe_header(temp_path):
        temp_path.replace(dest_path)
        return dest_path
    try:
        temp_path.unlink()
    except OSError:
        pass
    raise RuntimeError(
        "Download did not return an EXE (got HTML/redirect). Share link may not be truly anonymous or tenant blocks direct download."
    )


def _normalize_version_string(value: str | None) -> str | None:
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
        return _normalize_version_string(match.group(1))
    return cleaned


def _version_tuple(version: str | None) -> tuple[int, ...]:
    if not version:
        return tuple()
    parts = version.split(".")
    if not all(part.isdigit() for part in parts):
        return tuple()
    return tuple(int(part) for part in parts)


def _safe_file_part(value: str) -> str:
    cleaned = value.strip().replace(" ", "_")
    cleaned = re.sub(r"[^a-zA-Z0-9._-]", "", cleaned)
    return cleaned or "unknown"


def _winget_hash_mismatch(message: str) -> bool:
    lowered = message.lower()
    return "installer hash does not match" in lowered or "hash does not match" in lowered


def _extract_version_from_filename(filename: str) -> str | None:
    match = re.search(r"_([0-9]+(?:\.[0-9]+){1,3})\.(exe|msi)$", filename, re.IGNORECASE)
    if match:
        return _normalize_version_string(match.group(1))
    return None


def _local_version_from_path(path: Path | None) -> str | None:
    if not path or not path.exists() or not path.is_file():
        return None
    from_name = _extract_version_from_filename(path.name)
    if from_name:
        return from_name
    file_version = _get_file_version(path)
    if not file_version:
        return None
    return _normalize_version_string(file_version) or file_version


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


def _pick_best_candidate(files: Sequence[Path]) -> Path | None:
    best: Path | None = None
    best_version: tuple[int, ...] | None = None
    for candidate in sorted({file for file in files}, key=lambda p: p.name.lower()):
        version_token = _extract_version_from_filename(candidate.name)
        version_tuple = _version_tuple(version_token) if version_token else tuple()
        if version_tuple:
            if best_version is None or version_tuple > best_version:
                best_version = version_tuple
                best = candidate
        elif best is None:
            best = candidate
    return best


def _list_installers(target_root: Path) -> list[Path]:
    return list(target_root.glob("*.exe")) + list(target_root.glob("*.msi"))


def _best_versioned_installer(candidates: Sequence[Path]) -> tuple[Path | None, tuple[int, ...] | None]:
    best_path: Path | None = None
    best_version: tuple[int, ...] | None = None
    for candidate in candidates:
        version = _local_version_from_path(candidate)
        version_tuple = _version_tuple(version)
        if not version_tuple:
            continue
        if best_version is None or version_tuple > best_version:
            best_version = version_tuple
            best_path = candidate
    return best_path, best_version


def _remove_versioned_files(target_root: Path, stem: str, *, keep_name: str | None = None) -> None:
    pattern = f"{stem}_*"
    for candidate in target_root.glob(pattern):
        if candidate.suffix.lower() not in {".exe", ".msi"}:
            continue
        if keep_name and candidate.name == keep_name:
            continue
        try:
            candidate.unlink()
        except OSError:
            continue


def _remove_installers_in_dir(target_root: Path) -> None:
    for candidate in target_root.glob("*"):
        if candidate.suffix.lower() not in {".exe", ".msi"}:
            continue
        try:
            candidate.unlink()
        except OSError:
            continue


def _safe_name(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9_-]+", "_", name).lower()


def _is_64bit() -> bool:
    return sys.maxsize > 2**32


def _chrome_download_url() -> str:
    if _is_64bit():
        return "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi"
    return "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise.msi"


def _filename_from_url(url: str) -> str | None:
    parsed = urllib.parse.urlparse(url)
    name = Path(parsed.path).name
    if not name:
        return None
    if Path(name).suffix.lower() not in {".exe", ".msi"}:
        return None
    return name


def _directory_size_bytes(path: Path) -> int:
    if path.is_file():
        try:
            return path.stat().st_size
        except OSError:
            return 0
    total = 0
    for root, _, files in os.walk(path):
        for filename in files:
            try:
                total += (Path(root) / filename).stat().st_size
            except OSError:
                continue
    return total


def _directory_payload_size_bytes(path: Path) -> int:
    if path.is_file():
        try:
            return path.stat().st_size
        except OSError:
            return 0
    total = 0
    ignore_exts = {".log", ".txt"}
    ignore_names = {"setup.exe", "config.xml"}
    for root, _, files in os.walk(path):
        for filename in files:
            name = filename.lower()
            if name in ignore_names:
                continue
            if Path(name).suffix.lower() in ignore_exts:
                continue
            try:
                total += (Path(root) / filename).stat().st_size
            except OSError:
                continue
    return total


def _dir_has_entries(path: Path) -> bool:
    if not path.exists() or not path.is_dir():
        return False
    try:
        next(path.iterdir())
    except StopIteration:
        return False
    except OSError:
        return False
    return True


def _office_payload_present(office_dir: Path) -> bool:
    if _dir_has_entries(office_dir / "Office"):
        return True
    return _dir_has_entries(office_dir / "Data")


def _office_payload_versions(office_dir: Path) -> list[str]:
    versions: list[str] = []
    for data_root in (office_dir / "Office" / "Data", office_dir / "Data"):
        if not data_root.exists() or not data_root.is_dir():
            continue
        try:
            entries = list(data_root.iterdir())
        except OSError:
            continue
        for entry in entries:
            if not entry.is_dir():
                continue
            normalized = _normalize_version_string(entry.name)
            if not normalized:
                continue
            if not _version_tuple(normalized):
                continue
            versions.append(normalized)
    return versions


def _format_speed(value: float) -> str:
    speed = max(value, 0.0)
    units = ["B/s", "KB/s", "MB/s", "GB/s", "TB/s"]
    for unit in units:
        if speed < 1024 or unit == units[-1]:
            return f"{speed:.1f} {unit}"
        speed /= 1024
    return f"{speed:.1f} B/s"


def _format_speed_label(label: str, speed_bytes_per_sec: float) -> str:
    return f"Downloading {label}: {_format_speed(speed_bytes_per_sec)}"


def _run_office_download(
    cmd: Sequence[str],
    working_dir: Path,
    *,
    label: str,
    status_callback: Callable[[str], None] | None,
    idle_timeout: float = 180.0,
    complete_timeout: float = 60.0,
    interval: float = 1.0,
) -> CommandExecutionResult:
    proc = subprocess.Popen(cmd, cwd=working_dir, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    payload_root = working_dir / "Office"
    if not payload_root.exists():
        payload_root = working_dir
    last_size = _directory_payload_size_bytes(payload_root)
    last_time = time.monotonic()
    last_change = last_time
    had_activity = False
    terminated_for_idle = False
    terminated_for_complete = False
    while True:
        if proc.poll() is not None:
            break
        time.sleep(interval)
        payload_root = working_dir / "Office"
        if not payload_root.exists():
            payload_root = working_dir
        size = _directory_payload_size_bytes(payload_root)
        now = time.monotonic()
        delta_bytes = size - last_size
        _emit_speed_update(
            status_callback,
            label,
            delta_bytes=delta_bytes,
            elapsed=now - last_time,
        )
        if delta_bytes != 0:
            had_activity = True
            last_change = now
        if complete_timeout and had_activity and (now - last_change) >= complete_timeout:
            terminated_for_complete = True
            proc.terminate()
            try:
                proc.wait(timeout=15)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.wait(timeout=5)
            break
        if idle_timeout and (now - last_change) >= idle_timeout:
            terminated_for_idle = True
            proc.terminate()
            try:
                proc.wait(timeout=15)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.wait(timeout=5)
            break
        last_size = size
        last_time = now
    stdout, stderr = proc.communicate()
    returncode = proc.returncode or 0
    if terminated_for_complete:
        returncode = 0
        note = "ODT_DOWNLOAD_COMPLETE: download activity stopped; setup.exe terminated."
        stderr = f"{stderr}\n{note}" if stderr else note
    if terminated_for_idle:
        returncode = 0 if had_activity else (proc.returncode or 1)
        note = f"ODT_IDLE_TIMEOUT: no download activity for {int(idle_timeout)}s; setup.exe terminated."
        stderr = f"{stderr}\n{note}" if stderr else note
    return CommandExecutionResult(cmd, returncode, stdout or "", stderr or "")


def _monitor_directory_speed(
    monitor_path: Path,
    label: str,
    status_callback: Callable[[str], None],
    stop_event: threading.Event,
    *,
    interval: float = 1.0,
) -> None:
    last_size = _directory_size_bytes(monitor_path)
    last_time = time.monotonic()
    while not stop_event.wait(interval):
        size = _directory_size_bytes(monitor_path)
        now = time.monotonic()
        delta_bytes = size - last_size
        _emit_speed_update(
            status_callback,
            label,
            delta_bytes=delta_bytes,
            elapsed=now - last_time,
        )
        last_size = size
        last_time = now


def _emit_speed_update(
    status_callback: Callable[[str], None] | None,
    label: str,
    *,
    delta_bytes: int,
    elapsed: float,
) -> None:
    if not status_callback:
        return
    if delta_bytes <= 0:
        return
    speed = delta_bytes / max(elapsed, 0.001)
    status_callback(_format_speed_label(label, speed))


def _start_speed_monitor(
    monitor_path: Path,
    label: str,
    status_callback: Callable[[str], None] | None,
) -> tuple[threading.Event, threading.Thread] | None:
    if not status_callback:
        return None
    stop_event = threading.Event()
    thread = threading.Thread(
        target=_monitor_directory_speed,
        args=(monitor_path, label, status_callback, stop_event),
        daemon=True,
    )
    thread.start()
    return stop_event, thread


def _stop_speed_monitor(monitor: tuple[threading.Event, threading.Thread] | None) -> None:
    if not monitor:
        return
    stop_event, thread = monitor
    stop_event.set()
    thread.join(timeout=1.0)


def _has_crowdstrike_cid(app: AppEntry, settings: UserSettings) -> bool:
    if "CID=" in app.args:
        return True
    return bool(settings.crowdstrike_cid.strip())


def _apply_winrar_license(settings: UserSettings) -> LicenseApplyResult:
    license_path = settings.winrar_license_path.strip()
    if not license_path:
        return LicenseApplyResult(False, "WinRAR license path not configured")
    source = Path(license_path)
    if not source.exists() or not source.is_file():
        return LicenseApplyResult(False, f"WinRAR license file not found: {source}")
    install_dir = _find_winrar_install_dir()
    if not install_dir:
        return LicenseApplyResult(False, "WinRAR install directory not found")
    destination = install_dir / source.name
    try:
        shutil.copy2(source, destination)
    except OSError as exc:
        return LicenseApplyResult(False, f"WinRAR license copy failed: {exc}")
    return LicenseApplyResult(True, f"WinRAR license applied ({destination})")


def _find_winrar_install_dir() -> Path | None:
    for env_key in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env_key)
        if not base:
            continue
        candidate = Path(base) / "WinRAR"
        if (candidate / "WinRAR.exe").exists():
            return candidate
    return None
