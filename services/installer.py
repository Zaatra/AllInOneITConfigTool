"""Application installation orchestration."""
from __future__ import annotations

import os
import re
import shlex
import shutil
import subprocess
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Mapping, Protocol, Sequence

from allinone_it_config.app_registry import AppEntry
from allinone_it_config.constants import IMMUTABLE_CONFIG, OfficeTemplate


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
        silent: bool = True,
        force: bool = True,
    ) -> CommandExecutionResult:
        cmd = self._build_base_command("install", package_id, source, force)
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
        force: bool = True,
    ) -> CommandExecutionResult:
        destination.mkdir(parents=True, exist_ok=True)
        cmd = self._build_base_command("download", package_id, source, force)
        cmd.extend(["-d", str(destination)])
        return self._run(cmd)

    def show_package_version(self, package_id: str, *, source: str | None = None) -> str | None:
        if not self._executable:
            raise WingetError("winget executable not found in PATH")
        cmd = [str(self._executable), "show", "--id", package_id, "--exact", "--accept-source-agreements", "--locale", "en-US"]
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
        if force:
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

    TEMPLATE_KEYS = {
        "Office 2024 LTSC": "office_2024_ltsc",
        "Office 365 Ent": "office_365_enterprise",
    }

    def __init__(
        self,
        working_dir: Path,
        *,
        winget_client: WingetClient | None = None,
        templates: dict[str, OfficeTemplate] | None = None,
    ) -> None:
        self._working_dir = Path(working_dir)
        self._winget = winget_client or WingetClient()
        self._templates = templates or IMMUTABLE_CONFIG.office_templates
        self._setup_path = self._working_dir / "setup.exe"
        self._staging_dir = self._working_dir / "OfficeSetup"

    def ensure_setup(self) -> CommandExecutionResult | None:
        if self._setup_path.exists():
            return None
        if not self._winget.is_available():
            raise WingetError("winget unavailable to fetch Office Deployment Tool")
        self._staging_dir.mkdir(parents=True, exist_ok=True)
        override = f"/quiet /extract:{self._staging_dir}"
        result = self._winget.install_package(
            "Microsoft.OfficeDeploymentTool",
            override=override,
        )
        candidate = self._staging_dir / "setup.exe"
        if candidate.exists():
            shutil.move(candidate, self._setup_path)
        if not self._setup_path.exists():
            raise FileNotFoundError("setup.exe missing after Office Deployment Tool extraction")
        return result

    def install(self, app_name: str) -> CommandExecutionResult:
        key = self.TEMPLATE_KEYS.get(app_name)
        if not key:
            raise KeyError(f"No Office template mapped for {app_name}")
        template = self._templates[key]
        self._staging_dir.mkdir(parents=True, exist_ok=True)
        config_path = self._staging_dir / "config.xml"
        config_path.write_text(template.xml, encoding="utf-8")
        self.ensure_setup()
        cmd = [str(self._setup_path), "/configure", str(config_path)]
        completed = subprocess.run(cmd, capture_output=True, text=True, check=False)
        return CommandExecutionResult(cmd, completed.returncode, completed.stdout, completed.stderr)


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


class DirectDownloader(Protocol):
    def fetch(self) -> DirectDownloadInfo:
        ...


class IVMSDownloader:
    """Scrapes Hikvision's site for the latest iVMS-4200 build."""

    SOURCE_URL = "https://www.hikvision.com/us-en/support/download/software/"
    VERSION_PATTERNS = (
        re.compile(r"iVMS-4200V(\d+\.\d+\.\d+\.\d+)_E", re.IGNORECASE),
        re.compile(r"iVMS-4200\s+V(\d+\.\d+\.\d+\.\d+)", re.IGNORECASE),
    )

    def fetch(self) -> DirectDownloadInfo:
        request = urllib.request.Request(
            self.SOURCE_URL,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120",
                "Accept-Language": "en-US,en;q=0.9",
            },
        )
        with urllib.request.urlopen(request, timeout=20) as response:
            html = response.read().decode("utf-8", errors="ignore")
        versions = self._extract_versions(html)
        if not versions:
            raise RuntimeError("Unable to determine latest iVMS-4200 version")
        version = max(versions, key=_version_tuple)
        dash = "v" + version.replace(".", "-")
        url = (
            "https://www.hikvision.com/content/dam/hikvision/en/support/download/vms/"
            "ivms4200-series/software-download/{dash}/iVMS-4200V{version}_E.exe".format(dash=dash, version=version)
        )
        return DirectDownloadInfo(version=version, url=url)

    def _extract_versions(self, html: str) -> set[str]:
        found: set[str] = set()
        for pattern in self.VERSION_PATTERNS:
            for match in pattern.finditer(html):
                ver = match.group(1)
                if ver:
                    found.add(ver)
        return found


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
    ) -> None:
        self._apps = list(apps)
        self._working_dir = Path(working_dir or Path.cwd())
        self._downloads_dir = self._working_dir / "downloads"
        self._winget = winget_client or WingetClient()
        self._office = office_installer or OfficeInstaller(self._working_dir, winget_client=self._winget)
        default_direct = {"iVMS-4200": IVMSDownloader()}
        self._direct_downloaders = dict(default_direct)
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
        if app.download_mode == "onlineonly":
            return LocalInstallerInfo(False)
        if app.download_mode == "office":
            setup_path = self._working_dir / "setup.exe"
            return LocalInstallerInfo(setup_path.exists(), path=setup_path if setup_path.exists() else None)
        search_dirs = [self._working_dir]
        if include_downloads:
            search_dirs.append(self._downloads_dir)
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
        exact_names = tuple(app.local_alt_names) if app.download_mode == "localonly" else ()
        path = self._best_local_by_patterns(search_dirs, patterns, exact_names)
        return LocalInstallerInfo(bool(path), path=path)

    def install_selected(self, selection: Iterable[str]) -> list[OperationResult]:
        results: list[OperationResult] = []
        for app in self._selected_apps(selection):
            results.append(self._install_app(app))
        return results

    def download_selected(self, selection: Iterable[str]) -> list[OperationResult]:
        results: list[OperationResult] = []
        for app in self._selected_apps(selection):
            results.append(self._download_app(app))
        return results

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
        if app.download_mode in {"winget", "onlineonly"}:
            local_info = self.get_local_installer_info(app, include_downloads=True)
            if local_info.exists:
                return self._install_from_local(app, local_info)
            return self._install_via_winget(app)
        if app.download_mode == "office":
            try:
                result = self._office.install(app.name)
                return OperationResult(app, "install", result.succeeded, "Office deployment finished" if result.succeeded else "Office deployment failed", result.stdout, result.stderr)
            except Exception as exc:
                return OperationResult(app, "install", False, str(exc))
        if app.download_mode == "direct":
            local_info = self.get_local_installer_info(app, include_downloads=True)
            if local_info.exists:
                return self._install_from_local(app, local_info)
            download_result = self._download_direct(app)
            if not download_result.success:
                return OperationResult(app, "install", False, download_result.message, download_result.stdout, download_result.stderr)
            local_info = self.get_local_installer_info(app, include_downloads=True)
            if local_info.exists:
                return self._install_from_local(app, local_info)
            return OperationResult(app, "install", False, "Downloaded installer missing after download")
        if app.download_mode == "localonly":
            local_info = self.get_local_installer_info(app, include_downloads=False)
            if not local_info.exists:
                return OperationResult(app, "install", False, "Local installer not found in working directory")
            return self._install_from_local(app, local_info)
        return OperationResult(app, "install", False, f"Download mode {app.download_mode} not implemented")

    def _download_app(self, app: AppEntry) -> OperationResult:
        if app.download_mode == "onlineonly":
            return OperationResult(app, "download", True, "Online-only package; offline download not available")
        if app.download_mode == "winget":
            return self._download_via_winget(app)
        if app.download_mode == "office":
            try:
                result = self._office.ensure_setup()
                message = "setup.exe already present" if result is None else "Downloaded Office Deployment Tool"
                stdout = result.stdout if result else ""
                stderr = result.stderr if result else ""
                return OperationResult(app, "download", True, message, stdout, stderr)
            except Exception as exc:
                return OperationResult(app, "download", False, str(exc))
        if app.download_mode == "direct":
            return self._download_direct(app)
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
        for package_id in package_ids:
            try:
                result = self._winget.install_package(
                    package_id,
                    source=app.source,
                    override=app.args or None,
                )
            except WingetError as exc:
                return OperationResult(app, "install", False, str(exc))
            stdout_parts.append(result.stdout)
            stderr_parts.append(result.stderr)
            success = success and result.succeeded
        message = "Installed via winget" if success else "winget install failed"
        return OperationResult(app, "install", success, message, "\n".join(stdout_parts), "\n".join(stderr_parts))

    def _download_via_winget(self, app: AppEntry) -> OperationResult:
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
            version_raw = None
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
            try:
                result = self._winget.download_package(
                    package_id,
                    destination=temp_dir,
                    source=app.source,
                )
            except WingetError as exc:
                success = False
                messages.append(f"{stem}: {exc}")
                continue
            stdout_parts.append(result.stdout)
            stderr_parts.append(result.stderr)
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
            messages.append(f"{stem}: downloaded {dest_path.name}")
        message = "; ".join(messages) if messages else "No packages downloaded"
        return OperationResult(app, "download", success, message, "\n".join(stdout_parts), "\n".join(stderr_parts))

    def _download_direct(self, app: AppEntry) -> OperationResult:
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
        if dest_path.exists():
            return OperationResult(app, "download", True, f"Installer already present: {dest_path.name}")
        try:
            self._download_file(info.url, dest_path)
        except Exception as exc:
            return OperationResult(app, "download", False, f"Download error: {exc}")
        return OperationResult(app, "download", True, f"Downloaded {dest_path.name}")

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
        cmd = [str(path)]
        if app.args:
            cmd.extend(shlex.split(app.args))
        completed = subprocess.run(cmd, capture_output=True, text=True, check=False)
        success = completed.returncode == 0
        message = "Local install completed" if success else "Local install failed"
        return OperationResult(app, "install", success, message, completed.stdout, completed.stderr)

    def _download_file(self, url: str, destination: Path) -> None:
        request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(request, timeout=60) as response:
            destination.write_bytes(response.read())


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


def _extract_version_from_filename(filename: str) -> str | None:
    match = re.search(r"_([0-9]+(?:\.[0-9]+){1,3})\.(exe|msi)$", filename, re.IGNORECASE)
    if match:
        return _normalize_version_string(match.group(1))
    return None


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


def _safe_name(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9_-]+", "_", name).lower()


def _is_64bit() -> bool:
    return sys.maxsize > 2**32
