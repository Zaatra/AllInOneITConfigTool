from __future__ import annotations

from pathlib import Path
from typing import Iterable, Sequence

import pytest

from services.installer import (
    CommandExecutionResult,
    DirectDownloadInfo,
    InstallerService,
    OperationResult,
    WingetClient,
)
from trustpal.app_registry import AppEntry


class DummyWingetClient(WingetClient):
    def __init__(self) -> None:
        super().__init__(executable="winget")
        self._available = True
        self.installs: list[tuple[str, str | None, str | None]] = []
        self.downloads: list[tuple[str, Path]] = []

    def is_available(self) -> bool:  # type: ignore[override]
        return self._available

    def install_package(self, package_id: str, *, source: str | None = None, override: str | None = None, silent: bool = True, force: bool = True) -> CommandExecutionResult:  # type: ignore[override]
        self.installs.append((package_id, source, override))
        return CommandExecutionResult(["winget", "install", package_id], 0, "ok", "")

    def download_package(self, package_id: str, destination: Path, *, source: str | None = None, force: bool = True) -> CommandExecutionResult:  # type: ignore[override]
        destination.mkdir(parents=True, exist_ok=True)
        self.downloads.append((package_id, destination))
        return CommandExecutionResult(["winget", "download", package_id], 0, "ok", "")


class DummyOfficeInstaller:
    def __init__(self) -> None:
        self.install_calls: list[str] = []
        self.setup_downloaded = False

    def install(self, app_name: str) -> CommandExecutionResult:
        self.install_calls.append(app_name)
        return CommandExecutionResult(["setup.exe", "/configure", "config.xml"], 0, "done", "")

    def ensure_setup(self) -> CommandExecutionResult | None:
        self.setup_downloaded = True
        return CommandExecutionResult(["winget", "install", "Microsoft.OfficeDeploymentTool"], 0, "ok", "")


@pytest.fixture()
def chrome_entry() -> AppEntry:
    return AppEntry(
        category="Core",
        name="Chrome",
        download_mode="winget",
        winget_id="Google.Chrome",
    )


@pytest.fixture()
def office_entry() -> AppEntry:
    return AppEntry(
        category="Core",
        name="Office 2024 LTSC",
        download_mode="office",
        winget_id="Microsoft.OfficeDeploymentTool",
    )


def _service(apps: Iterable[AppEntry], winget_client: DummyWingetClient | None = None, office_installer: DummyOfficeInstaller | None = None) -> InstallerService:
    return InstallerService(
        apps,
        working_dir=Path.cwd(),
        winget_client=winget_client,
        office_installer=office_installer,
    )


def test_install_via_winget_uses_client(chrome_entry: AppEntry) -> None:
    fake_client = DummyWingetClient()
    service = _service([chrome_entry], winget_client=fake_client)
    results = service.install_selected(["Chrome"])
    assert results[0].success
    assert fake_client.installs == [("Google.Chrome", None, None)]


def test_download_via_winget_uses_client(chrome_entry: AppEntry, tmp_path: Path) -> None:
    fake_client = DummyWingetClient()
    service = InstallerService(
        [chrome_entry],
        working_dir=tmp_path,
        winget_client=fake_client,
    )
    results = service.download_selected(["Chrome"])
    assert results[0].success
    assert fake_client.downloads[0][0] == "Google.Chrome"
    assert fake_client.downloads[0][1].exists()


def test_office_install_delegates_to_office_installer(office_entry: AppEntry) -> None:
    fake_office = DummyOfficeInstaller()
    service = _service([office_entry], office_installer=fake_office)
    results = service.install_selected(["Office 2024 LTSC"])
    assert results[0].success
    assert fake_office.install_calls == ["Office 2024 LTSC"]


def test_office_download_triggers_setup_download(office_entry: AppEntry) -> None:
    fake_office = DummyOfficeInstaller()
    service = _service([office_entry], office_installer=fake_office)
    results = service.download_selected(["Office 2024 LTSC"])
    assert results[0].success
    assert fake_office.setup_downloaded is True


class DummyDirectDownloader:
    def __init__(self) -> None:
        self.fetch_count = 0

    def fetch(self) -> DirectDownloadInfo:
        self.fetch_count += 1
        return DirectDownloadInfo(version="1.2.3.4", url="https://example.com/ivms.exe", filename="ivms.exe")


def test_direct_download_uses_downloader(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    chrome_entry = AppEntry(
        category="Core",
        name="iVMS-4200",
        download_mode="direct",
        file_stem="ivms4200",
    )
    downloader = DummyDirectDownloader()

    def fake_download(self: InstallerService, url: str, destination: Path) -> None:  # type: ignore[override]
        destination.write_bytes(b"test")

    monkeypatch.setattr(InstallerService, "_download_file", fake_download, raising=False)
    service = InstallerService(
        [chrome_entry],
        working_dir=tmp_path,
        direct_downloaders={"iVMS-4200": downloader},
    )
    result = service.download_selected(["iVMS-4200"])[0]
    assert result.success
    assert downloader.fetch_count == 1
    downloads = list((tmp_path / "downloads").rglob("ivms.exe"))
    assert downloads, "Downloaded installer missing"


def test_local_install_finds_installer(tmp_path: Path) -> None:
    installer_path = tmp_path / "crowdstrike_falcon_sensor_1.0.exe"
    installer_path.write_text("#!/bin/sh\nexit 0\n")
    installer_path.chmod(0o755)

    app = AppEntry(
        category="Security",
        name="CrowdStrike Falcon Sensor",
        download_mode="localonly",
        file_stem="crowdstrike_falcon_sensor",
    )
    service = InstallerService([app], working_dir=tmp_path)
    result = service.install_selected([app.name])[0]
    assert result.success
