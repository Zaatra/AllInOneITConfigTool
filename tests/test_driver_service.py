from __future__ import annotations

import subprocess
from pathlib import Path
from typing import Sequence

import pytest

from services.drivers import (
    DriverOperationResult,
    DriverRecord,
    DriverService,
    HPSystemInfo,
)


class FakeRunner:
    def __init__(self) -> None:
        self.commands: list[Sequence[str]] = []

    def run(self, command: Sequence[str]) -> subprocess.CompletedProcess[str]:
        self.commands.append(tuple(command))
        return subprocess.CompletedProcess(command, 0, "", "")


class FakeHPIAClient:
    def __init__(self) -> None:
        self.available = True
        self.scanned = False
        self.download_calls: list[list[str]] = []

    def is_available(self) -> bool:
        return self.available

    def scan(self) -> list[DriverRecord]:
        self.scanned = True
        return [
            DriverRecord(
                name="HPIA Driver",
                status="Critical",
                source="HPIA",
                installed_version="1",
                latest_version="2",
                softpaq_id="SP123",
            )
        ]

    def download(self, softpaq_ids: list[str]) -> dict[str, Path]:
        self.download_calls.append(list(softpaq_ids))
        return {sp: Path(f"/tmp/{sp}.exe") for sp in softpaq_ids}


class FakeCMSLClient:
    def __init__(self) -> None:
        self.scan_calls: list[str | None] = []

    def is_available(self) -> bool:
        return True

    def scan(self, platform_id: str | None) -> list[DriverRecord]:
        self.scan_calls.append(platform_id)
        return [
            DriverRecord(
                name="CMSL Driver",
                status="Available",
                source="CMSL",
                installed_version=None,
                latest_version="1.0",
                softpaq_id="SP999",
            )
        ]

    def download(self, softpaq_id: str, destination: Path) -> Path:
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text("echo")
        return destination


class FakeLegacyRepo:
    def list_packages(self, platform_id: str | None, model: str | None) -> list[DriverRecord]:
        return [
            DriverRecord(
                name="Legacy Driver",
                status="Legacy",
                source="Legacy",
                installed_version=None,
                latest_version="1.0",
                download_url=str(Path("/repo/legacy.exe")),
            )
        ]


def _service(tmp_path: Path) -> DriverService:
    fake_hpia = FakeHPIAClient()
    fake_cmsl = FakeCMSLClient()
    legacy = FakeLegacyRepo()
    info = HPSystemInfo(platform_id="1234", model="HP Test", supports_hpia=True, supports_cmsl=True)
    return DriverService(
        working_dir=tmp_path,
        hpia_client=fake_hpia,
        cmsl_client=fake_cmsl,
        legacy_repo=legacy,
        system_info_provider=lambda: info,
    )


def test_scan_combines_hpia_and_cmsl(tmp_path: Path) -> None:
    service = _service(tmp_path)
    records = service.scan()
    assert any(r.source == "HPIA" for r in records)
    assert any(r.source == "CMSL" for r in records)


def test_download_routes_by_source(tmp_path: Path) -> None:
    service = _service(tmp_path)
    records = [
        DriverRecord("HPIA Driver", "Critical", "HPIA", "1", "2", softpaq_id="SP111"),
        DriverRecord("CMSL Driver", "Available", "CMSL", None, "1", softpaq_id="SP222"),
        DriverRecord("Legacy Driver", "Legacy", "Legacy", None, "1", download_url=str(Path(tmp_path, "legacy.exe"))),
    ]
    Path(tmp_path, "legacy.exe").write_text("legacy")
    ops = service.download(records)
    assert len([op for op in ops if op.success]) == 3
    assert all(record.output_path for record in records)


def test_install_runs_downloaded_packages(tmp_path: Path) -> None:
    runner = FakeRunner()
    service = DriverService(
        working_dir=tmp_path,
        hpia_client=FakeHPIAClient(),
        cmsl_client=FakeCMSLClient(),
        legacy_repo=FakeLegacyRepo(),
        command_runner=runner,
        system_info_provider=lambda: HPSystemInfo(),
    )
    record = DriverRecord("Legacy Driver", "Legacy", "Legacy", None, "1")
    record.output_path = tmp_path / "legacy.exe"
    record.output_path.write_text("echo")
    results = service.install([record])
    assert results[0].success
    assert (str(record.output_path), "/s") in runner.commands
