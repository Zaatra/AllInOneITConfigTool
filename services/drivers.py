"""Driver scanning and installation services using HP tooling."""
from __future__ import annotations

import json
import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Protocol, Sequence

from trustpal.constants import IMMUTABLE_CONFIG


@dataclass
class DriverRecord:
    name: str
    status: str
    source: str
    installed_version: str | None
    latest_version: str | None
    category: str | None = None
    softpaq_id: str | None = None
    download_url: str | None = None
    output_path: Path | None = None


@dataclass
class DriverOperationResult:
    driver: DriverRecord
    operation: str
    success: bool
    message: str


@dataclass
class HPSystemInfo:
    platform_id: str | None = None
    model: str | None = None
    supports_hpia: bool = False
    supports_cmsl: bool = False
    supports_legacy_repo: bool = True


class CommandRunner(Protocol):
    def run(self, command: Sequence[str]) -> subprocess.CompletedProcess[str]:  # pragma: no cover - protocol
        ...


class SubprocessRunner:
    def run(self, command: Sequence[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.run(command, capture_output=True, text=True, check=False)


class HPIAClient:
    def __init__(
        self,
        working_dir: Path,
        *,
        executable: str | None = None,
        command_runner: CommandRunner | None = None,
    ) -> None:
        self._working_dir = Path(working_dir)
        self._runner = command_runner or SubprocessRunner()
        self._executable = Path(executable) if executable else self._auto_detect()
        self._download_dir = self._working_dir / "hpia_softpaqs"
        self._report_dir = self._working_dir / "hpia_reports"

    def is_available(self) -> bool:
        return self._executable is not None and self._executable.exists()

    def scan(self) -> list[DriverRecord]:
        exe = self._require_executable()
        if self._report_dir.exists():
            shutil.rmtree(self._report_dir)
        self._report_dir.mkdir(parents=True, exist_ok=True)
        args = [
            str(exe),
            "/Operation:Analyze",
            "/Category:All",
            "/Selection:All",
            "/Action:List",
            f"/ReportFolder:{self._report_dir}",
            "/Silent",
        ]
        result = self._runner.run(args)
        if result.returncode != 0:
            raise RuntimeError(f"HPIA scan failed: {result.stderr}")
        report_file = next(self._report_dir.rglob("*.json"), None)
        if not report_file:
            return []
        data = json.loads(report_file.read_text(encoding="utf-8"))
        recommendations = data.get("HPIA", {}).get("Recommendations", [])
        records: list[DriverRecord] = []
        for rec in recommendations:
            status = rec.get("RecommendationValue", "Optional")
            records.append(
                DriverRecord(
                    name=rec.get("Name", "Unknown"),
                    status=status,
                    source="HPIA",
                    installed_version=rec.get("CurrentVersion"),
                    latest_version=rec.get("Version"),
                    category=rec.get("Category"),
                    softpaq_id=rec.get("SoftPaqId"),
                    download_url=rec.get("ReleaseNotesUrl"),
                )
            )
        return records

    def download(self, softpaq_ids: Sequence[str]) -> dict[str, Path]:
        if not softpaq_ids:
            return {}
        exe = self._require_executable()
        self._download_dir.mkdir(parents=True, exist_ok=True)
        args = [
            str(exe),
            "/Operation:Download",
            "/Selection:SoftPaq",
            f"/Softpaq:{';'.join(softpaq_ids)}",
            f"/ReportFolder:{self._download_dir}",
            "/Silent",
        ]
        result = self._runner.run(args)
        if result.returncode != 0:
            raise RuntimeError(f"HPIA download failed: {result.stderr}")
        mapping: dict[str, Path] = {}
        for spid in softpaq_ids:
            candidate = next(self._download_dir.glob(f"{spid}*.exe"), None)
            if candidate:
                mapping[spid] = candidate
        return mapping

    def _auto_detect(self) -> Path | None:
        candidates = [
            Path("C:/Program Files/HP/HPIA/HPImageAssistant.exe"),
            Path("C:/Program Files (x86)/HP/HPIA/HPImageAssistant.exe"),
            self._working_dir / "HPIA" / "HPImageAssistant.exe",
            self._working_dir / "HPImageAssistant.exe",
        ]
        for candidate in candidates:
            if candidate.exists():
                return candidate
        return None

    def _require_executable(self) -> Path:
        if not self.is_available():
            raise FileNotFoundError("HPImageAssistant.exe not found")
        return self._executable  # type: ignore[return-value]


class CMSLClient:
    def __init__(
        self,
        *,
        powershell: str = "powershell",
        command_runner: CommandRunner | None = None,
    ) -> None:
        self._powershell = powershell
        self._runner = command_runner or SubprocessRunner()

    def is_available(self) -> bool:
        return shutil.which(self._powershell) is not None

    def scan(self, platform_id: str | None) -> list[DriverRecord]:
        if not platform_id:
            return []
        script = (
            "Import-Module HPCMSL -ErrorAction Stop; "
            f"$sp = Get-SoftpaqList -Platform '{platform_id}' -Os Win11 -OsVer 24H2 -ErrorAction Stop; "
            "$sp | ConvertTo-Json -Depth 4"
        )
        result = self._runner.run([self._powershell, "-NoProfile", "-Command", script])
        if result.returncode != 0 or not result.stdout:
            return []
        try:
            data = json.loads(result.stdout)
        except json.JSONDecodeError:
            return []
        records: list[DriverRecord] = []
        if isinstance(data, dict):
            data = [data]
        for item in data or []:
            if not isinstance(item, dict):
                continue
            category = item.get("Category", "")
            if "driver" not in category.lower() and "bios" not in category.lower() and "firmware" not in category.lower():
                continue
            records.append(
                DriverRecord(
                    name=item.get("Name", "Unknown"),
                    status="Available",
                    source="CMSL",
                    installed_version=None,
                    latest_version=item.get("Version"),
                    category=category,
                    softpaq_id=item.get("Id") or item.get("SoftPaqId"),
                    download_url=item.get("Url"),
                )
            )
        return records

    def download(self, softpaq_id: str, destination: Path) -> Path:
        destination.parent.mkdir(parents=True, exist_ok=True)
        script = (
            "Import-Module HPCMSL -ErrorAction Stop; "
            f"Get-Softpaq -Number {softpaq_id} -SaveAs '{destination}' -Overwrite -ErrorAction Stop"
        )
        result = self._runner.run([self._powershell, "-NoProfile", "-Command", script])
        if result.returncode != 0:
            raise RuntimeError(f"CMSL download failed for {softpaq_id}: {result.stderr}")
        return destination


class LegacyRepository:
    def __init__(self, root: str | Path | None = None) -> None:
        repo_root = root or IMMUTABLE_CONFIG.ids.hp_legacy_repo_root
        self._root = Path(repo_root)

    def list_packages(self, platform_id: str | None, model: str | None) -> list[DriverRecord]:
        candidates = []
        if platform_id:
            candidates.append(self._root / platform_id)
        if model:
            candidates.append(self._root / model)
            clean = model.replace("HP ", "").replace("Hewlett-Packard ", "")
            candidates.append(self._root / clean)
        records: list[DriverRecord] = []
        for candidate in candidates:
            manifest = candidate / "manifest.json"
            if not manifest.exists():
                continue
            try:
                data = json.loads(manifest.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                continue
            for item in data:
                file_name = item.get("File") or item.get("Path")
                if not file_name:
                    continue
                file_path = candidate / file_name
                records.append(
                    DriverRecord(
                        name=item.get("Name", "Legacy Driver"),
                        status="Legacy",
                        source="Legacy",
                        installed_version=None,
                        latest_version=item.get("Version"),
                        category=item.get("Category"),
                        softpaq_id=item.get("SoftPaqId"),
                        download_url=str(file_path),
                        output_path=file_path,
                    )
                )
            if records:
                break
        return records


class DriverService:
    def __init__(
        self,
        *,
        working_dir: Path | str | None = None,
        hpia_client: HPIAClient | None = None,
        cmsl_client: CMSLClient | None = None,
        legacy_repo: LegacyRepository | None = None,
        command_runner: CommandRunner | None = None,
        system_info_provider: Callable[[], HPSystemInfo] | None = None,
    ) -> None:
        self._working_dir = Path(working_dir or Path.cwd())
        self._runner = command_runner or SubprocessRunner()
        self._hpia = hpia_client or HPIAClient(self._working_dir)
        self._cmsl = cmsl_client or CMSLClient()
        self._legacy = legacy_repo or LegacyRepository()
        self._system_info_provider = system_info_provider or (lambda: HPSystemInfo())

    def scan(self) -> list[DriverRecord]:
        info = self._system_info_provider()
        records: list[DriverRecord] = []
        if info.supports_hpia and self._hpia.is_available():
            try:
                records.extend(self._hpia.scan())
            except Exception:
                pass
        if info.supports_cmsl and self._cmsl.is_available():
            try:
                records.extend(self._cmsl.scan(info.platform_id))
            except Exception:
                pass
        if not records and info.supports_legacy_repo:
            records.extend(self._legacy.list_packages(info.platform_id, info.model))
        return records

    def download(self, records: Iterable[DriverRecord]) -> list[DriverOperationResult]:
        ops: list[DriverOperationResult] = []
        hpia_targets = [r for r in records if r.source == "HPIA" and r.softpaq_id]
        cmsl_targets = [r for r in records if r.source == "CMSL" and r.softpaq_id]
        legacy_targets = [r for r in records if r.source == "Legacy" and r.download_url]

        if hpia_targets:
            try:
                mapping = self._hpia.download([r.softpaq_id for r in hpia_targets if r.softpaq_id])
                for record in hpia_targets:
                    record.output_path = mapping.get(record.softpaq_id)
                    success = record.output_path is not None
                    ops.append(DriverOperationResult(record, "download", success, "Downloaded" if success else "Missing output"))
            except Exception as exc:
                for record in hpia_targets:
                    ops.append(DriverOperationResult(record, "download", False, str(exc)))

        for record in cmsl_targets:
            try:
                dest = self._working_dir / "cmsl_softpaqs" / f"{record.softpaq_id}.exe"
                record.output_path = self._cmsl.download(record.softpaq_id or "", dest)
                ops.append(DriverOperationResult(record, "download", True, "Downloaded"))
            except Exception as exc:
                ops.append(DriverOperationResult(record, "download", False, str(exc)))

        for record in legacy_targets:
            try:
                src = Path(record.download_url or "")
                dest_dir = self._working_dir / "legacy_drivers"
                dest_dir.mkdir(parents=True, exist_ok=True)
                dest = dest_dir / src.name
                shutil.copy2(src, dest)
                record.output_path = dest  # type: ignore[assignment]
                ops.append(DriverOperationResult(record, "download", True, "Copied"))
            except Exception as exc:
                ops.append(DriverOperationResult(record, "download", False, str(exc)))

        return ops

    def install(self, records: Iterable[DriverRecord]) -> list[DriverOperationResult]:
        ops: list[DriverOperationResult] = []
        for record in records:
            if not record.output_path:
                ops.append(DriverOperationResult(record, "install", False, "No installer downloaded"))
                continue
            cmd = [str(record.output_path), "/s"]
            result = self._runner.run(cmd)
            success = result.returncode in {0, 3010}
            message = "Installed" if success else f"Installer exit {result.returncode}"
            ops.append(DriverOperationResult(record, "install", success, message))
        return ops
