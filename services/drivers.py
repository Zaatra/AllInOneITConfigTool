"""Driver scanning and installation services using HP tooling."""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Protocol, Sequence

from allinone_it_config.constants import IMMUTABLE_CONFIG
from allinone_it_config.paths import get_application_directory

try:
    import winreg  # type: ignore[import-not-found]
except ImportError:
    winreg = None  # type: ignore[assignment]


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
    manufacturer: str | None = None
    serial_number: str | None = None
    sku: str | None = None
    generation: int | None = None
    os_version: str | None = None
    os_build: str | None = None
    supports_hpia: bool = False
    supports_cmsl: bool = False
    supports_legacy_repo: bool = True


@dataclass
class InstalledItem:
    name: str
    version: str
    publisher: str | None = None


def _normalize_version(version_str: str | None) -> str:
    if not version_str:
        return "0.0.0.0"
    parts = version_str.strip().split(".")
    while len(parts) < 4:
        parts.append("0")
    return ".".join(parts[:4])


def _compare_versions(installed: str | None, available: str | None) -> int | None:
    if not installed or not available:
        return None
    norm_installed = _normalize_version(installed)
    norm_available = _normalize_version(available)
    try:
        inst_parts = [int(p) for p in norm_installed.split(".")]
        avail_parts = [int(p) for p in norm_available.split(".")]
        if avail_parts > inst_parts:
            return 1
        if avail_parts < inst_parts:
            return -1
        return 0
    except (ValueError, AttributeError):
        return None


def _compare_version_strings(left: str | None, right: str | None) -> int | None:
    if not left or not right:
        return None
    norm_left = _normalize_version(left)
    norm_right = _normalize_version(right)
    try:
        left_parts = [int(p) for p in norm_left.split(".")]
        right_parts = [int(p) for p in norm_right.split(".")]
        if left_parts < right_parts:
            return -1
        if left_parts > right_parts:
            return 1
        return 0
    except (ValueError, AttributeError):
        return None


def _normalize_name(value: str) -> str:
    text = value.lower()
    text = text.replace("wi-fi", "wifi").replace("wi fi", "wifi")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return text.strip()


def _categorize_cmsl(category: str | None, name: str | None) -> str:
    raw = f"{category or ''} {name or ''}".lower()
    if "bios" in raw or "firmware" in raw or "uefi" in raw:
        return "BIOS/Firmware"
    if "audio" in raw or "sound" in raw:
        return "Audio"
    if "video" in raw or "graphics" in raw or "display" in raw:
        return "Video"
    if "network" in raw or "ethernet" in raw or "lan" in raw:
        return "Network"
    if "wireless" in raw or "wifi" in raw or "wlan" in raw or "bluetooth" in raw:
        return "Network"
    if "storage" in raw or "sata" in raw or "raid" in raw or "rst" in raw or "nvme" in raw:
        return "Storage"
    if "chipset" in raw or "serial" in raw or "usb" in raw:
        return "Chipset"
    if "input" in raw or "keyboard" in raw or "touchpad" in raw or "mouse" in raw:
        return "Input"
    if "security" in raw or "tpm" in raw:
        return "Security"
    if "software" in raw or "utility" in raw or "management" in raw:
        return "Software"
    return "Other"


def _dedupe_latest_records(records: list[DriverRecord]) -> list[DriverRecord]:
    best: dict[str, DriverRecord] = {}
    for record in records:
        key = _normalize_name(record.name)
        current = best.get(key)
        if not current:
            best[key] = record
            continue
        if current.latest_version is None and record.latest_version:
            best[key] = record
            continue
        cmp_result = _compare_version_strings(current.latest_version, record.latest_version)
        if cmp_result is None:
            continue
        if cmp_result < 0:
            best[key] = record
    return list(best.values())


def get_hp_system_info(*, powershell: str = "powershell") -> HPSystemInfo:
    info = HPSystemInfo()
    if not shutil.which(powershell):
        return info
    script = """
    $cs = Get-WmiObject Win32_ComputerSystem
    $bios = Get-WmiObject Win32_BIOS
    $bb = Get-WmiObject Win32_BaseBoard
    $os = Get-WmiObject Win32_OperatingSystem
    $csProduct = Get-WmiObject Win32_ComputerSystemProduct
    $result = @{
        Manufacturer = $cs.Manufacturer
        Model = $cs.Model
        SerialNumber = $bios.SerialNumber
        ProductCode = $bb.Product
        OSVersion = $os.Caption
        OSBuild = $os.BuildNumber
        SKU = $csProduct.SKUNumber
    }
    $regPath = 'HKLM:\\HARDWARE\\DESCRIPTION\\System\\BIOS'
    $biosReg = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
    if ($biosReg -and $biosReg.SystemSKU) {
        $result.SKU = $biosReg.SystemSKU
        if (-not $result.ProductCode -or $result.ProductCode.Length -lt 4) {
            $result.ProductCode = $biosReg.SystemSKU
        }
    }
    $result | ConvertTo-Json -Compress
    """
    try:
        result = subprocess.run([powershell, "-NoProfile", "-Command", script], capture_output=True, text=True, check=False, timeout=10)
        if result.returncode == 0 and result.stdout:
            data = json.loads(result.stdout)
            manufacturer = data.get("Manufacturer", "")
            model = data.get("Model", "")
            os_version = data.get("OSVersion", "")
            if re.search(r"HP|Hewlett|Packard", manufacturer, re.IGNORECASE):
                info.manufacturer = manufacturer
                info.model = model
                info.serial_number = data.get("SerialNumber")
                info.platform_id = data.get("ProductCode") or data.get("SKU")
                info.sku = data.get("SKU")
                info.os_version = os_version
                info.os_build = data.get("OSBuild")
                gen_match = re.search(r"G(\d+)", model)
                if gen_match:
                    info.generation = int(gen_match.group(1))
                info.supports_hpia = (info.generation is not None and info.generation >= 3) or bool(
                    re.search(r"Z[0-9]+ G|ZBook.*G[3-9]|Elite.*G[3-9]|Pro.*G[3-9]", model, re.IGNORECASE)
                )
                if re.search(r"Windows 7|Windows 8", os_version, re.IGNORECASE):
                    info.supports_cmsl = False
                else:
                    info.supports_cmsl = True
                if re.search(r"Compaq|Pro3?500|dc\d{4}|8[0-3]00", model, re.IGNORECASE):
                    info.supports_hpia = False
                    info.supports_cmsl = False
    except (subprocess.TimeoutExpired, json.JSONDecodeError, ValueError):
        pass
    return info


def get_installed_drivers_and_software(*, powershell: str = "powershell") -> dict[str, InstalledItem]:
    installed: dict[str, InstalledItem] = {}
    if not shutil.which(powershell):
        return installed
    script = """
    $items = @()
    $regPaths = @(
        'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*',
        'HKLM:\\SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*'
    )
    foreach ($path in $regPaths) {
        Get-ItemProperty $path -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and $_.DisplayVersion } | ForEach-Object {
            $items += @{
                Name = $_.DisplayName
                Version = $_.DisplayVersion
                Publisher = $_.Publisher
                Type = 'Registry'
            }
        }
    }
    Get-WmiObject Win32_PnPSignedDriver -ErrorAction SilentlyContinue | Where-Object { $_.DeviceName -and $_.DriverVersion } | ForEach-Object {
        $items += @{
            Name = $_.DeviceName
            Version = $_.DriverVersion
            Publisher = $_.Manufacturer
            Type = 'Driver'
        }
    }
    $bios = Get-WmiObject Win32_BIOS -ErrorAction SilentlyContinue
    if ($bios) {
        $items += @{
            Name = 'System BIOS'
            Version = $bios.SMBIOSBIOSVersion
            Publisher = $bios.Manufacturer
            Type = 'BIOS'
        }
    }
    $items | ConvertTo-Json -Depth 2 -Compress
    """
    try:
        result = subprocess.run([powershell, "-NoProfile", "-Command", script], capture_output=True, text=True, check=False, timeout=30)
        if result.returncode == 0 and result.stdout:
            data = json.loads(result.stdout)
            if not isinstance(data, list):
                data = [data]
            for item in data:
                if not isinstance(item, dict):
                    continue
                name = item.get("Name", "").lower().strip()
                version = item.get("Version", "")
                publisher = item.get("Publisher")
                if name and version and name not in installed:
                    installed[name] = InstalledItem(name=item.get("Name", ""), version=version, publisher=publisher)
    except (subprocess.TimeoutExpired, json.JSONDecodeError, ValueError):
        pass
    return installed


def find_installed_version(driver_name: str, category: str | None, installed_cache: dict[str, InstalledItem]) -> str | None:
    driver_lower = driver_name.lower()
    driver_norm = _normalize_name(driver_name)
    is_bios = bool(re.search(r"\bbios\b", driver_lower)) or (category and "bios" in category.lower())
    if is_bios:
        bios_item = installed_cache.get("system bios")
        if bios_item and bios_item.version:
            return bios_item.version
    search_terms: list[str] = []
    if "intel" in driver_lower:
        search_terms.append("intel")
    if "realtek" in driver_lower:
        search_terms.append("realtek")
    if "nvidia" in driver_lower:
        search_terms.append("nvidia")
    if "amd" in driver_lower:
        search_terms.append("amd")
    if "bluetooth" in driver_lower:
        search_terms.append("bluetooth")
    if re.search(r"wireless|wlan|wifi|wi-fi", driver_lower):
        search_terms.extend(["wireless", "wlan", "wifi"])
    if re.search(r"graphics|video|display", driver_lower):
        search_terms.extend(["graphics", "video", "display"])
    if re.search(r"audio|sound", driver_lower):
        search_terms.extend(["audio", "sound"])
    if re.search(r"ethernet|nic|network", driver_lower):
        search_terms.extend(["ethernet", "network"])
    if "chipset" in driver_lower:
        search_terms.append("chipset")
    if re.search(r"storage|raid|rst|rapid", driver_lower):
        search_terms.extend(["storage", "rapid", "rst"])
    if "bios" in driver_lower:
        search_terms.append("bios")
    if "firmware" in driver_lower:
        search_terms.append("firmware")
    if re.search(r"management engine|me driver", driver_lower):
        search_terms.append("management engine")
    if "thunderbolt" in driver_lower:
        search_terms.append("thunderbolt")
    if re.search(r"serial io|serialio", driver_lower):
        search_terms.append("serial")
    if re.search(r"arc|a380|a770", driver_lower):
        search_terms.append("arc")
    if "usb 3" in driver_lower:
        search_terms.append("usb 3")
    is_wireless_driver = bool(re.search(r"\b(wlan|wifi|wireless)\b", driver_norm))
    best_match: InstalledItem | None = None
    best_score = 0
    for item_name, item_data in installed_cache.items():
        item_norm = _normalize_name(item_name)
        if is_bios and not re.search(r"\bbios\b", item_norm):
            continue
        if is_wireless_driver and "manageability" in item_norm and "manageability" not in driver_norm:
            continue
        score = 0
        for term in search_terms:
            if term in item_norm:
                score += 1
        if category:
            cat_lower = category.lower()
            if "graphics" in cat_lower and re.search(r"graphics|display|video", item_norm):
                score += 2
            if "audio" in cat_lower and re.search(r"audio|sound|realtek", item_norm):
                score += 2
            if "network" in cat_lower and re.search(r"network|ethernet|wireless|wifi|bluetooth", item_norm):
                score += 2
            if "chipset" in cat_lower and re.search(r"chipset|serial|management|usb", item_norm):
                score += 2
            if "storage" in cat_lower and re.search(r"storage|rapid|rst|raid|optane", item_norm):
                score += 2
            if "bios" in cat_lower and re.search(r"\bbios\b", item_norm):
                score += 2
            elif "firmware" in cat_lower and re.search(r"firmware", item_norm):
                score += 2
        if score > best_score:
            best_score = score
            best_match = item_data
    if best_match and best_score >= 2:
        return best_match.version
    return None


def get_driver_status(driver_name: str, category: str | None, available_version: str | None, installed_cache: dict[str, InstalledItem]) -> tuple[str, str | None]:
    installed_ver = find_installed_version(driver_name, category, installed_cache)
    if not installed_ver:
        return ("Not Installed", None)
    cmp_result = _compare_versions(installed_ver, available_version)
    if cmp_result is None:
        return ("Installed", installed_ver)
    if cmp_result > 0:
        return ("Update Available", installed_ver)
    if cmp_result == 0:
        return ("Up to Date", installed_ver)
    return ("Installed", installed_ver)


class CommandRunner(Protocol):
    def run(self, command: Sequence[str]) -> subprocess.CompletedProcess[str]:  # pragma: no cover - protocol
        ...


class SubprocessRunner:
    def run(self, command: Sequence[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.run(command, capture_output=True, text=True, check=False)


def _resolve_legacy_repo_root(root: str | Path | None) -> Path | None:
    if root is None:
        cleaned_default = IMMUTABLE_CONFIG.ids.hp_legacy_repo_root.strip()
        return Path(cleaned_default) if cleaned_default else None
    if isinstance(root, str):
        cleaned = root.strip()
        if not cleaned:
            return None
        return Path(cleaned)
    return Path(root)


def _download_file(url: str, destination: Path) -> None:
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_path = destination.with_suffix(destination.suffix + ".download")
    try:
        with urllib.request.urlopen(request, timeout=60) as response, temp_path.open("wb") as handle:
            shutil.copyfileobj(response, handle)
        temp_path.replace(destination)
    except urllib.error.URLError as exc:
        if temp_path.exists():
            temp_path.unlink()
        raise RuntimeError(f"Download failed for {url}: {exc}") from exc


class HPIAClient:
    def __init__(
        self,
        working_dir: Path,
        *,
        executable: str | None = None,
        download_url: str | None = None,
        command_runner: CommandRunner | None = None,
    ) -> None:
        self._working_dir = Path(working_dir)
        self._runner = command_runner or SubprocessRunner()
        self._executable = Path(executable) if executable else self._auto_detect()
        self._download_dir = self._working_dir / "hpia_softpaqs"
        self._report_dir = self._working_dir / "hpia_reports"
        self._download_url = download_url or os.getenv(
            "HPIA_DOWNLOAD_URL",
            "https://hpia.hpcloud.hp.com/downloads/hpia/hp-hpia-5.2.1.exe",
        )

    def is_available(self) -> bool:
        return self._executable is not None and self._executable.exists()

    def ensure_available(self) -> bool:
        if self.is_available():
            return True
        hpia_dir = self._working_dir / "HPIA"
        hpia_dir.mkdir(parents=True, exist_ok=True)
        existing = next(hpia_dir.rglob("HPImageAssistant.exe"), None)
        if existing:
            self._executable = existing
            return True
        installer = self._try_winget_download(hpia_dir)
        if installer and installer.suffix.lower() != ".exe":
            installer = None
        if not installer:
            installer = hpia_dir / "hp-hpia-setup.exe"
            if not installer.exists():
                try:
                    _download_file(self._download_url, installer)
                except Exception:
                    if self._try_winget_install():
                        refreshed = self._auto_detect()
                        if refreshed and refreshed.exists():
                            self._executable = refreshed
                            return True
                    raise
        try:
            result = self._runner.run([str(installer), "/s", "/e", "/f", str(hpia_dir)])
            if result.returncode != 0:
                raise RuntimeError(f"HPIA extract failed: {result.stderr}")
        except Exception:
            if self._try_winget_install():
                refreshed = self._auto_detect()
                if refreshed and refreshed.exists():
                    self._executable = refreshed
                    return True
            raise
        extracted = next(hpia_dir.rglob("HPImageAssistant.exe"), None)
        if extracted:
            self._executable = extracted
            return True
        if self._try_winget_install():
            refreshed = self._auto_detect()
            if refreshed and refreshed.exists():
                self._executable = refreshed
                return True
        return False

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
        installed_cache = get_installed_drivers_and_software()
        records: list[DriverRecord] = []
        for rec in recommendations:
            rec_value = rec.get("RecommendationValue", "Optional")
            driver_name = rec.get("Name", "Unknown")
            category = rec.get("Category")
            available_ver = rec.get("Version")
            hpia_installed_ver = rec.get("CurrentVersion")
            status_result, detected_installed_ver = get_driver_status(driver_name, category, available_ver, installed_cache)
            final_installed = detected_installed_ver or hpia_installed_ver
            if rec_value == "Critical":
                status = "Critical"
            elif rec_value == "Recommended":
                status = "Recommended" if status_result in ("Not Installed", "Update Available") else status_result
            elif status_result == "Update Available":
                status = "Update Available"
            elif status_result == "Not Installed":
                status = "Optional"
            else:
                status = status_result
            records.append(
                DriverRecord(
                    name=driver_name,
                    status=status,
                    source="HPIA",
                    installed_version=final_installed,
                    latest_version=available_ver,
                    category=category,
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

    def _try_winget_download(self, target_dir: Path) -> Path | None:
        if shutil.which("winget") is None:
            return None
        args = [
            "winget",
            "download",
            "--id",
            "HP.ImageAssistant",
            "--source",
            "winget",
            "--accept-source-agreements",
            "--accept-package-agreements",
            "--download-directory",
            str(target_dir),
        ]
        result = self._runner.run(args)
        if result.returncode != 0:
            return None
        candidates = sorted(target_dir.glob("*.exe"), key=lambda p: p.stat().st_mtime, reverse=True)
        return candidates[0] if candidates else None

    def _try_winget_install(self) -> bool:
        if shutil.which("winget") is None:
            return False
        args = [
            "winget",
            "install",
            "--id",
            "HP.ImageAssistant",
            "--source",
            "winget",
            "--silent",
            "--accept-source-agreements",
            "--accept-package-agreements",
        ]
        result = self._runner.run(args)
        return result.returncode == 0

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
        installed_cache = get_installed_drivers_and_software()
        records: list[DriverRecord] = []
        if isinstance(data, dict):
            data = [data]
        for item in data or []:
            if not isinstance(item, dict):
                continue
            category = item.get("Category", "")
            if "driver" not in category.lower() and "bios" not in category.lower() and "firmware" not in category.lower():
                continue
            driver_name = item.get("Name", "Unknown")
            available_ver = item.get("Version")
            status_result, installed_ver = get_driver_status(driver_name, category, available_ver, installed_cache)
            records.append(
                DriverRecord(
                    name=driver_name,
                    status=status_result,
                    source="CMSL",
                    installed_version=installed_ver,
                    latest_version=available_ver,
                    category=category,
                    softpaq_id=item.get("Id") or item.get("SoftPaqId"),
                    download_url=item.get("Url"),
                )
            )
        return records

    def scan_catalog(self, platform_id: str | None) -> list[DriverRecord]:
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
        if isinstance(data, dict):
            data = [data]
        records: list[DriverRecord] = []
        for item in data or []:
            if not isinstance(item, dict):
                continue
            category = item.get("Category", "")
            if "driver" not in category.lower() and "bios" not in category.lower() and "firmware" not in category.lower():
                continue
            bucket = _categorize_cmsl(category, item.get("Name", ""))
            records.append(
                DriverRecord(
                    name=item.get("Name", "Unknown"),
                    status="Catalog",
                    source="CMSL",
                    installed_version=None,
                    latest_version=item.get("Version"),
                    category=bucket,
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
        self._root = _resolve_legacy_repo_root(root)
        self.last_match_detail: str | None = None

    def _normalize_folder_name(self, value: str) -> str:
        return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()

    def _score_candidate(self, folder_name: str, platform_id: str | None, model: str | None) -> int:
        score = 0
        folder_norm = self._normalize_folder_name(folder_name)
        folder_tokens = set(folder_norm.split())

        if platform_id:
            pid_norm = self._normalize_folder_name(platform_id)
            if pid_norm:
                if folder_norm == pid_norm:
                    score += 100
                elif pid_norm in folder_norm:
                    score += 60

        variants: list[str] = []
        if model:
            variants.append(model)
            variants.append(model.replace("HP ", "").replace("Hewlett-Packard ", ""))

        for variant in variants:
            model_norm = self._normalize_folder_name(variant)
            if not model_norm:
                continue
            if folder_norm == model_norm:
                score += 80
            elif model_norm in folder_norm:
                score += 50
            model_tokens = set(model_norm.split())
            if model_tokens and folder_tokens:
                score += len(folder_tokens & model_tokens) * 5

        return score

    def is_configured(self) -> bool:
        return self._root is not None

    def list_packages(self, platform_id: str | None, model: str | None) -> list[DriverRecord]:
        self.last_match_detail = None
        if self._root is None:
            return []
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
            installed_cache = get_installed_drivers_and_software()
            for item in data:
                file_name = item.get("File") or item.get("Path") or item.get("FileName")
                if not file_name:
                    continue
                file_path = candidate / file_name
                driver_name = item.get("Name", "Legacy Driver")
                category = item.get("Category")
                available_ver = item.get("Version")
                status_result, installed_ver = get_driver_status(driver_name, category, available_ver, installed_cache)
                if "bios" in category.lower() and status_result == "Update Available":
                    status_result = "Critical"
                records.append(
                    DriverRecord(
                        name=driver_name,
                        status=status_result,
                        source="Legacy",
                        installed_version=installed_ver,
                        latest_version=available_ver,
                        category=category,
                        softpaq_id=item.get("SoftPaqId"),
                        download_url=str(file_path),
                        output_path=file_path,
                    )
                )
            if records:
                break
        if records:
            return records

        # Fallback: scan all immediate subfolders for a manifest and pick the best match.
        matches: list[tuple[int, Path, list[DriverRecord]]] = []
        try:
            subdirs = [p for p in self._root.iterdir() if p.is_dir()]
        except OSError:
            subdirs = []
        for subdir in subdirs:
            manifest = subdir / "manifest.json"
            if not manifest.exists():
                continue
            try:
                data = json.loads(manifest.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                continue
            installed_cache = get_installed_drivers_and_software()
            sub_records: list[DriverRecord] = []
            for item in data:
                file_name = item.get("File") or item.get("Path") or item.get("FileName")
                if not file_name:
                    continue
                file_path = subdir / file_name
                driver_name = item.get("Name", "Legacy Driver")
                category = item.get("Category")
                available_ver = item.get("Version")
                status_result, installed_ver = get_driver_status(driver_name, category, available_ver, installed_cache)
                if "bios" in category.lower() and status_result == "Update Available":
                    status_result = "Critical"
                sub_records.append(
                    DriverRecord(
                        name=driver_name,
                        status=status_result,
                        source="Legacy",
                        installed_version=installed_ver,
                        latest_version=available_ver,
                        category=category,
                        softpaq_id=item.get("SoftPaqId"),
                        download_url=str(file_path),
                        output_path=file_path,
                    )
                )
            if not sub_records:
                continue
            score = self._score_candidate(subdir.name, platform_id, model)
            matches.append((score, subdir, sub_records))

        if not matches:
            return []
        matches.sort(key=lambda item: (item[0], item[1].name.lower()), reverse=True)
        best_score, best_dir, best_records = matches[0]
        if len(matches) > 1 and matches[1][0] == best_score:
            self.last_match_detail = (
                "Multiple legacy manifest folders matched equally; "
                f"using '{best_dir.name}'. Consider creating a platform ID folder."
            )
        else:
            self.last_match_detail = f"Legacy repo fallback selected '{best_dir.name}'."
        return best_records


class DriverService:
    def __init__(
        self,
        *,
        working_dir: Path | str | None = None,
        legacy_repo_root: str | Path | None = None,
        hpia_client: HPIAClient | None = None,
        cmsl_client: CMSLClient | None = None,
        legacy_repo: LegacyRepository | None = None,
        command_runner: CommandRunner | None = None,
        system_info_provider: Callable[[], HPSystemInfo] | None = None,
    ) -> None:
        self._working_dir = Path(working_dir) if working_dir is not None else get_application_directory()
        self._runner = command_runner or SubprocessRunner()
        self._hpia = hpia_client or HPIAClient(self._working_dir)
        self._cmsl = cmsl_client or CMSLClient()
        self._legacy = legacy_repo or LegacyRepository(legacy_repo_root)
        self._system_info_provider = system_info_provider or get_hp_system_info
        self.last_scan_warnings: list[str] = []
        self.last_system_info: HPSystemInfo | None = None

    def scan(self) -> list[DriverRecord]:
        info = self._system_info_provider()
        self.last_system_info = info
        self.last_scan_warnings = []
        records: list[DriverRecord] = []
        hpia_ready = self._hpia.is_available()
        attempted_hpia = False
        auto_download_failed = False
        if not hpia_ready and (info.supports_hpia or info.manufacturer or info.model or info.platform_id):
            attempted_hpia = True
            try:
                hpia_ready = self._hpia.ensure_available()
            except Exception as exc:
                auto_download_failed = True
                self.last_scan_warnings.append(f"HPIA auto-download failed: {exc}")
        if hpia_ready:
            try:
                records.extend(self._hpia.scan())
            except Exception as exc:
                self.last_scan_warnings.append(f"HPIA scan failed: {exc}")
        elif info.supports_hpia or info.manufacturer or info.model or info.platform_id:
            if not auto_download_failed:
                message = "HPIA not found after auto-download attempt."
                if not attempted_hpia:
                    message = "HPIA not found. Install HP Image Assistant or place HPImageAssistant.exe in the working directory."
                self.last_scan_warnings.append(message)
        if info.supports_cmsl:
            if self._cmsl.is_available():
                try:
                    records.extend(self._cmsl.scan(info.platform_id))
                except Exception as exc:
                    self.last_scan_warnings.append(f"CMSL scan failed: {exc}")
            else:
                self.last_scan_warnings.append("CMSL not available. Install the HPCMSL PowerShell module.")
        if not records and info.supports_legacy_repo:
            records.extend(self._legacy.list_packages(info.platform_id, info.model))
        return records

    def scan_hpia(self) -> list[DriverRecord]:
        info = self._system_info_provider()
        self.last_system_info = info
        self.last_scan_warnings = []
        records: list[DriverRecord] = []
        hpia_ready = self._hpia.is_available()
        attempted_hpia = False
        auto_download_failed = False
        if not hpia_ready and (info.supports_hpia or info.manufacturer or info.model or info.platform_id):
            attempted_hpia = True
            try:
                hpia_ready = self._hpia.ensure_available()
            except Exception as exc:
                auto_download_failed = True
                self.last_scan_warnings.append(f"HPIA auto-download failed: {exc}")
        if hpia_ready:
            try:
                records = self._hpia.scan()
            except Exception as exc:
                self.last_scan_warnings.append(f"HPIA scan failed: {exc}")
        else:
            if not auto_download_failed:
                message = "HPIA not found after auto-download attempt."
                if not attempted_hpia:
                    message = "HPIA not found. Install HP Image Assistant or place HPImageAssistant.exe in the working directory."
                self.last_scan_warnings.append(message)
        return records

    def scan_cmsl_catalog(self) -> list[DriverRecord]:
        info = self._system_info_provider()
        self.last_system_info = info
        self.last_scan_warnings = []
        if not info.supports_cmsl:
            self.last_scan_warnings.append("CMSL not supported on this system/OS.")
            return []
        if not self._cmsl.is_available():
            self.last_scan_warnings.append("CMSL not available. Install the HPCMSL PowerShell module.")
            return []
        try:
            records = self._cmsl.scan_catalog(info.platform_id)
        except Exception as exc:
            self.last_scan_warnings.append(f"CMSL scan failed: {exc}")
            return []
        deduped = _dedupe_latest_records(records)
        return sorted(deduped, key=lambda r: (r.category or "", r.name))

    def scan_legacy(self) -> list[DriverRecord]:
        info = self._system_info_provider()
        self.last_system_info = info
        self.last_scan_warnings = []
        if not self._legacy.is_configured():
            self.last_scan_warnings.append("Legacy repository root not configured. Set it in Drivers -> Legacy -> Settings.")
            return []
        if not info.supports_legacy_repo:
            self.last_scan_warnings.append("Legacy repository not supported on this system.")
            return []
        records = self._legacy.list_packages(info.platform_id, info.model)
        if self._legacy.last_match_detail:
            self.last_scan_warnings.append(self._legacy.last_match_detail)
        if not records:
            self.last_scan_warnings.append("No legacy repository manifest found for this device.")
        return records

    def download(
        self,
        records: Iterable[DriverRecord],
        *,
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> list[DriverOperationResult]:
        ops: list[DriverOperationResult] = []
        record_list = list(records)
        total = len(record_list)
        completed = 0

        def _emit(message: str) -> None:
            nonlocal completed
            completed += 1
            if progress_callback:
                progress_callback(completed, total, message)

        hpia_targets = [r for r in record_list if r.source == "HPIA" and r.softpaq_id]
        cmsl_targets = [r for r in record_list if r.source == "CMSL" and r.softpaq_id]
        legacy_targets = [r for r in record_list if r.source == "Legacy" and r.download_url]

        if hpia_targets:
            try:
                if progress_callback:
                    progress_callback(completed, total, f"Downloading {len(hpia_targets)} HPIA package(s)...")
                mapping = self._hpia.download([r.softpaq_id for r in hpia_targets if r.softpaq_id])
                for record in hpia_targets:
                    record.output_path = mapping.get(record.softpaq_id)
                    success = record.output_path is not None
                    ops.append(DriverOperationResult(record, "download", success, "Downloaded" if success else "Missing output"))
                    _emit(f"Downloaded: {record.name}" if success else f"Failed: {record.name}")
            except Exception as exc:
                for record in hpia_targets:
                    ops.append(DriverOperationResult(record, "download", False, str(exc)))
                    _emit(f"Failed: {record.name}")

        for record in cmsl_targets:
            try:
                dest = self._working_dir / "cmsl_softpaqs" / f"{record.softpaq_id}.exe"
                record.output_path = self._cmsl.download(record.softpaq_id or "", dest)
                ops.append(DriverOperationResult(record, "download", True, "Downloaded"))
                _emit(f"Downloaded: {record.name}")
            except Exception as exc:
                ops.append(DriverOperationResult(record, "download", False, str(exc)))
                _emit(f"Failed: {record.name}")

        for record in legacy_targets:
            try:
                src = Path(record.download_url or "")
                dest_dir = self._working_dir / "legacy_drivers"
                dest_dir.mkdir(parents=True, exist_ok=True)
                dest = dest_dir / src.name
                shutil.copy2(src, dest)
                record.output_path = dest  # type: ignore[assignment]
                ops.append(DriverOperationResult(record, "download", True, "Copied"))
                _emit(f"Copied: {record.name}")
            except Exception as exc:
                ops.append(DriverOperationResult(record, "download", False, str(exc)))
                _emit(f"Failed: {record.name}")

        return ops

    def install(
        self,
        records: Iterable[DriverRecord],
        *,
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> list[DriverOperationResult]:
        ops: list[DriverOperationResult] = []
        record_list = list(records)
        total = len(record_list)
        completed = 0

        def _emit(message: str) -> None:
            nonlocal completed
            completed += 1
            if progress_callback:
                progress_callback(completed, total, message)

        for record in record_list:
            if not record.output_path:
                ops.append(DriverOperationResult(record, "install", False, "No installer downloaded"))
                _emit(f"Skipped: {record.name}")
                continue
            cmd = [str(record.output_path), "/s"]
            result = self._runner.run(cmd)
            success = result.returncode in {0, 3010}
            message = "Installed" if success else f"Installer exit {result.returncode}"
            ops.append(DriverOperationResult(record, "install", success, message))
            _emit(f"{message}: {record.name}")
        return ops
