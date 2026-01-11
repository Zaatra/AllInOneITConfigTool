"""System configuration logic (timezone, locale, power, icons)."""
from __future__ import annotations

import platform
import shlex
import subprocess
from dataclasses import dataclass
from typing import Iterable, Protocol, Sequence

from trustpal.constants import FixedSystemConfig

try:  # Windows-only dependency, optional for test doubles
    import winreg  # type: ignore
except ImportError:  # pragma: no cover - not available on Linux runners
    winreg = None  # type: ignore


@dataclass
class ConfigCheckResult:
    name: str
    expected: str
    actual: str
    in_desired_state: bool


class CommandRunner(Protocol):
    def run(self, command: Sequence[str]) -> subprocess.CompletedProcess[str]:  # pragma: no cover - protocol
        ...


class SubprocessRunner:
    def run(self, command: Sequence[str]) -> subprocess.CompletedProcess[str]:
        return subprocess.run(command, capture_output=True, text=True, check=False)


class RegistryAccessor(Protocol):
    def get_value(self, path: str, value_name: str) -> str | int | None:  # pragma: no cover - protocol
        ...

    def set_value(self, path: str, value_name: str, value: str | int) -> None:  # pragma: no cover - protocol
        ...


class WindowsRegistryAccessor:
    """Minimal registry helper backed by winreg."""

    def __init__(self) -> None:
        if winreg is None:
            raise RuntimeError("winreg not available on this platform")

    def get_value(self, path: str, value_name: str) -> str | int | None:
        hive, subkey = self._split_path(path)
        try:
            with winreg.OpenKey(hive, subkey) as key:  # type: ignore[arg-type]
                value, _ = winreg.QueryValueEx(key, value_name)
                return value
        except FileNotFoundError:
            return None

    def set_value(self, path: str, value_name: str, value: str | int) -> None:
        hive, subkey = self._split_path(path)
        value_type = winreg.REG_DWORD if isinstance(value, int) else winreg.REG_SZ
        with winreg.CreateKeyEx(hive, subkey) as key:  # type: ignore[arg-type]
            winreg.SetValueEx(key, value_name, 0, value_type, value)

    def _split_path(self, path: str) -> tuple[object, str]:
        cleaned = path.replace("/", "\\")
        marker = ":\\"
        if marker not in cleaned:
            raise ValueError(f"Invalid registry path: {path}")
        hive_name, subkey = cleaned.split(marker, 1)
        subkey = subkey.lstrip("\\")
        hive_map = {
            "HKLM": winreg.HKEY_LOCAL_MACHINE,
            "HKCU": winreg.HKEY_CURRENT_USER,
            "HKCR": winreg.HKEY_CLASSES_ROOT,
            "HKU": winreg.HKEY_USERS,
            "HKCC": winreg.HKEY_CURRENT_CONFIG,
        }
        try:
            hive = hive_map[hive_name.upper()]
        except KeyError as exc:  # pragma: no cover - invalid input handled upstream
            raise ValueError(f"Unsupported hive: {hive_name}") from exc
        return hive, subkey


class SystemConfigService:
    def __init__(
        self,
        config: FixedSystemConfig,
        *,
        command_runner: CommandRunner | None = None,
        registry: RegistryAccessor | None = None,
    ) -> None:
        self._config = config
        self._runner = command_runner or SubprocessRunner()
        self._registry = registry or WindowsRegistryAccessor()

    def check(self) -> list[ConfigCheckResult]:
        results = [
            self._check_timezone(),
            self._check_power_plan(),
            self._check_fast_boot(),
            self._check_desktop_icons(),
            self._check_locale(),
        ]
        return results

    def apply(self) -> None:
        self._set_timezone()
        self._set_power_plan()
        self._registry.set_value(
            self._config.fast_boot.path,
            self._config.fast_boot.value_name,
            int(self._config.fast_boot.desired_value),
        )
        self._registry.set_value(
            self._config.desktop_icons.path,
            self._config.desktop_icons.value_name,
            int(self._config.desktop_icons.desired_value),
        )
        self._set_locale()
        self._registry.set_value(
            r"HKCU:\Control Panel\International",
            "sShortDate",
            self._config.locale.short_date_format,
        )

    def _check_timezone(self) -> ConfigCheckResult:
        expected = self._config.timezone
        actual = self._run_and_capture(["tzutil", "/g"])
        return ConfigCheckResult("Timezone", expected, actual, actual == expected)

    def _check_power_plan(self) -> ConfigCheckResult:
        expected = self._config.power_plan.friendly_name
        output = self._run_and_capture(["powercfg", "/getactivescheme"])
        actual = self._extract_power_scheme_name(output)
        return ConfigCheckResult("Power Plan", expected, actual, expected.lower() in actual.lower())

    def _check_fast_boot(self) -> ConfigCheckResult:
        expected_value = int(self._config.fast_boot.desired_value)
        actual_value = self._registry.get_value(
            self._config.fast_boot.path,
            self._config.fast_boot.value_name,
        )
        actual_str = "Not Set" if actual_value is None else str(actual_value)
        return ConfigCheckResult("Fast Boot", str(expected_value), actual_str, actual_value == expected_value)

    def _check_desktop_icons(self) -> ConfigCheckResult:
        expected_value = int(self._config.desktop_icons.desired_value)
        actual_value = self._registry.get_value(
            self._config.desktop_icons.path,
            self._config.desktop_icons.value_name,
        )
        actual_str = "Not Set" if actual_value is None else str(actual_value)
        return ConfigCheckResult("Desktop Icons", str(expected_value), actual_str, actual_value == expected_value)

    def _check_locale(self) -> ConfigCheckResult:
        expected = self._config.locale.system_locale
        actual = self._run_and_capture(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "Get-WinSystemLocale | Select-Object -ExpandProperty Name",
            ]
        )
        ok = expected.lower() == actual.lower()
        if ok:
            date_val = self._registry.get_value(r"HKCU:\Control Panel\International", "sShortDate") or ""
            ok = ok and str(date_val).lower() == self._config.locale.short_date_format.lower()
            actual = f"{actual} / {date_val}"
        return ConfigCheckResult("Locale", f"{expected} / {self._config.locale.short_date_format}", actual, ok)

    def _set_timezone(self) -> None:
        self._runner.run(["tzutil", "/s", self._config.timezone])

    def _set_power_plan(self) -> None:
        self._runner.run(["powercfg", "/setactive", self._config.power_plan.scheme])

    def _set_locale(self) -> None:
        command = f"Set-WinSystemLocale -SystemLocale {shlex.quote(self._config.locale.system_locale)}"
        self._runner.run(["powershell", "-NoProfile", "-Command", command])

    def _extract_power_scheme_name(self, output: str) -> str:
        if "(" in output and ")" in output:
            return output.split("(")[-1].split(")")[0].strip()
        return output.strip()

    def _run_and_capture(self, command: Sequence[str]) -> str:
        completed = self._runner.run(command)
        if completed.stderr and not completed.stdout:
            return completed.stderr.strip()
        return completed.stdout.strip()
