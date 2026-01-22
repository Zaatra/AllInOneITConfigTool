"""System configuration logic (timezone, locale, power, icons)."""
from __future__ import annotations

import shlex
import subprocess
from dataclasses import dataclass
from typing import Callable, Iterable, Protocol, Sequence, TypeVar

from allinone_it_config.constants import FixedSystemConfig

try:  # Windows-only dependency, optional for test doubles
    import winreg  # type: ignore
except ImportError:  # pragma: no cover - not available on Linux runners
    winreg = None  # type: ignore

DEFAULT_USER_HIVE_KEY = "AIO_DefaultUser"
DEFAULT_USER_HIVE_PATH = r"C:\Users\Default\NTUSER.DAT"
HKCU_PREFIX = "HKCU:\\"

T = TypeVar("T")


@dataclass
class ConfigCheckResult:
    name: str
    expected: str
    actual: str
    in_desired_state: bool


@dataclass
class ApplyStepResult:
    name: str
    success: bool
    detail: str = ""


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
            self._check_default_user_profile(),
        ]
        return results

    def apply(self) -> None:
        self.apply_with_results()

    def apply_with_results(self) -> list[ApplyStepResult]:
        results = [
            self._apply_timezone(),
            self._apply_power_plan(),
            self._apply_fast_boot(),
            self._apply_locale(),
            self._apply_user_profile_settings(),
        ]
        return results

    def _apply_timezone(self) -> ApplyStepResult:
        expected = self._config.timezone
        completed = self._runner.run(["tzutil", "/s", expected])
        detail = self._format_command_detail(completed)
        actual = self._run_and_capture(["tzutil", "/g"])
        if actual:
            detail = f"{detail}; current: {actual}"
        success = completed.returncode == 0 and (not actual or actual == expected)
        return ApplyStepResult("Timezone", success, detail)

    def _apply_power_plan(self) -> ApplyStepResult:
        expected = self._config.power_plan.friendly_name
        completed = self._runner.run(["powercfg", "/setactive", self._config.power_plan.scheme])
        detail = self._format_command_detail(completed)
        active_output = self._run_and_capture(["powercfg", "/getactivescheme"])
        active = self._extract_power_scheme_name(active_output)
        if active:
            detail = f"{detail}; active: {active}"
        success = completed.returncode == 0 and (not active or expected.lower() in active.lower())
        return ApplyStepResult("Power Plan", success, detail)

    def _apply_fast_boot(self) -> ApplyStepResult:
        try:
            desired = int(self._config.fast_boot.desired_value)
            self._registry.set_value(
                self._config.fast_boot.path,
                self._config.fast_boot.value_name,
                desired,
            )
            actual = self._registry.get_value(self._config.fast_boot.path, self._config.fast_boot.value_name)
            detail = f"set to {desired}; current: {actual}"
            return ApplyStepResult("Fast Boot", actual == desired, detail)
        except Exception as exc:  # pragma: no cover - surfaced via UI logging
            return ApplyStepResult("Fast Boot", False, str(exc))

    def _apply_locale(self) -> ApplyStepResult:
        command = f"Set-WinSystemLocale -SystemLocale {shlex.quote(self._config.locale.system_locale)}"
        completed = self._runner.run(["powershell", "-NoProfile", "-Command", command])
        detail = self._format_command_detail(completed)
        actual_locale = self._run_and_capture(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "Get-WinSystemLocale | Select-Object -ExpandProperty Name",
            ]
        )
        date_val = self._registry.get_value(r"HKCU:\Control Panel\International", "sShortDate") or ""
        if actual_locale:
            detail = f"{detail}; current: {actual_locale} / {date_val}"
        success = completed.returncode == 0 and (
            not actual_locale or actual_locale.lower() == self._config.locale.system_locale.lower()
        )
        return ApplyStepResult("Locale", success, detail)

    def _apply_user_profile_settings(self) -> ApplyStepResult:
        try:
            self._apply_user_profile_settings_inner()
        except Exception as exc:  # pragma: no cover - surfaced via UI logging
            return ApplyStepResult("Default User Profile", False, str(exc))
        result = self._check_default_user_profile()
        return ApplyStepResult("Default User Profile", result.in_desired_state, result.actual)

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

    def _check_default_user_profile(self) -> ConfigCheckResult:
        expected_hide = int(self._config.desktop_icons.desired_value)
        expected_date = self._config.locale.short_date_format
        expected = f"HideIcons={expected_hide}, sShortDate={expected_date}"

        def check(root: str) -> ConfigCheckResult:
            hide_path = self._map_user_path(self._config.desktop_icons.path, root)
            date_path = self._map_user_path(r"HKCU:\Control Panel\International", root)
            hide_val = self._registry.get_value(hide_path, self._config.desktop_icons.value_name)
            date_val = self._registry.get_value(date_path, "sShortDate")
            hide_str = "Not Set" if hide_val is None else str(hide_val)
            date_str = "Not Set" if date_val is None else str(date_val)
            actual = f"HideIcons={hide_str}, sShortDate={date_str}"
            ok = hide_val == expected_hide and str(date_val).lower() == expected_date.lower()
            return ConfigCheckResult("Default User Profile", expected, actual, ok)

        try:
            return self._with_default_user_hive(check)
        except RuntimeError as exc:
            return ConfigCheckResult("Default User Profile", expected, str(exc), False)

    def _apply_user_profile_settings_inner(self) -> None:
        self._registry.set_value(
            self._config.desktop_icons.path,
            self._config.desktop_icons.value_name,
            int(self._config.desktop_icons.desired_value),
        )
        self._registry.set_value(
            r"HKCU:\Control Panel\International",
            "sShortDate",
            self._config.locale.short_date_format,
        )

        def apply_to_default(root: str) -> None:
            self._registry.set_value(
                self._map_user_path(self._config.desktop_icons.path, root),
                self._config.desktop_icons.value_name,
                int(self._config.desktop_icons.desired_value),
            )
            self._registry.set_value(
                self._map_user_path(r"HKCU:\Control Panel\International", root),
                "sShortDate",
                self._config.locale.short_date_format,
            )

        self._with_default_user_hive(apply_to_default)

    def _format_command_detail(self, completed: subprocess.CompletedProcess[str]) -> str:
        detail_parts = [f"exit={completed.returncode}"]
        stdout = (completed.stdout or "").strip()
        stderr = (completed.stderr or "").strip()
        if stdout:
            detail_parts.append(f"stdout: {stdout}")
        if stderr:
            detail_parts.append(f"stderr: {stderr}")
        return ", ".join(detail_parts)

    def _extract_power_scheme_name(self, output: str) -> str:
        if "(" in output and ")" in output:
            return output.split("(")[-1].split(")")[0].strip()
        return output.strip()

    def _run_and_capture(self, command: Sequence[str]) -> str:
        completed = self._runner.run(command)
        if completed.stderr and not completed.stdout:
            return completed.stderr.strip()
        return completed.stdout.strip()

    def _run_and_check(self, command: Sequence[str], step: str) -> None:
        completed = self._runner.run(command)
        if completed.returncode == 0:
            return
        detail = (completed.stderr or completed.stdout or "").strip()
        if detail:
            raise RuntimeError(f"{step} failed: {detail}")
        raise RuntimeError(f"{step} failed with exit code {completed.returncode}")

    def _with_default_user_hive(self, action: Callable[[str], T]) -> T:
        load = self._runner.run(["reg", "load", fr"HKU\{DEFAULT_USER_HIVE_KEY}", DEFAULT_USER_HIVE_PATH])
        if load.returncode != 0:
            detail = (load.stderr or load.stdout or "").strip() or "Unknown error"
            raise RuntimeError(f"Default user profile load failed: {detail}")
        try:
            return action(fr"HKU:\{DEFAULT_USER_HIVE_KEY}")
        finally:
            self._runner.run(["reg", "unload", fr"HKU\{DEFAULT_USER_HIVE_KEY}"])

    def _map_user_path(self, path: str, root: str) -> str:
        if not path.upper().startswith(HKCU_PREFIX):
            raise ValueError(f"Expected HKCU path, got: {path}")
        suffix = path[len(HKCU_PREFIX) :]
        return f"{root}\\{suffix}"
