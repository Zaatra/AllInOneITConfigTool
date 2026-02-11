"""System configuration logic (timezone, locale, power, icons)."""
from __future__ import annotations

import re
import shlex
import subprocess
import tempfile
import time
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Protocol, Sequence, TypeVar

from allinone_it_config.constants import FixedSystemConfig

try:  # Windows-only dependency, optional for test doubles
    import winreg  # type: ignore
except ImportError:  # pragma: no cover - not available on Linux runners
    winreg = None  # type: ignore

DEFAULT_USER_HIVE_KEY = "AIO_DefaultUser"
DEFAULT_USER_HIVE_PATH = r"C:\Users\Default\NTUSER.DAT"
HKCU_PREFIX = "HKCU:\\"
POWERCFG_GUID_PATTERN = re.compile(r"Power Scheme GUID:\s*([0-9a-fA-F-]{36})\s*\((.*?)\)\s*(\*)?")
KNOWN_POWER_SCHEMES = {
    "SCHEME_BALANCED": "381b4222-f694-41f0-9685-ff5bb260df2e",
    "SCHEME_MIN": "a1841308-3541-4fab-bc81-f71556f20b4a",
    "SCHEME_MAX": "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c",
}
DESKTOP_POLICY_PATH = r"HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
DESKTOP_POLICY_VALUE = "NoDesktop"
DESKTOP_ICON_VISIBILITY_PATHS = (
    r"HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel",
    r"HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu",
)
DESKTOP_ICON_GUIDS = (
    "{20D04FE0-3AEA-1069-A2D8-08002B30309D}",  # This PC
    "{59031a47-3f72-44a7-89c5-5595fe6b30ee}",  # User files
    "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}",  # Network
    "{645FF040-5081-101B-9F08-00AA002F954E}",  # Recycle Bin
    "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}",  # Control Panel
)
DEFAULT_APPS_POLICY_PATH = r"HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
DEFAULT_APPS_POLICY_VALUE = "DefaultAssociationsConfiguration"
DEFAULT_APPS_XML_PATH_WINDOWS = r"C:\ProgramData\AllInOneITConfigTool\DefaultAppAssociations.xml"
DEFAULT_APP_ASSOCIATIONS = (
    (".htm", "ChromeHTML", "Google Chrome"),
    (".html", "ChromeHTML", "Google Chrome"),
    (".mhtml", "ChromeHTML", "Google Chrome"),
    (".pdf", "ChromePDF", "Google Chrome"),
    (".svg", "ChromeHTML", "Google Chrome"),
    (".xht", "ChromeHTML", "Google Chrome"),
    (".xhtml", "ChromeHTML", "Google Chrome"),
    ("ftp", "ChromeHTML", "Google Chrome"),
    ("http", "ChromeHTML", "Google Chrome"),
    ("https", "ChromeHTML", "Google Chrome"),
    ("mailto", "Outlook.URL.mailto.15", "Outlook (classic)"),
)
LANGUAGE_CAPABILITY_PREFIXES = {
    "en-US": ("Language.Basic", "Language.Handwriting", "Language.TextToSpeech", "Language.Speech"),
    "ar-SA": ("Language.Basic", "Language.Handwriting", "Language.TextToSpeech", "Language.Speech"),
}

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
            self._check_default_apps(),
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
            self._apply_default_apps(),
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
        schemes = self._list_power_schemes()
        target_guid, target_name = self._resolve_power_scheme(schemes)
        target = target_guid or self._config.power_plan.scheme
        completed = self._runner.run(["powercfg", "/setactive", target])
        detail = self._format_command_detail(completed)
        if schemes:
            schemes_summary = ", ".join(
                f"{name}={guid}{'*' if active else ''}" for guid, name, active in schemes
            )
            detail = f"{detail}; schemes: {schemes_summary}"
        if target_guid:
            detail = f"{detail}; target: {target_guid}"
        elif target_name:
            detail = f"{detail}; target: {target_name}"
        else:
            detail = f"{detail}; target: {target}"
        active_guid, active_name = self._wait_for_active_scheme(target_guid)
        if active_name or active_guid:
            active_label = active_name or ""
            if active_guid:
                active_label = f"{active_label} ({active_guid})".strip()
            detail = f"{detail}; active: {active_label}"
        if target_guid and active_guid:
            success = completed.returncode == 0 and active_guid.lower() == target_guid.lower()
        elif target_name:
            success = completed.returncode == 0 and target_name.lower() in (active_name or "").lower()
        else:
            success = completed.returncode == 0 and expected.lower() in (active_name or "").lower()
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
        detail_parts: list[str] = []
        command = f"Set-WinSystemLocale -SystemLocale {shlex.quote(self._config.locale.system_locale)}"
        completed = self._runner.run(["powershell", "-NoProfile", "-Command", command])
        detail_parts.append(f"system locale: {self._format_command_detail(completed)}")
        success = completed.returncode == 0

        feature_success, feature_detail = self._apply_language_packs_and_features()
        success = success and feature_success
        if feature_detail:
            detail_parts.append(feature_detail)

        ui_detail = self._apply_ui_languages()
        if ui_detail:
            detail_parts.append(f"ui languages: {ui_detail}")

        primary_language = self._primary_ui_language()
        culture_script = "; ".join(
            [
                f"Set-WinUILanguageOverride -Language {shlex.quote(primary_language)}",
                f"Set-Culture -CultureInfo {shlex.quote(primary_language)}",
                "Set-ItemProperty -Path 'HKCU:\\Control Panel\\International' -Name 'iDate' -Value '1'",
                "Set-ItemProperty -Path 'HKCU:\\Control Panel\\International' -Name 'sDate' -Value '/'",
                f"Set-ItemProperty -Path 'HKCU:\\Control Panel\\International' -Name 'sShortDate' -Value {shlex.quote(self._config.locale.short_date_format)}",
            ]
        )
        culture_completed = self._runner.run(["powershell", "-NoProfile", "-Command", culture_script])
        detail_parts.append(f"culture: {self._format_command_detail(culture_completed)}")
        success = success and culture_completed.returncode == 0

        date_error = None
        try:
            self._registry.set_value(
                r"HKCU:\Control Panel\International",
                "sShortDate",
                self._config.locale.short_date_format,
            )
            self._registry.set_value(r"HKCU:\Control Panel\International", "sDate", "/")
            self._registry.set_value(r"HKCU:\Control Panel\International", "iDate", "1")
        except Exception as exc:  # pragma: no cover - surfaced via UI logging
            date_error = str(exc)
            detail_parts.append(f"date: {date_error}")

        speech_detail = self._apply_speech_preferences()
        if speech_detail:
            detail_parts.append(speech_detail)
            success = False

        actual_locale = self._wait_for_system_locale(self._config.locale.system_locale)
        date_val = self._registry.get_value(r"HKCU:\Control Panel\International", "sShortDate") or ""
        display_language = self._run_and_capture(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "(Get-WinUILanguageOverride).Name",
            ]
        )
        current_parts = [actual_locale, str(date_val)]
        if display_language:
            current_parts.append(f"display={display_language}")
        detail_parts.append(f"current: {' / '.join(part for part in current_parts if part)}")

        if actual_locale:
            success = success and actual_locale.lower() == self._config.locale.system_locale.lower()
        if date_val:
            success = success and str(date_val).lower() == self._config.locale.short_date_format.lower()
        if display_language:
            success = success and display_language.lower() == primary_language.lower()
        if date_error:
            success = False
        return ApplyStepResult("Locale", success, "; ".join(part for part in detail_parts if part))

    def _apply_user_profile_settings(self) -> ApplyStepResult:
        try:
            self._apply_user_profile_settings_inner()
        except Exception as exc:  # pragma: no cover - surfaced via UI logging
            return ApplyStepResult("Default User Profile", False, str(exc))
        result = self._check_default_user_profile()
        refresh = self._refresh_desktop_shell()
        detail = result.actual
        if refresh:
            detail = f"{detail}; explorer refresh: {refresh}"
        return ApplyStepResult("Default User Profile", result.in_desired_state, detail)

    def _check_timezone(self) -> ConfigCheckResult:
        expected = self._config.timezone
        actual = self._run_and_capture(["tzutil", "/g"])
        return ConfigCheckResult("Timezone", expected, actual, actual == expected)

    def _check_power_plan(self) -> ConfigCheckResult:
        expected = self._config.power_plan.friendly_name
        active_guid, active_name = self._get_active_power_scheme()
        schemes = self._list_power_schemes()
        target_guid, target_name = self._resolve_power_scheme(schemes)
        actual = active_name or ""
        if active_guid and active_name:
            actual = f"{active_name} ({active_guid})"
        if target_guid and active_guid:
            ok = active_guid.lower() == target_guid.lower()
        elif target_name:
            ok = target_name.lower() in (active_name or "").lower()
        else:
            ok = expected.lower() in (active_name or "").lower()
        return ConfigCheckResult("Power Plan", expected, actual, ok)

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
        no_desktop = self._registry.get_value(DESKTOP_POLICY_PATH, DESKTOP_POLICY_VALUE)
        actual_str = f"HideIcons={actual_value if actual_value is not None else 'Not Set'}, NoDesktop={no_desktop if no_desktop is not None else 'Not Set'}"
        policy_ok = True
        if no_desktop is not None:
            try:
                policy_ok = int(no_desktop) == 0
            except (TypeError, ValueError):
                policy_ok = False
        ok = actual_value == expected_value and policy_ok
        return ConfigCheckResult("Desktop Icons", f"HideIcons={expected_value}, NoDesktop=0", actual_str, ok)

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
        date_val = self._registry.get_value(r"HKCU:\Control Panel\International", "sShortDate") or ""
        display_language = self._run_and_capture(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "(Get-WinUILanguageOverride).Name",
            ]
        )
        actual_display = f"{actual} / {date_val}" if date_val else actual
        if display_language:
            actual_display = f"{actual_display} / display={display_language}"
        ok = expected.lower() == actual.lower()
        if date_val:
            ok = ok and str(date_val).lower() == self._config.locale.short_date_format.lower()
        if display_language:
            ok = ok and display_language.lower() == self._primary_ui_language().lower()
        expected_display = f"{expected} / {self._config.locale.short_date_format} / display={self._primary_ui_language()}"
        return ConfigCheckResult("Locale", expected_display, actual_display, ok)

    def _check_default_user_profile(self) -> ConfigCheckResult:
        expected_hide = int(self._config.desktop_icons.desired_value)
        expected_date = self._config.locale.short_date_format
        expected = f"HideIcons={expected_hide}, NoDesktop=0, sShortDate={expected_date}"

        def check(root: str) -> ConfigCheckResult:
            hide_path = self._map_user_path(self._config.desktop_icons.path, root)
            policy_path = self._map_user_path(DESKTOP_POLICY_PATH, root)
            date_path = self._map_user_path(r"HKCU:\Control Panel\International", root)
            hide_val = self._registry.get_value(hide_path, self._config.desktop_icons.value_name)
            no_desktop = self._registry.get_value(policy_path, DESKTOP_POLICY_VALUE)
            date_val = self._registry.get_value(date_path, "sShortDate")
            hide_str = "Not Set" if hide_val is None else str(hide_val)
            policy_str = "Not Set" if no_desktop is None else str(no_desktop)
            date_str = "Not Set" if date_val is None else str(date_val)
            actual = f"HideIcons={hide_str}, NoDesktop={policy_str}, sShortDate={date_str}"
            icons_ok = self._check_desktop_icon_registry_values(root)
            policy_ok = True
            if no_desktop is not None:
                try:
                    policy_ok = int(no_desktop) == 0
                except (TypeError, ValueError):
                    policy_ok = False
            ok = (
                hide_val == expected_hide
                and policy_ok
                and str(date_val).lower() == expected_date.lower()
                and icons_ok
            )
            return ConfigCheckResult("Default User Profile", expected, actual, ok)

        try:
            return self._with_default_user_hive(check)
        except RuntimeError as exc:
            return ConfigCheckResult("Default User Profile", expected, str(exc), False)

    def _check_default_apps(self) -> ConfigCheckResult:
        expected_path = str(self._default_apps_xml_path())
        expected = "Chrome defaults for web/file types + MAILTO mapped to Outlook classic"
        actual = self._registry.get_value(DEFAULT_APPS_POLICY_PATH, DEFAULT_APPS_POLICY_VALUE)
        actual_str = "" if actual is None else str(actual)
        ok = actual_str.strip().lower() == expected_path.lower()
        return ConfigCheckResult("Default Apps", expected, actual_str or "Not Set", ok)

    def _apply_user_profile_settings_inner(self) -> None:
        desired = int(self._config.desktop_icons.desired_value)

        def apply_to_root(root: str | None) -> None:
            map_path = (lambda value: value) if root is None else (lambda value: self._map_user_path(value, root))
            self._registry.set_value(
                map_path(self._config.desktop_icons.path),
                self._config.desktop_icons.value_name,
                desired,
            )
            self._registry.set_value(map_path(DESKTOP_POLICY_PATH), DESKTOP_POLICY_VALUE, 0)
            self._registry.set_value(
                map_path(r"HKCU:\Control Panel\International"),
                "sShortDate",
                self._config.locale.short_date_format,
            )
            self._registry.set_value(map_path(r"HKCU:\Control Panel\International"), "sDate", "/")
            self._registry.set_value(map_path(r"HKCU:\Control Panel\International"), "iDate", "1")
            self._set_desktop_icon_registry_values(map_path)

        apply_to_root(None)
        self._with_default_user_hive(lambda root: apply_to_root(root))

    def _apply_default_apps(self) -> ApplyStepResult:
        try:
            path = self._write_default_apps_association_file()
            self._registry.set_value(DEFAULT_APPS_POLICY_PATH, DEFAULT_APPS_POLICY_VALUE, str(path))
        except Exception as exc:  # pragma: no cover - surfaced via UI logging
            return ApplyStepResult("Default Apps", False, str(exc))
        command = ["dism", "/Online", f"/Import-DefaultAppAssociations:{path}"]
        completed = self._runner.run(command)
        detail = self._format_command_detail(completed)
        policy_value = self._registry.get_value(DEFAULT_APPS_POLICY_PATH, DEFAULT_APPS_POLICY_VALUE)
        if policy_value:
            detail = f"{detail}; policy: {policy_value}"
        success = completed.returncode == 0 and str(policy_value or "").strip().lower() == str(path).lower()
        return ApplyStepResult("Default Apps", success, detail)

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
        match = POWERCFG_GUID_PATTERN.search(output)
        if match:
            return match.group(2).strip()
        if "(" in output and ")" in output:
            return output.split("(")[-1].split(")")[0].strip()
        return output.strip()

    def _list_power_schemes(self) -> list[tuple[str, str, bool]]:
        output = self._run_and_capture(["powercfg", "/list"])
        schemes: list[tuple[str, str, bool]] = []
        for match in POWERCFG_GUID_PATTERN.finditer(output):
            guid = match.group(1).strip()
            name = match.group(2).strip()
            active = bool(match.group(3))
            schemes.append((guid, name, active))
        return schemes

    def _get_active_power_scheme(self) -> tuple[str, str]:
        output = self._run_and_capture(["powercfg", "/getactivescheme"])
        match = POWERCFG_GUID_PATTERN.search(output)
        if match:
            return match.group(1).strip(), match.group(2).strip()
        return "", self._extract_power_scheme_name(output)

    def _resolve_power_scheme(
        self, schemes: Iterable[tuple[str, str, bool]]
    ) -> tuple[str, str]:
        scheme = self._config.power_plan.scheme.strip()
        friendly = self._config.power_plan.friendly_name.strip()
        if POWERCFG_GUID_PATTERN.search(f"Power Scheme GUID: {scheme} (x)"):
            return scheme, friendly
        alias_guid = KNOWN_POWER_SCHEMES.get(scheme.upper())
        if alias_guid:
            return alias_guid, friendly
        for guid, name, _active in schemes:
            if name.lower() == friendly.lower():
                return guid, name
        return "", ""

    def _wait_for_active_scheme(self, target_guid: str) -> tuple[str, str]:
        active_guid, active_name = self._get_active_power_scheme()
        if not target_guid:
            return active_guid, active_name
        for _ in range(5):
            if active_guid and active_guid.lower() == target_guid.lower():
                return active_guid, active_name
            time.sleep(0.3)
            active_guid, active_name = self._get_active_power_scheme()
        return active_guid, active_name

    def _wait_for_system_locale(self, expected: str) -> str:
        actual = self._run_and_capture(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "Get-WinSystemLocale | Select-Object -ExpandProperty Name",
            ]
        )
        if not expected:
            return actual
        for _ in range(5):
            if actual and actual.lower() == expected.lower():
                return actual
            time.sleep(0.3)
            actual = self._run_and_capture(
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    "Get-WinSystemLocale | Select-Object -ExpandProperty Name",
                ]
            )
        return actual

    def _apply_ui_languages(self) -> str | None:
        languages = tuple(lang for lang in self._config.locale.ui_languages if lang)
        if not languages:
            return None
        first, *rest = languages
        parts = [f"$list = New-WinUserLanguageList -Language {_ps_quote(first)}"]
        for lang in rest:
            parts.append(f"$list.Add({_ps_quote(lang)})")
        parts.append("Set-WinUserLanguageList -LanguageList $list -Force")
        script = "; ".join(parts)
        completed = self._runner.run(["powershell", "-NoProfile", "-Command", script])
        detail = self._format_command_detail(completed)
        if completed.returncode != 0 or completed.stdout or completed.stderr:
            return detail
        return None

    def _apply_language_packs_and_features(self) -> tuple[bool, str]:
        languages = tuple(lang for lang in self._config.locale.ui_languages if lang)
        if not languages:
            return True, ""
        detail_parts: list[str] = []
        success = True
        for language in languages:
            install_script = "; ".join(
                [
                    "$cmd = Get-Command Install-Language -ErrorAction SilentlyContinue",
                    f"if ($cmd) {{ Install-Language -Language {_ps_quote(language)} -ErrorAction Stop | Out-Null }}",
                ]
            )
            install = self._runner.run(["powershell", "-NoProfile", "-Command", install_script])
            detail_parts.append(f"{language} pack: {self._format_command_detail(install)}")
            success = success and install.returncode == 0
            prefixes = LANGUAGE_CAPABILITY_PREFIXES.get(language, ())
            for prefix in prefixes:
                capability_ok, capability_detail = self._ensure_language_capability(language, prefix)
                success = success and capability_ok
                detail_parts.append(capability_detail)
        return success, "language features: " + " | ".join(detail_parts)

    def _ensure_language_capability(self, language: str, prefix: str) -> tuple[bool, str]:
        pattern = f"{prefix}~~~{language}~*"
        script = "; ".join(
            [
                f"$cap = Get-WindowsCapability -Online | Where-Object {{ $_.Name -like {_ps_quote(pattern)} }} | Select-Object -First 1",
                "if (-not $cap) { Write-Output 'missing'; exit 7 }",
                "if ($cap.State -ne 'Installed') { Add-WindowsCapability -Online -Name $cap.Name -ErrorAction Stop | Out-Null }",
                "$state = (Get-WindowsCapability -Online -Name $cap.Name).State",
                "if ($state -ne 'Installed') { Write-Output \"state=$state\"; exit 8 }",
                "Write-Output $cap.Name",
            ]
        )
        completed = self._runner.run(["powershell", "-NoProfile", "-Command", script])
        detail = f"{language} {prefix}: {self._format_command_detail(completed)}"
        return completed.returncode == 0, detail

    def _apply_speech_preferences(self) -> str | None:
        script = "; ".join(
            [
                "Set-ItemProperty -Path 'HKCU:\\Software\\Microsoft\\Speech_OneCore\\Settings\\OnlineSpeechPrivacy' -Name 'HasAccepted' -Value 1 -Type DWord",
                "Set-ItemProperty -Path 'HKCU:\\Software\\Microsoft\\Speech_OneCore\\Settings\\SpeechRecognition' -Name 'PreferOffline' -Value 0 -Type DWord",
            ]
        )
        completed = self._runner.run(["powershell", "-NoProfile", "-Command", script])
        detail = self._format_command_detail(completed)
        if completed.returncode != 0:
            return f"speech preferences: {detail}"
        return None

    def _set_desktop_icon_registry_values(self, map_path: Callable[[str], str]) -> None:
        for icon_path in DESKTOP_ICON_VISIBILITY_PATHS:
            target = map_path(icon_path)
            for guid in DESKTOP_ICON_GUIDS:
                self._registry.set_value(target, guid, 0)

    def _check_desktop_icon_registry_values(self, root: str) -> bool:
        for icon_path in DESKTOP_ICON_VISIBILITY_PATHS:
            path = self._map_user_path(icon_path, root)
            for guid in DESKTOP_ICON_GUIDS:
                value = self._registry.get_value(path, guid)
                if value is None:
                    return False
                try:
                    if int(value) != 0:
                        return False
                except (TypeError, ValueError):
                    return False
        return True

    def _refresh_desktop_shell(self) -> str:
        command = [
            "powershell",
            "-NoProfile",
            "-Command",
            "Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue; Start-Process explorer.exe",
        ]
        completed = self._runner.run(command)
        return self._format_command_detail(completed)

    def _primary_ui_language(self) -> str:
        languages = tuple(lang for lang in self._config.locale.ui_languages if lang)
        if languages:
            return languages[0]
        return self._config.locale.system_locale

    def _default_apps_xml_path(self) -> Path:
        if winreg is None:
            return Path(tempfile.gettempdir()) / "AIO_DefaultAppAssociations.xml"
        return Path(DEFAULT_APPS_XML_PATH_WINDOWS)

    def _write_default_apps_association_file(self) -> Path:
        path = self._default_apps_xml_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        root = ET.Element("DefaultAssociations")
        for identifier, prog_id, app_name in DEFAULT_APP_ASSOCIATIONS:
            ET.SubElement(
                root,
                "Association",
                attrib={"Identifier": identifier, "ProgId": prog_id, "ApplicationName": app_name},
            )
        tree = ET.ElementTree(root)
        tree.write(path, encoding="utf-8", xml_declaration=True)
        return path

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


def _ps_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"
