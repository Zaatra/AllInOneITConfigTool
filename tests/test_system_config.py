from __future__ import annotations

import subprocess
from typing import Sequence

from allinone_it_config.constants import IMMUTABLE_CONFIG
from services.system_config import (
    ARABIC_SPELLING_REG_PATH,
    ARABIC_SPELLING_RULES,
    DEFAULT_APPS_POLICY_PATH,
    DEFAULT_APPS_POLICY_VALUE,
    DEFAULT_USER_HIVE_KEY,
    DESKTOP_ICON_GUIDS,
    DESKTOP_ICON_VISIBILITY_PATHS,
    DESKTOP_POLICY_PATH,
    DESKTOP_POLICY_VALUE,
    TARGET_HOME_GEO_ID,
    ConfigCheckResult,
    RegistryAccessor,
    SystemConfigService,
)


class FakeRunner:
    def __init__(self, stdouts: dict[tuple[str, ...], str] | None = None) -> None:
        self.stdouts = stdouts or {}
        self.commands: list[Sequence[str]] = []

    def run(self, command: Sequence[str]) -> subprocess.CompletedProcess[str]:
        self.commands.append(tuple(command))
        stdout = self.stdouts.get(tuple(command), "")
        return subprocess.CompletedProcess(command, 0, stdout, "")


class FakeRegistry(RegistryAccessor):
    def __init__(self, initial: dict[tuple[str, str], str | int] | None = None) -> None:
        self.values = initial or {}

    def get_value(self, path: str, value_name: str) -> str | int | None:
        return self.values.get((path, value_name))

    def set_value(self, path: str, value_name: str, value: str | int) -> None:
        self.values[(path, value_name)] = value


def _desired_state_runner() -> FakeRunner:
    config = IMMUTABLE_CONFIG.system
    return FakeRunner(
        {
            ("tzutil", "/g"): f"{config.timezone}\n",
            ("powercfg", "/getactivescheme"): "Power Scheme GUID: 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c  (High performance)",
            (
                "powershell",
                "-NoProfile",
                "-Command",
                "Get-WinSystemLocale | Select-Object -ExpandProperty Name",
            ): f"{config.locale.system_locale}\n",
            (
                "powershell",
                "-NoProfile",
                "-Command",
                "(Get-Culture).Name",
            ): f"{config.locale.ui_languages[0]}\n",
            (
                "powershell",
                "-NoProfile",
                "-Command",
                "(Get-WinHomeLocation).GeoId",
            ): f"{TARGET_HOME_GEO_ID}\n",
            (
                "powershell",
                "-NoProfile",
                "-Command",
                "(Get-WinUserLanguageList).LanguageTag",
            ): "en-US\nar-SA\n",
            (
                "powershell",
                "-NoProfile",
                "-Command",
                "(Get-WinUILanguageOverride).Name",
            ): f"{config.locale.ui_languages[0]}\n",
        }
    )


def _desired_state_registry() -> FakeRegistry:
    config = IMMUTABLE_CONFIG.system
    default_root = fr"HKU:\{DEFAULT_USER_HIVE_KEY}"
    initial_registry: dict[tuple[str, str], str | int] = {
        (config.fast_boot.path, config.fast_boot.value_name): int(config.fast_boot.desired_value),
        (config.desktop_icons.path, config.desktop_icons.value_name): int(config.desktop_icons.desired_value),
        (DESKTOP_POLICY_PATH, DESKTOP_POLICY_VALUE): 0,
        (r"HKCU:\Control Panel\International", "sShortDate"): config.locale.short_date_format,
        (
            fr"{default_root}\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced",
            config.desktop_icons.value_name,
        ): int(config.desktop_icons.desired_value),
        (
            fr"{default_root}\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer",
            DESKTOP_POLICY_VALUE,
        ): 0,
        (fr"{default_root}\Control Panel\International", "sShortDate"): config.locale.short_date_format,
    }
    for value_name, expected in ARABIC_SPELLING_RULES.items():
        initial_registry[(ARABIC_SPELLING_REG_PATH, value_name)] = expected
    for icon_path in DESKTOP_ICON_VISIBILITY_PATHS:
        suffix = icon_path.split("HKCU:\\", 1)[1]
        mapped = fr"{default_root}\{suffix}"
        for guid in DESKTOP_ICON_GUIDS:
            initial_registry[(mapped, guid)] = 0
    return FakeRegistry(initial_registry)


def test_check_reports_desired_state() -> None:
    runner = _desired_state_runner()
    registry = _desired_state_registry()
    service = SystemConfigService(IMMUTABLE_CONFIG.system, command_runner=runner, registry=registry)
    registry.set_value(
        DEFAULT_APPS_POLICY_PATH,
        DEFAULT_APPS_POLICY_VALUE,
        str(service._default_apps_xml_path()),
    )
    results = service.check()
    assert all(isinstance(res, ConfigCheckResult) and res.in_desired_state for res in results)


def test_apply_runs_commands_and_sets_registry() -> None:
    config = IMMUTABLE_CONFIG.system
    runner = FakeRunner()
    registry = FakeRegistry()
    service = SystemConfigService(config, command_runner=runner, registry=registry)
    service.apply()

    assert ("tzutil", "/s", config.timezone) in runner.commands
    assert (
        ("powercfg", "/setactive", config.power_plan.scheme) in runner.commands
        or ("powercfg", "/setactive", "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c") in runner.commands
    )
    locale_cmd = (
        "powershell",
        "-NoProfile",
        "-Command",
        f"Set-WinSystemLocale -SystemLocale {config.locale.system_locale}",
    )
    assert locale_cmd in runner.commands
    assert any(cmd[0] == "dism" and "/Import-DefaultAppAssociations:" in cmd[2] for cmd in runner.commands if len(cmd) >= 3)
    assert ("reg", "load", fr"HKU\{DEFAULT_USER_HIVE_KEY}", r"C:\Users\Default\NTUSER.DAT") in runner.commands
    assert ("reg", "unload", fr"HKU\{DEFAULT_USER_HIVE_KEY}") in runner.commands
    assert registry.get_value(config.fast_boot.path, config.fast_boot.value_name) == int(config.fast_boot.desired_value)
    assert registry.get_value(config.desktop_icons.path, config.desktop_icons.value_name) == int(config.desktop_icons.desired_value)
    assert registry.get_value(DESKTOP_POLICY_PATH, DESKTOP_POLICY_VALUE) == 0
    assert registry.get_value(r"HKCU:\Control Panel\International", "sShortDate") == config.locale.short_date_format
    assert registry.get_value(r"HKCU:\Control Panel\International", "iDate") == "1"
    assert registry.get_value(r"HKCU:\Control Panel\International", "sDate") == "/"
    for value_name, expected in ARABIC_SPELLING_RULES.items():
        assert registry.get_value(ARABIC_SPELLING_REG_PATH, value_name) == expected
    assert (
        registry.get_value(
            fr"HKU:\{DEFAULT_USER_HIVE_KEY}\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced",
            config.desktop_icons.value_name,
        )
        == int(config.desktop_icons.desired_value)
    )
    assert (
        registry.get_value(
            fr"HKU:\{DEFAULT_USER_HIVE_KEY}\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer",
            DESKTOP_POLICY_VALUE,
        )
        == 0
    )
    assert (
        registry.get_value(fr"HKU:\{DEFAULT_USER_HIVE_KEY}\Control Panel\International", "sShortDate")
        == config.locale.short_date_format
    )
    for icon_path in DESKTOP_ICON_VISIBILITY_PATHS:
        suffix = icon_path.split("HKCU:\\", 1)[1]
        mapped_current = fr"HKCU:\{suffix}"
        mapped_default = fr"HKU:\{DEFAULT_USER_HIVE_KEY}\{suffix}"
        for guid in DESKTOP_ICON_GUIDS:
            assert registry.get_value(mapped_current, guid) == 0
            assert registry.get_value(mapped_default, guid) == 0
    assert registry.get_value(DEFAULT_APPS_POLICY_PATH, DEFAULT_APPS_POLICY_VALUE) == str(service._default_apps_xml_path())


def test_apply_selected_only_runs_requested_steps() -> None:
    config = IMMUTABLE_CONFIG.system
    runner = FakeRunner({("tzutil", "/g"): config.timezone})
    registry = FakeRegistry()
    service = SystemConfigService(config, command_runner=runner, registry=registry)

    results = service.apply_with_results(["Timezone", "Fast Boot"])

    assert [result.name for result in results] == ["Timezone", "Fast Boot"]
    assert ("tzutil", "/s", config.timezone) in runner.commands
    assert any(cmd[0] == "tzutil" and cmd[1] == "/g" for cmd in runner.commands)
    assert not any(cmd[0] == "powercfg" and "/setactive" in cmd for cmd in runner.commands)
    assert not any(cmd[0] == "dism" for cmd in runner.commands)
    assert registry.get_value(config.fast_boot.path, config.fast_boot.value_name) == int(config.fast_boot.desired_value)


def test_diagnostics_include_time_and_locale_checks() -> None:
    runner = _desired_state_runner()
    registry = _desired_state_registry()
    service = SystemConfigService(IMMUTABLE_CONFIG.system, command_runner=runner, registry=registry)

    results = service.diagnostics()

    assert any(result.name == "Locale" for result in results)
    assert any(result.name == "Diagnostic Timezone" for result in results)
    assert any(result.name == "Diagnostic Arabic Spelling" for result in results)
