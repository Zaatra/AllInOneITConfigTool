from __future__ import annotations

import subprocess
from typing import Sequence

from services.system_config import ConfigCheckResult, RegistryAccessor, SystemConfigService
from trustpal.constants import IMMUTABLE_CONFIG


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


def test_check_reports_desired_state() -> None:
    config = IMMUTABLE_CONFIG.system
    runner = FakeRunner(
        {
            ("tzutil", "/g"): f"{config.timezone}\n",
            ("powercfg", "/getactivescheme"): "Power Scheme GUID: 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c  (High performance)",
            (
                "powershell",
                "-NoProfile",
                "-Command",
                "Get-WinSystemLocale | Select-Object -ExpandProperty Name",
            ): f"{config.locale.system_locale}\n",
        }
    )
    registry = FakeRegistry(
        {
            (config.fast_boot.path, config.fast_boot.value_name): int(config.fast_boot.desired_value),
            (config.desktop_icons.path, config.desktop_icons.value_name): int(config.desktop_icons.desired_value),
            (r"HKCU:\Control Panel\International", "sShortDate"): config.locale.short_date_format,
        }
    )
    service = SystemConfigService(config, command_runner=runner, registry=registry)
    results = service.check()
    assert all(isinstance(res, ConfigCheckResult) and res.in_desired_state for res in results)


def test_apply_runs_commands_and_sets_registry() -> None:
    config = IMMUTABLE_CONFIG.system
    runner = FakeRunner()
    registry = FakeRegistry()
    service = SystemConfigService(config, command_runner=runner, registry=registry)
    service.apply()

    assert ("tzutil", "/s", config.timezone) in runner.commands
    assert ("powercfg", "/setactive", config.power_plan.scheme) in runner.commands
    locale_cmd = (
        "powershell",
        "-NoProfile",
        "-Command",
        f"Set-WinSystemLocale -SystemLocale {config.locale.system_locale}",
    )
    assert locale_cmd in runner.commands
    assert registry.get_value(config.fast_boot.path, config.fast_boot.value_name) == int(config.fast_boot.desired_value)
    assert registry.get_value(config.desktop_icons.path, config.desktop_icons.value_name) == int(config.desktop_icons.desired_value)
    assert registry.get_value(r"HKCU:\Control Panel\International", "sShortDate") == config.locale.short_date_format
