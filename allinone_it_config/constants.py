"""Immutable settings mirrored from the PowerShell implementation."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Tuple


@dataclass(frozen=True)
class LocaleSetting:
    system_locale: str
    short_date_format: str
    ui_languages: Tuple[str, ...]


@dataclass(frozen=True)
class PowerPlanSetting:
    scheme: str
    friendly_name: str


@dataclass(frozen=True)
class RegistrySetting:
    path: str
    value_name: str
    desired_value: int | str


@dataclass(frozen=True)
class FixedSystemConfig:
    timezone: str
    locale: LocaleSetting
    power_plan: PowerPlanSetting
    fast_boot: RegistrySetting
    desktop_icons: RegistrySetting


@dataclass(frozen=True)
class ImmutableConfig:
    system: FixedSystemConfig


CONFIG_ROOT = Path(__file__).resolve().parent

FIXED_SYSTEM_CONFIG = FixedSystemConfig(
    timezone="West Bank Standard Time",
    locale=LocaleSetting(
        system_locale="en-US",
        short_date_format="dd/MM/yyyy",
        ui_languages=("en-US", "ar-SA"),
    ),
    power_plan=PowerPlanSetting(
        scheme="SCHEME_MAX",
        friendly_name="High performance",
    ),
    fast_boot=RegistrySetting(
        path=r"HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power",
        value_name="HiberbootEnabled",
        desired_value=0,
    ),
    desktop_icons=RegistrySetting(
        path=r"HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced",
        value_name="HideIcons",
        desired_value=0,
    ),
)

IMMUTABLE_CONFIG = ImmutableConfig(
    system=FIXED_SYSTEM_CONFIG,
)
