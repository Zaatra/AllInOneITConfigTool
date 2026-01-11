"""Immutable settings mirrored from the PowerShell implementation."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Tuple


@dataclass(frozen=True)
class OfficeTemplate:
    name: str
    xml: str


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
class GlobalIds:
    crowdstrike_cid: str
    hp_legacy_repo_root: str


@dataclass(frozen=True)
class ImmutableConfig:
    office_templates: Dict[str, OfficeTemplate]
    system: FixedSystemConfig
    ids: GlobalIds


CONFIG_ROOT = Path(__file__).resolve().parent

OFFICE_TEMPLATES: Dict[str, OfficeTemplate] = {
    "office_2024_ltsc": OfficeTemplate(
        name="Office 2024 LTSC",
        xml="""<Configuration ID=\"917267f9-54e9-4c93-8a91-b6df3e3f2b1d\">
  <Add OfficeClientEdition=\"64\" Channel=\"PerpetualVL2024\" MigrateArch=\"TRUE\">
    <Product ID=\"ProPlus2024Volume\" PIDKEY=\"BQG8N-4Y7XV-92QPV-K26X6-QYM2Q\">
      <Language ID=\"en-us\"/>
      <Language ID=\"ar-sa\"/>
      <ExcludeApp ID=\"Lync\"/>
      <ExcludeApp ID=\"OneDrive\"/>
      <ExcludeApp ID=\"Publisher\"/>
    </Product>
  </Add>
  <Property Name=\"SharedComputerLicensing\" Value=\"0\"/>
  <Property Name=\"FORCEAPPSHUTDOWN\" Value=\"TRUE\"/>
  <Property Name=\"AUTOACTIVATE\" Value=\"1\"/>
  <Updates Enabled=\"TRUE\"/>
  <Remove All=\"TRUE\"/>
  <Display Level=\"None\" AcceptEULA=\"TRUE\"/>
</Configuration>
""",
    ),
    "office_365_enterprise": OfficeTemplate(
        name="Office 365 Enterprise",
        xml="""<Configuration ID=\"eb27020c-89bd-4519-96d9-7c652bf678eb\">
<Info Description=\"\"/>
<Add OfficeClientEdition=\"64\" Channel=\"Current\" MigrateArch=\"TRUE\">
<Product ID=\"O365ProPlusRetail\">
<Language ID=\"en-us\"/>
<Language ID=\"ar-sa\"/>
<ExcludeApp ID=\"Groove\"/>
<ExcludeApp ID=\"Lync\"/>
<ExcludeApp ID=\"OneDrive\"/>
<ExcludeApp ID=\"OneNote\"/>
<ExcludeApp ID=\"OutlookForWindows\"/>
<ExcludeApp ID=\"Publisher\"/>
</Product>
</Add>
<Property Name=\"SharedComputerLicensing\" Value=\"0\"/>
<Property Name=\"FORCEAPPSHUTDOWN\" Value=\"TRUE\"/>
<Property Name=\"DeviceBasedLicensing\" Value=\"0\"/>
<Property Name=\"SCLCacheOverride\" Value=\"0\"/>
<Updates Enabled=\"TRUE\"/>
<Remove All=\"TRUE\"/>
<RemoveMSI/>
<AppSettings>
<Setup Name=\"Company\" Value=\"Trust\"/>
<User Key=\"software\\microsoft\\office\\16.0\\excel\\options\" Name=\"defaultformat\" Value=\"51\" Type=\"REG_DWORD\" App=\"excel16\" Id=\"L_SaveExcelfilesas\"/>
<User Key=\"software\\microsoft\\office\\16.0\\powerpoint\\options\" Name=\"defaultformat\" Value=\"27\" Type=\"REG_DWORD\" App=\"ppt16\" Id=\"L_SavePowerPointfilesas\"/>
<User Key=\"software\\microsoft\\office\\16.0\\word\\options\" Name=\"defaultformat\" Value=\"\" Type=\"REG_SZ\" App=\"word16\" Id=\"L_SaveWordfilesas\"/>
</AppSettings>
<Display Level=\"None\" AcceptEULA=\"TRUE\"/>
</Configuration>
""",
    ),
}

FIXED_SYSTEM_CONFIG = FixedSystemConfig(
    timezone="West Bank Standard Time",
    locale=LocaleSetting(
        system_locale="ar-SA",
        short_date_format="dd/MM/yyyy",
        ui_languages=("ar-SA", "en-US"),
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

GLOBAL_IDS = GlobalIds(
    crowdstrike_cid="753B2C62DBF944A59165207E3FE2FF2A-27",
    hp_legacy_repo_root=r"\\192.168.168.6\Admin Tools\Drivers-Repo",
)

IMMUTABLE_CONFIG = ImmutableConfig(
    office_templates=OFFICE_TEMPLATES,
    system=FIXED_SYSTEM_CONFIG,
    ids=GLOBAL_IDS,
)
