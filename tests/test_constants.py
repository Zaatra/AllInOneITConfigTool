"""Regression checks for immutable data migrated from PowerShell."""
from __future__ import annotations

from allinone_it_config.app_registry import build_registry
from allinone_it_config.constants import IMMUTABLE_CONFIG
from allinone_it_config.user_settings import UserSettings


def test_timezone_is_west_bank() -> None:
    assert IMMUTABLE_CONFIG.system.timezone == "West Bank Standard Time"


def test_locale_defaults() -> None:
    locale = IMMUTABLE_CONFIG.system.locale
    assert locale.system_locale == "en-US"
    assert locale.short_date_format == "dd/MM/yyyy"
    assert locale.ui_languages == ("en-US", "ar-SA")


def test_crowdstrike_args_empty_by_default() -> None:
    registry = build_registry(UserSettings())
    target = next(app for app in registry.entries if app.name == "CrowdStrike Falcon Sensor")
    assert target.args == ""


def test_hp_legacy_repo_path() -> None:
    assert IMMUTABLE_CONFIG.ids.hp_legacy_repo_root == ""


def test_winget_ids_preserved() -> None:
    registry = build_registry(UserSettings())
    chrome = next(app for app in registry.entries if app.name == "Chrome")
    assert chrome.winget_id == "Google.Chrome"
    vc2015 = next(app for app in registry.entries if app.name == "VC++ 2015+")
    assert vc2015.winget_id_x64 == "Microsoft.VCRedist.2015+.x64"
    assert vc2015.winget_id_x86 == "Microsoft.VCRedist.2015+.x86"


def test_registry_has_expected_count() -> None:
    registry = build_registry(UserSettings())
    assert len(registry.entries) == 21
