"""Regression checks for immutable data migrated from PowerShell."""
from __future__ import annotations

from trustpal.app_registry import REGISTRY
from trustpal.constants import IMMUTABLE_CONFIG


def test_timezone_is_west_bank() -> None:
    assert IMMUTABLE_CONFIG.system.timezone == "West Bank Standard Time"


def test_locale_defaults() -> None:
    locale = IMMUTABLE_CONFIG.system.locale
    assert locale.system_locale == "ar-SA"
    assert locale.short_date_format == "dd/MM/yyyy"
    assert locale.ui_languages == ("ar-SA", "en-US")


def test_crowdstrike_cid_matches_args() -> None:
    target = next(app for app in REGISTRY.entries if app.name == "CrowdStrike Falcon Sensor")
    assert IMMUTABLE_CONFIG.ids.crowdstrike_cid in target.args


def test_hp_legacy_repo_path() -> None:
    assert IMMUTABLE_CONFIG.ids.hp_legacy_repo_root == r"\\192.168.168.6\Admin Tools\Drivers-Repo"


def test_winget_ids_preserved() -> None:
    chrome = next(app for app in REGISTRY.entries if app.name == "Chrome")
    assert chrome.winget_id == "Google.Chrome"
    vc2015 = next(app for app in REGISTRY.entries if app.name == "VC++ 2015+")
    assert vc2015.winget_id_x64 == "Microsoft.VCRedist.2015+.x64"
    assert vc2015.winget_id_x86 == "Microsoft.VCRedist.2015+.x86"


def test_registry_has_expected_count() -> None:
    assert len(REGISTRY.entries) == 20
