"""Helpers for Office Deployment Tool templates."""
from __future__ import annotations

from trustpal.constants import OfficeTemplate


def get_template(name: str, templates: dict[str, OfficeTemplate]) -> OfficeTemplate:
    try:
        return templates[name]
    except KeyError as exc:
        raise KeyError(f"Unknown Office template: {name}") from exc
