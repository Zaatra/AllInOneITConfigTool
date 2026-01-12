"""Helpers for Office Deployment Tool templates."""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class OfficeTemplate:
    name: str
    xml: str


def get_template(name: str, templates: dict[str, OfficeTemplate]) -> OfficeTemplate:
    try:
        return templates[name]
    except KeyError as exc:
        raise KeyError(f"Unknown Office template: {name}") from exc
