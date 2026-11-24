"""Load bot configuration from YAML + environment overrides."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

WEEKDAY_ALIASES = {
    "mon": 0,
    "monday": 0,
    "tue": 1,
    "tues": 1,
    "tuesday": 1,
    "wed": 2,
    "wednesday": 2,
    "thu": 3,
    "thur": 3,
    "thurs": 3,
    "thursday": 3,
    "fri": 4,
    "friday": 4,
    "sat": 5,
    "saturday": 5,
    "sun": 6,
    "sunday": 6,
}


@dataclass
class SharePointConfig:
    file_url: str = ""
    site_hostname: str = "hdrinc.sharepoint.com"
    site_path: str = "/teams/LasVegasOfficeHomepage"
    drive_item_id: str = ""
    worksheet_name: str = "Master Bookings"
    header_row: int = 1


@dataclass
class BotConfig:
    display_name: str
    preferred_desk: str
    office_weekdays: list[int]
    timezone: str = "America/Los_Angeles"
    clear_dates: list[str] = field(default_factory=list)
    sharepoint: SharePointConfig = field(default_factory=SharePointConfig)
    dry_run: bool = False


def _parse_weekdays(raw: list[Any]) -> list[int]:
    out: list[int] = []
    for item in raw:
        if isinstance(item, int):
            if item < 0 or item > 6:
                raise ValueError(f"Invalid weekday int: {item}")
            out.append(item)
            continue
        key = str(item).strip().lower()
        if key not in WEEKDAY_ALIASES:
            raise ValueError(f"Unknown weekday: {item}")
        out.append(WEEKDAY_ALIASES[key])
    return sorted(set(out))


def load_config(path: str | Path | None = None) -> BotConfig:
    """Load config.yaml (or CONFIG_PATH / path). Env vars override SharePoint IDs and dry-run."""
    config_path = Path(path or os.environ.get("CONFIG_PATH", "config.yaml"))
    if not config_path.exists():
        example = config_path.with_name("config.example.yaml")
        if example.exists() and path is None and "CONFIG_PATH" not in os.environ:
            config_path = example
        else:
            raise FileNotFoundError(f"Config not found: {config_path}")

    with config_path.open(encoding="utf-8") as fh:
        data = yaml.safe_load(fh) or {}

    sp_raw = data.get("sharepoint") or {}
    sharepoint = SharePointConfig(
        file_url=str(sp_raw.get("file_url") or ""),
        site_hostname=str(
            os.environ.get("SHAREPOINT_SITE_HOSTNAME")
            or sp_raw.get("site_hostname")
            or "hdrinc.sharepoint.com"
        ),
        site_path=str(
            os.environ.get("SHAREPOINT_SITE_PATH")
            or sp_raw.get("site_path")
            or "/teams/LasVegasOfficeHomepage"
        ),
        drive_item_id=str(
            os.environ.get("SHAREPOINT_DRIVE_ITEM_ID")
            or sp_raw.get("drive_item_id")
            or ""
        ),
        worksheet_name=str(sp_raw.get("worksheet_name") or "Master Bookings"),
        header_row=int(sp_raw.get("header_row") or 1),
    )

    dry_run = str(os.environ.get("DRY_RUN", "")).lower() in {"1", "true", "yes"}
    if data.get("dry_run"):
        dry_run = True

    return BotConfig(
        display_name=str(data.get("display_name") or "").strip(),
        preferred_desk=str(data.get("preferred_desk") or "").strip(),
        office_weekdays=_parse_weekdays(list(data.get("office_weekdays") or [])),
        timezone=str(data.get("timezone") or "America/Los_Angeles"),
        clear_dates=[str(d) for d in (data.get("clear_dates") or [])],
        sharepoint=sharepoint,
        dry_run=dry_run,
    )
