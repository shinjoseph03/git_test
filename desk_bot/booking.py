"""Book / clear orchestration against Master Bookings."""

from __future__ import annotations

import logging
from datetime import date, datetime
from zoneinfo import ZoneInfo

from desk_bot.config import BotConfig
from desk_bot.sheet import (
    apply_book,
    apply_clear,
    find_rows_for_desk_date,
    get_worksheet,
    iter_booking_rows,
    load_workbook_bytes,
    workbook_to_bytes,
)

log = logging.getLogger(__name__)


def today_in_tz(timezone: str, override: str | None = None) -> date:
    if override:
        return date.fromisoformat(override)
    return datetime.now(ZoneInfo(timezone)).date()


def is_office_weekday(cfg: BotConfig, on: date) -> bool:
    return on.weekday() in cfg.office_weekdays


def book_for_date(
    cfg: BotConfig,
    workbook_bytes: bytes,
    on: date,
) -> tuple[bytes, list[str]]:
    """
    Book preferred desk for `on` if Available AM+PM.
    Returns (possibly updated workbook bytes, list of log/result lines).
    """
    messages: list[str] = []
    if not cfg.display_name:
        messages.append("error: display_name is empty")
        return workbook_bytes, messages
    if not cfg.preferred_desk:
        messages.append("error: preferred_desk is empty")
        return workbook_bytes, messages

    if not is_office_weekday(cfg, on):
        messages.append(
            f"skip: {on.isoformat()} is not an office weekday "
            f"(configured={cfg.office_weekdays})"
        )
        return workbook_bytes, messages

    wb = load_workbook_bytes(workbook_bytes)
    ws = get_worksheet(wb, cfg.sharepoint.worksheet_name)
    rows = iter_booking_rows(ws, cfg.sharepoint.header_row)
    matches = find_rows_for_desk_date(rows, desk=cfg.preferred_desk, on=on)

    if not matches:
        messages.append(
            f"skip: no row for desk={cfg.preferred_desk!r} on {on.isoformat()}"
        )
        return workbook_bytes, messages

    changed = False
    for row in matches:
        try:
            result = apply_book(
                ws, row, cfg.display_name, header_row=cfg.sharepoint.header_row
            )
        except Exception as exc:  # noqa: BLE001 — lock/blocked cells surface here
            result = f"skip_write_failed:{exc}"
        messages.append(
            f"{result}: desk={row.desk} date={row.date.isoformat()} "
            f"area={row.area} row={row.excel_row}"
        )
        if result == "booked":
            changed = True

    if changed:
        return workbook_to_bytes(wb), messages
    return workbook_bytes, messages


def clear_for_dates(
    cfg: BotConfig,
    workbook_bytes: bytes,
    dates: list[date],
) -> tuple[bytes, list[str]]:
    """Clear configured display_name from AM/PM on the given dates (all desks)."""
    messages: list[str] = []
    if not cfg.display_name:
        messages.append("error: display_name is empty")
        return workbook_bytes, messages
    if not dates:
        messages.append("skip: no clear_dates provided")
        return workbook_bytes, messages

    wb = load_workbook_bytes(workbook_bytes)
    ws = get_worksheet(wb, cfg.sharepoint.worksheet_name)
    rows = iter_booking_rows(ws, cfg.sharepoint.header_row)
    target = set(dates)
    changed = False

    for row in rows:
        if row.date not in target:
            continue
        try:
            result = apply_clear(
                ws, row, cfg.display_name, header_row=cfg.sharepoint.header_row
            )
        except Exception as exc:  # noqa: BLE001
            result = f"skip_write_failed:{exc}"
        messages.append(
            f"{result}: desk={row.desk} date={row.date.isoformat()} "
            f"area={row.area} row={row.excel_row}"
        )
        if result == "cleared":
            changed = True

    if not any(m.startswith(("cleared", "already_clear", "skip_")) for m in messages):
        messages.append(
            f"skip: no rows matched clear dates {[d.isoformat() for d in dates]}"
        )

    if changed:
        return workbook_to_bytes(wb), messages
    return workbook_bytes, messages
