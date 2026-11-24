"""Dry-run unit tests — no SharePoint / Azure secrets required."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest
import yaml

from desk_bot.booking import book_for_date, clear_for_dates, is_office_weekday
from desk_bot.config import BotConfig, SharePointConfig, load_config
from desk_bot.sheet import (
    apply_book,
    find_rows_for_desk_date,
    get_worksheet,
    iter_booking_rows,
    load_workbook_bytes,
    make_fixture_workbook,
)

FIXTURES = Path(__file__).parent / "fixtures"


def _cfg(**overrides) -> BotConfig:
    base = BotConfig(
        display_name="Joseph Shin",
        preferred_desk="Desk 01",
        office_weekdays=[0, 1, 3, 4],  # Mon Tue Thu Fri
        timezone="America/Los_Angeles",
        clear_dates=[],
        sharepoint=SharePointConfig(
            drive_item_id="4CC41F6E-391F-4BBC-B75B-3D0BABCA2BE5",
            worksheet_name="Master Bookings",
        ),
        dry_run=True,
    )
    for k, v in overrides.items():
        setattr(base, k, v)
    return base


def _wb(rows):
    return make_fixture_workbook(rows)


def test_load_config_example():
    cfg = load_config(Path(__file__).resolve().parents[1] / "config.example.yaml")
    assert cfg.preferred_desk == "Desk 01"
    assert 0 in cfg.office_weekdays  # Mon
    assert cfg.sharepoint.drive_item_id.upper() == "4CC41F6E-391F-4BBC-B75B-3D0BABCA2BE5"
    assert "LasVegasOfficeHomepage" in cfg.sharepoint.file_url
    assert "Desk%20Reservation%20Master.xlsx" in cfg.sharepoint.file_url


def test_office_weekday():
    cfg = _cfg()
    assert is_office_weekday(cfg, date(2026, 11, 24))  # Tuesday
    assert not is_office_weekday(cfg, date(2026, 11, 25))  # Wednesday
    assert not is_office_weekday(cfg, date(2026, 11, 28))  # Saturday


def test_book_when_available():
    on = date(2026, 11, 24)  # Tue
    data = _wb(
        [
            {
                "Date": on,
                "Desk": "Desk 01",
                "Area": "North",
                "AM Employee": None,
                "PM Employee": None,
                "AM Status": "Available",
                "PM Status": "Available",
            }
        ]
    )
    updated, msgs = book_for_date(_cfg(), data, on)
    assert any(m.startswith("booked:") for m in msgs)
    wb = load_workbook_bytes(updated)
    rows = iter_booking_rows(get_worksheet(wb, "Master Bookings"))
    assert rows[0].am_employee == "Joseph Shin"
    assert rows[0].pm_employee == "Joseph Shin"


def test_book_skip_not_available():
    on = date(2026, 11, 24)
    data = _wb(
        [
            {
                "Date": on,
                "Desk": "Desk 01",
                "AM Status": "Blocked",
                "PM Status": "Available",
            }
        ]
    )
    updated, msgs = book_for_date(_cfg(), data, on)
    assert updated == data
    assert any("skip_am_status" in m for m in msgs)


def test_book_skip_held_by_other():
    on = date(2026, 11, 24)
    data = _wb(
        [
            {
                "Date": on,
                "Desk": "Desk 01",
                "AM Employee": "Alex Other",
                "PM Employee": "Alex Other",
                "AM Status": "Available",
                "PM Status": "Available",
            }
        ]
    )
    updated, msgs = book_for_date(_cfg(), data, on)
    assert updated == data
    assert any("skip_am_held_by" in m for m in msgs)


def test_book_already_mine_noop():
    on = date(2026, 11, 24)
    data = _wb(
        [
            {
                "Date": on,
                "Desk": "Desk 01",
                "AM Employee": "Joseph Shin",
                "PM Employee": "Joseph Shin",
            }
        ]
    )
    updated, msgs = book_for_date(_cfg(), data, on)
    assert updated == data
    assert any(m.startswith("already_mine:") for m in msgs)


def test_book_skip_non_office_day():
    on = date(2026, 11, 25)  # Wed
    data = _wb([{"Date": on, "Desk": "Desk 01"}])
    updated, msgs = book_for_date(_cfg(), data, on)
    assert updated == data
    assert any("not an office weekday" in m for m in msgs)


def test_clear_only_my_name():
    d1 = date(2026, 11, 24)
    d2 = date(2026, 11, 26)
    data = _wb(
        [
            {
                "Date": d1,
                "Desk": "Desk 01",
                "AM Employee": "Joseph Shin",
                "PM Employee": "Joseph Shin",
            },
            {
                "Date": d1,
                "Desk": "Desk 02",
                "AM Employee": "Alex Other",
                "PM Employee": "Alex Other",
            },
            {
                "Date": d2,
                "Desk": "Desk 01",
                "AM Employee": "Joseph Shin",
                "PM Employee": None,
            },
        ]
    )
    updated, msgs = clear_for_dates(_cfg(), data, [d1])
    assert any(m.startswith("cleared:") and "Desk 01" in m for m in msgs)
    assert any("skip_not_mine" in m and "Desk 02" in m for m in msgs)

    wb = load_workbook_bytes(updated)
    rows = iter_booking_rows(get_worksheet(wb, "Master Bookings"))
    by_desk = {(r.date, r.desk): r for r in rows}
    assert by_desk[(d1, "Desk 01")].am_employee == ""
    assert by_desk[(d1, "Desk 01")].pm_employee == ""
    assert by_desk[(d1, "Desk 02")].am_employee == "Alex Other"
    # d2 not in clear list — untouched
    assert by_desk[(d2, "Desk 01")].am_employee == "Joseph Shin"


def test_fixture_file_roundtrip(tmp_path: Path):
    """Generate fixture xlsx and run book via CLI-equivalent path."""
    on = date(2026, 11, 27)  # Fri
    raw = _wb(
        [
            {
                "Date": on,
                "Desk": "Desk 01",
                "AM Status": "Available",
                "PM Status": "Available",
            }
        ]
    )
    path = tmp_path / "master.xlsx"
    path.write_bytes(raw)
    updated, msgs = book_for_date(_cfg(), path.read_bytes(), on)
    assert any("booked:" in m for m in msgs)
    out = tmp_path / "out.xlsx"
    out.write_bytes(updated)
    rows = iter_booking_rows(
        get_worksheet(load_workbook_bytes(out.read_bytes()), "Master Bookings")
    )
    match = find_rows_for_desk_date(rows, desk="Desk 01", on=on)
    assert match[0].am_employee == "Joseph Shin"


def test_apply_book_direct():
    data = _wb(
        [
            {
                "Date": date(2026, 11, 24),
                "Desk": "Desk 01",
                "AM Status": "Available",
                "PM Status": "Available",
            }
        ]
    )
    wb = load_workbook_bytes(data)
    ws = get_worksheet(wb, "Master Bookings")
    row = iter_booking_rows(ws)[0]
    assert apply_book(ws, row, "Joseph Shin") == "booked"
