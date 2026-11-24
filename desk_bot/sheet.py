"""Parse and mutate Master Bookings rows (openpyxl workbook in memory)."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

log = logging.getLogger(__name__)

# Canonical column headers (case-insensitive match)
COLUMNS = (
    "Date",
    "Desk",
    "Area",
    "AM Employee",
    "PM Employee",
    "AM Status",
    "PM Status",
)

AVAILABLE = "available"


@dataclass
class BookingRow:
    excel_row: int  # 1-based sheet row
    date: date
    desk: str
    area: str
    am_employee: str
    pm_employee: str
    am_status: str
    pm_status: str


class SheetError(Exception):
    """Workbook / sheet structure problem."""


def _norm(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _parse_date(value: Any) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _header_map(ws: Worksheet, header_row: int) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for cell in ws[header_row]:
        if cell.value is None:
            continue
        name = _norm(cell.value)
        mapping[name.lower()] = cell.column
    missing = [c for c in COLUMNS if c.lower() not in mapping]
    if missing:
        raise SheetError(f"Missing columns on header row {header_row}: {missing}")
    return mapping


def load_workbook_bytes(data: bytes) -> Workbook:
    return load_workbook(BytesIO(data))


def workbook_to_bytes(wb: Workbook) -> bytes:
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def get_worksheet(wb: Workbook, name: str) -> Worksheet:
    if name in wb.sheetnames:
        return wb[name]
    # Soft fallback: first sheet if only one, else error
    if len(wb.sheetnames) == 1:
        log.warning("Worksheet %r not found; using sole sheet %r", name, wb.sheetnames[0])
        return wb[wb.sheetnames[0]]
    raise SheetError(
        f"Worksheet {name!r} not found. Available: {wb.sheetnames}"
    )


def iter_booking_rows(ws: Worksheet, header_row: int = 1) -> list[BookingRow]:
    headers = _header_map(ws, header_row)
    rows: list[BookingRow] = []
    for excel_row in range(header_row + 1, ws.max_row + 1):
        raw_date = ws.cell(excel_row, headers["date"]).value
        parsed = _parse_date(raw_date)
        if parsed is None:
            continue
        rows.append(
            BookingRow(
                excel_row=excel_row,
                date=parsed,
                desk=_norm(ws.cell(excel_row, headers["desk"]).value),
                area=_norm(ws.cell(excel_row, headers["area"]).value),
                am_employee=_norm(ws.cell(excel_row, headers["am employee"]).value),
                pm_employee=_norm(ws.cell(excel_row, headers["pm employee"]).value),
                am_status=_norm(ws.cell(excel_row, headers["am status"]).value),
                pm_status=_norm(ws.cell(excel_row, headers["pm status"]).value),
            )
        )
    return rows


def find_rows_for_desk_date(
    rows: list[BookingRow], *, desk: str, on: date
) -> list[BookingRow]:
    desk_l = desk.strip().lower()
    return [r for r in rows if r.date == on and r.desk.lower() == desk_l]


def names_match(a: str, b: str) -> bool:
    return a.strip().lower() == b.strip().lower()


def is_available(status: str) -> bool:
    return status.strip().lower() == AVAILABLE


def apply_book(
    ws: Worksheet,
    row: BookingRow,
    display_name: str,
    header_row: int = 1,
) -> str:
    """
    Attempt to write display_name into AM/PM Employee for the row.
    Returns a status reason: booked | already_mine | skip_* .
    Does not raise on 'skip' business cases.
    """
    headers = _header_map(ws, header_row)
    am_emp = row.am_employee
    pm_emp = row.pm_employee

    if names_match(am_emp, display_name) and names_match(pm_emp, display_name):
        return "already_mine"

    if am_emp and not names_match(am_emp, display_name):
        return f"skip_am_held_by:{am_emp}"
    if pm_emp and not names_match(pm_emp, display_name):
        return f"skip_pm_held_by:{pm_emp}"

    if not is_available(row.am_status):
        return f"skip_am_status:{row.am_status or 'empty'}"
    if not is_available(row.pm_status):
        return f"skip_pm_status:{row.pm_status or 'empty'}"

    ws.cell(row.excel_row, headers["am employee"]).value = display_name
    ws.cell(row.excel_row, headers["pm employee"]).value = display_name
    return "booked"


def apply_clear(
    ws: Worksheet,
    row: BookingRow,
    display_name: str,
    header_row: int = 1,
) -> str:
    """Clear display_name from AM/PM only; never touch other people's names."""
    headers = _header_map(ws, header_row)
    changed = False

    if names_match(row.am_employee, display_name):
        ws.cell(row.excel_row, headers["am employee"]).value = None
        changed = True
    if names_match(row.pm_employee, display_name):
        ws.cell(row.excel_row, headers["pm employee"]).value = None
        changed = True

    if changed:
        return "cleared"
    if not row.am_employee and not row.pm_employee:
        return "already_clear"
    return f"skip_not_mine:am={row.am_employee or '-'} pm={row.pm_employee or '-'}"


def make_fixture_workbook(rows: list[dict[str, Any]], sheet_name: str = "Master Bookings") -> bytes:
    """Build a minimal Master Bookings xlsx for tests."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    for col, name in enumerate(COLUMNS, start=1):
        ws.cell(1, col).value = name
    for i, raw in enumerate(rows, start=2):
        ws.cell(i, 1).value = raw["Date"]
        ws.cell(i, 2).value = raw["Desk"]
        ws.cell(i, 3).value = raw.get("Area", "A")
        ws.cell(i, 4).value = raw.get("AM Employee")
        ws.cell(i, 5).value = raw.get("PM Employee")
        ws.cell(i, 6).value = raw.get("AM Status", "Available")
        ws.cell(i, 7).value = raw.get("PM Status", "Available")
    return workbook_to_bytes(wb)
