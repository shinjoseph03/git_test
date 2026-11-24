"""CLI entry: python -m desk_bot --mode book|clear."""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import date
from pathlib import Path

from desk_bot.booking import book_for_date, clear_for_dates, today_in_tz
from desk_bot.config import load_config
from desk_bot.graph import GraphClient, GraphCredentials, GraphError, acquire_token

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
log = logging.getLogger("desk_bot")


def _parse_dates(raw: list[str] | None) -> list[date]:
    if not raw:
        return []
    return [date.fromisoformat(s.strip()) for s in raw if s.strip()]


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Desk reservation SharePoint bot")
    p.add_argument(
        "--mode",
        choices=("book", "clear"),
        default="book",
        help="book preferred desk for today, or clear configured name on dates",
    )
    p.add_argument(
        "--config",
        default=None,
        help="Path to config.yaml (default: CONFIG_PATH or config.yaml)",
    )
    p.add_argument(
        "--date",
        default=None,
        help="Override 'today' for book mode (YYYY-MM-DD)",
    )
    p.add_argument(
        "--clear-date",
        action="append",
        dest="clear_dates",
        default=None,
        help="Date to clear (repeatable). Overrides config clear_dates when set.",
    )
    p.add_argument(
        "--fixture",
        default=None,
        help="Local xlsx path for dry-run / offline (skips Graph)",
    )
    p.add_argument(
        "--write-fixture",
        default=None,
        help="When using --fixture, write updated workbook here",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Download (or use fixture), compute changes, do not upload",
    )
    return p


def run(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    cfg = load_config(args.config)
    if args.dry_run:
        cfg.dry_run = True

    if args.fixture:
        workbook = Path(args.fixture).read_bytes()
        site_id = item_id = None
        client = None
    else:
        try:
            creds = GraphCredentials.from_env()
            token = acquire_token(creds)
            client = GraphClient(token)
            site_id = client.resolve_site_id(
                cfg.sharepoint.site_hostname, cfg.sharepoint.site_path
            )
            item_id = cfg.sharepoint.drive_item_id
            if not item_id:
                raise GraphError(
                    "SHAREPOINT_DRIVE_ITEM_ID / sharepoint.drive_item_id is required"
                )
            log.info(
                "Downloading drive item %s from site %s",
                item_id,
                cfg.sharepoint.site_path,
            )
            workbook = client.download_drive_item(site_id, item_id)
        except GraphError as exc:
            log.error("%s", exc)
            return 2

    if args.mode == "book":
        on = today_in_tz(cfg.timezone, args.date)
        updated, messages = book_for_date(cfg, workbook, on)
    else:
        clear_list = _parse_dates(args.clear_dates) or [
            date.fromisoformat(d) for d in cfg.clear_dates
        ]
        updated, messages = clear_for_dates(cfg, workbook, clear_list)

    for msg in messages:
        level = logging.ERROR if msg.startswith("error:") else logging.INFO
        log.log(level, "%s", msg)

    changed = updated != workbook
    if args.fixture and args.write_fixture and changed:
        Path(args.write_fixture).write_bytes(updated)
        log.info("Wrote fixture output to %s", args.write_fixture)

    if not changed:
        log.info("No workbook changes to upload")
        return 0

    if cfg.dry_run or args.fixture:
        log.info("Dry-run / fixture mode: skipping upload (%d bytes would change)", len(updated))
        return 0

    assert client is not None and site_id is not None and item_id is not None
    try:
        log.info("Uploading updated workbook")
        client.upload_drive_item(site_id, item_id, updated)
    except GraphError as exc:
        # Locked / blocked cells or file lock — skip, don't crash the schedule
        log.warning("skip_upload_failed: %s", exc)
        return 0

    log.info("Upload complete")
    return 0


if __name__ == "__main__":
    sys.exit(run())
