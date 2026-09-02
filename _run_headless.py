"""
Headless entry point for the eTender Scraper in Docker / Azure App Service.

Runs the eTenders.gov.za scrape unattended. Configuration comes from
(in order of precedence): function args > CLI args > environment variables >
auto-defaults derived from today's weekday via suggested_batch().

Only the `etenders` scrape mode is supported here. Watchlist / cyber /
full-batch modes were retired 2026-08-27 (see project memory) and their
scrapers are no longer wired to this entrypoint.

Environment variables:
    BATCH_TYPE    T | M                                          (default: auto)
    DATE_FROM     YYYY-MM-DD                                     (default: auto)
    DATE_TO       YYYY-MM-DD                                     (default: auto)
    TEST_MODE     1 / true to enable test mode (visible Chrome)  (default: off)

CLI:
    python _run_headless.py [--batch T|M] [--from YYYY-MM-DD] [--to YYYY-MM-DD] [--test]

Programmatic (used by app.py Flask handler):
    from _run_headless import run_scrape
    summary = run_scrape(overrides={"batch_type": "T", ...})
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from datetime import date, datetime, timedelta
from typing import Optional

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except (AttributeError, ValueError):
    pass

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("etender.headless")

VALID_BATCH_TYPES = ("T", "M")
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")


def suggested_batch(today: date) -> tuple:
    """Return (batch_type, date_from, date_to) — the most recent COMPLETED batch."""
    wd = today.weekday()
    if wd <= 2:  # Mon/Tue/Wed
        last_sunday = today - timedelta(days=wd + 1)
        thursday = last_sunday - timedelta(days=3)
        return "M", thursday, last_sunday
    monday = today - timedelta(days=wd)
    return "T", monday, monday + timedelta(days=2)


def _parse_date(value: str, field: str) -> str:
    try:
        datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        raise ValueError(f"[config] {field} must be YYYY-MM-DD, got: {value!r}")
    return value


def _env_bool(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in ("1", "true", "yes", "y", "on")


def resolve_config(
    overrides: Optional[dict] = None,
    argv: Optional[list] = None,
) -> dict:
    """Resolve batch/date/test config from overrides > CLI > env > auto-defaults."""
    overrides = overrides or {}

    parser = argparse.ArgumentParser(
        prog="etender-headless",
        description="Run the eTenders.gov.za scraper unattended.",
    )
    parser.add_argument("--batch", choices=VALID_BATCH_TYPES)
    parser.add_argument("--from", dest="date_from")
    parser.add_argument("--to", dest="date_to")
    parser.add_argument("--test", action="store_true")
    args, _ = parser.parse_known_args(argv if argv is not None else [])

    today = date.today()
    auto_type, auto_from, auto_to = suggested_batch(today)

    batch_type = (
        overrides.get("batch_type")
        or args.batch
        or os.environ.get("BATCH_TYPE", auto_type)
    ).strip().upper()
    date_from = (
        overrides.get("date_from")
        or args.date_from
        or os.environ.get("DATE_FROM", auto_from.strftime("%Y-%m-%d"))
    )
    date_to = (
        overrides.get("date_to")
        or args.date_to
        or os.environ.get("DATE_TO", auto_to.strftime("%Y-%m-%d"))
    )
    test_mode = bool(overrides.get("test_mode") or args.test or _env_bool("TEST_MODE"))

    if batch_type not in VALID_BATCH_TYPES:
        raise ValueError(f"BATCH_TYPE must be T or M, got: {batch_type!r}")
    date_from = _parse_date(date_from, "DATE_FROM")
    date_to = _parse_date(date_to, "DATE_TO")
    if date_from > date_to:
        raise ValueError(f"DATE_FROM ({date_from}) is after DATE_TO ({date_to})")

    return {
        "batch_type": batch_type,
        "date_from": date_from,
        "date_to": date_to,
        "test_mode": test_mode,
    }


def _patch_config_json(date_from: str, date_to: str, headless: bool) -> None:
    with open(CONFIG_PATH) as f:
        config = json.load(f)
    config.setdefault("scraping", {})
    config["scraping"]["dateFrom"] = date_from
    config["scraping"]["dateTo"] = date_to
    config.setdefault("browser", {})
    config["browser"]["headless"] = headless
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f, indent=2)


def run_scrape(overrides: Optional[dict] = None) -> dict:
    """
    Execute the eTenders scrape end-to-end. Returns a JSON-serialisable summary.

    Intended for both CLI and Flask-handler callers. Callers that need
    SharePoint sync (download master before, upload after) should wrap this
    function — see app.py.
    """
    cfg = resolve_config(overrides=overrides, argv=[])
    batch_type = cfg["batch_type"]
    date_from = cfg["date_from"]
    date_to = cfg["date_to"]
    test_mode = cfg["test_mode"]

    log.info("=" * 60)
    log.info("Amidel eTender Scraper — %s run", "TEST" if test_mode else "headless")
    log.info("batch=%s | %s → %s | chrome=%s",
             batch_type, date_from, date_to,
             "visible" if test_mode else "headless")
    log.info("=" * 60)

    _patch_config_json(date_from, date_to, headless=not test_mode)

    import pandas as pd
    from BatchProcessor import (
        create_batch_folder, save_daily_file, create_end_product,
        update_equation_file, calculate_counts, update_power_bi_export,
        update_master_tenders, merge_and_flag_duplicates, MASTER_FILE,
    )
    from TenderSummary import create_tender_summary
    from TenderAnalysisGenerator import create_tender_analysis
    from TenderScraper import TenderScraper

    report_date = datetime.strptime(date_to, "%Y-%m-%d")
    report_date_str = report_date.strftime("%d %b %Y").lstrip("0")

    rows_before = 0
    if os.path.exists(MASTER_FILE):
        try:
            rows_before = len(pd.read_excel(MASTER_FILE, dtype=str))
        except Exception:
            rows_before = 0

    batch_folder = create_batch_folder(
        date_from, date_to, batch_type,
        root_dir=os.path.join("data", "etenders.gov.za"),
    )
    scraper = TenderScraper(CONFIG_PATH)
    scraper.run(export=False)
    etender_tenders = scraper.tenderData
    save_daily_file(etender_tenders, date_to, batch_folder)
    log.info("eTenders: %d tenders scraped", len(etender_tenders))

    if not etender_tenders:
        log.warning("No tenders scraped.")
        return {
            "status": "ok",
            "rows_scraped": 0,
            "rows_before": rows_before,
            "rows_after": rows_before,
            "rows_added": 0,
            "batch_folder": batch_folder,
        }

    deduped = merge_and_flag_duplicates(etender_tenders)
    df = pd.DataFrame(deduped)

    create_tender_summary(df, batch_folder)
    create_tender_analysis(df, batch_folder, report_date_str)

    counts = calculate_counts(df)
    update_equation_file(counts, batch_type, report_date, batch_folder)
    create_end_product(df, date_from, date_to, batch_type, report_date, batch_folder)
    update_power_bi_export(batch_folder, date_from, date_to, batch_type)
    update_master_tenders(batch_folder)

    rows_after = 0
    if os.path.exists(MASTER_FILE):
        try:
            rows_after = len(pd.read_excel(MASTER_FILE, dtype=str))
        except Exception:
            rows_after = 0

    log.info("Done. %d tenders processed. Output: %s", len(deduped), batch_folder)
    return {
        "status": "ok",
        "rows_scraped": len(deduped),
        "rows_before": rows_before,
        "rows_after": rows_after,
        "rows_added": max(0, rows_after - rows_before),
        "batch_folder": batch_folder,
    }


if __name__ == "__main__":
    try:
        summary = run_scrape()
    except Exception as e:
        log.exception("Scrape failed: %s", e)
        sys.exit(1)
    print(json.dumps(summary, indent=2))
    sys.exit(0)
