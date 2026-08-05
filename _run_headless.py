"""
Headless entry point for the eTender Scraper in Docker / Azure.

Runs unattended by default — no stdin required. Configuration comes from
(in order of precedence): CLI args > environment variables > auto-defaults
derived from today's weekday via suggested_batch().

Environment variables:
    SCRAPE_MODE   one of: etenders | watchlist | full | cyber   (default: etenders)
    BATCH_TYPE    T | M                                          (default: auto from today)
    DATE_FROM     YYYY-MM-DD                                     (default: auto from today)
    DATE_TO       YYYY-MM-DD                                     (default: auto from today)
    TEST_MODE     1 / true to enable test mode (see below)       (default: off)

CLI:
    python _run_headless.py [--mode MODE] [--batch T|M]
                            [--from YYYY-MM-DD] [--to YYYY-MM-DD]
                            [--test]

Test mode (--test or TEST_MODE=1):
    Chrome runs VISIBLE (not headless) so you can watch the scrape.
    Each config prompt is shown with a 10-second timeout; if no input is
    given, the auto-default is used. Intended for local sanity-checks only —
    Azure/Docker scheduled runs should never use this flag.

The GUI (main.py) is unaffected — this file is only used by Docker / Azure.
"""

import argparse
import json
import logging
import os
import sys
import time
from datetime import date, datetime, timedelta

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except (AttributeError, ValueError):
    pass

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("etender.headless")

VALID_MODES = ("etenders", "watchlist", "full", "cyber")
VALID_BATCH_TYPES = ("T", "M")
PROMPT_TIMEOUT_SEC = 10.0


# ── Batch date auto-calculation ───────────────────────────────────────────────
def suggested_batch(today: date) -> tuple:
    """Return (batch_type, date_from, date_to) — the most recent COMPLETED batch.

    A batch only becomes eligible once its final day has ended:
        Mon/Tue/Wed -> M batch (Thu-Sun of previous week — completed last Sun)
        Thu/Fri/Sat/Sun -> T batch (Mon-Wed of this week — completed Wed night)

    Scheduled runs on Mon 06:00 and Thu 06:00 land on exactly the right batch.
    """
    wd = today.weekday()  # Mon=0 … Sun=6
    if wd <= 2:  # Mon/Tue/Wed
        last_sunday = today - timedelta(days=wd + 1)
        thursday = last_sunday - timedelta(days=3)
        return "M", thursday, last_sunday
    # Thu/Fri/Sat/Sun
    monday = today - timedelta(days=wd)
    return "T", monday, monday + timedelta(days=2)


# ── Timed prompt (cross-platform, degrades gracefully without a TTY) ──────────
def _timed_input(prompt: str, default: str, timeout: float = PROMPT_TIMEOUT_SEC) -> str:
    """Prompt for input with a timeout. Returns default on empty/timeout/no-tty."""
    if not sys.stdin.isatty():
        print(f"{prompt}[no tty - using default: {default}]", flush=True)
        return default

    print(prompt, end="", flush=True)

    if sys.platform == "win32":
        import msvcrt
        buf: list = []
        deadline = time.monotonic() + timeout
        typing_started = False
        while True:
            if msvcrt.kbhit():
                if not typing_started:
                    typing_started = True  # first keypress cancels the timer
                ch = msvcrt.getwche()
                if ch in ("\r", "\n"):
                    print()
                    return "".join(buf).strip() or default
                if ch == "\x08":  # backspace
                    if buf:
                        buf.pop()
                        print(" \b", end="", flush=True)
                elif ch == "\x03":
                    raise KeyboardInterrupt
                else:
                    buf.append(ch)
            else:
                if not typing_started and time.monotonic() > deadline:
                    print(f"\n  [timeout - using default: {default}]", flush=True)
                    return default
                time.sleep(0.05)

    # POSIX
    import select
    rlist, _, _ = select.select([sys.stdin], [], [], timeout)
    if rlist:
        return sys.stdin.readline().strip() or default
    print(f"\n  [timeout - using default: {default}]", flush=True)
    return default


def _prompt_choice(question: str, options: list, default_value: str) -> str:
    """Show numbered options, return the value chosen (or default on timeout)."""
    default_idx = next((i for i, (_, v) in enumerate(options, 1) if v == default_value), 1)
    print(f"\n{question}  [default: {default_value}]")
    for i, (label, _) in enumerate(options, 1):
        marker = "*" if i == default_idx else " "
        print(f"  {marker} {i}. {label}")
    raw = _timed_input(f"Enter 1-{len(options)} ({PROMPT_TIMEOUT_SEC:g}s): ", str(default_idx))
    if raw.isdigit() and 1 <= int(raw) <= len(options):
        return options[int(raw) - 1][1]
    print(f"  invalid choice '{raw}' - using default: {default_value}")
    return default_value


def _prompt_date(question: str, default_value: str) -> str:
    raw = _timed_input(f"{question} [{default_value}]: ", default_value)
    try:
        datetime.strptime(raw, "%Y-%m-%d")
        return raw
    except ValueError:
        print(f"  invalid date '{raw}' - using default: {default_value}")
        return default_value


def _parse_date(value: str, field: str) -> str:
    try:
        datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        raise SystemExit(f"[config] {field} must be YYYY-MM-DD, got: {value!r}")
    return value


def _env_bool(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in ("1", "true", "yes", "y", "on")


def resolve_config() -> dict:
    parser = argparse.ArgumentParser(
        prog="etender-headless",
        description="Run the eTender scraper unattended (no stdin) or in local --test mode.",
    )
    parser.add_argument("--mode", choices=VALID_MODES, help="scrape mode")
    parser.add_argument("--batch", choices=VALID_BATCH_TYPES, help="batch type")
    parser.add_argument("--from", dest="date_from", help="start date YYYY-MM-DD")
    parser.add_argument("--to", dest="date_to", help="end date YYYY-MM-DD")
    parser.add_argument("--test", action="store_true",
                        help="local test mode: visible Chrome + 10s timed prompts falling to defaults")
    args = parser.parse_args()

    test_mode = args.test or _env_bool("TEST_MODE")

    today = date.today()
    auto_type, auto_from, auto_to = suggested_batch(today)
    auto_from_str = auto_from.strftime("%Y-%m-%d")
    auto_to_str = auto_to.strftime("%Y-%m-%d")
    auto_mode = "etenders"

    # Non-interactive resolution (CLI > env > auto)
    mode = args.mode or os.environ.get("SCRAPE_MODE", auto_mode).strip().lower()
    batch_type = args.batch or os.environ.get("BATCH_TYPE", auto_type).strip().upper()
    date_from = args.date_from or os.environ.get("DATE_FROM", auto_from_str)
    date_to = args.date_to or os.environ.get("DATE_TO", auto_to_str)

    # In test mode, only prompt for fields that weren't explicitly set via CLI
    if test_mode:
        print("\n" + "=" * 60)
        print("  TEST MODE - visible Chrome, 10s timed prompts")
        print("=" * 60)
        if not args.mode:
            mode = _prompt_choice("Scrape mode?", [
                ("eTenders.gov.za only", "etenders"),
                ("Watchlist (all but eTenders)", "watchlist"),
                ("Full batch (both sources)", "full"),
                ("Cybersecurity tenders only", "cyber"),
            ], mode)
        if not args.batch:
            batch_type = _prompt_choice("Batch type?", [
                ("T - Thursday batch (Mon-Wed)", "T"),
                ("M - Monday batch   (Thu-Sun)", "M"),
            ], batch_type)
        if not args.date_from:
            date_from = _prompt_date("Date from (YYYY-MM-DD)", date_from)
        if not args.date_to:
            date_to = _prompt_date("Date to   (YYYY-MM-DD)", date_to)

    if mode not in VALID_MODES:
        raise SystemExit(f"[config] SCRAPE_MODE must be one of {VALID_MODES}, got: {mode!r}")
    if batch_type not in VALID_BATCH_TYPES:
        raise SystemExit(f"[config] BATCH_TYPE must be T or M, got: {batch_type!r}")
    date_from = _parse_date(date_from, "DATE_FROM")
    date_to = _parse_date(date_to, "DATE_TO")
    if date_from > date_to:
        raise SystemExit(f"[config] DATE_FROM ({date_from}) is after DATE_TO ({date_to})")

    return {
        "mode": mode,
        "batch_type": batch_type,
        "date_from": date_from,
        "date_to": date_to,
        "test_mode": test_mode,
    }


# ── Resolve config and log the run plan ───────────────────────────────────────
cfg = resolve_config()
mode = cfg["mode"]
batch_type = cfg["batch_type"]
date_from = cfg["date_from"]
date_to = cfg["date_to"]
test_mode = cfg["test_mode"]

log.info("=" * 60)
log.info("Amidel eTender Scraper - %s run", "TEST" if test_mode else "headless")
log.info("mode=%s | batch=%s | %s -> %s | chrome=%s",
         mode, batch_type, date_from, date_to,
         "visible" if test_mode else "headless")
log.info("=" * 60)

# ── Patch config.json ─────────────────────────────────────────────────────────
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")
with open(CONFIG_PATH) as f:
    config = json.load(f)

config.setdefault("scraping", {})
config["scraping"]["dateFrom"] = date_from
config["scraping"]["dateTo"] = date_to
config.setdefault("browser", {})
config["browser"]["headless"] = not test_mode

with open(CONFIG_PATH, "w") as f:
    json.dump(config, f, indent=2)

log.info("Config patched - headless=%s", not test_mode)

# ── Imports ───────────────────────────────────────────────────────────────────
import pandas as pd
from BatchProcessor import (
    create_batch_folder, save_daily_file, create_end_product,
    update_equation_file, calculate_counts, update_power_bi_export,
    update_master_tenders, merge_and_flag_duplicates,
)
from TenderSummary import create_tender_summary
from TenderAnalysisGenerator import create_tender_analysis
from CybersecurityTenders import create_cybersecurity_tenders

report_date = datetime.strptime(date_to, "%Y-%m-%d")
report_date_str = report_date.strftime("%d %b %Y").lstrip("0")

all_tenders: list = []

# ── eTenders.gov.za ───────────────────────────────────────────────────────────
if mode in ("etenders", "full"):
    from TenderScraper import TenderScraper
    log.info("Running eTenders.gov.za scraper...")
    batch_folder = create_batch_folder(date_from, date_to, batch_type,
                                       root_dir=os.path.join("data", "etenders.gov.za"))
    scraper = TenderScraper(CONFIG_PATH)
    scraper.run(export=False)
    etender_tenders = scraper.tenderData
    save_daily_file(etender_tenders, date_to, batch_folder)
    all_tenders.extend(etender_tenders)
    log.info("eTenders: %d tenders scraped", len(etender_tenders))

# ── Watchlist scrapers ────────────────────────────────────────────────────────
if mode in ("watchlist", "full"):
    from WatchlistScrapers import run_watchlist_scrapers, SCRAPER_REGISTRY
    log.info("Running watchlist scrapers...")
    watch_folder = create_batch_folder(date_from, date_to, batch_type,
                                       root_dir=os.path.join("data", "all_but_etenders"))
    all_sources = set(SCRAPER_REGISTRY.keys())
    watch_tenders = run_watchlist_scrapers(date_from, date_to, all_sources)
    save_daily_file(watch_tenders, date_to, watch_folder)
    all_tenders.extend(watch_tenders)
    log.info("Watchlist: %d tenders scraped", len(watch_tenders))

# ── Cybersecurity filter only ─────────────────────────────────────────────────
if mode == "cyber":
    from TenderScraper import TenderScraper
    log.info("Running eTenders.gov.za scraper for cybersecurity filter...")
    batch_folder = create_batch_folder(date_from, date_to, batch_type,
                                       root_dir=os.path.join("data", "etenders.gov.za"))
    scraper = TenderScraper(CONFIG_PATH)
    scraper.run(export=False)
    all_tenders.extend(scraper.tenderData)

# ── Dedup + outputs ───────────────────────────────────────────────────────────
if all_tenders:
    deduped = merge_and_flag_duplicates(all_tenders)
    df = pd.DataFrame(deduped)

    if mode == "etenders":
        output_folder = batch_folder
    elif mode == "watchlist":
        output_folder = watch_folder
    elif mode == "cyber":
        output_folder = batch_folder
    else:
        output_folder = create_batch_folder(date_from, date_to, batch_type,
                                            root_dir=os.path.join("data", "All_Tenders"))

    create_tender_summary(df, output_folder)
    create_tender_analysis(df, output_folder, report_date_str)
    create_cybersecurity_tenders(df, output_folder)

    if mode != "cyber":
        counts = calculate_counts(df)
        update_equation_file(counts, batch_type, report_date, output_folder)
        create_end_product(df, date_from, date_to, batch_type, report_date, output_folder)
        update_power_bi_export(output_folder, date_from, date_to, batch_type)
        update_master_tenders(output_folder)

    log.info("Done. %d tenders processed. Output: %s", len(deduped), output_folder)
    sys.exit(0)
else:
    log.warning("No tenders scraped.")
    sys.exit(0)
