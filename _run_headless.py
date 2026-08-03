"""
Headless interactive entry point for the eTender Scraper in Docker / Azure.
Prompts the user for scrape options, then runs the full pipeline.
The GUI (main.py) is unaffected — this file is only used by Docker.
"""

import json
import logging
import os
import sys
from datetime import date, datetime, timedelta

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)

# ── Batch date auto-calculation (mirrors GUI logic) ───────────────────────────
def suggested_batch(today: date) -> tuple:
    """Return (batch_type, date_from, date_to) based on today's weekday."""
    wd = today.weekday()  # Mon=0 … Sun=6
    if wd == 0:           # Monday → M batch (Thu–Sun of previous week)
        batch_type = "M"
        thursday = today - timedelta(days=4)
        return batch_type, thursday, thursday + timedelta(days=3)
    else:                 # Any other day → T batch (Mon–Wed of current week)
        batch_type = "T"
        monday = today - timedelta(days=wd)
        return batch_type, monday, monday + timedelta(days=2)

# ── Prompt helpers ────────────────────────────────────────────────────────────
def prompt_choice(question: str, options: list) -> str:
    print(f"\n{question}")
    for i, (label, _) in enumerate(options, 1):
        print(f"  {i}. {label}")
    while True:
        raw = input("Enter number: ").strip()
        if raw.isdigit() and 1 <= int(raw) <= len(options):
            return options[int(raw) - 1][1]
        print(f"  Please enter a number between 1 and {len(options)}.")

def prompt_date(prompt: str, default: str) -> str:
    while True:
        raw = input(f"{prompt} [{default}]: ").strip() or default
        try:
            datetime.strptime(raw, "%Y-%m-%d")
            return raw
        except ValueError:
            print("  Invalid date. Use YYYY-MM-DD format.")

# ── Interactive prompts ───────────────────────────────────────────────────────
print("\n" + "="*50)
print("  Amidel eTender Scraper — Docker Runner")
print("="*50)

today = date.today()
auto_type, auto_from, auto_to = suggested_batch(today)
auto_from_str = auto_from.strftime("%Y-%m-%d")
auto_to_str   = auto_to.strftime("%Y-%m-%d")
auto_label    = f"{auto_from.strftime('%-d %b')}–{auto_to.strftime('%-d %b %Y')}" \
                if sys.platform != "win32" else \
                f"{auto_from.strftime('%d %b').lstrip('0')}–{auto_to.strftime('%d %b %Y').lstrip('0')}"

print(f"\n  Suggested batch: ({auto_type}) {auto_label}")

mode = prompt_choice("What would you like to scrape?", [
    ("eTenders.gov.za only",          "etenders"),
    ("All but eTenders (watchlist)",  "watchlist"),
    ("Full batch (both sources)",     "full"),
    ("Cybersecurity tenders only",    "cyber"),
])

batch_type = prompt_choice("Batch type?", [
    (f"T — Tuesday batch (Mon–Wed)  [suggested: {'✓' if auto_type == 'T' else ' '}]", "T"),
    (f"M — Monday batch  (Thu–Sun)  [suggested: {'✓' if auto_type == 'M' else ' '}]", "M"),
])

date_from = prompt_date("Date from (YYYY-MM-DD)", auto_from_str)
date_to   = prompt_date("Date to   (YYYY-MM-DD)", auto_to_str)

print(f"\nStarting: mode={mode} | batch={batch_type} | {date_from} → {date_to}\n")

# ── Patch config.json ─────────────────────────────────────────────────────────
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")
with open(CONFIG_PATH) as f:
    config = json.load(f)

config.setdefault("scraping", {})
config["scraping"]["dateFrom"] = date_from
config["scraping"]["dateTo"]   = date_to
config.setdefault("browser", {})
config["browser"]["headless"]  = True

with open(CONFIG_PATH, "w") as f:
    json.dump(config, f, indent=2)

logging.info("Config patched — headless Chrome enabled")

# ── Imports ───────────────────────────────────────────────────────────────────
import pandas as pd
from BatchProcessor import (
    create_batch_folder, save_daily_file, create_end_product,
    update_equation_file, calculate_counts, update_power_bi_export,
    update_master_tenders, deduplicate_tenders,
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
    logging.info("Running eTenders.gov.za scraper...")
    batch_folder = create_batch_folder(date_from, date_to, batch_type,
                                       root_dir=os.path.join("data", "etenders.gov.za"))
    scraper = TenderScraper(CONFIG_PATH)
    scraper.run(export=False)
    etender_tenders = scraper.tenderData
    save_daily_file(etender_tenders, date_to, batch_folder)
    all_tenders.extend(etender_tenders)
    logging.info(f"eTenders: {len(etender_tenders)} tenders scraped")

# ── Watchlist scrapers ────────────────────────────────────────────────────────
if mode in ("watchlist", "full"):
    from WatchlistScrapers import run_watchlist_scrapers, SCRAPER_REGISTRY
    logging.info("Running watchlist scrapers...")
    watch_folder = create_batch_folder(date_from, date_to, batch_type,
                                       root_dir=os.path.join("data", "all_but_etenders"))
    all_sources = set(SCRAPER_REGISTRY.keys())
    watch_tenders = run_watchlist_scrapers(date_from, date_to, all_sources)
    save_daily_file(watch_tenders, date_to, watch_folder)
    all_tenders.extend(watch_tenders)
    logging.info(f"Watchlist: {len(watch_tenders)} tenders scraped")

# ── Cybersecurity filter only ─────────────────────────────────────────────────
if mode == "cyber":
    from TenderScraper import TenderScraper
    logging.info("Running eTenders.gov.za scraper for cybersecurity filter...")
    batch_folder = create_batch_folder(date_from, date_to, batch_type,
                                       root_dir=os.path.join("data", "etenders.gov.za"))
    scraper = TenderScraper(CONFIG_PATH)
    scraper.run(export=False)
    all_tenders.extend(scraper.tenderData)

# ── Dedup + outputs ───────────────────────────────────────────────────────────
if all_tenders:
    deduped = deduplicate_tenders(all_tenders)
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

    logging.info(f"Done. {len(deduped)} tenders processed. Output: {output_folder}")
else:
    logging.warning("No tenders scraped.")
