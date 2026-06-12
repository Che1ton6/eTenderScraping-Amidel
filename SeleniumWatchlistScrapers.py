#!/usr/bin/env python3
"""
Selenium-based watchlist scrapers for sites that require JavaScript rendering.
Currently covers:
  - Amahlathi LM  (JS tab filtering)
  - City Power JHB (JS-rendered, SSL issue)
"""

import logging
import os
import re
import time
from datetime import datetime, date
from typing import Optional, List

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from WatchlistScrapers import _parse_date, _parse_closing, _infer_type


def _make_driver(ignore_ssl: bool = False) -> webdriver.Chrome:
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    if ignore_ssl:
        opts.add_argument("--ignore-certificate-errors")
        opts.add_argument("--allow-insecure-localhost")
    return webdriver.Chrome(options=opts)


def _blank(report_date, department, province, source_key) -> dict:
    return {
        "REPORT_DATE": report_date, "RECORD_ID": None,
        "TENDER_ID": "", "PUBLICATION_DATE": "", "CLOSING_DATE": "",
        "CLOSING_TIME": "", "TENDER_TYPE": "Request for Bid",
        "TENDER_DESCRIPTION": "", "TENDER_SOURCE": source_key,
        "DEPARTMENT": department, "PROVINCE": province,
        "ESUBMISSION": "", "CATEGORY": "", "IS_THERE_A_BRIEFING_SESSION": "",
        "BRIEFING_DATE": "", "COMPULSORY_BRIEFING": "", "BRIEFING_SESSION_VENUE": "",
        "LINK": "", "SOE": "", "COST_OF_SALES_ESTIMATE": "",
        "CAPABILITY_AVAILABLE": "", "CAPABILITY_GROUP": "", "REQUIREMENTS": "",
    }


# ── Amahlathi LM ─────────────────────────────────────────────────────────────

def scrape_amahlathi(date_from: str, date_to: str, log_queue=None) -> List[dict]:
    """
    https://amahlathi.gov.za/tenders-rfqs/
    Clicks the Open Tenders tab then scrapes visible cards.
    """
    report_date = date_to.replace("-", "/")
    pub_date_str = datetime.strptime(date_to, "%Y-%m-%d").date().strftime("%Y/%m/%d")
    tenders = []
    driver = None

    logging.info("Amahlathi LM: launching browser")
    try:
        driver = _make_driver()
        driver.get("https://amahlathi.gov.za/tenders-rfqs/")
        time.sleep(5)

        # Try to click an "Open" tab/filter button
        for text in ("Open", "open", "Open Tenders", "Current"):
            try:
                btn = driver.find_element(
                    By.XPATH,
                    f"//a[normalize-space()='{text}'] | //button[normalize-space()='{text}'] | "
                    f"//li[normalize-space()='{text}'] | //*[@data-filter='.open']",
                )
                btn.click()
                time.sleep(3)
                logging.info(f"Amahlathi: clicked '{text}' tab")
                break
            except NoSuchElementException:
                continue

        # Collect all visible card headings
        cards = driver.find_elements(By.CSS_SELECTOR, "article, .tribe_events_cat-tenders, .post")
        if not cards:
            cards = driver.find_elements(By.XPATH, "//h3/.. | //h2/..")

        seen = set()
        for card in cards:
            try:
                heading = None
                for tag in ("h3", "h2", "h4"):
                    try:
                        heading = card.find_element(By.TAG_NAME, tag)
                        break
                    except NoSuchElementException:
                        continue
                if not heading:
                    continue

                desc = heading.text.strip()
                if not desc or desc in seen:
                    continue
                seen.add(desc)

                t = _blank(report_date, "Amahlathi Local Municipality", "Eastern Cape",
                           "AMAHLATHI.GOV.ZA")
                t["TENDER_DESCRIPTION"] = desc
                t["PUBLICATION_DATE"]   = pub_date_str
                t["TENDER_TYPE"]        = _infer_type(desc)

                # Link
                try:
                    a = card.find_element(By.XPATH, ".//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'detail') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'view')]")
                    t["LINK"] = a.get_attribute("href") or ""
                except NoSuchElementException:
                    try:
                        a = card.find_element(By.TAG_NAME, "a")
                        t["LINK"] = a.get_attribute("href") or ""
                    except NoSuchElementException:
                        pass

                # Closing date from card text
                card_text = card.text
                close_m = re.search(
                    r"[Cc]los(?:ing|ed?)\s*(?:[Dd]ate)?:?\s*"
                    r"(\d{1,2}\s+[A-Za-z]+\s+\d{4}|\d{1,2}/\d{1,2}/\d{4})"
                    r"(?:\s*[-–]\s*(\d{1,2}:\d{2}(?:\s*[AaPp][Mm])?))?",
                    card_text,
                )
                if close_m:
                    cd = _parse_date(close_m.group(1))
                    if cd:
                        t["CLOSING_DATE"] = cd.strftime("%Y/%m/%d")
                    t["CLOSING_TIME"] = close_m.group(2) or ""

                tenders.append(t)
            except Exception as e:
                logging.debug(f"Amahlathi card error: {e}")

    except Exception as e:
        logging.error(f"Amahlathi LM scraper error: {e}")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    logging.info(f"Amahlathi LM: {len(tenders)} tender(s)")
    return tenders


# ── City Power JHB ────────────────────────────────────────────────────────────

def scrape_city_power(date_from: str, date_to: str, log_queue=None) -> List[dict]:
    """
    https://www.citypower.co.za/tender-bulletin/open-tenders
    SSL cert issue — uses --ignore-certificate-errors.
    """
    report_date = date_to.replace("-", "/")
    date_from_d = datetime.strptime(date_from, "%Y-%m-%d").date()
    date_to_d   = datetime.strptime(date_to,   "%Y-%m-%d").date()
    pub_proxy   = date_to_d.strftime("%Y/%m/%d")
    tenders = []
    driver = None

    logging.info("City Power JHB: launching browser (SSL errors ignored)")
    try:
        driver = _make_driver(ignore_ssl=True)
        driver.get("https://www.citypower.co.za/tender-bulletin/open-tenders")

        # Wait for tender content to load (table or list)
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//table | //article | //li[contains(@class,'tender')]")
                )
            )
        except TimeoutException:
            logging.warning("City Power: timed out waiting for content — scraping what loaded")
        time.sleep(3)

        page_source = driver.page_source
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(page_source, "html.parser")

        # Try table first
        table = soup.find("table")
        if table:
            rows = table.find_all("tr")
            header = [th.get_text(strip=True).lower()
                      for th in (rows[0].find_all(["th", "td"]) if rows else [])]

            def _col(*keys):
                for k in keys:
                    for i, h in enumerate(header):
                        if k in h:
                            return i
                return None

            id_idx    = _col("tender no", "number", "ref")
            desc_idx  = _col("description")
            pub_idx   = _col("advert", "publish", "date")
            close_idx = _col("closing", "close")

            for tr in rows[1:]:
                tds = tr.find_all("td")
                if not tds:
                    continue
                t = _blank(report_date, "City Power Johannesburg", "Gauteng",
                           "CITYPOWER.CO.ZA")

                def _val(idx):
                    return tds[idx].get_text(strip=True) if idx is not None and idx < len(tds) else ""

                t["TENDER_ID"]          = _val(id_idx)
                t["TENDER_DESCRIPTION"] = _val(desc_idx) or _val(1)
                if not t["TENDER_DESCRIPTION"]:
                    continue

                pub = _parse_date(_val(pub_idx)) if pub_idx is not None else None
                t["PUBLICATION_DATE"] = pub.strftime("%Y/%m/%d") if pub else pub_proxy

                if close_idx is not None:
                    t["CLOSING_DATE"], t["CLOSING_TIME"] = _parse_closing(_val(close_idx))

                a = tr.find("a", href=True)
                if a:
                    t["LINK"] = a["href"]

                t["TENDER_TYPE"] = _infer_type(t["TENDER_DESCRIPTION"])
                tenders.append(t)
        else:
            # Fallback: generic card/list extraction
            for card in soup.select("article, .tender, .post, li"):
                heading = card.find(["h2", "h3", "h4"])
                if not heading:
                    continue
                desc = heading.get_text(strip=True)
                if not desc:
                    continue
                t = _blank(report_date, "City Power Johannesburg", "Gauteng",
                           "CITYPOWER.CO.ZA")
                t["TENDER_DESCRIPTION"] = desc
                t["PUBLICATION_DATE"]   = pub_proxy
                a = card.find("a", href=True)
                if a:
                    t["LINK"] = a["href"]
                t["TENDER_TYPE"] = _infer_type(desc)

                text = card.get_text(" ", strip=True)
                close_m = re.search(
                    r"[Cc]los\w*\s*:?\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4}|\d{1,2}/\d{1,2}/\d{4})",
                    text,
                )
                if close_m:
                    t["CLOSING_DATE"], t["CLOSING_TIME"] = _parse_closing(close_m.group(1))

                tenders.append(t)

    except Exception as e:
        logging.error(f"City Power JHB scraper error: {e}")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    logging.info(f"City Power JHB: {len(tenders)} tender(s)")
    return tenders


# ── Registry ──────────────────────────────────────────────────────────────────

SELENIUM_REGISTRY = {
    "Amahlathi LM": scrape_amahlathi,
    "CP JHB":       scrape_city_power,
}


def run_selenium_watchlist_scrapers(date_from: str, date_to: str,
                                    watchlist_sources: set,
                                    log_queue=None) -> list:
    all_tenders = []
    for src_name, fn in SELENIUM_REGISTRY.items():
        if src_name not in watchlist_sources:
            continue
        logging.info(f"Selenium watchlist: starting {src_name}")
        try:
            tenders = fn(date_from, date_to, log_queue)
            all_tenders.extend(tenders)
        except Exception as e:
            logging.error(f"Selenium watchlist {src_name} error: {e}")
    return all_tenders
