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

from WatchlistScrapers import _parse_date, _parse_closing, _infer_type, wp_pub_date_from_url


def _make_driver(ignore_ssl: bool = False) -> webdriver.Chrome:
    opts = Options()
    # opts.add_argument("--headless=new")  # visible mode — user requested to watch scraping
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-renderer-backgrounding")
    opts.add_argument("--disable-background-timer-throttling")
    opts.add_argument("--disable-backgrounding-occluded-windows")
    opts.add_argument("--remote-debugging-port=0")
    opts.add_argument("--disable-extensions")
    if ignore_ssl:
        opts.add_argument("--ignore-certificate-errors")
        opts.add_argument("--allow-insecure-localhost")
    driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(60)
    return driver


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
                # Try tribe events title class first (The Events Calendar plugin)
                for css in (".tribe-events-list-event-title",
                            ".tribe-event-title",
                            ".entry-title"):
                    try:
                        heading = card.find_element(By.CSS_SELECTOR, css)
                        break
                    except NoSuchElementException:
                        continue
                # Fall back to standard heading tags
                if not heading:
                    for tag in ("h3", "h2", "h4"):
                        try:
                            heading = card.find_element(By.TAG_NAME, tag)
                            break
                        except NoSuchElementException:
                            continue

                desc = ""
                if heading:
                    # Tribe events wraps title in an <a>; get the anchor text first
                    try:
                        a_el = heading.find_element(By.TAG_NAME, "a")
                        desc = a_el.text.strip() or a_el.get_attribute("title") or ""
                    except NoSuchElementException:
                        desc = heading.text.strip()

                # Last resort: first anchor in the card
                if not desc:
                    try:
                        a_el = card.find_element(By.TAG_NAME, "a")
                        desc = a_el.text.strip() or a_el.get_attribute("title") or ""
                    except NoSuchElementException:
                        pass

                if not desc or desc in seen:
                    continue
                seen.add(desc)

                t = _blank(report_date, "Amahlathi Local Municipality", "Eastern Cape",
                           "AMAHLATHI.GOV.ZA")
                t["TENDER_DESCRIPTION"] = desc
                t["PUBLICATION_DATE"]   = ""
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

                # Fallback publication date from URL patterns (WP /uploads/YYYY/MM/, /YYYY/MM/DD/, /YYYY/MM/)
                if not t["PUBLICATION_DATE"] and t["LINK"]:
                    url_pub = wp_pub_date_from_url(t["LINK"], date_from, date_to)
                    if url_pub:
                        try:
                            pd_ = datetime.strptime(url_pub, "%Y/%m/%d").date()
                            df_ = datetime.strptime(date_from, "%Y-%m-%d").date()
                            dt_ = datetime.strptime(date_to,   "%Y-%m-%d").date()
                            if not (df_ <= pd_ <= dt_):
                                continue
                            t["PUBLICATION_DATE"] = url_pub
                        except ValueError:
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

    # Second pass: scrape the Events Calendar category page directly so that
    # re-adverts (published under /event/ URLs, outside the "Open" tab) are not missed.
    events_driver = None
    try:
        events_driver = _make_driver()
        events_driver.get("https://amahlathi.gov.za/events/category/tenders/list/")
        time.sleep(5)

        from bs4 import BeautifulSoup as _BS4
        soup2 = _BS4(events_driver.page_source, "html.parser")

        existing_descs = {t["TENDER_DESCRIPTION"].lower() for t in tenders}
        for article in soup2.select("article, .tribe-event, .tribe-events-list-event"):
            try:
                heading = (
                    article.find(class_="tribe-events-list-event-title")
                    or article.find(class_="tribe-event-title")
                    or article.find(["h2", "h3", "h4"])
                )
                if not heading:
                    continue
                a_el = heading.find("a")
                desc = (a_el.get_text(strip=True) if a_el else heading.get_text(strip=True))
                if not desc or desc.lower() in existing_descs:
                    continue
                existing_descs.add(desc.lower())

                t2 = _blank(report_date, "Amahlathi Local Municipality", "Eastern Cape",
                            "AMAHLATHI.GOV.ZA")
                t2["TENDER_DESCRIPTION"] = desc
                t2["TENDER_TYPE"]        = _infer_type(desc)
                if a_el:
                    t2["LINK"] = a_el.get("href", "")

                text2 = article.get_text(" ", strip=True)
                close_m2 = re.search(
                    r"[Cc]los(?:ing|ed?)\s*(?:[Dd]ate)?:?\s*"
                    r"(\d{1,2}\s+[A-Za-z]+\s+\d{4}|\d{1,2}/\d{1,2}/\d{4})"
                    r"(?:\s*[-–]\s*(\d{1,2}:\d{2}(?:\s*[AaPp][Mm])?))?",
                    text2,
                )
                if close_m2:
                    cd2 = _parse_date(close_m2.group(1))
                    if cd2:
                        t2["CLOSING_DATE"] = cd2.strftime("%Y/%m/%d")
                    t2["CLOSING_TIME"] = close_m2.group(2) or ""

                tenders.append(t2)
            except Exception as e:
                logging.debug(f"Amahlathi events-page card error: {e}")

    except Exception as e:
        logging.warning(f"Amahlathi events-page scrape failed: {e}")
    finally:
        if events_driver:
            try:
                events_driver.quit()
            except Exception:
                pass

    logging.info(f"Amahlathi LM: {len(tenders)} tender(s)")
    return tenders


# ── Matatiele LM ─────────────────────────────────────────────────────────────

def scrape_matatiele(date_from: str, date_to: str, log_queue=None) -> List[dict]:
    """
    https://www.matatiele.gov.za/tenders/
    WP Job Manager renders listings via JavaScript — requires full browser.
    """
    report_date = date_to.replace("-", "/")
    tenders = []
    driver = None

    logging.info("Matatiele LM: launching browser (WP Job Manager JS rendering)")
    try:
        driver = _make_driver()
        driver.get("https://www.matatiele.gov.za/tenders/")

        # Wait for job listings to appear
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "ul.jobs, .job_listings, li.job_listing")
                )
            )
        except TimeoutException:
            logging.warning("Matatiele: timed out waiting for job listings — scraping what loaded")
        time.sleep(3)

        # WP Job Manager "Load more jobs" pagination — click until exhausted
        for _attempt in range(40):
            try:
                load_more = driver.find_element(
                    By.CSS_SELECTOR,
                    "a.load_more_jobs, .load_more_jobs, a[class*='load_more'], "
                    "button.load_more_jobs, .job_listings_pagination a",
                )
                if load_more.is_displayed() and load_more.is_enabled():
                    driver.execute_script("arguments[0].click();", load_more)
                    time.sleep(3)
                    logging.info(f"Matatiele: loaded more jobs (attempt {_attempt + 1})")
                else:
                    break
            except NoSuchElementException:
                break
            except Exception as e:
                logging.debug(f"Matatiele load-more error: {e}")
                break

        from bs4 import BeautifulSoup
        soup = BeautifulSoup(driver.page_source, "html.parser")

        listings = soup.select("li.job_listing, article.job_listing")
        if not listings:
            listings = soup.select("ul.jobs li, .job_listings li")

        seen = set()
        for item in listings:
            heading = item.find(["h3", "h2", "h4"])
            if not heading:
                continue
            desc = heading.get_text(strip=True)
            if not desc or desc in seen:
                continue
            seen.add(desc)

            t = _blank(report_date, "Matatiele Local Municipality", "Eastern Cape",
                       "MATATIELE.GOV.ZA")
            t["TENDER_DESCRIPTION"] = desc
            t["TENDER_TYPE"]        = _infer_type(desc)

            a = item.find("a", href=True)
            link = a["href"] if a else ""
            if link:
                t["LINK"] = link

            text = item.get_text(" ", strip=True)
            close_m = re.search(
                r"(?:[Ee]xpir|[Cc]los)\w*\s*:?\s*"
                r"(\d{1,2}\s+[A-Za-z]+\s+\d{4}|\d{1,2}/\d{1,2}/\d{4}|[A-Za-z]+\s+\d{1,2},?\s+\d{4})",
                text,
            )
            if close_m:
                cd = _parse_date(close_m.group(1))
                if cd:
                    t["CLOSING_DATE"] = cd.strftime("%Y/%m/%d")

            # Fetch individual page for publication date ("Posted on June 11, 2026")
            if link:
                try:
                    import requests as _req
                    from bs4 import BeautifulSoup as _BS
                    resp = _req.get(link, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
                    if resp.ok:
                        detail = _BS(resp.text, "html.parser")
                        posted = detail.find(string=re.compile(r"Posted on", re.I))
                        if posted:
                            pub_m = re.search(
                                r"Posted\s+on\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4}|\d{1,2}\s+[A-Za-z]+\s+\d{4})",
                                str(posted.parent), re.I
                            )
                            if pub_m:
                                pd = _parse_date(pub_m.group(1))
                                if pd:
                                    t["PUBLICATION_DATE"] = pd.strftime("%Y/%m/%d")
                except Exception as _e:
                    logging.debug(f"Matatiele: could not fetch detail page {link}: {_e}")

            # Fallback publication date via URL patterns
            if not t["PUBLICATION_DATE"] and link:
                url_pub = wp_pub_date_from_url(link, date_from, date_to)
                if url_pub:
                    try:
                        pd_ = datetime.strptime(url_pub, "%Y/%m/%d").date()
                        df_ = datetime.strptime(date_from, "%Y-%m-%d").date()
                        dt_ = datetime.strptime(date_to,   "%Y-%m-%d").date()
                        if not (df_ <= pd_ <= dt_):
                            continue
                        t["PUBLICATION_DATE"] = url_pub
                    except ValueError:
                        pass

            tenders.append(t)

    except Exception as e:
        logging.error(f"Matatiele LM scraper error: {e}")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    logging.info(f"Matatiele LM: {len(tenders)} tender(s)")
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
                t["PUBLICATION_DATE"] = pub.strftime("%Y/%m/%d") if pub else ""

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
                t["PUBLICATION_DATE"]   = ""
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
    "Matatiele LM": scrape_matatiele,
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
