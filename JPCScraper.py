#!/usr/bin/env python3
"""
Scraper for Joburg Property Company (JPC) tenders.
Targets legacy.jhbproperty.co.za RFQ and RFP pages.
"""

import logging
import os
import re
from datetime import datetime, date
from typing import Optional, List

import requests
from bs4 import BeautifulSoup

DEPARTMENT    = "Joburg Property Company SOC Ltd"
TENDER_SOURCE = "JPC"
PROVINCE      = "Gauteng"

RFQ_URL = "https://legacy.jhbproperty.co.za/supply-chain-management-scm/rfqs/"
RFP_URL = "https://legacy.jhbproperty.co.za/supply-chain-management-scm/rfps/"

_PAGES = [
    (RFQ_URL, "Request for Quotation"),
    (RFP_URL, "Request for Proposal"),
]

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36"
    )
}


class JPCScraper:
    def __init__(
        self,
        date_from: str,
        date_to: str,
        log_queue=None,
    ):
        self._setup_logging(log_queue)
        self.date_from    = datetime.strptime(date_from, "%Y-%m-%d").date()
        self.date_to      = datetime.strptime(date_to,   "%Y-%m-%d").date()
        self.report_date  = date_to.replace("-", "/")
        self.tenderData: List[dict] = []

    def _setup_logging(self, log_queue):
        root = logging.getLogger()
        if not root.handlers:
            os.makedirs("logs", exist_ok=True)
            root.setLevel(logging.INFO)
            root.addHandler(logging.FileHandler("logs/scraper.log", encoding='utf-8'))

        if log_queue is not None:
            class _QueueHandler(logging.Handler):
                def __init__(self, q):
                    super().__init__()
                    self.q = q
                def emit(self, record):
                    self.q.put(self.format(record))

            qh = _QueueHandler(log_queue)
            qh.setFormatter(logging.Formatter(
                "%(asctime)s  %(levelname)-8s  %(message)s", datefmt="%H:%M:%S"))
            root.addHandler(qh)

    # ── Date helpers ─────────────────────────────────────────────────────────

    def _parse_pub_date(self, raw: str) -> Optional[date]:
        raw = (raw or "").strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d %B %Y", "%B %d, %Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(raw, fmt).date()
            except ValueError:
                continue
        return None

    def _parse_closing(self, raw: str):
        """Return (CLOSING_DATE 'YYYY/MM/DD', CLOSING_TIME str) or ('', '')."""
        raw = (raw or "").strip()
        patterns = [
            (r"(\d{4}-\d{2}-\d{2})\s+(\d{1,2}:\d{2}(?::\d{2})?(?:\s*[aApP][mM])?)", "%Y-%m-%d"),
            (r"(\d{4}/\d{2}/\d{2})\s+(\d{1,2}:\d{2}(?::\d{2})?(?:\s*[aApP][mM])?)", "%Y/%m/%d"),
            (r"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2}(?::\d{2})?(?:\s*[aApP][mM])?)", "%d/%m/%Y"),
        ]
        for pat, fmt in patterns:
            m = re.search(pat, raw, re.IGNORECASE)
            if m:
                date_part = m.group(1)
                time_part = m.group(2).strip()
                try:
                    dt = datetime.strptime(date_part, fmt)
                    return dt.strftime("%Y/%m/%d"), time_part
                except ValueError:
                    continue
        # Date only — no time
        d = self._parse_pub_date(raw)
        if d:
            return d.strftime("%Y/%m/%d"), ""
        return "", ""

    # ── Scraping ─────────────────────────────────────────────────────────────

    def _scrape_page(self, url: str, tender_type: str):
        logging.info(f"JPC: fetching {url}")
        try:
            resp = requests.get(url, headers=_HEADERS, timeout=30)
            resp.raise_for_status()
        except Exception as e:
            logging.error(f"JPC: failed to fetch {url}: {e}")
            return

        soup = BeautifulSoup(resp.text, "html.parser")
        table = soup.find("table")
        if not table:
            logging.warning(f"JPC: no table found at {url}")
            return

        page_count = 0
        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) < 4:
                continue

            date_raw    = tds[0].get_text(strip=True)
            bid_number  = tds[1].get_text(strip=True)
            desc_td     = tds[2]
            description = desc_td.get_text(strip=True)
            link_tag    = desc_td.find("a", href=True)
            link        = link_tag["href"] if link_tag else ""
            closing_raw = tds[3].get_text(strip=True)

            pub_date = self._parse_pub_date(date_raw)
            if pub_date is None:
                continue

            if pub_date < self.date_from or pub_date > self.date_to:
                continue

            closing_date, closing_time = self._parse_closing(closing_raw)

            self.tenderData.append({
                "REPORT_DATE":               self.report_date,
                "RECORD_ID":                 None,
                "TENDER_ID":                 bid_number,
                "PUBLICATION_DATE":          pub_date.strftime("%Y/%m/%d"),
                "CLOSING_DATE":              closing_date,
                "CLOSING_TIME":              closing_time,
                "TENDER_TYPE":               tender_type,
                "TENDER_DESCRIPTION":        description,
                "TENDER_SOURCE":             TENDER_SOURCE,
                "DEPARTMENT":                DEPARTMENT,
                "PROVINCE":                  PROVINCE,
                "ESUBMISSION":               "",
                "CATEGORY":                  "",
                "IS_THERE_A_BRIEFING_SESSION": "",
                "BRIEFING_DATE":             "",
                "COMPULSORY_BRIEFING":       "",
                "BRIEFING_SESSION_VENUE":    "",
                "LINK":                      link,
                "SOE":                       "",
                "COST_OF_SALES_ESTIMATE":    "",
                "CAPABILITY_AVAILABLE":      "",
                "CAPABILITY_GROUP":          "",
                "REQUIREMENTS":              "",
            })
            page_count += 1

        logging.info(f"JPC: {url} → {page_count} tender(s) in date range")

    def run(self) -> List[dict]:
        for url, tender_type in _PAGES:
            self._scrape_page(url, tender_type)
        logging.info(
            f"JPC scrape complete: {len(self.tenderData)} tender(s) "
            f"({self.date_from} → {self.date_to})"
        )
        return self.tenderData
