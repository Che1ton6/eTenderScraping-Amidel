#!/usr/bin/env python3
"""
Scraper for Raymond Mhlaba Local Municipality tenders.
Target: www.raymondmhlaba.gov.za/documents/tenders/
Scrapes the "Current Advert and Tender" section.
No publication date is shown on the site — uses the batch end date (date_to) as proxy.
"""

import logging
import os
import re
from datetime import datetime, date
from typing import Optional, List

import requests
from bs4 import BeautifulSoup

DEPARTMENT    = "Raymond Mhlaba Local Municipality"
TENDER_SOURCE = "RAYMONDMHLABA.GOV.ZA"
PROVINCE      = "Eastern Cape"
SOURCE_URL    = "https://www.raymondmhlaba.gov.za/documents/tenders/"

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36"
    )
}

_MONTH_MAP = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
}


class RaymondMhlabaScraper:
    def __init__(self, date_to: str, log_queue=None):
        self._setup_logging(log_queue)
        self.date_to     = datetime.strptime(date_to, "%Y-%m-%d").date()
        self.report_date = date_to.replace("-", "/")
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

    def _parse_closing(self, raw: str):
        """
        Parse closing text like:
          "Closing on Friday, 27 June 2025"
          "Closing on 27 June 2025"
          "27/06/2025"
        Returns (CLOSING_DATE 'YYYY/MM/DD', CLOSING_TIME '').
        """
        raw = (raw or "").strip()
        # Remove "Closing on" prefix and day name
        cleaned = re.sub(r"(?i)closing\s+on\s+", "", raw).strip()
        cleaned = re.sub(r"(?i)^(monday|tuesday|wednesday|thursday|friday|saturday|sunday),?\s*", "", cleaned).strip()

        formats = ["%d %B %Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d"]
        for fmt in formats:
            try:
                dt = datetime.strptime(cleaned, fmt)
                return dt.strftime("%Y/%m/%d"), ""
            except ValueError:
                continue

        # Fallback: regex for "27 June 2025"
        m = re.search(
            r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", cleaned
        )
        if m:
            day  = int(m.group(1))
            mon  = _MONTH_MAP.get(m.group(2).lower())
            year = int(m.group(3))
            if mon:
                return datetime(year, mon, day).strftime("%Y/%m/%d"), ""

        return "", ""

    def _infer_type(self, description: str) -> str:
        desc = (description or "").upper()
        if "EXPRESSION OF INTEREST" in desc or "EOI" in desc:
            return "Expression of Interest"
        if any(k in desc for k in ("QUOTATION", "RFQ", "PRICE")):
            return "Request for Quotation"
        if any(k in desc for k in ("PROPOSAL", "RFP")):
            return "Request for Proposal"
        return "Request for Bid"

    # ── Scraping ─────────────────────────────────────────────────────────────

    def run(self) -> List[dict]:
        logging.info(f"Raymond Mhlaba: fetching {SOURCE_URL}")
        try:
            resp = requests.get(SOURCE_URL, headers=_HEADERS, timeout=30)
            resp.raise_for_status()
        except Exception as e:
            logging.error(f"Raymond Mhlaba: failed to fetch page: {e}")
            return self.tenderData

        soup = BeautifulSoup(resp.text, "html.parser")

        # Find the "Current Advert and Tender" section heading, then
        # iterate sibling/child entries. The site lists tenders as
        # individual items (divs, list items, or paragraphs) with:
        #   - title/description text
        #   - "Closing on ..." text
        #   - "View Full Advert" / PDF link
        #
        # We cast a wide net: collect all anchor tags and look for
        # surrounding context that contains closing date text.
        self._parse_entries(soup)

        logging.info(
            f"Raymond Mhlaba scrape complete: {len(self.tenderData)} tender(s)"
        )
        return self.tenderData

    def _parse_entries(self, soup: BeautifulSoup):
        """
        Strategy: find every element that contains "Closing on" text —
        that marks a tender entry. Walk up to find description and link.
        """
        pub_date_str = self.date_to.strftime("%Y/%m/%d")

        # Collect all text nodes / elements mentioning "Closing on"
        closing_elements = soup.find_all(
            string=re.compile(r"(?i)closing\s+on")
        )

        seen_ids = set()

        for closing_el in closing_elements:
            closing_text = closing_el.strip()
            closing_date, closing_time = self._parse_closing(closing_text)

            # Walk up to find the nearest block-level parent
            parent = closing_el.parent
            for _ in range(5):
                if parent is None:
                    break
                if parent.name in ("div", "li", "article", "section", "tr", "p"):
                    break
                parent = parent.parent

            if parent is None:
                continue

            # Extract description: first text block in parent that is not the closing line
            description = ""
            link = ""
            for child in parent.descendants:
                text = child.string
                if text and text.strip() and "closing" not in text.lower() \
                        and "view" not in text.lower() and len(text.strip()) > 10:
                    description = text.strip()
                    break

            # Extract link
            a_tag = parent.find("a", href=True)
            if a_tag:
                link = a_tag["href"]
                if link.startswith("/"):
                    link = "https://www.raymondmhlaba.gov.za" + link

            if not description:
                description = parent.get_text(" ", strip=True)
                description = re.sub(r"(?i)closing\s+on.*", "", description).strip()
                description = re.sub(r"(?i)view\s+full\s+advert.*", "", description).strip()

            # Deduplicate by description
            dedup_key = description[:80].lower()
            if dedup_key in seen_ids or not description:
                continue
            seen_ids.add(dedup_key)

            tender_type = self._infer_type(description)

            self.tenderData.append({
                "REPORT_DATE":               self.report_date,
                "RECORD_ID":                 None,
                "TENDER_ID":                 "",
                "PUBLICATION_DATE":          pub_date_str,
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
