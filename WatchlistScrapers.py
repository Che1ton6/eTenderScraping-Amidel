#!/usr/bin/env python3
"""
Watchlist website scrapers — all static HTML / requests+BS4 sites.
Selenium-required sites (Amahlathi, GDoH) are excluded here.
Blocked sites (Eskom=403, Transnet=ECONNREFUSED) are logged and skipped.

SCRAPER_REGISTRY maps Websites.xlsx source names to scraper classes.
Call run_watchlist_scrapers() from main.py to execute all applicable scrapers.
"""

import logging
import os
import re
from datetime import datetime, date
from typing import Optional, List

import requests
from bs4 import BeautifulSoup

# ── Shared constants ──────────────────────────────────────────────────────────

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-ZA,en;q=0.9",
}

_MONTH_MAP = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
    "jan": 1, "feb": 2, "mar": 3, "apr": 4,
    "jun": 6, "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


# ── Shared utilities ──────────────────────────────────────────────────────────

def _parse_date(raw: str) -> Optional[date]:
    raw = re.sub(r"\s+", " ", (raw or "").strip())
    raw = re.sub(r"\s*(at\s*)?\d{1,2}[h:]\d{2}.*$", "", raw, flags=re.IGNORECASE).strip()
    raw = re.sub(r"\s*[-–]\s*\d{1,2}:\d{2}.*$", "", raw).strip()
    for fmt in ("%d %B %Y", "%B %d, %Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d",
                "%d-%m-%Y", "%m/%d/%Y", "%d %b %Y", "%b %d, %Y", "%d %B, %Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    m = re.search(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", raw)
    if m:
        mon = _MONTH_MAP.get(m.group(2).lower())
        if mon:
            try:
                return date(int(m.group(3)), mon, int(m.group(1)))
            except ValueError:
                pass
    m = re.search(r"([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})", raw)
    if m:
        mon = _MONTH_MAP.get(m.group(1).lower())
        if mon:
            try:
                return date(int(m.group(3)), mon, int(m.group(2)))
            except ValueError:
                pass
    return None


def _parse_closing(raw: str):
    """Return ('YYYY/MM/DD', time_str) or ('', '')."""
    raw = (raw or "").strip()
    tm = re.search(
        r"(\d{1,2}[h:]\d{2}(?:\s*[AaPp][Mm])?|\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AaPp][Mm])?)",
        raw, re.IGNORECASE,
    )
    time_str = tm.group(1).strip() if tm else ""
    d = _parse_date(raw)
    return (d.strftime("%Y/%m/%d"), time_str) if d else ("", time_str)


def _infer_type(description: str) -> str:
    desc = (description or "").upper()
    if any(k in desc for k in ("EXPRESSION OF INTEREST", " EOI")):
        return "Expression of Interest"
    if any(k in desc for k in ("QUOTATION", " RFQ", "PRICE QUOTATION")):
        return "Request for Quotation"
    if any(k in desc for k in ("PROPOSAL", " RFP")):
        return "Request for Proposal"
    return "Request for Bid"


def _get(url: str, verify=True, timeout=30) -> Optional[BeautifulSoup]:
    try:
        resp = requests.get(url, headers=_HEADERS, timeout=timeout, verify=verify)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding or "utf-8"
        return BeautifulSoup(resp.text, "html.parser")
    except Exception as e:
        logging.error(f"Fetch failed {url}: {e}")
        return None


# ── Base scraper ──────────────────────────────────────────────────────────────

class _Base:
    DEPARTMENT = ""
    PROVINCE   = ""
    SOURCE_KEY = ""

    def __init__(self, date_from: str, date_to: str, log_queue=None):
        self._setup_log(log_queue)
        self.date_from   = datetime.strptime(date_from, "%Y-%m-%d").date()
        self.date_to     = datetime.strptime(date_to,   "%Y-%m-%d").date()
        self.report_date = date_to.replace("-", "/")
        self.tenderData: List[dict] = []

    def _setup_log(self, log_queue):
        root = logging.getLogger()
        if not root.handlers:
            os.makedirs("logs", exist_ok=True)
            root.setLevel(logging.INFO)
            root.addHandler(logging.FileHandler("logs/scraper.log"))
        if log_queue is not None:
            class _QH(logging.Handler):
                def __init__(self, q): super().__init__(); self.q = q
                def emit(self, r): self.q.put(self.format(r))
            qh = _QH(log_queue)
            qh.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s",
                                              datefmt="%H:%M:%S"))
            root.addHandler(qh)

    def _blank(self) -> dict:
        return {
            "REPORT_DATE": self.report_date, "RECORD_ID": None,
            "TENDER_ID": "", "PUBLICATION_DATE": "", "CLOSING_DATE": "",
            "CLOSING_TIME": "", "TENDER_TYPE": "Request for Bid",
            "TENDER_DESCRIPTION": "", "TENDER_SOURCE": self.SOURCE_KEY,
            "DEPARTMENT": self.DEPARTMENT, "PROVINCE": self.PROVINCE,
            "ESUBMISSION": "", "CATEGORY": "", "IS_THERE_A_BRIEFING_SESSION": "",
            "BRIEFING_DATE": "", "COMPULSORY_BRIEFING": "", "BRIEFING_SESSION_VENUE": "",
            "LINK": "", "SOE": "", "COST_OF_SALES_ESTIMATE": "",
            "CAPABILITY_AVAILABLE": "", "CAPABILITY_GROUP": "", "REQUIREMENTS": "",
        }

    def _in_range(self, pub_date: Optional[date]) -> bool:
        if pub_date is None:
            return False
        return self.date_from <= pub_date <= self.date_to

    def run(self) -> List[dict]:
        raise NotImplementedError


# ── Individual scrapers ───────────────────────────────────────────────────────

class MatatieleScraper(_Base):
    """https://www.matatiele.gov.za/tenders/ — WP Job Manager (JavaScript-rendered).
    Static scraping returns nothing; tenders fetched via Selenium scraper instead.
    """
    DEPARTMENT = "Matatiele Local Municipality"
    PROVINCE   = "Eastern Cape"
    SOURCE_KEY = "MATATIELE.GOV.ZA"

    def run(self):
        logging.info(
            "Matatiele: page requires JavaScript (WP Job Manager renders listings "
            "client-side) — static scraper skipped; Selenium scraper handles this source."
        )
        return self.tenderData


class NtabankuluScraper(_Base):
    """https://www.ntabankulu.gov.za/category/tenders/open-tenders/ — WordPress archive."""
    DEPARTMENT = "Ntabankulu Local Municipality"
    PROVINCE   = "Eastern Cape"
    SOURCE_KEY = "NTABANKULU.GOV.ZA"

    def run(self):
        logging.info("Ntabankulu: fetching tender page")
        url = "https://www.ntabankulu.gov.za/category/tenders/open-tenders/"
        soup = _get(url)
        if not soup:
            return self.tenderData

        # WordPress category archive — each post is an <article> or .post
        articles = soup.select("article, .post, .hentry")
        for art in articles:
            t = self._blank()
            title_el = art.find(["h1", "h2", "h3"])
            if not title_el:
                continue
            a = title_el.find("a", href=True)
            t["TENDER_DESCRIPTION"] = (a or title_el).get_text(strip=True)
            if a:
                t["LINK"] = a["href"]

            # Publication date
            time_el = art.find("time")
            if time_el:
                pub = _parse_date(time_el.get("datetime", "") or time_el.get_text(strip=True))
                if pub:
                    if not self._in_range(pub):
                        continue
                    t["PUBLICATION_DATE"] = pub.strftime("%Y/%m/%d")
            if not t["PUBLICATION_DATE"]:
                t["PUBLICATION_DATE"] = ""

            t["TENDER_TYPE"] = _infer_type(t["TENDER_DESCRIPTION"])
            self.tenderData.append(t)

        logging.info(f"Ntabankulu: {len(self.tenderData)} tender(s)")
        return self.tenderData


class UmzimvubuScraper(_Base):
    """https://www.umzimvubu.gov.za/tender-documents/ — 3-column table."""
    DEPARTMENT = "Umzimvubu Local Municipality"
    PROVINCE   = "Eastern Cape"
    SOURCE_KEY = "UMZIMVUBU.GOV.ZA"

    def run(self):
        logging.info("Umzimvubu: fetching tender page")
        soup = _get("https://www.umzimvubu.gov.za/tender-documents/")
        if not soup:
            return self.tenderData

        table = soup.find("table")
        if table:
            for tr in table.find_all("tr"):
                tds = tr.find_all("td")
                if not tds:
                    continue
                t = self._blank()
                t["TENDER_DESCRIPTION"] = tds[0].get_text(strip=True)
                if not t["TENDER_DESCRIPTION"]:
                    continue
                a = tr.find("a", href=True)
                if a:
                    t["LINK"] = a["href"]
                t["PUBLICATION_DATE"] = ""
                t["TENDER_TYPE"] = _infer_type(t["TENDER_DESCRIPTION"])
                self.tenderData.append(t)
        else:
            # Fallback: catfolders plugin rows
            for item in soup.select(".catfolders-item, .wp-block-file, a[href$='.pdf']"):
                t = self._blank()
                t["TENDER_DESCRIPTION"] = item.get_text(strip=True)
                if not t["TENDER_DESCRIPTION"]:
                    continue
                if item.get("href"):
                    t["LINK"] = item["href"]
                t["PUBLICATION_DATE"] = ""
                self.tenderData.append(t)

        logging.info(f"Umzimvubu: {len(self.tenderData)} tender(s)")
        return self.tenderData


class WinnieMMLScraper(_Base):
    """https://www.winniemmlm.gov.za/tenders/ — tabbed HTML table (Open Tenders tab)."""
    DEPARTMENT = "Winnie Madikizela-Mandela Local Municipality"
    PROVINCE   = "Eastern Cape"
    SOURCE_KEY = "WINNIEMMLM.GOV.ZA"

    def run(self):
        logging.info("Winnie MM: fetching tender page")
        soup = _get("https://www.winniemmlm.gov.za/tenders/")
        if not soup:
            return self.tenderData

        # Find all tables — the first (Open Tenders) tab is target
        tables = soup.find_all("table")
        target = None
        for tbl in tables:
            hdrs = [th.get_text(strip=True).lower() for th in tbl.find_all("th")]
            if any(h in hdrs for h in ("description", "closing", "advert")):
                target = tbl
                break
        if not target and tables:
            target = tables[0]
        if not target:
            logging.warning("Winnie MM: no table found")
            return self.tenderData

        header_cells = [th.get_text(strip=True).lower() for th in target.find_all("th")]
        col = {h: i for i, h in enumerate(header_cells)}

        def _idx(*keys):
            for k in keys:
                for h, i in col.items():
                    if k in h:
                        return i
            return None

        advert_idx  = _idx("advert", "date", "publication")
        desc_idx    = _idx("description")
        closing_idx = _idx("closing")
        doc_idx     = _idx("document", "download", "link")

        for tr in target.find_all("tr")[1:]:
            tds = tr.find_all("td")
            if not tds:
                continue
            t = self._blank()

            if desc_idx is not None and desc_idx < len(tds):
                t["TENDER_DESCRIPTION"] = tds[desc_idx].get_text(strip=True)
            else:
                t["TENDER_DESCRIPTION"] = tds[0].get_text(strip=True)

            if not t["TENDER_DESCRIPTION"]:
                continue

            if advert_idx is not None and advert_idx < len(tds):
                pub = _parse_date(tds[advert_idx].get_text(strip=True))
                if pub:
                    t["PUBLICATION_DATE"] = pub.strftime("%Y/%m/%d")
            if not t["PUBLICATION_DATE"]:
                t["PUBLICATION_DATE"] = ""

            if closing_idx is not None and closing_idx < len(tds):
                t["CLOSING_DATE"], t["CLOSING_TIME"] = _parse_closing(
                    tds[closing_idx].get_text(strip=True))

            link_td = tds[doc_idx] if doc_idx is not None and doc_idx < len(tds) else tr
            a = link_td.find("a", href=True)
            if a:
                t["LINK"] = a["href"]

            t["TENDER_TYPE"] = _infer_type(t["TENDER_DESCRIPTION"])
            self.tenderData.append(t)

        logging.info(f"Winnie MM: {len(self.tenderData)} tender(s)")
        return self.tenderData


class MnqumaScraper(_Base):
    """https://www.mnquma.gov.za/supply-chain/ — card listing of current tenders.
    Publication date and closing date are on each tender's individual post page,
    not on the listing page. Fetches up to 50 individual pages per run.
    """
    DEPARTMENT = "Mnquma Local Municipality"
    PROVINCE   = "Eastern Cape"
    SOURCE_KEY = "MNQUMA.GOV.ZA"

    _LISTING = "https://www.mnquma.gov.za/supply-chain/"

    def run(self):
        logging.info("Mnquma: fetching supply-chain listing")
        soup = _get(self._LISTING)
        if not soup:
            return self.tenderData

        # Collect all unique "VIEW TENDER" link targets
        links = []
        seen_hrefs: set = set()
        for a in soup.find_all("a", href=True):
            if a.get_text(strip=True).upper() == "VIEW TENDER":
                href = a["href"]
                if href not in seen_hrefs:
                    seen_hrefs.add(href)
                    links.append(href)

        logging.info(f"Mnquma: {len(links)} tender page(s) to fetch")

        for url in links[:50]:
            try:
                post = _get(url)
                if not post:
                    continue

                # Page title contains description + SCM reference number.
                # Mnquma uses Elementor (h4.elementor-heading-title); fall back to any heading.
                heading = (
                    post.find(class_="elementor-heading-title")
                    or post.find(["h1", "h2", "h3", "h4"])
                )
                if not heading:
                    continue
                desc = heading.get_text(strip=True)
                if not desc:
                    continue

                t = self._blank()
                t["TENDER_DESCRIPTION"] = desc
                t["LINK"] = url

                # Extract SCM reference from title, e.g. "... MNQ/SCM/94/25-26"
                ref_m = re.search(
                    r"((?:MNQ|SCM)[/\-](?:SCM|MLM)[/\-])(\d+)[/\-](\d{2,4})[/\-](\d{2,4})",
                    desc, re.I,
                )
                if ref_m:
                    prefix = ref_m.group(1).upper().replace("-", "/")
                    t["TENDER_ID"] = f"{prefix}{ref_m.group(2)}/{ref_m.group(3)}-{ref_m.group(4)}"
                else:
                    # Fallback: try URL slug e.g. "...-mnq-scm-94-25-26"
                    slug = url.rstrip("/").split("/")[-1]
                    slug_m = re.search(
                        r"(?:mnq|scm|mlm)[/-]scm[/-](\d+)[/-](\d{2,4})[/-](\d{2,4})$", slug, re.I
                    )
                    if slug_m:
                        t["TENDER_ID"] = f"MNQ/SCM/{slug_m.group(1)}/{slug_m.group(2)}-{slug_m.group(3)}"

                # Pub date and closing date are in the post body as labelled fields
                body = post.get_text(" ", strip=True)
                pub_m = re.search(r"Publish\s+Date:\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})", body, re.I)
                if pub_m:
                    pub = _parse_date(pub_m.group(1))
                    if pub:
                        t["PUBLICATION_DATE"] = pub.strftime("%Y/%m/%d")

                close_m = re.search(r"Closing\s+Date:\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})", body, re.I)
                if close_m:
                    cd = _parse_date(close_m.group(1))
                    if cd:
                        t["CLOSING_DATE"] = cd.strftime("%Y/%m/%d")

                # First document download link (PDF/DOC preferred over landing page)
                doc_a = post.find("a", href=re.compile(r"\.(pdf|docx?|xlsx?)$", re.I))
                if doc_a:
                    t["LINK"] = doc_a["href"]

                t["TENDER_TYPE"] = _infer_type(desc)
                self.tenderData.append(t)

            except Exception as e:
                logging.debug(f"Mnquma post error {url}: {e}")

        logging.info(f"Mnquma: {len(self.tenderData)} tender(s)")
        return self.tenderData


class GreatKeiScraper(_Base):
    """https://greatkeilm.gov.za/web/category/tenders/open-tenders/ — WP post list."""
    DEPARTMENT = "Great Kei Local Municipality"
    PROVINCE   = "Eastern Cape"
    SOURCE_KEY = "GREATKEILM.GOV.ZA"

    _BASE_URL = "https://greatkeilm.gov.za/web/category/tenders/open-tenders/"

    def run(self):
        logging.info("Great Kei: fetching open tenders (paginated)")
        seen_ids: set = set()
        pages_fetched = 0

        for page in range(1, 6):  # fetch up to 5 pages
            url = self._BASE_URL if page == 1 else f"{self._BASE_URL}page/{page}/"
            soup = _get(url)
            if not soup:
                break
            articles = soup.select("article, .post, .hentry")
            if not articles:
                break
            pages_fetched += 1
            found_in_range = False

            for art in articles:
                t = self._blank()
                heading = art.find(["h4", "h3", "h2"])
                if not heading:
                    continue
                a = heading.find("a", href=True) or art.find("a", href=True)
                t["TENDER_DESCRIPTION"] = heading.get_text(strip=True)
                if a:
                    t["LINK"] = a["href"]

                text = art.get_text(" ", strip=True)

                # Reference number e.g. "RFQ/BTO/09/2025/26"
                ref_m = re.search(r"(RFQ|BID|SCM|RFP|T|B)[/\-]\w+[/\-]\w+", text, re.I)
                if ref_m:
                    t["TENDER_ID"] = ref_m.group(0)

                # Skip already-seen IDs across pages
                dedup_key = t["TENDER_ID"] or t["TENDER_DESCRIPTION"][:60]
                if dedup_key in seen_ids:
                    continue
                seen_ids.add(dedup_key)

                # Publication date
                time_el = art.find("time")
                if time_el:
                    pub = _parse_date(time_el.get("datetime", "") or time_el.get_text())
                    if pub:
                        if self._in_range(pub):
                            found_in_range = True
                        t["PUBLICATION_DATE"] = pub.strftime("%Y/%m/%d")
                if not t["PUBLICATION_DATE"]:
                    t["PUBLICATION_DATE"] = ""

                # Closing date — "12 June 2026 - 11:00 am" pattern
                close_m = re.search(
                    r"[Cc]los(?:ing|ed?)\s*(?:date)?:?\s*"
                    r"(\d{1,2}\s+[A-Za-z]+\s+\d{4}|\d{1,2}/\d{1,2}/\d{4})"
                    r"(?:\s*[-–]\s*(\d{1,2}[h:]\d{2}(?:\s*[AaPp][Mm])?))?",
                    text,
                )
                if close_m:
                    cd = _parse_date(close_m.group(1))
                    if cd:
                        t["CLOSING_DATE"] = cd.strftime("%Y/%m/%d")
                    t["CLOSING_TIME"] = close_m.group(2) or ""

                t["TENDER_TYPE"] = _infer_type(t["TENDER_DESCRIPTION"])
                self.tenderData.append(t)

            # Stop paginating if no in-range items found on this page
            if page > 1 and not found_in_range:
                break

        logging.info(f"Great Kei: {len(self.tenderData)} tender(s) across {pages_fetched} page(s)")
        return self.tenderData


class JOSHCOScraper(_Base):
    """https://www.joshco.co.za/about-joshco/tenders/ — nested list by year, Azure CDN PDFs."""
    DEPARTMENT = "Johannesburg Social Housing Company"
    PROVINCE   = "Gauteng"
    SOURCE_KEY = "JOSHCO.CO.ZA"

    def run(self):
        logging.info("JOSHCO: fetching tender page")
        soup = _get("https://www.joshco.co.za/about-joshco/tenders/")
        if not soup:
            return self.tenderData

        current_year = date.today().year
        # Year sections are <h3> followed by <ul>
        for h3 in soup.find_all("h3"):
            year_text = h3.get_text(strip=True)
            try:
                year = int(re.search(r"\d{4}", year_text).group())
            except (AttributeError, ValueError):
                year = current_year

            ul = h3.find_next_sibling("ul")
            if not ul:
                continue

            for li in ul.find_all("li", recursive=False):
                t = self._blank()
                a = li.find("a", href=True)
                text = li.get_text(" ", strip=True)

                # Remove "Additional File" sub-items from description
                sub_a = li.find("ul")
                main_text = text
                if sub_a:
                    main_text = li.contents[0].string or text if li.contents else text

                t["TENDER_DESCRIPTION"] = re.sub(r"\s+", " ", main_text).strip()
                if not t["TENDER_DESCRIPTION"]:
                    continue
                if a:
                    t["LINK"] = a["href"]

                # Closing date — "June 12, 2025, 11:00 AM" or similar
                close_m = re.search(
                    r"([A-Za-z]+\s+\d{1,2},?\s+\d{4}|\d{1,2}\s+[A-Za-z]+\s+\d{4})"
                    r"(?:,?\s+(\d{1,2}:\d{2}(?:\s*[AaPp][Mm])?))?",
                    text,
                )
                if close_m:
                    cd = _parse_date(close_m.group(1))
                    if cd:
                        t["CLOSING_DATE"] = cd.strftime("%Y/%m/%d")
                    t["CLOSING_TIME"] = close_m.group(2) or ""

                t["PUBLICATION_DATE"] = ""
                t["TENDER_TYPE"] = _infer_type(t["TENDER_DESCRIPTION"])
                self.tenderData.append(t)

        logging.info(f"JOSHCO: {len(self.tenderData)} tender(s)")
        return self.tenderData


class GPLScraper(_Base):
    """https://www.gpl.gov.za/advertised-tenders/ — simple WordPress post list."""
    DEPARTMENT = "Gauteng Provincial Legislature"
    PROVINCE   = "Gauteng"
    SOURCE_KEY = "GPL.GOV.ZA"

    def run(self):
        logging.info("GPL: fetching advertised tenders")
        soup = _get("https://www.gpl.gov.za/advertised-tenders/")
        if not soup:
            return self.tenderData

        articles = soup.select("article, .post, .hentry, .entry")
        for art in articles:
            t = self._blank()
            heading = art.find(["h2", "h3", "h4"])
            if not heading:
                continue
            a = heading.find("a", href=True) or art.find("a", href=True)
            t["TENDER_DESCRIPTION"] = heading.get_text(strip=True)
            if not t["TENDER_DESCRIPTION"]:
                continue
            if a:
                t["LINK"] = a["href"]

            # Tender number — e.g. "GPL019/2022"
            ref_m = re.search(r"GPL\d+/\d+", t["TENDER_DESCRIPTION"], re.I)
            if ref_m:
                t["TENDER_ID"] = ref_m.group(0)

            time_el = art.find("time")
            pub = None
            if time_el:
                pub = _parse_date(time_el.get("datetime", "") or time_el.get_text())
            if pub:
                if not self._in_range(pub):
                    continue
                t["PUBLICATION_DATE"] = pub.strftime("%Y/%m/%d")
            else:
                t["PUBLICATION_DATE"] = ""

            t["TENDER_TYPE"] = _infer_type(t["TENDER_DESCRIPTION"])
            self.tenderData.append(t)

        logging.info(f"GPL: {len(self.tenderData)} tender(s)")
        return self.tenderData


class RAFScraper(_Base):
    """https://www.raf.co.za/procurement/Pages/Tenders-running.aspx
    SharePoint portal — requires JavaScript to render tender listings.
    Static scraping returns no content; manual monitoring required.
    """
    DEPARTMENT = "Road Accident Fund"
    PROVINCE   = "National"
    SOURCE_KEY = "RAF.CO.ZA"

    def run(self):
        logging.warning(
            "RAF: portal is SharePoint-based and requires JavaScript to render "
            "tender listings — static scraping not possible. "
            "Manual monitoring at https://www.raf.co.za/procurement/Pages/Tenders-running.aspx"
        )
        return self.tenderData


class NMBMMScraper(_Base):
    """https://www.nelsonmandelabay.gov.za/tenders/ — multiple tables, one per tender type.
    Each table: row[0]=section heading (e.g. 'FORMAL TENDERS'), row[1]=column headers,
    row[2+]=data rows.  Columns: SCM No. | Tender Description | Tender Fee.
    Closing dates and publication dates are NOT shown on the website listing.
    """
    DEPARTMENT = "Nelson Mandela Bay Metropolitan Municipality"
    PROVINCE   = "Eastern Cape"
    SOURCE_KEY = "NELSONMANDELABAY.GOV.ZA"

    def run(self):
        logging.info("NMB MM: fetching tenders")
        base_url = "https://www.nelsonmandelabay.gov.za/tenders/"
        page = 1
        seen_descs: set = set()

        while True:
            url = base_url if page == 1 else f"{base_url}?page={page}"
            soup = _get(url)
            if not soup:
                break

            tables = soup.find_all("table")
            if not tables:
                break

            found_any = False
            for table in tables:
                rows = table.find_all("tr")
                if len(rows) < 3:
                    continue

                # Find the column header row — the row whose cells contain "description" / "scm"
                header_idx = None
                header: list = []
                for i, tr in enumerate(rows[:3]):
                    cells = [c.get_text(strip=True).lower() for c in tr.find_all(["th", "td"])]
                    if any("description" in c or "scm" in c for c in cells):
                        header_idx = i
                        header = cells
                        break
                if header_idx is None:
                    continue

                def _col(*keys):
                    for k in keys:
                        for i, h in enumerate(header):
                            if k in h:
                                return i
                    return None

                id_idx   = _col("scm no", "scm", "tender no", "number", "ref")
                desc_idx = _col("description", "subject")

                for tr in rows[header_idx + 1:]:
                    tds = tr.find_all("td")
                    if not tds:
                        continue

                    def _val(idx):
                        return tds[idx].get_text(strip=True) if idx is not None and idx < len(tds) else ""

                    tender_id = _val(id_idx)
                    desc      = _val(desc_idx) or (tds[1].get_text(strip=True) if len(tds) > 1 else "")
                    if not desc:
                        continue

                    dedup_key = desc[:60].lower()
                    if dedup_key in seen_descs:
                        continue
                    seen_descs.add(dedup_key)

                    t = self._blank()
                    t["TENDER_ID"]          = tender_id
                    t["TENDER_DESCRIPTION"] = desc
                    t["PUBLICATION_DATE"]   = ""
                    # No closing date column on NMB website — field will be blank

                    a = tr.find("a", href=True)
                    if a:
                        t["LINK"] = a["href"]

                    t["TENDER_TYPE"] = _infer_type(desc)
                    self.tenderData.append(t)
                    found_any = True

            next_link = soup.find("a", string=re.compile(r"[Nn]ext|»|›"))
            if next_link and found_any:
                page += 1
            else:
                break

        logging.info(f"NMB MM: {len(self.tenderData)} tender(s)")
        return self.tenderData


class DELScraper(_Base):
    """https://www.labour.gov.za/tenders/available-tenders — SharePoint table.
    Page renders two tables: Table 0 has a merged mega-cell in the header (SharePoint
    accessibility row), Table 1 is the clean version.
    Columns: (blank) | Tender Ref | Tender Description | Tender Location | Closing Date
    Date format is DD/MM/YYYY (South African), closing time as HH:MM.
    """
    DEPARTMENT = "Department of Employment and Labour"
    PROVINCE   = "National"
    SOURCE_KEY = "LABOUR.GOV.ZA"
    _URL       = "https://www.labour.gov.za/tenders/available-tenders"
    _BASE      = "https://www.labour.gov.za"

    def run(self):
        logging.info("DEL: fetching available tenders")
        soup = _get(self._URL)
        if not soup:
            return self.tenderData

        # Use the last table whose header row contains "tender ref" and "closing"
        # (avoids the SharePoint merged-cell Table 0)
        target = None
        for table in soup.find_all("table"):
            rows = table.find_all("tr")
            if not rows:
                continue
            hdrs = [c.get_text(strip=True).lower() for c in rows[0].find_all(["th", "td"])]
            if any("tender ref" in h for h in hdrs) or any("closing" in h for h in hdrs):
                target = table

        if not target:
            logging.warning("DEL: no usable table found")
            return self.tenderData

        rows = target.find_all("tr")
        header = [c.get_text(strip=True).lower() for c in rows[0].find_all(["th", "td"])]

        def _col(*keys):
            for k in keys:
                for i, h in enumerate(header):
                    if k in h:
                        return i
            return None

        ref_idx   = _col("tender ref", "ref")
        desc_idx  = _col("description")
        close_idx = _col("closing")

        for tr in rows[1:]:
            tds = tr.find_all("td")
            if not tds:
                continue
            t = self._blank()

            def _val(idx):
                return tds[idx].get_text(strip=True) if idx is not None and idx < len(tds) else ""

            t["TENDER_ID"]          = _val(ref_idx)
            t["TENDER_DESCRIPTION"] = _val(desc_idx)
            if not t["TENDER_DESCRIPTION"]:
                continue

            # Closing date — DEL stores as "DD/MM/YYYY HH:MM" (South African format)
            raw_close = _val(close_idx)
            if raw_close:
                sp_m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})\s*(\d{1,2}:\d{2})?", raw_close)
                if sp_m:
                    try:
                        cd = date(int(sp_m.group(3)), int(sp_m.group(2)), int(sp_m.group(1)))
                        t["CLOSING_DATE"] = cd.strftime("%Y/%m/%d")
                        t["CLOSING_TIME"] = sp_m.group(4) or ""
                    except ValueError:
                        pass
                else:
                    t["CLOSING_DATE"], t["CLOSING_TIME"] = _parse_closing(raw_close)

            a = tr.find("a", href=True)
            if a:
                href = a["href"]
                if href.startswith("/"):
                    href = self._BASE + href
                if not href.startswith("javascript"):
                    t["LINK"] = href

            t["PUBLICATION_DATE"] = ""
            t["TENDER_TYPE"] = _infer_type(t["TENDER_DESCRIPTION"])
            self.tenderData.append(t)

        logging.info(f"DEL: {len(self.tenderData)} tender(s)")
        return self.tenderData


class DIRCOScraper(_Base):
    """https://dirco.gov.za/tenders — Elementor-based page.
    DIRCO uses Elementor text-editor widgets, not a standard content div.
    Tenders appear as headings (h3/h4) inside Elementor widgets with sibling
    paragraphs containing dates.  Page is genuinely empty until DIRCO lists
    2026/27 tenders; scraper returns 0 correctly in that case.
    """
    DEPARTMENT = "Department of International Relations and Cooperation"
    PROVINCE   = "National"
    SOURCE_KEY = "DIRCO.GOV.ZA"

    def run(self):
        logging.info("DIRCO: fetching tenders")
        soup = _get("https://dirco.gov.za/tenders")
        if not soup:
            return self.tenderData

        # DIRCO uses Elementor — content is inside .elementor-widget-container divs,
        # NOT inside a standard entry/content div.
        # Collect all text from Elementor text-editor and heading widgets.
        widgets = soup.find_all(
            "div",
            class_=re.compile(r"elementor-widget-(text-editor|heading)", re.I),
        )
        # Fall back to body if Elementor not found (future redesign resilience)
        if not widgets:
            widgets = [soup.find("main") or soup.body]

        seen: set = set()
        for widget in widgets:
            for heading in widget.find_all(["h3", "h4"]):
                title = heading.get_text(strip=True)
                if not title or len(title) < 10:
                    continue
                # Skip year-navigation and section headings
                if re.match(r"DIRCO tenders for \d{4}", title, re.I):
                    continue

                t = self._blank()
                t["TENDER_DESCRIPTION"] = title

                # Reference: "DIRCO 02 2026-2027" or "DIRCO/02/2026-2027"
                ref_m = re.search(r"DIRCO[\s/]*\d+[\s/]*\d{4}[-/]\d{4}", title, re.I)
                if ref_m:
                    t["TENDER_ID"] = ref_m.group(0).strip()

                # Sibling paragraphs hold dates
                text_parts = []
                for sib in heading.next_siblings:
                    if sib.name in ("h3", "h4"):
                        break
                    if hasattr(sib, "get_text"):
                        text_parts.append(sib.get_text(" ", strip=True))
                combined = " ".join(text_parts)

                pub_m = re.search(
                    r"[Pp]ublish(?:ed)?\s*[Dd]ate:?\s*"
                    r"(\d{1,2}\s+[A-Za-z]+\s+\d{4}|\d{1,2}/\d{1,2}/\d{4})",
                    combined,
                )
                if pub_m:
                    pub = _parse_date(pub_m.group(1))
                    if pub:
                        t["PUBLICATION_DATE"] = pub.strftime("%Y/%m/%d")

                close_m = re.search(
                    r"[Cc]losing\s*[Dd]ate:?\s*"
                    r"(\d{1,2}\s+[A-Za-z]+\s+\d{4}|\d{1,2}/\d{1,2}/\d{4})"
                    r"(?:\s+at\s+(\d{1,2}h\d{2}))?",
                    combined,
                )
                if close_m:
                    t["CLOSING_DATE"], _ = _parse_closing(close_m.group(1))
                    t["CLOSING_TIME"] = close_m.group(2) or ""

                a = heading.find("a", href=True) or widget.find("a", href=re.compile(r"\.pdf", re.I))
                if a:
                    t["LINK"] = a["href"]

                dedup = title[:60].lower()
                if dedup in seen:
                    continue
                seen.add(dedup)

                t["TENDER_TYPE"] = _infer_type(title)
                self.tenderData.append(t)

        if not self.tenderData:
            logging.info("DIRCO: no tenders listed on page (2026/27 section is currently empty)")
        else:
            logging.info(f"DIRCO: {len(self.tenderData)} tender(s)")
        return self.tenderData


class DOJScraper(_Base):
    """https://www.justice.gov.za/cfo_tender/tender.htm — page only lists contract extensions, not open bids."""
    DEPARTMENT = "Department of Justice and Constitutional Development"
    PROVINCE   = "National"
    SOURCE_KEY = "JUSTICE.GOV.ZA"

    def run(self):
        logging.info("DOJ: page contains no open tenders (contract extensions only) — skipping")
        return self.tenderData


class WJHBScraper(_Base):
    """https://scm.johannesburgwater.co.za/supply-chain/tenders/all-open-tenders/ — HTML table."""
    DEPARTMENT = "Johannesburg Water SOC Ltd"
    PROVINCE   = "Gauteng"
    SOURCE_KEY = "JOHANNESBURGWATER.CO.ZA"

    def run(self):
        logging.info("W JHB: fetching open tenders")
        import requests as _req
        try:
            r = _req.get(
                "https://scm.johannesburgwater.co.za/supply-chain/tenders/all-open-tenders/",
                headers={"User-Agent": "Mozilla/5.0"},
                timeout=30,
            )
            if r.status_code == 403:
                logging.warning(
                    "W JHB: 403 Forbidden — IP blocked by WP Defender security plugin. "
                    "Tenders skipped this run; block should auto-reset."
                )
                return self.tenderData
            r.raise_for_status()
            r.encoding = r.apparent_encoding or "utf-8"
            from bs4 import BeautifulSoup as _BS
            soup = _BS(r.text, "html.parser")
        except Exception as e:
            logging.error(f"W JHB: fetch failed — {e}")
            return self.tenderData

        table = soup.find("table")
        if not table:
            logging.warning("W JHB: no table found")
            return self.tenderData

        rows = table.find_all("tr")
        header = [th.get_text(strip=True).lower() for th in
                  (rows[0].find_all(["th", "td"]) if rows else [])]

        def _col(*keys):
            for k in keys:
                for i, h in enumerate(header):
                    if k in h:
                        return i
            return None

        id_idx       = _col("tender no", "number", "ref")
        desc_idx     = _col("description")
        advert_idx   = _col("advert", "publish", "date")
        closing_idx  = _col("closing", "close")
        briefing_idx = _col("briefing")
        dl_idx       = _col("download", "link", "pdf")

        for tr in rows[1:]:
            tds = tr.find_all("td")
            if not tds:
                continue
            t = self._blank()

            def _val(idx):
                return tds[idx].get_text(strip=True) if idx is not None and idx < len(tds) else ""

            t["TENDER_ID"]          = _val(id_idx)
            t["TENDER_DESCRIPTION"] = _val(desc_idx) or _val(1)
            if not t["TENDER_DESCRIPTION"]:
                continue

            pub = _parse_date(_val(advert_idx)) if advert_idx is not None else None
            if pub:
                if not self._in_range(pub):
                    continue
                t["PUBLICATION_DATE"] = pub.strftime("%Y/%m/%d")
            else:
                t["PUBLICATION_DATE"] = ""

            if closing_idx is not None:
                t["CLOSING_DATE"], t["CLOSING_TIME"] = _parse_closing(_val(closing_idx))

            if briefing_idx is not None:
                t["BRIEFING_SESSION_VENUE"] = _val(briefing_idx)

            # First PDF link in the download cell (or entire row)
            link_td = tds[dl_idx] if dl_idx is not None and dl_idx < len(tds) else tr
            a = link_td.find("a", href=True)
            if a:
                t["LINK"] = a["href"]

            t["TENDER_TYPE"] = _infer_type(t["TENDER_DESCRIPTION"])
            self.tenderData.append(t)

        logging.info(f"W JHB: {len(self.tenderData)} tender(s)")
        return self.tenderData


class SITAScraper(_Base):
    """https://rfq.sita.co.za/TendersAdministration/invitations.asp
    Page has 2 tables: table[0]=contact info, table[1]=tenders.
    Tender table columns: Description | Tender Number | Closing Date(at 11:00) | Download Documents
    The Description cell starts with 'Published Date: DD/MM/YYYY'.
    """
    DEPARTMENT = "State Information Technology Agency"
    PROVINCE   = "National"
    SOURCE_KEY = "SITA.CO.ZA"
    _URL       = "https://rfq.sita.co.za/TendersAdministration/invitations.asp"
    _BASE      = "https://rfq.sita.co.za/TendersAdministration/"

    def run(self):
        logging.info("SITA: fetching tender portal")
        soup = _get(self._URL)
        if not soup:
            return self.tenderData

        tables = soup.find_all("table")
        if len(tables) < 2:
            logging.warning("SITA: expected 2 tables, got %d", len(tables))
            return self.tenderData

        tender_table = tables[1]
        rows = tender_table.find_all("tr")
        if not rows:
            return self.tenderData

        for tr in rows[1:]:  # skip header row
            tds = tr.find_all("td")
            if len(tds) < 3:
                continue

            desc_raw = tds[0].get_text(" ", strip=True)
            tender_no = tds[1].get_text(strip=True)
            close_raw = tds[2].get_text(strip=True)

            # Extract "Published Date: DD/MM/YYYY" from start of description cell
            pub = None
            pub_m = re.search(r"Published Date:\s*(\d{2}/\d{2}/\d{4})", desc_raw, re.I)
            if pub_m:
                pub = _parse_date(pub_m.group(1))
                desc = desc_raw[pub_m.end():].strip()
            else:
                desc = desc_raw.strip()

            if not desc:
                continue

            if pub and not self._in_range(pub):
                continue

            t = self._blank()
            t["TENDER_ID"]          = tender_no
            t["TENDER_DESCRIPTION"] = desc
            t["PUBLICATION_DATE"]   = pub.strftime("%Y/%m/%d") if pub else ""

            cd = _parse_date(close_raw)
            if cd:
                t["CLOSING_DATE"] = cd.strftime("%Y/%m/%d")
                t["CLOSING_TIME"] = "11:00"  # always 11:00 per column header

            # Link: point to the portal page (documents require POST, no direct URL)
            t["LINK"] = self._URL
            t["TENDER_TYPE"] = _infer_type(desc)
            self.tenderData.append(t)

        logging.info(f"SITA: {len(self.tenderData)} tender(s)")
        return self.tenderData


class GDoHScraper(_Base):
    """https://e-tenders.gauteng.gov.za/ — SharePoint table, server-rendered (no JS needed)."""
    DEPARTMENT = "Gauteng Department of Health"
    PROVINCE   = "Gauteng"
    SOURCE_KEY = "E-TENDERS.GAUTENG.GOV.ZA"
    _URL       = "https://e-tenders.gauteng.gov.za/Pages/Advertised-Open-Tenders.aspx"

    def run(self):
        logging.info("GDoH: fetching Gauteng e-tenders")
        soup = _get(self._URL)
        if not soup:
            return self.tenderData

        # SharePoint GUID table ID
        table = soup.find("table", id=re.compile(r"g_bd032e40", re.I))
        if not table:
            table = soup.find("table")
        if not table:
            logging.warning("GDoH: no table found")
            return self.tenderData

        # Columns: Tender Number | Department | Description | Briefing Session Date | Closing Date
        for tr in table.find_all("tr")[1:]:
            tds = tr.find_all("td")
            if len(tds) < 3:
                continue
            t = self._blank()
            t["TENDER_ID"]          = tds[0].get_text(strip=True)
            t["TENDER_DESCRIPTION"] = tds[2].get_text(strip=True)
            if not t["TENDER_DESCRIPTION"]:
                continue

            if len(tds) > 3:
                # Use space separator — cell may contain date, time, and "Compulsory" run together
                briefing_raw = tds[3].get_text(" ", strip=True)
                if briefing_raw.strip():
                    t["IS_THERE_A_BRIEFING_SESSION"] = "Yes"
                    bd = _parse_date(briefing_raw)
                    if bd:
                        t["BRIEFING_DATE"] = bd.strftime("%Y/%m/%d")
                    # Extract compulsory flag embedded in the date cell
                    if re.search(r"\bnon[- ]compulsory\b", briefing_raw, re.I):
                        t["COMPULSORY_BRIEFING"] = "No"
                    elif re.search(r"\bcompulsory\b", briefing_raw, re.I):
                        t["COMPULSORY_BRIEFING"] = "Yes"

            if len(tds) > 4:
                t["CLOSING_DATE"], t["CLOSING_TIME"] = _parse_closing(tds[4].get_text(strip=True))

            t["PUBLICATION_DATE"] = ""
            t["TENDER_TYPE"]      = _infer_type(t["TENDER_DESCRIPTION"])
            self.tenderData.append(t)

        logging.info(f"GDoH: {len(self.tenderData)} tender(s)")
        return self.tenderData


class BuffaloCityMMScraper(_Base):
    """https://www.buffalocity.gov.za/tender_documents.php — 3-col table, year+type params."""
    DEPARTMENT = "Buffalo City Metropolitan Municipality"
    PROVINCE   = "Eastern Cape"
    SOURCE_KEY = "BUFFALOCITY.GOV.ZA"

    _TYPES = [
        ("Formal+Tender",                    "Request for Bid"),
        ("Formal+Request+for+Quotation",     "Request for Quotation"),
    ]

    def run(self):
        year = self.date_to.year
        for type_param, tender_type in self._TYPES:
            url = (f"http://www.buffalocity.gov.za/tender_documents.php"
                   f"?year={year}&type={type_param}")
            logging.info(f"Buffalo City: fetching {url}")
            soup = _get(url)
            if not soup:
                continue

            table = soup.find("table")
            if not table:
                continue

            for tr in table.find_all("tr")[1:]:
                tds = tr.find_all("td")
                if len(tds) < 2:
                    continue
                t = self._blank()
                t["TENDER_ID"]          = tds[0].get_text(strip=True)
                t["TENDER_DESCRIPTION"] = tds[1].get_text(strip=True)
                if not t["TENDER_DESCRIPTION"]:
                    continue

                if len(tds) > 2:
                    cd = _parse_date(tds[2].get_text(strip=True))
                    if cd:
                        t["CLOSING_DATE"] = cd.strftime("%Y/%m/%d")

                a = tr.find("a", href=True)
                if a:
                    href = a["href"]
                    if not href.startswith("http"):
                        href = "http://www.buffalocity.gov.za/" + href.lstrip("/")
                    t["LINK"] = href

                t["PUBLICATION_DATE"] = ""
                t["TENDER_TYPE"]      = tender_type
                self.tenderData.append(t)

        logging.info(f"Buffalo City MM: {len(self.tenderData)} tender(s)")
        return self.tenderData


class SIUScraper(_Base):
    """https://www.siu.org.za/current/
    5-column table: TENDER No. | TENDER | BRIEFING SESSION INFO | DATE | VIEW
    DATE column contains closing date (handles 'extended to' and 'Closes on' prefixes).
    """
    DEPARTMENT = "Special Investigating Unit"
    PROVINCE   = "National"
    SOURCE_KEY = "SIU.ORG.ZA"
    _URL       = "https://www.siu.org.za/current/"

    def run(self):
        logging.info("SIU: fetching current tenders")
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        soup = _get(self._URL, verify=False)
        if not soup:
            return self.tenderData

        table = soup.find("table")
        if not table:
            logging.warning("SIU: no table found on page")
            return self.tenderData

        rows = table.find_all("tr")
        for tr in rows[1:]:  # skip header
            tds = tr.find_all(["td", "th"])
            if len(tds) < 4:
                continue

            tender_no  = tds[0].get_text(strip=True)
            desc       = tds[1].get_text(" ", strip=True)
            briefing   = tds[2].get_text(" ", strip=True)
            date_text  = tds[3].get_text(" ", strip=True)

            if not desc or not tender_no:
                continue

            t = self._blank()
            t["TENDER_ID"]          = tender_no
            t["TENDER_DESCRIPTION"] = desc
            t["PUBLICATION_DATE"]   = ""  # not published on page

            # Closing date: prefer "extended to X" if present, else first date found
            ext_m = re.search(
                r"extended to\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s*@\s*(\d{1,2}[h:]\d{2}(?:\s*[AaPp][Mm])?)",
                date_text, re.I
            )
            if ext_m:
                cd = _parse_date(ext_m.group(1))
                if cd:
                    t["CLOSING_DATE"] = cd.strftime("%Y/%m/%d")
                    t["CLOSING_TIME"] = ext_m.group(2)
            else:
                # "Closes on X" or "X @ Y"
                plain = re.sub(r"[Cc]loses?\s+on\s+", "", date_text)
                t["CLOSING_DATE"], t["CLOSING_TIME"] = _parse_closing(plain)

            # Briefing info
            if briefing and briefing.upper() not in ("N/A", "NA", ""):
                brie_m = re.search(
                    r"(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s*@?\s*(\d{1,2}[h:]\d{2})",
                    briefing, re.I
                )
                if brie_m:
                    bd = _parse_date(brie_m.group(1))
                    if bd:
                        t["BRIEFING_DATE"]              = bd.strftime("%Y/%m/%d")
                        t["IS_THERE_A_BRIEFING_SESSION"] = "Yes"
                        compulsory = "non-compulsory" not in briefing.lower() and "compulsory" in briefing.lower()
                        t["COMPULSORY_BRIEFING"] = "Yes" if compulsory else "No"

            # First PDF/doc link in VIEW column
            if len(tds) > 4:
                a = tds[4].find("a", href=True)
                if a:
                    t["LINK"] = a["href"]

            t["TENDER_TYPE"] = _infer_type(desc)
            self.tenderData.append(t)

        logging.info(f"SIU: {len(self.tenderData)} tender(s)")
        return self.tenderData


# ── Registry ──────────────────────────────────────────────────────────────────

# Maps source names from Websites.xlsx to scraper classes.
# Blocked sites (Eskom=403, Transnet=ECONNREFUSED) are excluded.
# Selenium-required sites (Amahlathi, CP JHB, Matatiele) are handled in SeleniumWatchlistScrapers.
SCRAPER_REGISTRY: dict = {
    "Matatiele LM":          MatatieleScraper,   # stub — JS-required; real scraper is Selenium
    "Ntabankulu LM":         NtabankuluScraper,
    "Umzimvubu LM":          UmzimvubuScraper,
    "Winnie Mandela LM":     WinnieMMLScraper,
    "Mnquma LM":             MnqumaScraper,
    "Great Kei LM":          GreatKeiScraper,
    "JOSHCO":                JOSHCOScraper,
    "GPL":                   GPLScraper,
    "RAF":                   RAFScraper,          # stub — SharePoint JS-required
    "Nelson Mandela Bay MM": NMBMMScraper,
    "DEL":                   DELScraper,
    "DIRCO":                 DIRCOScraper,
    "DOJ":                   DOJScraper,
    "W JHB":                 WJHBScraper,
    "SIU":                   SIUScraper,
    "SITA":                  SITAScraper,
    "GDoH":                  GDoHScraper,
    "Buffalo City MM":       BuffaloCityMMScraper,
}

_BLOCKED = {
    "Eskom":    "403 Forbidden (WAF/Cloudflare) — manual monitoring required",
    "Transnet": "ECONNREFUSED — portal may be IP-restricted",
}

# Selenium-based scrapers handled in SeleniumWatchlistScrapers.py
_SELENIUM_HANDLED = {"Amahlathi LM", "CP JHB", "Matatiele LM"}


def run_watchlist_scrapers(date_from: str, date_to: str,
                           watchlist_sources: set,
                           log_queue=None) -> list:
    """
    Run all registered scrapers for sources present in watchlist_sources.
    Returns combined list of tender dicts.
    """
    all_tenders = []

    for src_name, scraper_cls in SCRAPER_REGISTRY.items():
        if src_name not in watchlist_sources:
            continue
        logging.info(f"Watchlist: starting {src_name}")
        try:
            s = scraper_cls(date_from=date_from, date_to=date_to, log_queue=log_queue)
            tenders = s.run()
            all_tenders.extend(tenders)
            logging.info(f"Watchlist {src_name}: {len(tenders)} tender(s) collected")
        except Exception as e:
            logging.error(f"Watchlist {src_name} error: {e}")

    for src, reason in _BLOCKED.items():
        if src in watchlist_sources:
            logging.warning(f"Watchlist: '{src}' skipped — {reason}")

    for src in _SELENIUM_HANDLED:
        if src in watchlist_sources:
            logging.info(f"Watchlist: '{src}' handled by Selenium scraper")

    return all_tenders
