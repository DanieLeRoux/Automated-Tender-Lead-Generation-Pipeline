"""
dhs_tender_pull_and_clean (Department of Human Settlements / Tenders)

Scrapes tender notices from the public Department of Human Settlements tender page,
normalises notice detail pages, filters to currently open tenders, scores entries,
and maintains:

1) MASTER Excel: latest current snapshot per opportunity_key
2) DELTA Excel: brand new tenders discovered this run only
3) UPDATES Excel: existing tenders whose key operational fields changed
4) TOP DELTA Excel: highest-scoring new tenders from the current run
5) RAW CSV/XLSX/HTML: unfiltered source output for audit/debugging

Typical daily run:
    python dhs_tender_pull_and_clean.py

Optional:
    python dhs_tender_pull_and_clean.py --include-expired
    python dhs_tender_pull_and_clean.py --include-cancellations
    python dhs_tender_pull_and_clean.py --top-delta-limit 50

Notes:
- The DHS tenders page uses tabbed/dynamic content. The scraper first tries
  normal requests-based extraction, including the Drupal quicktabs/nojs URLs.
  If that does not return tender detail links, it falls back to Selenium.
- The scraper parses individual tender detail pages, because those pages contain
  the useful fields such as Bid Number, Closing Date, Principle, tender
  information and supporting document links.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import re
import sys
import time
from io import StringIO
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# pylint: disable=abstract-class-instantiated

SOURCE_NAME = "Department of Human Settlements"
SOURCE_URL = "https://www.dhs.gov.za/tenders"
BASE_URL = "https://www.dhs.gov.za"

# Drupal quicktabs fallback endpoints shown by the public tender page.
# 0 = Latest, 1 = Received, 2 = Awarded, 3 = Cancelled, 4 = Closed.
QUICKTAB_URLS = {
    "Latest": "https://www.dhs.gov.za/quicktabs/nojs/tenders/0",
    "Received": "https://www.dhs.gov.za/quicktabs/nojs/tenders/1",
    "Awarded": "https://www.dhs.gov.za/quicktabs/nojs/tenders/2",
    "Cancelled": "https://www.dhs.gov.za/quicktabs/nojs/tenders/3",
    "Closed": "https://www.dhs.gov.za/quicktabs/nojs/tenders/4",
}

# === Folder configuration ===
# Expected location after migration:
#   02_TenderScrapers\02_Sources\DHS\00_Scripts\dhs_tender_pull_and_clean.py
# The paths below are derived from the script location so the scraper can move
# with the framework without hard-coded user folders.
SCRIPT_PATH = Path(__file__).resolve()
SOURCE_ROOT = SCRIPT_PATH.parents[1] if len(SCRIPT_PATH.parents) >= 2 else SCRIPT_PATH.parent
TENDER_SCRAPERS_ROOT = SCRIPT_PATH.parents[3] if len(SCRIPT_PATH.parents) >= 4 else SOURCE_ROOT.parent
BASE_ROOT = str(SOURCE_ROOT)

RAW_DIR = str(SOURCE_ROOT / "01_RawData")
OUTPUT_DIR = str(SOURCE_ROOT / "02_Outputs")
LOG_DIR = str(SOURCE_ROOT / "03_Logs")
STATE_DIR = str(SOURCE_ROOT / "04_State")
METADATA_DIR = str(SOURCE_ROOT / "05_Metadata")

LOG_PATH = os.path.join(LOG_DIR, "dhs_run.log")

MASTER_XLSX = os.path.join(OUTPUT_DIR, "dhs_master.xlsx")
DELTA_XLSX = os.path.join(OUTPUT_DIR, "dhs_delta.xlsx")
UPDATES_XLSX = os.path.join(OUTPUT_DIR, "dhs_updates.xlsx")
TOP_DELTA_XLSX = os.path.join(OUTPUT_DIR, "dhs_top_delta.xlsx")

RAW_CSV = os.path.join(RAW_DIR, "dhs_raw.csv")
RAW_XLSX = os.path.join(RAW_DIR, "dhs_raw_full.xlsx")
RAW_HTML = os.path.join(RAW_DIR, "dhs_raw_page.html")
RAW_RENDERED_HTML = os.path.join(RAW_DIR, "dhs_raw_page_rendered.html")

STATE_FILE_DEFAULT = os.path.join(STATE_DIR, "seen_ids.json")
HASH_VERSION_FILE = os.path.join(STATE_DIR, "hash_version.txt")
LOCK_FILE = os.path.join(STATE_DIR, "run.lock")

TIMEOUT_S = 60
SLEEP_S = 0.25
SELENIUM_WAIT_S = 25

TOP_DELTA_LIMIT_DEFAULT = 20
TOP_DELTA_MIN_SCORE_DEFAULT = 0

# Fields that define a meaningful operational update.
HASH_VERSION = "dhs_hash_v2_no_document_urls_20260704"

HASH_COLS = [
    "bid_number",
    "description",
    "closing_date",
    "principle",
    "current_tenders_information",
    "document_names",
    "cancellation_flag",
]

# Scoring constants. Source is a priority source, then keyword matches rank within it.
WEIGHT_PRIORITY_SOURCE = 10
WEIGHT_DESCRIPTION_KEYWORD = 1

DESCRIPTION_KEYWORD_RULES = {
    "Human Settlements": ["Human Settlements", "Housing", "Settlement"],
    "Title Deeds": ["Title Deed", "Title Deeds", "Deeds"],
    "Conveyancer": ["Conveyancer", "Conveyancing"],
    "Development": ["Development", "Develop"],
    "Programme Management": ["Programme Management", "Program Management"],
    "Project Management": ["Project Management"],
    "Professional Service Provider": ["Professional Service Provider", "Service Provider", "PSP"],
    "Advisory Services": ["Advisory", "Advisory Services"],
    "Feasibility Study": ["Feasibility Study", "Feasibility"],
    "Infrastructure": ["Infrastructure"],
    "Construction": ["Construction", "Construct", "Build"],
    "Civil Engineering": ["Civil Engineering", "Civil"],
    "Structural Engineering": ["Structural Engineering", "Structural"],
    "Engineering": ["Engineering"],
    "Planning": ["Planning", "Spatial Planning", "Town Planning"],
    "Design": ["Design"],
    "Refurbishment": ["Refurbishment", "Refurbish", "Renovation", "Maintenance"],
    "Data / ICT": ["Database", "Data", "ICT", "Information Technology", "Web Interface", "Software"],
    "Research": ["Research", "Study"],
    "Audit / Compliance": ["Audit", "Compliance", "Risk"],
}

DEFAULT_COLUMNS = [
    "source",
    "opportunity_key",
    "bid_number",
    "notice_type",
    "description",
    "client",
    "principle",
    "closing_date",
    "current_tenders_information",
    "document_names",
    "document_urls",
    "document_count",
    "bids_received_names",
    "bids_received_urls",
    "source_url",
    "listing_section",
    "cancellation_flag",
    "priority_source_match",
    "matched_description_keywords",
    "description_keyword_match_count",
    "weight_score",
    "row_hash",
    "run_timestamp",
    "change_type",
    "previous_row_hash",
]


def setup_logging() -> None:
    """Create log directory and configure root logger with file + stdout handlers."""
    os.makedirs(LOG_DIR, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(LOG_PATH, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
        force=True,
    )


def build_session() -> requests.Session:
    """Requests session with retry/backoff for flaky CMS endpoints."""
    sess = requests.Session()
    sess.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-ZA,en;q=0.9",
        }
    )
    retry = Retry(
        total=5,
        connect=5,
        read=5,
        status=5,
        backoff_factor=0.8,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry)
    sess.mount("https://", adapter)
    sess.mount("http://", adapter)
    return sess


def acquire_lock(lock_path: str) -> None:
    os.makedirs(os.path.dirname(lock_path), exist_ok=True)
    if os.path.exists(lock_path):
        raise RuntimeError(f"Lock file exists (another run may be active): {lock_path}")
    with open(lock_path, "w", encoding="utf-8") as fh:
        fh.write(f"pid={os.getpid()}\n")


def release_lock(lock_path: str) -> None:
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
    except OSError:
        logging.exception("Failed to remove lock file: %s", lock_path)


def _normalise_key(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).casefold().strip()
    text = re.sub(r"\s+", " ", text)
    text = text.replace("\\", "/")
    return text


def _normalise_search_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    cleaned = re.sub(r"[^a-z0-9]+", " ", str(value).casefold())
    return " ".join(cleaned.split())


def _contains_phrase(text_norm: str, phrase: str) -> bool:
    phrase_norm = _normalise_search_text(phrase)
    if not phrase_norm:
        return False
    return f" {phrase_norm} " in f" {text_norm} "


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"[ \t\r\f\v]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n", text)
    return text.strip()


def _load_seen_ids(path: str) -> Set[str]:
    try:
        if not os.path.exists(path):
            return set()
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return {_normalise_key(x) for x in (data or []) if _normalise_key(x)}
    except (OSError, ValueError, TypeError) as e:
        logging.exception("Failed to load state file: %s -- %s", path, e)
        return set()


def _read_hash_version(path: str) -> str:
    try:
        if not os.path.exists(path):
            return ""
        with open(path, "r", encoding="utf-8") as fh:
            return fh.read().strip()
    except OSError:
        logging.exception("Failed to read hash version file: %s", path)
        return ""


def _write_hash_version(path: str, version: str) -> None:
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            fh.write(version)
        os.replace(tmp, path)
    except OSError:
        logging.exception("Failed to write hash version file: %s", path)


def _save_seen_ids(path: str, ids: Iterable[str]) -> None:
    try:
        norm = sorted(_normalise_key(x) for x in ids if _normalise_key(x))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(norm, fh, ensure_ascii=False)
        os.replace(tmp, path)
    except (OSError, TypeError) as e:
        logging.exception("Failed to save state file: %s -- %s", path, e)


def fetch_url(session: requests.Session, url: str) -> str:
    response = session.get(url, timeout=TIMEOUT_S)
    response.raise_for_status()
    return response.text


def save_html(path: str, html: str, label: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)
    logging.info("Saved %s HTML: %s", label, path)


def save_quicktab_html(section: str, html: str) -> None:
    safe = re.sub(r"[^a-z0-9]+", "_", section.casefold()).strip("_") or "section"
    path = os.path.join(RAW_DIR, f"dhs_raw_quicktab_{safe}.html")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)
    logging.info("Saved quicktab HTML for %s: %s", section, path)


def is_same_site(url: str) -> bool:
    parsed = urlparse(url)
    return parsed.netloc in {"", "www.dhs.gov.za", "dhs.gov.za"}


def extract_node_id(url: str) -> str:
    match = re.search(r"/node/(\d+)", url)
    return match.group(1) if match else ""


def looks_like_tender_link(text: str, href: str) -> bool:
    combined = _normalise_search_text(f"{text} {href}")
    if not re.search(r"/node/\d+", href):
        return False
    tender_terms = [
        "bid",
        "tender",
        "terms of reference",
        "tor",
        "appointment",
        "service provider",
        "rfp",
        "rfq",
        "rfi",
        "erratum",
        "cancellation",
        "cancelled",
    ]
    return any(_contains_phrase(combined, term) for term in tender_terms)


def extract_detail_links_from_html(html: str, section: str) -> List[Tuple[str, str]]:
    """Return unique (detail_url, listing_section) candidates from a listing HTML page."""
    soup = BeautifulSoup(html or "", "html.parser")
    links: List[Tuple[str, str]] = []
    seen: Set[str] = set()

    for a in soup.find_all("a", href=True):
        href = str(a.get("href") or "").strip()
        text = _clean_text(a.get_text(" ", strip=True))
        if not href:
            continue
        url = urljoin(BASE_URL, href)
        if not is_same_site(url):
            continue
        if not looks_like_tender_link(text, url):
            continue
        if url not in seen:
            links.append((url, section))
            seen.add(url)

    return links


def fetch_listing_htmls() -> List[Tuple[str, str]]:
    """Fetch main and quicktabs/nojs listing pages and return (html, section)."""
    session = build_session()
    htmls: List[Tuple[str, str]] = []

    main_html = fetch_url(session, SOURCE_URL)
    save_html(RAW_HTML, main_html, "static")
    htmls.append((main_html, "Main"))

    for section, url in QUICKTAB_URLS.items():
        try:
            html = fetch_url(session, url)
            save_quicktab_html(section, html)
            htmls.append((html, section))
            time.sleep(SLEEP_S)
        except requests.RequestException as exc:
            logging.warning("Failed to fetch quicktab %s (%s): %s", section, url, exc)

    return htmls


def fetch_rendered_html_with_selenium() -> str:
    """Render the dynamic DHS tender page with Selenium and return page source."""
    try:
        from selenium import webdriver  # type: ignore
        from selenium.webdriver.chrome.options import Options  # type: ignore
        from selenium.webdriver.common.by import By  # type: ignore
        from selenium.webdriver.support import expected_conditions as EC  # type: ignore
        from selenium.webdriver.support.ui import WebDriverWait  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "No tender detail links were found in static HTML and Selenium is not installed. "
            "Install it with: pip install selenium"
        ) from exc

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1600,1200")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--log-level=3")

    driver = webdriver.Chrome(options=options)
    try:
        driver.get(SOURCE_URL)
        wait = WebDriverWait(driver, SELENIUM_WAIT_S)
        wait.until(
            lambda d: (
                len(d.find_elements(By.XPATH, "//a[contains(@href, '/node/')]") ) > 0
                or "Loading content" not in d.page_source
            )
        )
        time.sleep(2)
        html = driver.page_source
        save_html(RAW_RENDERED_HTML, html, "rendered")
        return html
    finally:
        driver.quit()


def collect_detail_links() -> List[Tuple[str, str]]:
    """Collect tender detail URLs from static/quicktabs pages, with Selenium fallback."""
    links: List[Tuple[str, str]] = []
    seen: Set[str] = set()

    for html, section in fetch_listing_htmls():
        for url, listing_section in extract_detail_links_from_html(html, section):
            if url not in seen:
                links.append((url, listing_section))
                seen.add(url)

    logging.info("Found tender detail links before Selenium fallback: %d", len(links))

    if not links:
        rendered_html = fetch_rendered_html_with_selenium()
        for url, listing_section in extract_detail_links_from_html(rendered_html, "Rendered"):
            if url not in seen:
                links.append((url, listing_section))
                seen.add(url)
        logging.info("Found tender detail links after Selenium fallback: %d", len(links))

    return links


def extract_value_after_label(text: str, label: str, next_labels: Sequence[str]) -> str:
    """Extract text after a label until one of the next labels."""
    labels_pattern = "|".join(re.escape(x) for x in next_labels)
    pattern = re.compile(
        rf"(?:^|\n)\s*{re.escape(label)}\s*\n+\s*(.*?)(?=\n\s*(?:{labels_pattern})\s*(?:\n|$)|$)",
        flags=re.IGNORECASE | re.DOTALL,
    )
    match = pattern.search(text)
    if not match:
        return ""
    return _clean_text(match.group(1))


def extract_section_text(text: str, start_label: str, end_labels: Sequence[str]) -> str:
    ends = "|".join(re.escape(x) for x in end_labels)
    pattern = re.compile(
        rf"(?:^|\n)\s*{re.escape(start_label)}\s*\n+(.*?)(?=\n\s*(?:{ends})\s*(?:\n|$)|$)",
        flags=re.IGNORECASE | re.DOTALL,
    )
    match = pattern.search(text)
    return _clean_text(match.group(1)) if match else ""


def parse_date(value: Any) -> pd.Timestamp:
    if value is None or pd.isna(value):
        return pd.NaT
    text = _clean_text(value)
    if not text:
        return pd.NaT

    # Keep the first plausible date, e.g. "22 April 2025 @11:00" -> "22 April 2025".
    date_patterns = [
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\b",
        r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b",
        r"\b\d{1,2}[-/]\d{1,2}[-/]\d{4}\b",
    ]
    for pattern in date_patterns:
        match = re.search(pattern, text)
        if match:
            parsed = pd.to_datetime(match.group(0), errors="coerce", dayfirst=True)
            if pd.notna(parsed):
                return parsed

    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def infer_notice_type(bid_number: str, title: str, info: str) -> str:
    combined = _normalise_search_text(f"{bid_number} {title} {info}")
    if _contains_phrase(combined, "erratum"):
        return "ERRATUM"
    if _contains_phrase(combined, "cancellation") or _contains_phrase(combined, "cancelled"):
        return "CANCELLATION"
    for term in ["RFP", "RFQ", "RFI", "BID", "Tender"]:
        if _contains_phrase(combined, term):
            return term.upper() if term != "Tender" else "TENDER"
    if str(bid_number).strip():
        return "BID"
    return "TENDER"


def find_title(soup: BeautifulSoup) -> str:
    for selector in ["h1", "h2", ".page-title", "title"]:
        node = soup.select_one(selector)
        if node:
            title = _clean_text(node.get_text(" ", strip=True))
            if title and title.casefold() not in {"department of human settlement", "tenders"}:
                return title
    return ""


def collect_document_links(soup: BeautifulSoup, detail_url: str) -> Tuple[str, str, int, str, str]:
    doc_names: List[str] = []
    doc_urls: List[str] = []
    bids_received_names: List[str] = []
    bids_received_urls: List[str] = []

    for a in soup.find_all("a", href=True):
        href = str(a.get("href") or "").strip()
        if not href:
            continue
        url = urljoin(detail_url, href)
        text = _clean_text(a.get_text(" ", strip=True)) or os.path.basename(urlparse(url).path)
        href_norm = href.casefold()
        text_norm = _normalise_search_text(text)

        # Keep common tender-document links. DHS usually stores PDFs, but this is broader.
        if not re.search(r"\.(pdf|doc|docx|xls|xlsx|zip|rar)(\?|#|$)", href_norm):
            continue

        if _contains_phrase(text_norm, "bids received") or _contains_phrase(text_norm, "bid opening"):
            bids_received_names.append(text)
            bids_received_urls.append(url)
        else:
            doc_names.append(text)
            doc_urls.append(url)

    return (
        "; ".join(doc_names),
        "; ".join(doc_urls),
        len(doc_urls),
        "; ".join(bids_received_names),
        "; ".join(bids_received_urls),
    )


def parse_detail_page(html: str, detail_url: str, listing_section: str) -> Dict[str, Any]:
    soup = BeautifulSoup(html or "", "html.parser")
    title = find_title(soup)
    page_text = _clean_text(soup.get_text("\n", strip=True))

    known_labels = [
        "Bid Number",
        "Closing Date",
        "Principle",
        "Current tenders information",
        "Supporting documents",
        "Bids received",
        "CONTACT US",
        "FIND US",
        "EXPLORE",
        "MAP",
    ]

    bid_number = extract_value_after_label(page_text, "Bid Number", known_labels)
    closing_raw = extract_value_after_label(page_text, "Closing Date", known_labels)
    principle = extract_value_after_label(page_text, "Principle", known_labels)
    current_info = extract_section_text(
        page_text,
        "Current tenders information",
        ["Supporting documents", "Bids received", "CONTACT US", "FIND US", "EXPLORE", "MAP"],
    )

    # Fallbacks if labels are not captured exactly.
    if not bid_number:
        match = re.search(r"\b(?:BID|RFP|RFQ|RFI)\s+[A-Z0-9/._ -]+", page_text, flags=re.IGNORECASE)
        bid_number = _clean_text(match.group(0)) if match else ""

    if not closing_raw:
        match = re.search(r"Closing\s+Date\s*[:\-]?\s*([^\n]+)", page_text, flags=re.IGNORECASE)
        closing_raw = _clean_text(match.group(1)) if match else ""

    closing_date = parse_date(closing_raw)
    doc_names, doc_urls, document_count, bid_names, bid_urls = collect_document_links(soup, detail_url)

    description = title or current_info.split("\n", 1)[0] if current_info else title
    description = _clean_text(description)

    combined = f"{title}\n{bid_number}\n{current_info}\n{doc_names}"
    cancellation_flag = bool(
        re.search(r"\b(cancelled|cancellation|cancelled due to|withdrawn)\b", combined, flags=re.IGNORECASE)
    )

    node_id = extract_node_id(detail_url)
    key_base = bid_number or node_id or detail_url
    opportunity_key = "dhs:" + re.sub(r"[^a-z0-9]+", "-", key_base.casefold()).strip("-")

    return {
        "source": SOURCE_NAME,
        "opportunity_key": opportunity_key,
        "bid_number": bid_number,
        "notice_type": infer_notice_type(bid_number, title, current_info),
        "description": description,
        "client": SOURCE_NAME,
        "principle": principle,
        "closing_date": closing_date,
        "current_tenders_information": current_info,
        "document_names": doc_names,
        "document_urls": doc_urls,
        "document_count": document_count,
        "bids_received_names": bid_names,
        "bids_received_urls": bid_urls,
        "source_url": detail_url,
        "listing_section": listing_section,
        "cancellation_flag": cancellation_flag,
    }


def fetch_and_parse() -> List[Dict[str, Any]]:
    """Fetch listing links, parse each detail page and return raw row dicts."""
    links = collect_detail_links()
    session = build_session()
    rows: List[Dict[str, Any]] = []
    seen: Set[str] = set()

    for detail_url, section in links:
        if detail_url in seen:
            continue
        seen.add(detail_url)
        try:
            html = fetch_url(session, detail_url)
            row = parse_detail_page(html, detail_url, section)
            rows.append(row)
            time.sleep(SLEEP_S)
        except requests.RequestException as exc:
            logging.warning("Failed to fetch tender detail page %s: %s", detail_url, exc)
        except Exception as exc:  # deliberately broad for parser resilience
            logging.warning("Failed to parse tender detail page %s: %s", detail_url, exc)

    logging.info("Parsed DHS tender detail rows: %d", len(rows))
    return rows


def ensure_columns(df: pd.DataFrame, columns: Sequence[str]) -> pd.DataFrame:
    out = df.copy()
    for col in columns:
        if col not in out.columns:
            out[col] = pd.NA
    return out[list(columns)]


def filter_current(df: pd.DataFrame, include_expired: bool, include_cancellations: bool) -> pd.DataFrame:
    out = df.copy()

    if not include_cancellations and "cancellation_flag" in out.columns:
        before = len(out)
        out = out[~out["cancellation_flag"].fillna(False).astype(bool)].copy()
        logging.info("Filtered cancellation/withdrawal notices: %d -> %d rows", before, len(out))

    if not include_expired:
        before = len(out)
        if "closing_date" not in out.columns:
            out = out.iloc[0:0].copy()
        else:
            closing = pd.to_datetime(out["closing_date"], errors="coerce")
            today = pd.Timestamp.today().normalize()
            out = out[pd.notna(closing) & (closing >= today)].copy()
        logging.info("Filtered current tenders by closing date: %d -> %d rows", before, len(out))

    return out


def add_priority_scoring(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["priority_source_match"] = True

    def matched_keywords(row: pd.Series) -> str:
        combined = " ".join(
            str(row.get(col, "") or "")
            for col in ["description", "current_tenders_information", "document_names", "bid_number"]
        )
        text_norm = _normalise_search_text(combined)
        matches = []
        for label, phrases in DESCRIPTION_KEYWORD_RULES.items():
            if any(_contains_phrase(text_norm, phrase) for phrase in phrases):
                matches.append(label)
        return "; ".join(matches)

    if len(out) == 0:
        out["matched_description_keywords"] = pd.Series(dtype="object")
        out["description_keyword_match_count"] = pd.Series(dtype="int64")
        out["weight_score"] = pd.Series(dtype="int64")
        return out

    out["matched_description_keywords"] = out.apply(matched_keywords, axis=1)
    out["description_keyword_match_count"] = out["matched_description_keywords"].apply(
        lambda x: 0 if not x else len(str(x).split("; "))
    )
    out["weight_score"] = (
        out["priority_source_match"].astype(int) * WEIGHT_PRIORITY_SOURCE
        + out["description_keyword_match_count"].astype(int) * WEIGHT_DESCRIPTION_KEYWORD
    )

    sort_cols = [c for c in ["weight_score", "closing_date"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

    return out


def compute_row_hash(df: pd.DataFrame, cols: Sequence[str]) -> pd.Series:
    use_cols = [c for c in cols if c in df.columns]
    if len(df) == 0:
        return pd.Series(dtype="object", index=df.index)
    if not use_cols:
        return pd.Series([""] * len(df), index=df.index)

    tmp = df[use_cols].copy()
    for c in use_cols:
        if pd.api.types.is_datetime64_any_dtype(tmp[c]):
            tmp[c] = tmp[c].dt.strftime("%Y-%m-%d %H:%M:%S").fillna("")
        else:
            tmp[c] = tmp[c].fillna("").astype(str).str.strip()

    joined = tmp.apply(lambda r: "||".join(r.values.tolist()), axis=1)
    return joined.apply(lambda s: hashlib.sha256(s.encode("utf-8")).hexdigest())


def write_excel_with_table(df: pd.DataFrame, path: str, sheet_name: str, table_name: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp.xlsx"
    out_df = df.copy() if df is not None else pd.DataFrame(columns=DEFAULT_COLUMNS)

    if out_df.shape[1] == 0:
        out_df = pd.DataFrame(columns=DEFAULT_COLUMNS)

    with pd.ExcelWriter(tmp, engine="xlsxwriter", datetime_format="yyyy-mm-dd", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        out_df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        nrows, ncols = out_df.shape
        columns = [{"header": col} for col in out_df.columns]
        if nrows > 0 and ncols > 0:
            worksheet.add_table(
                0,
                0,
                nrows,
                ncols - 1,
                {"name": table_name, "columns": columns},
            )
        elif ncols > 0:
            worksheet.autofilter(0, 0, 0, ncols - 1)
        worksheet.freeze_panes(1, 0)
    os.replace(tmp, path)


def load_master(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame(columns=DEFAULT_COLUMNS)
    try:
        return pd.read_excel(path, sheet_name="Master", engine="openpyxl")
    except Exception as exc:
        logging.exception("Failed to read master file (%s): %s", path, exc)
        return pd.DataFrame(columns=DEFAULT_COLUMNS)


def append_and_dedupe_master(master: pd.DataFrame, incoming: pd.DataFrame) -> pd.DataFrame:
    if master.empty:
        out = incoming.copy()
        out["_source_priority"] = 1
    else:
        master_part = master.copy()
        incoming_part = incoming.copy()
        master_part["_source_priority"] = 0
        incoming_part["_source_priority"] = 1
        out = pd.concat([master_part, incoming_part], ignore_index=True)

    if "opportunity_key" not in out.columns:
        return out.drop(columns=["_source_priority"], errors="ignore")

    out["opportunity_key"] = out["opportunity_key"].fillna("").astype(str).str.strip().str.lower()
    out = out[out["opportunity_key"].ne("")]

    if "closing_date" in out.columns:
        out["closing_date"] = pd.to_datetime(out["closing_date"], errors="coerce")
        out = out.sort_values(
            ["closing_date", "_source_priority"],
            ascending=[True, False],
            na_position="last",
            kind="mergesort",
        )
    else:
        out = out.sort_values("_source_priority", ascending=False, kind="mergesort")

    out = out.drop_duplicates(subset=["opportunity_key"], keep="first")
    out = out.drop(columns=["_source_priority"], errors="ignore")
    return out


def build_top_delta(df: pd.DataFrame, limit: int, min_score: int) -> pd.DataFrame:
    out = df.copy()
    if "weight_score" not in out.columns:
        out["weight_score"] = 0
    if min_score is not None:
        out = out[out["weight_score"].fillna(0).astype(int) >= int(min_score)].copy()
    sort_cols = [c for c in ["weight_score", "closing_date"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")
    if limit and int(limit) > 0:
        out = out.head(int(limit)).copy()
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description="Fetch and clean DHS tender data")
    parser.add_argument("--include-expired", action="store_true", help="keep expired/closed tenders in clean outputs")
    parser.add_argument("--include-cancellations", action="store_true", help="keep cancellation/withdrawal notices")
    parser.add_argument("--top-delta-limit", type=int, default=TOP_DELTA_LIMIT_DEFAULT,
                        help="maximum rows for dhs_top_delta.xlsx. Use 0 for all qualifying rows.")
    parser.add_argument("--top-delta-min-score", type=int, default=TOP_DELTA_MIN_SCORE_DEFAULT,
                        help="minimum weight_score required for dhs_top_delta.xlsx")
    args = parser.parse_args()

    for path in [RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR]:
        os.makedirs(path, exist_ok=True)
    setup_logging()
    logging.info("Using tender scrapers root: %s", TENDER_SCRAPERS_ROOT)
    logging.info("Using source root: %s", SOURCE_ROOT)
    logging.info("Using raw directory: %s", RAW_DIR)
    logging.info("Using output directory: %s", OUTPUT_DIR)
    logging.info("Using log directory: %s", LOG_DIR)
    logging.info("Using state directory: %s", STATE_DIR)

    run_ts = pd.Timestamp.now()

    try:
        acquire_lock(LOCK_FILE)

        rows = fetch_and_parse()
        raw_df = pd.DataFrame(rows)
        raw_df = ensure_columns(raw_df, [c for c in DEFAULT_COLUMNS if c not in {"row_hash", "run_timestamp", "change_type", "previous_row_hash", "weight_score", "matched_description_keywords", "description_keyword_match_count", "priority_source_match"}])
        logging.info("Parsed DHS raw rows: %d | cols: %d", len(raw_df), len(raw_df.columns))

        raw_df.to_csv(RAW_CSV, index=False)
        logging.info("Saved raw CSV: %s", RAW_CSV)
        with pd.ExcelWriter(RAW_XLSX, engine="xlsxwriter", datetime_format="yyyy-mm-dd", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
            raw_df.to_excel(writer, index=False, sheet_name="Raw_Data")
        logging.info("Saved raw Excel: %s", RAW_XLSX)

        mapped = raw_df.copy()
        if "closing_date" in mapped.columns:
            mapped["closing_date"] = pd.to_datetime(mapped["closing_date"], errors="coerce")

        mapped = filter_current(
            mapped,
            include_expired=args.include_expired,
            include_cancellations=args.include_cancellations,
        )
        mapped = add_priority_scoring(mapped)
        mapped["row_hash"] = compute_row_hash(mapped, HASH_COLS)
        mapped["run_timestamp"] = run_ts

        master_df = load_master(MASTER_XLSX)
        master_df = ensure_columns(master_df, DEFAULT_COLUMNS)
        if not master_df.empty:
            if "opportunity_key" in master_df.columns:
                master_df["opportunity_key"] = master_df["opportunity_key"].fillna("").astype(str).str.strip().str.lower()
                master_df = master_df[master_df["opportunity_key"].ne("")]
            if "row_hash" not in master_df.columns or master_df["row_hash"].isna().all():
                master_df["row_hash"] = compute_row_hash(master_df, HASH_COLS)

        seen_ids = _load_seen_ids(STATE_FILE_DEFAULT)
        recorded_hash_version = _read_hash_version(HASH_VERSION_FILE)
        suppress_alerts = False
        if recorded_hash_version != HASH_VERSION:
            logging.info("DHS hash/state version changed or was not recorded. Re-baselining and suppressing delta/update alerts for this run.")
            suppress_alerts = True
        elif not seen_ids and not master_df.empty:
            logging.info("No DHS seen_ids.json found. Bootstrapping state from existing master and suppressing delta/update alerts for this run.")
            suppress_alerts = True

        if suppress_alerts:
            delta_df = mapped.iloc[0:0].copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""
            updates_df = mapped.iloc[0:0].copy()
            updates_df["change_type"] = "UPDATED"
            updates_df["previous_row_hash"] = ""
        elif not master_df.empty and "opportunity_key" in master_df.columns:
            master_ids = set(master_df["opportunity_key"].fillna("").astype(str).str.lower().tolist())
            master_hash_map = dict(
                zip(
                    master_df["opportunity_key"].fillna("").astype(str).str.lower().tolist(),
                    master_df["row_hash"].fillna("").astype(str).tolist(),
                )
            )

            is_new = ~mapped["opportunity_key"].isin(master_ids) if "opportunity_key" in mapped.columns else pd.Series(False, index=mapped.index)
            delta_df = mapped[is_new].copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""

            exists = ~is_new
            mapped_prev_hash = mapped["opportunity_key"].map(master_hash_map).fillna("") if "opportunity_key" in mapped.columns else pd.Series("", index=mapped.index)
            is_updated = exists & (mapped["row_hash"] != mapped_prev_hash)
            updates_df = mapped[is_updated].copy()
            updates_df["change_type"] = "UPDATED"
            updates_df["previous_row_hash"] = updates_df["opportunity_key"].map(master_hash_map).fillna("") if "opportunity_key" in updates_df.columns else ""
        else:
            delta_df = mapped.copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""
            updates_df = mapped.iloc[0:0].copy()
            updates_df["change_type"] = "UPDATED"
            updates_df["previous_row_hash"] = ""

        sort_cols = [c for c in ["weight_score", "closing_date"] if c in delta_df.columns]
        if sort_cols:
            ascending = [False if c == "weight_score" else True for c in sort_cols]
            delta_df = delta_df.sort_values(sort_cols, ascending=ascending, na_position="last")
            updates_df = updates_df.sort_values(sort_cols, ascending=ascending, na_position="last")

        top_delta_df = build_top_delta(delta_df, args.top_delta_limit, args.top_delta_min_score)

        logging.info("Delta rows this run after filters: %d", len(delta_df))
        logging.info("Updates rows this run after filters: %d", len(updates_df))
        logging.info("Top delta rows this run: %d", len(top_delta_df))

        write_excel_with_table(ensure_columns(delta_df, DEFAULT_COLUMNS), DELTA_XLSX, "Delta", "tbl_delta")
        logging.info("Saved delta Excel: %s | rows: %d", DELTA_XLSX, len(delta_df))

        write_excel_with_table(ensure_columns(updates_df, DEFAULT_COLUMNS), UPDATES_XLSX, "Updates", "tbl_updates")
        logging.info("Saved updates Excel: %s | rows: %d", UPDATES_XLSX, len(updates_df))

        write_excel_with_table(ensure_columns(top_delta_df, DEFAULT_COLUMNS), TOP_DELTA_XLSX, "Top_Delta", "tbl_top_delta")
        logging.info("Saved top delta Excel: %s | rows: %d", TOP_DELTA_XLSX, len(top_delta_df))

        updated_master = append_and_dedupe_master(master_df, mapped)
        updated_master = ensure_columns(updated_master, DEFAULT_COLUMNS)
        write_excel_with_table(updated_master, MASTER_XLSX, "Master", "tbl_master")
        logging.info("Saved master Excel: %s | rows: %d", MASTER_XLSX, len(updated_master))

        if "opportunity_key" in updated_master.columns:
            _save_seen_ids(STATE_FILE_DEFAULT, updated_master["opportunity_key"].fillna("").astype(str).tolist())
        _write_hash_version(HASH_VERSION_FILE, HASH_VERSION)

        return 0

    except (requests.RequestException, ValueError, RuntimeError, OSError) as exc:
        logging.exception("FAILED: %s", exc)
        return 1
    finally:
        release_lock(LOCK_FILE)


if __name__ == "__main__":
    raise SystemExit(main())
