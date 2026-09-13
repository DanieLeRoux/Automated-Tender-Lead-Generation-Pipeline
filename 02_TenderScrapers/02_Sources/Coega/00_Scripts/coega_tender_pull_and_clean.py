"""
coega_tender_pull_and_clean (Delta + Updates + Master + Top Delta)

Scrapes current tenders from the Coega Development Corporation tender page,
cleans/normalises them, and maintains:

1) MASTER Excel: latest snapshot per opportunity_key
2) DELTA Excel: brand new tenders discovered this run only
3) UPDATES Excel: existing tenders whose key operational fields changed
4) TOP DELTA Excel: highest-scoring new tenders from the current run

Typical run:
    python coega_tender_pull_and_clean.py

Optional:
    python coega_tender_pull_and_clean.py --include-expired
    python coega_tender_pull_and_clean.py --include-cancellations
    python coega_tender_pull_and_clean.py --top-delta-limit 50

Notes:
- The Coega page is JavaScript-rendered. The scraper first tries a normal
  requests fetch, then falls back to Selenium rendering if no rows are found.
- Raw HTML and raw tender rows are saved for audit/debugging.
"""

from __future__ import annotations

import os
import re
import sys
import time
import json
import logging
import argparse
import hashlib
from io import StringIO
from pathlib import Path
from urllib.parse import urljoin
from typing import Any, Dict, List, Optional, Sequence

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# -----------------------------------------------------------------------------
# Source and folder configuration
# -----------------------------------------------------------------------------
SOURCE = "COEGA"
SOURCE_URL = "https://www.coega.co.za/site/tenders/current-tenders"

# Folder layout:
# 02_TenderScrapers/02_Sources/Coega/00_Scripts/coega_tender_pull_and_clean.py
SCRIPT_PATH = Path(__file__).resolve()
SOURCE_ROOT = SCRIPT_PATH.parents[1]
TENDER_SCRAPERS_ROOT = SCRIPT_PATH.parents[3] if len(SCRIPT_PATH.parents) >= 4 else SOURCE_ROOT

BASE_ROOT = str(SOURCE_ROOT)
RAW_DIR = str(SOURCE_ROOT / "01_RawData")
OUTPUT_DIR = str(SOURCE_ROOT / "02_Outputs")
LOG_DIR = str(SOURCE_ROOT / "03_Logs")
STATE_DIR = str(SOURCE_ROOT / "04_State")
METADATA_DIR = str(SOURCE_ROOT / "05_Metadata")

LOG_PATH = os.path.join(LOG_DIR, "coega_run.log")
LOCK_FILE = os.path.join(STATE_DIR, "run.lock")
STATE_FILE_DEFAULT = os.path.join(STATE_DIR, "seen_ids.json")
HASH_VERSION_FILE = os.path.join(STATE_DIR, "hash_version.txt")
HASH_VERSION = "coega_hash_v2"

RAW_CSV = os.path.join(RAW_DIR, "coega_raw.csv")
RAW_XLSX = os.path.join(RAW_DIR, "coega_raw_full.xlsx")
RAW_STATIC_HTML = os.path.join(RAW_DIR, "coega_raw_page_static.html")
RAW_RENDERED_HTML = os.path.join(RAW_DIR, "coega_raw_page_rendered.html")

MASTER_XLSX = os.path.join(OUTPUT_DIR, "coega_master.xlsx")
DELTA_XLSX = os.path.join(OUTPUT_DIR, "coega_delta.xlsx")
UPDATES_XLSX = os.path.join(OUTPUT_DIR, "coega_updates.xlsx")
TOP_DELTA_XLSX = os.path.join(OUTPUT_DIR, "coega_top_delta.xlsx")

TIMEOUT_S = 60
SLEEP_S = 0.4
SELENIUM_WAIT_S = 12

# Key fields that define a meaningful update.
HASH_COLS = [
    "description",
    "tender_number",
    "notice_type",
    "date_published",
    "closing_date",
    "closing_datetime",
    "document_names",
    "document_count",
]

TOP_DELTA_LIMIT_DEFAULT = 20
TOP_DELTA_MIN_SCORE_DEFAULT = 0

WEIGHT_DESCRIPTION_KEYWORD = 1
WEIGHT_NOTICE_TYPE_PRIORITY = 3
WEIGHT_COEGA_MATCH = 2

DESCRIPTION_KEYWORD_RULES = {
    "Water and Sanitation": ["Water and Sanitation"],
    "Housing": ["Housing"],
    "Education": ["Education", "School", "Primary School", "Special School"],
    "Transport": ["Transport", "Road", "Roads", "Bridge"],
    "Nuclear": ["Nuclear"],
    "Bus Rapid Transit (BRT)": ["Bus Rapid Transit", "BRT"],
    "Design": ["Design"],
    "Refurbishment": ["Refurbishment", "Refurbish", "Repairs", "Maintenance"],
    "Engineering, Procurement, and Construction (EPC)": [
        "Engineering, Procurement, and Construction",
        "Engineering Procurement and Construction",
        "EPC",
    ],
    "Fire Engineering": ["Fire Engineering", "Fire Extinguishing", "Fire Detection", "Sprinkler"],
    "HVAC": ["HVAC", "Heating Ventilation", "Air Conditioning", "Refrigeration"],
    "Mechanical": ["Mechanical"],
    "Electrical": ["Electrical", "Emergency Power", "Generator"],
    "Electronic": ["Electronic"],
    "Wet Services": ["Wet Services"],
    "Project Management": ["Project Management"],
    "Programme Management": ["Programme Management"],
    "Project Management Office (PMO)": ["Project Management Office", "PMO"],
    "Advisory Services": ["Advisory Services", "Consulting", "Consultancy"],
    "Structural Engineering": ["Structural Engineering"],
    "Civil Engineering": ["Civil Engineering", "Civil Works"],
    "Feasibility Study": ["Feasibility Study"],
    "Construction": ["Construction", "Construct", "Completion Contract"],
}

DEFAULT_COLUMNS = [
    "source",
    "opportunity_key",
    "tender_number",
    "notice_type",
    "description",
    "client",
    "province",
    "date_published",
    "closing_date",
    "closing_datetime",
    "document_names",
    "document_urls",
    "document_count",
    "source_url",
    "cancellation_flag",
    "priority_notice_type_match",
    "matched_description_keywords",
    "description_keyword_match_count",
    "weight_score",
    "row_hash",
    "run_timestamp",
    "change_type",
    "previous_row_hash",
]


# -----------------------------------------------------------------------------
# Logging, requests and lock handling
# -----------------------------------------------------------------------------
def setup_logging() -> None:
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


def ensure_directories() -> None:
    for directory in [RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR]:
        os.makedirs(directory, exist_ok=True)


def build_session() -> requests.Session:
    sess = requests.Session()
    sess.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/122 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
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
        raise RuntimeError(f"Lock file exists; another run may be active: {lock_path}")
    with open(lock_path, "w", encoding="utf-8") as fh:
        fh.write(f"pid={os.getpid()}\n")


def release_lock(lock_path: str) -> None:
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
    except OSError:
        logging.exception("Failed to remove lock file: %s", lock_path)


def _load_seen_ids(path: str) -> set[str]:
    try:
        if not os.path.exists(path):
            return set()
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return set(clean_text(x).casefold() for x in (data or []) if clean_text(x))
    except (OSError, ValueError, TypeError) as exc:
        logging.exception("Failed to load state file: %s -- %s", path, exc)
        return set()


def _save_seen_ids(path: str, ids: Sequence[str]) -> None:
    try:
        norm = sorted(clean_text(x).casefold() for x in ids if clean_text(x))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(norm, fh, ensure_ascii=False)
        os.replace(tmp, path)
    except (OSError, TypeError) as exc:
        logging.exception("Failed to save state file: %s -- %s", path, exc)


def _read_text_file(path: str) -> str:
    try:
        if not os.path.exists(path):
            return ""
        with open(path, "r", encoding="utf-8") as fh:
            return fh.read().strip()
    except OSError:
        return ""


def _write_text_file(path: str, value: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        fh.write(value)
    os.replace(tmp, path)


def _empty_change_frame(mapped: pd.DataFrame, change_type: str) -> pd.DataFrame:
    out = mapped.iloc[0:0].copy()
    out["change_type"] = change_type
    out["previous_row_hash"] = ""
    return out


# -----------------------------------------------------------------------------
# Text/date utilities
# -----------------------------------------------------------------------------
def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalise_key_text(value: Any) -> str:
    text = clean_text(value).casefold()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def normalise_search_text(value: Any) -> str:
    return normalise_key_text(value)


def contains_phrase(text_norm: str, phrase: str) -> bool:
    phrase_norm = normalise_search_text(phrase)
    if not phrase_norm:
        return False
    return f" {phrase_norm} " in f" {text_norm} "


def parse_date(value: Any) -> pd.Timestamp:
    if value is None:
        return pd.NaT
    text = clean_text(value)
    if not text:
        return pd.NaT

    # Normalise common date formats: 2026-05-29, 29 May 2026, 29/05/2026, etc.
    text = re.sub(r"(\d{1,2})(st|nd|rd|th)", r"\1", text, flags=re.I)
    text = text.replace("@", " ")
    text = clean_text(text)

    # Avoid pandas warnings by parsing ISO-style dates with dayfirst=False.
    if re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}", text):
        return pd.to_datetime(text, errors="coerce", dayfirst=False)

    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def extract_first_date(text: str) -> str:
    text = clean_text(text)
    patterns = [
        r"\b\d{4}-\d{2}-\d{2}\b",
        r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",
        r"\b\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\s+\d{4}\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            return match.group(0)
    return ""


def extract_labelled_date(text: str, label: str) -> str:
    text = clean_text(text)
    escaped = re.escape(label)
    pattern = rf"{escaped}\s*[:\-]?\s*((?:\d{{4}}-\d{{2}}-\d{{2}})|(?:\d{{1,2}}[/-]\d{{1,2}}[/-]\d{{2,4}})|(?:\d{{1,2}}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\s+\d{{4}}))"
    match = re.search(pattern, text, flags=re.I)
    return match.group(1) if match else ""


def extract_tender_number(text: str) -> str:
    text = clean_text(text)
    patterns = [
        r"\bCDC\s*/\s*\d+\s*/\s*\d+(?:\s*(?:ET\s*AL|[A-Z]))?\b",
        r"\bCDC[-\s]\d+[-/]\d+(?:\s*(?:ET\s*AL|[A-Z]))?\b",
        r"\b(?:RFP|RFQ|RFI|EOI)\s*[-/]?\s*\d+[A-Z0-9/-]*\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            return clean_text(match.group(0)).upper().replace(" ", "")
    return ""


def infer_notice_type(text: str) -> str:
    norm = normalise_search_text(text)
    if contains_phrase(norm, "tender cancellation") or contains_phrase(norm, "cancelled"):
        return "Cancellation"
    if contains_phrase(norm, "request for proposal") or re.search(r"\brfp\b", norm):
        return "RFP"
    if contains_phrase(norm, "request for quotation") or re.search(r"\brfq\b", norm):
        return "RFQ"
    if contains_phrase(norm, "request for information") or re.search(r"\brfi\b", norm):
        return "RFI"
    if contains_phrase(norm, "expression of interest") or re.search(r"\beoi\b", norm):
        return "EOI"
    if contains_phrase(norm, "invitation to tender"):
        return "Invitation to Tender"
    return "Tender"


def is_cancellation(text: str) -> bool:
    norm = normalise_search_text(text)
    phrases = ["tender cancellation", "cancelled", "cancellation", "withdrawn", "withdrawal"]
    return any(contains_phrase(norm, phrase) for phrase in phrases)


# -----------------------------------------------------------------------------
# Fetching and parsing
# -----------------------------------------------------------------------------
def fetch_static_html() -> str:
    sess = build_session()
    response = sess.get(SOURCE_URL, timeout=TIMEOUT_S)
    response.raise_for_status()
    html = response.text
    os.makedirs(os.path.dirname(RAW_STATIC_HTML), exist_ok=True)
    with open(RAW_STATIC_HTML, "w", encoding="utf-8") as fh:
        fh.write(html)
    logging.info("Saved static HTML: %s", RAW_STATIC_HTML)
    return html


def fetch_rendered_html() -> str:
    """Render the JavaScript page using Selenium and return page_source."""
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
    except ImportError as exc:
        raise RuntimeError(
            "No tender rows found in static HTML and Selenium is not installed. "
            "Install it with: pip install selenium"
        ) from exc

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/122 Safari/537.36")

    driver = webdriver.Chrome(options=options)
    try:
        driver.get(SOURCE_URL)
        wait = WebDriverWait(driver, SELENIUM_WAIT_S)
        try:
            wait.until(
                lambda d: "PUBLISHED DATE" in d.page_source.upper()
                or "CLOSING DATE" in d.page_source.upper()
                or len(d.find_elements(By.TAG_NAME, "a")) > 20
            )
        except Exception:
            logging.warning("Timed out waiting for Coega tender content; saving whatever rendered.")
        time.sleep(2)
        html = driver.page_source
    finally:
        driver.quit()

    os.makedirs(os.path.dirname(RAW_RENDERED_HTML), exist_ok=True)
    with open(RAW_RENDERED_HTML, "w", encoding="utf-8") as fh:
        fh.write(html)
    logging.info("Saved rendered HTML: %s", RAW_RENDERED_HTML)
    return html


def candidate_elements(soup: BeautifulSoup) -> List[Any]:
    """Find likely tender-card elements using marker text and size heuristics."""
    candidates = []
    for el in soup.find_all(True):
        text = clean_text(el.get_text(" ", strip=True))
        upper = text.upper()
        if "PUBLISHED DATE" in upper and "CLOSING DATE" in upper:
            candidates.append((len(text), el))

    # Prefer small, card-like elements. Reject elements that clearly contain the whole page.
    candidates = [(size, el) for size, el in candidates if 80 <= size <= 5000]
    candidates.sort(key=lambda x: x[0])

    selected = []
    selected_texts = set()
    for _, el in candidates:
        text = clean_text(el.get_text(" ", strip=True))
        text_key = normalise_key_text(text)[:600]
        if not text_key or text_key in selected_texts:
            continue
        # Avoid selecting nested duplicates where an existing smaller selected block has same text.
        if any(text_key in normalise_key_text(clean_text(existing.get_text(" ", strip=True))) for existing in selected):
            continue
        selected.append(el)
        selected_texts.add(text_key)

    return selected


def parse_from_tables(soup: BeautifulSoup) -> List[Dict[str, Any]]:
    """Parse conventional HTML tables if the rendered page exposes one.

    Pandas 3 / newer lxml combinations can treat a raw HTML string as a
    file path. Wrapping the HTML in StringIO forces pandas to parse the
    string content instead of trying to open it as a local file.
    """
    rows: List[Dict[str, Any]] = []
    try:
        dfs = pd.read_html(StringIO(str(soup)))
    except (ValueError, OSError, FileNotFoundError) as exc:
        logging.info("No parseable Coega HTML tables found: %s", exc)
        return rows

    for df in dfs:
        if df.empty:
            continue
        df.columns = [clean_text(c).lower().replace(" ", "_") for c in df.columns]
        cols = set(df.columns)
        if not ({"published_date", "closing_date"} & cols or {"description", "closing_date"} <= cols):
            continue
        for _, rec in df.iterrows():
            text = " | ".join(clean_text(x) for x in rec.tolist())
            rows.append(
                {
                    "source": SOURCE,
                    "tender_number": extract_tender_number(text),
                    "notice_type": infer_notice_type(text),
                    "description": clean_text(
                        rec.get("description")
                        or rec.get("tender_description")
                        or rec.get("name")
                        or text
                    ),
                    "client": "Coega Development Corporation",
                    "province": "Eastern Cape",
                    "date_published": parse_date(rec.get("published_date") or extract_labelled_date(text, "Published Date")),
                    "closing_date": parse_date(rec.get("closing_date") or extract_labelled_date(text, "Closing Date")),
                    "closing_datetime": clean_text(rec.get("closing_date") or extract_labelled_date(text, "Closing Date")),
                    "document_names": "",
                    "document_urls": "",
                    "document_count": 0,
                    "source_url": SOURCE_URL,
                    "cancellation_flag": is_cancellation(text),
                }
            )
    return rows


def build_description(block_text: str, tender_number: str) -> str:
    """Derive a clean description from a tender block."""
    text = clean_text(block_text)

    # Remove common labels but keep tender title content.
    text = re.sub(r"PUBLISHED DATE\s*[:\-]?\s*", " ", text, flags=re.I)
    text = re.sub(r"CLOSING DATE\s*[:\-]?\s*", " ", text, flags=re.I)
    text = re.sub(r"\b\d{4}-\d{2}-\d{2}\b", " ", text)
    text = re.sub(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b", " ", text)
    text = re.sub(r"\b\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\s+\d{4}\b", " ", text, flags=re.I)

    # Remove repeated tender number from the front only; keep if part of official description.
    if tender_number:
        text = re.sub(re.escape(tender_number), " ", text, flags=re.I)
        spaced = tender_number.replace("/", " / ")
        text = re.sub(re.escape(spaced), " ", text, flags=re.I)

    # Remove obvious navigation/download words that can leak into cards.
    text = re.sub(r"\b(download|documents?|current tenders?|read more|view|click here)\b", " ", text, flags=re.I)
    text = clean_text(text)

    return text[:1500]



def is_coega_tender_document_url(url: str, link_text: str = "") -> bool:
    """Return True only for tender-specific Coega document/download URLs.

    The rendered Coega page contains many navigation, social-media and generic
    site links. Earlier versions captured those into document_urls, which caused
    very long Excel URL warnings and polluted the source master.
    """
    u = (url or "").strip().lower()
    t = (link_text or "").strip().lower()
    if not u:
        return False
    if u.startswith("mailto:"):
        return False
    if "/site/tenders/tender-download" in u:
        return True
    # Allow tender-specific files under the Coega tender document area, but
    # exclude static guidance documents and generic procurement resources.
    if "/site/tenders/assets/documents/tenders/" in u:
        if any(x in t or x in u for x in [
            "procurement-handbook",
            "tender-evaluation",
            "scoring",
            "handbook",
            "evaluation",
        ]):
            return False
        return any(u.endswith(ext) for ext in [".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip"])
    return False


def parse_card(el: Any) -> Dict[str, Any]:
    block_text = clean_text(el.get_text(" ", strip=True))
    anchors = el.find_all("a", href=True)

    document_names = []
    document_urls = []
    for a in anchors:
        href = a.get("href", "")
        url = urljoin(SOURCE_URL, href)
        name = clean_text(a.get_text(" ", strip=True)) or os.path.basename(url.split("?")[0])
        if not url or url == SOURCE_URL:
            continue
        if not is_coega_tender_document_url(url, name):
            continue
        document_names.append(name)
        document_urls.append(url)

    # Deduplicate documents while preserving order.
    seen_urls = set()
    dedup_names = []
    dedup_urls = []
    for name, url in zip(document_names, document_urls):
        if url in seen_urls:
            continue
        seen_urls.add(url)
        dedup_names.append(name)
        dedup_urls.append(url)

    published_raw = extract_labelled_date(block_text, "Published Date")
    closing_raw = extract_labelled_date(block_text, "Closing Date")

    # If labels are split from values, fall back to first two dates in the block.
    if not published_raw or not closing_raw:
        dates = []
        for pattern in [
            r"\b\d{4}-\d{2}-\d{2}\b",
            r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",
            r"\b\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\s+\d{4}\b",
        ]:
            dates.extend(re.findall(pattern, block_text, flags=re.I))
        if not published_raw and dates:
            published_raw = dates[0]
        if not closing_raw and len(dates) >= 2:
            closing_raw = dates[1]

    tender_number = extract_tender_number(block_text)
    notice_type = infer_notice_type(block_text)
    description = build_description(block_text, tender_number)

    if not tender_number:
        # Fall back to a stable hash if no official tender number is visible.
        tender_number = "COEGA-" + hashlib.sha1(block_text.encode("utf-8", errors="ignore")).hexdigest()[:10].upper()

    return {
        "source": SOURCE,
        "tender_number": tender_number,
        "notice_type": notice_type,
        "description": description,
        "client": "Coega Development Corporation",
        "province": "Eastern Cape",
        "date_published": parse_date(published_raw),
        "closing_date": parse_date(closing_raw),
        "closing_datetime": clean_text(closing_raw),
        "document_names": "; ".join(dedup_names),
        "document_urls": "; ".join(dedup_urls),
        "document_count": len(dedup_urls),
        "source_url": SOURCE_URL,
        "cancellation_flag": is_cancellation(block_text),
    }


def parse_text_blocks(soup: BeautifulSoup) -> List[Dict[str, Any]]:
    rows = []
    elements = candidate_elements(soup)
    logging.info("Candidate Coega tender blocks found: %d", len(elements))

    for el in elements:
        row = parse_card(el)
        if row.get("description") or row.get("tender_number"):
            rows.append(row)

    return rows


def parse_coega_html(html: str) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "html.parser")

    rows = parse_from_tables(soup)
    if rows:
        logging.info("Parsed Coega rows from HTML tables: %d", len(rows))
        return rows

    rows = parse_text_blocks(soup)
    if rows:
        logging.info("Parsed Coega rows from tender blocks: %d", len(rows))
        return rows

    return []


def fetch_and_parse() -> List[Dict[str, Any]]:
    html = fetch_static_html()
    rows = parse_coega_html(html)

    if rows:
        logging.info("Static HTML parse succeeded: %d rows", len(rows))
        return rows

    logging.info("No Coega rows found in static HTML; attempting Selenium render.")
    try:
        rendered_html = fetch_rendered_html()
    except Exception as exc:
        # Coega is JavaScript-rendered and the Selenium browser can occasionally
        # hang or time out. Do not fail the full daily run because one dynamic
        # source is temporarily unavailable. Returning no newly parsed rows will
        # leave the existing Coega master/delta/update files in place and allow
        # consolidation to continue using the last successful source outputs.
        logging.exception("Coega Selenium fallback failed; continuing with no newly parsed rows: %s", exc)
        return []
    rows = parse_coega_html(rendered_html)
    logging.info("Rendered HTML parse rows: %d", len(rows))
    return rows


# -----------------------------------------------------------------------------
# Cleaning, scoring, hash and output helpers
# -----------------------------------------------------------------------------
def add_opportunity_key(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if out.empty:
        out["opportunity_key"] = pd.Series(dtype="object")
        return out

    def key_for_row(row: pd.Series) -> str:
        tender_number = clean_text(row.get("tender_number"))
        if tender_number:
            return "coega:" + normalise_key_text(tender_number).replace(" ", "-")
        base = "|".join(
            [
                clean_text(row.get("description")),
                clean_text(row.get("closing_datetime")),
                clean_text(row.get("document_urls")),
            ]
        )
        return "coega:" + hashlib.sha1(base.encode("utf-8", errors="ignore")).hexdigest()[:16]

    out["opportunity_key"] = out.apply(key_for_row, axis=1)
    return out


def add_priority_scoring(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if out.empty:
        for col in [
            "priority_notice_type_match",
            "matched_description_keywords",
            "description_keyword_match_count",
            "weight_score",
        ]:
            out[col] = pd.Series(dtype="object")
        return out

    def matched_keywords(description: Any) -> str:
        desc_norm = normalise_search_text(description)
        matches = []
        for label, phrases in DESCRIPTION_KEYWORD_RULES.items():
            if any(contains_phrase(desc_norm, phrase) for phrase in phrases):
                matches.append(label)
        return "; ".join(matches)

    out["matched_description_keywords"] = out["description"].apply(matched_keywords)
    out["description_keyword_match_count"] = out["matched_description_keywords"].apply(
        lambda x: 0 if not x else len(str(x).split("; "))
    )

    priority_notice_types = {"RFP", "RFQ", "RFI", "EOI", "Invitation to Tender", "Tender"}
    out["priority_notice_type_match"] = out["notice_type"].isin(priority_notice_types)

    out["weight_score"] = (
        out["description_keyword_match_count"].fillna(0).astype(int) * WEIGHT_DESCRIPTION_KEYWORD
        + out["priority_notice_type_match"].astype(int) * WEIGHT_NOTICE_TYPE_PRIORITY
        + WEIGHT_COEGA_MATCH
    )

    sort_cols = [c for c in ["weight_score", "closing_date", "date_published"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

    return out


def filter_current(df: pd.DataFrame, include_expired: bool, include_cancellations: bool) -> pd.DataFrame:
    out = df.copy()

    if not include_cancellations and "cancellation_flag" in out.columns:
        before = len(out)
        out = out[~out["cancellation_flag"].fillna(False).astype(bool)].copy()
        logging.info("Filtered cancellation/withdrawal notices: %d -> %d rows", before, len(out))

    if not include_expired and "closing_date" in out.columns:
        before = len(out)
        today = pd.Timestamp.today().normalize()
        out = out[pd.to_datetime(out["closing_date"], errors="coerce") >= today].copy()
        logging.info("Filtered current tenders by closing date: %d -> %d rows", before, len(out))

    return out


def compute_row_hash(df: pd.DataFrame, cols: Sequence[str]) -> pd.Series:
    if df.empty:
        return pd.Series([], dtype="object", index=df.index)

    use_cols = [c for c in cols if c in df.columns]
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

    out_df = df.copy() if df is not None else pd.DataFrame()
    if out_df.shape[1] == 0:
        out_df = pd.DataFrame(columns=DEFAULT_COLUMNS)

    with pd.ExcelWriter(
        tmp,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        out_df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        nrows, ncols = out_df.shape
        columns = [{"header": col} for col in out_df.columns]

        if nrows > 0 and ncols > 0:
            worksheet.add_table(0, 0, nrows, ncols - 1, {"name": table_name, "columns": columns})
        elif ncols > 0:
            worksheet.autofilter(0, 0, 0, ncols - 1)
        worksheet.freeze_panes(1, 0)

    os.replace(tmp, path)


def load_master(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name="Master", engine="openpyxl")
    except Exception as exc:
        logging.exception("Failed to read master file %s: %s", path, exc)
        return pd.DataFrame()


def append_and_dedupe_master(master: pd.DataFrame, incoming: pd.DataFrame) -> pd.DataFrame:
    if master.empty:
        out = incoming.copy()
        out["_source_priority"] = 1
    else:
        master_part = master.copy()
        incoming_part = incoming.copy()
        master_part["_source_priority"] = 0
        incoming_part["_source_priority"] = 1
        out = pd.concat([master_part, incoming_part], ignore_index=True, sort=False)

    if "opportunity_key" not in out.columns:
        return out.drop(columns=["_source_priority"], errors="ignore")

    out["opportunity_key"] = out["opportunity_key"].fillna("").astype(str).str.strip()
    out = out[out["opportunity_key"].ne("")]

    sort_cols = ["_source_priority"]
    ascending = [False]
    if "date_published" in out.columns:
        out["date_published"] = pd.to_datetime(out["date_published"], errors="coerce")
        sort_cols.append("date_published")
        ascending.append(False)

    out = out.sort_values(sort_cols, ascending=ascending, na_position="last", kind="mergesort")
    out = out.drop_duplicates(subset=["opportunity_key"], keep="first")
    out = out.drop(columns=["_source_priority"], errors="ignore")
    return out


def build_top_delta(df: pd.DataFrame, limit: int, min_score: int) -> pd.DataFrame:
    out = df.copy()
    if "weight_score" not in out.columns:
        out["weight_score"] = 0

    out = out[out["weight_score"].fillna(0).astype(int) >= int(min_score)].copy()

    sort_cols = [c for c in ["weight_score", "closing_date", "date_published"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

    if limit and int(limit) > 0:
        out = out.head(int(limit)).copy()

    return out


def ensure_order(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    preferred = [c for c in DEFAULT_COLUMNS if c in out.columns]
    remaining = [c for c in out.columns if c not in preferred]
    return out[preferred + remaining]


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
def main() -> int:
    parser = argparse.ArgumentParser(description="Scrape and clean Coega current tenders")
    parser.add_argument("--include-expired", action="store_true", help="include tenders whose closing date has passed")
    parser.add_argument("--include-cancellations", action="store_true", help="include cancellation/withdrawal notices")
    parser.add_argument("--top-delta-limit", type=int, default=TOP_DELTA_LIMIT_DEFAULT, help="max rows in top delta; 0 means all")
    parser.add_argument("--top-delta-min-score", type=int, default=TOP_DELTA_MIN_SCORE_DEFAULT, help="minimum score for top delta")
    parser.add_argument("--state-file", default=STATE_FILE_DEFAULT, help="path to JSON file storing seen opportunity keys")
    args = parser.parse_args()

    ensure_directories()
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
        logging.info("Parsed Coega notice rows: %d", len(rows))

        raw_df = pd.DataFrame(rows)
        if raw_df.empty:
            raw_df = pd.DataFrame(columns=[c for c in DEFAULT_COLUMNS if c not in ["row_hash", "run_timestamp", "change_type", "previous_row_hash"]])
        raw_df = add_opportunity_key(raw_df)
        raw_df = ensure_order(raw_df)

        raw_df.to_csv(RAW_CSV, index=False)
        logging.info("Saved raw CSV: %s", RAW_CSV)
        with pd.ExcelWriter(
            RAW_XLSX,
            engine="xlsxwriter",
            datetime_format="yyyy-mm-dd",
            engine_kwargs={"options": {"strings_to_urls": False}},
        ) as writer:
            raw_df.to_excel(writer, index=False, sheet_name="Raw_Data")
        logging.info("Saved raw Excel: %s", RAW_XLSX)

        mapped = raw_df.copy()
        for c in ["date_published", "closing_date"]:
            if c in mapped.columns:
                mapped[c] = pd.to_datetime(mapped[c], errors="coerce")

        mapped = filter_current(mapped, args.include_expired, args.include_cancellations)
        mapped = add_priority_scoring(mapped)
        mapped["row_hash"] = compute_row_hash(mapped, HASH_COLS)
        mapped["run_timestamp"] = run_ts
        mapped = ensure_order(mapped)

        master_df = load_master(MASTER_XLSX)
        if not master_df.empty:
            if not args.include_cancellations and "cancellation_flag" in master_df.columns:
                master_df = master_df[~master_df["cancellation_flag"].fillna(False).astype(bool)].copy()
            if not args.include_expired and "closing_date" in master_df.columns:
                today = pd.Timestamp.today().normalize()
                master_df["closing_date"] = pd.to_datetime(master_df["closing_date"], errors="coerce")
                master_df = master_df[master_df["closing_date"] >= today].copy()
            if "row_hash" not in master_df.columns:
                master_df["row_hash"] = compute_row_hash(master_df, HASH_COLS)

        state_file_missing = not os.path.exists(args.state_file)
        stored_hash_version = _read_text_file(HASH_VERSION_FILE)
        hash_version_changed = stored_hash_version != HASH_VERSION
        suppress_alerts = False

        if not master_df.empty and state_file_missing:
            logging.info(
                "No Coega seen_ids.json found. Bootstrapping state from existing master and suppressing delta/update alerts for this run."
            )
            suppress_alerts = True
        elif not master_df.empty and hash_version_changed:
            logging.info(
                "Coega hash/state version changed or was not recorded. Re-baselining and suppressing delta/update alerts for this run."
            )
            suppress_alerts = True

        if suppress_alerts:
            delta_df = _empty_change_frame(mapped, "NEW")
            updates_df = _empty_change_frame(mapped, "UPDATED")
        elif not master_df.empty and "opportunity_key" in master_df.columns:
            master_df["opportunity_key"] = master_df["opportunity_key"].fillna("").astype(str).str.strip()
            master_ids = set(master_df["opportunity_key"].tolist())
            master_hash_map = dict(
                zip(master_df["opportunity_key"].tolist(), master_df["row_hash"].fillna("").astype(str).tolist())
            )

            is_new = ~mapped["opportunity_key"].isin(master_ids)
            delta_df = mapped[is_new].copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""

            prev_hash = mapped["opportunity_key"].map(master_hash_map).fillna("")
            is_updated = (~is_new) & (mapped["row_hash"] != prev_hash)
            updates_df = mapped[is_updated].copy()
            updates_df["change_type"] = "UPDATED"
            updates_df["previous_row_hash"] = updates_df["opportunity_key"].map(master_hash_map).fillna("")
        else:
            delta_df = mapped.copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""
            updates_df = mapped.iloc[0:0].copy()
            updates_df["change_type"] = "UPDATED"
            updates_df["previous_row_hash"] = ""

        for name, df in [("delta", delta_df), ("updates", updates_df)]:
            sort_cols = [c for c in ["weight_score", "closing_date", "date_published"] if c in df.columns]
            if sort_cols:
                ascending = [False if c == "weight_score" else True for c in sort_cols]
                sorted_df = df.sort_values(sort_cols, ascending=ascending, na_position="last")
                if name == "delta":
                    delta_df = sorted_df
                else:
                    updates_df = sorted_df

        top_delta_df = build_top_delta(delta_df, args.top_delta_limit, args.top_delta_min_score)

        logging.info("Delta rows this run: %d", len(delta_df))
        logging.info("Updates rows this run: %d", len(updates_df))
        logging.info("Top delta rows this run: %d", len(top_delta_df))

        write_excel_with_table(ensure_order(delta_df), DELTA_XLSX, "Delta", "tbl_delta")
        logging.info("Saved delta Excel: %s | rows: %d", DELTA_XLSX, len(delta_df))

        write_excel_with_table(ensure_order(updates_df), UPDATES_XLSX, "Updates", "tbl_updates")
        logging.info("Saved updates Excel: %s | rows: %d", UPDATES_XLSX, len(updates_df))

        write_excel_with_table(ensure_order(top_delta_df), TOP_DELTA_XLSX, "Top_Delta", "tbl_top_delta")
        logging.info("Saved top delta Excel: %s | rows: %d", TOP_DELTA_XLSX, len(top_delta_df))

        updated_master = append_and_dedupe_master(master_df, mapped)
        updated_master = ensure_order(updated_master)
        write_excel_with_table(updated_master, MASTER_XLSX, "Master", "tbl_master")
        logging.info("Saved master Excel: %s | rows: %d", MASTER_XLSX, len(updated_master))

        if "opportunity_key" in updated_master.columns:
            _save_seen_ids(args.state_file, updated_master["opportunity_key"].tolist())
        _write_text_file(HASH_VERSION_FILE, HASH_VERSION)

        return 0

    except (requests.RequestException, RuntimeError, OSError, ValueError) as exc:
        logging.exception("FAILED: %s", exc)
        return 1
    finally:
        release_lock(LOCK_FILE)


if __name__ == "__main__":
    raise SystemExit(main())
