r"""COGTA public tenders scraper.

Scrapes public COGTA tender-related pages from the official national COGTA website:
    https://www.cogta.gov.za/index.php/cogta-tenders/
    https://www.cogta.gov.za/index.php/new-tender-page/
    https://www.cogta.gov.za/index.php/category/cogta/dcog-cogta/documents/tenders/

The COGTA site is WordPress-based and is not always a clean structured tender register.
This scraper therefore uses a defensive approach:
    - scrape visible HTML tables if present;
    - scrape tender/post links from tender pages;
    - optionally enrich from detail pages;
    - preserve rows with missing closing dates so the central consolidation layer can
      resolve dates from eTenders duplicates where possible or route unresolved items
      to manual-date review.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Set, Tuple
from urllib.parse import urljoin, urlparse
from io import StringIO

import pandas as pd
import requests
from requests.exceptions import RequestException, SSLError
import urllib3
from bs4 import BeautifulSoup, Tag

LEGACY_BASE_ROOT = r"C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\COGTA"


def _default_source_root() -> Path:
    """Resolve source root from script location when installed under 02_Sources/COGTA/00_Scripts."""
    script_path = Path(__file__).resolve()
    if script_path.parent.name == "00_Scripts":
        return script_path.parents[1]
    return Path(LEGACY_BASE_ROOT)


BASE_ROOT = str(_default_source_root())
SOURCE_NAME = "COGTA"
CLIENT_NAME = "Cooperative Governance and Traditional Affairs"
COUNTRY_NAME = "South Africa"
PROVINCE = "National"
REGION = "South Africa"
SECTOR = "Government / Municipal Support"
CATEGORY = "Services: Professional"

SOURCE_URLS = [
    "https://www.cogta.gov.za/index.php/cogta-tenders/",
    "https://www.cogta.gov.za/index.php/new-tender-page/",
    "https://www.cogta.gov.za/index.php/category/cogta/dcog-cogta/documents/tenders/",
]

RAW_DIR_NAME = "01_RawData"
OUTPUT_DIR_NAME = "02_Outputs"
LOG_DIR_NAME = "03_Logs"
STATE_DIR_NAME = "04_State"
METADATA_DIR_NAME = "05_Metadata"

RAW_CSV = "cogta_raw.csv"
RAW_XLSX = "cogta_raw_full.xlsx"
RAW_HTML = "cogta_raw_page.html"
MASTER_XLSX = "cogta_master.xlsx"
DELTA_XLSX = "cogta_delta.xlsx"
UPDATES_XLSX = "cogta_updates.xlsx"
TOP_DELTA_XLSX = "cogta_top_delta.xlsx"
STATE_FILE = "seen_ids.json"
HASH_VERSION_FILE = "hash_version.txt"
LOCK_FILE = "run.lock"
HASH_VERSION = "cogta_v3_restricted_link_extraction_2026_07"

REQUEST_TIMEOUT = 90
REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}
DOC_EXTS = (".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".rar")

OUTPUT_COLUMNS = [
    "source", "opportunity_key", "tender_number", "reference_number", "notice_type", "title", "description",
    "client", "sector", "category", "country_name", "province", "region", "date_published", "issue_date",
    "briefing_date", "closing_date", "closing_datetime", "status", "is_current", "is_archived", "is_cancelled",
    "document_names", "document_urls", "document_count", "contact_person", "contact_email", "contact_telephone",
    "submission_method", "briefing_session_details", "conditions", "raw_category", "raw_status", "source_url",
    "run_timestamp", "row_hash", "change_type", "previous_row_hash", "weight_score", "matched_description_keywords",
]
HASH_COLS = ["reference_number", "title", "description", "closing_datetime", "source_url", "status", "document_names", "document_count"]

KEYWORDS = [
    "cogta", "cooperative governance", "traditional affairs", "municipal", "municipalities", "local government",
    "professional services", "professional service provider", "psp", "panel", "consultant", "consulting", "consultancy",
    "technical assistance", "technical advisory", "transaction advisory", "infrastructure", "project management",
    "programme management", "pmo", "implementation", "business case", "spatial", "gis", "dashboard", "planning",
    "development model", "district development model", "ddm", "disaster management", "engineering", "built environment",
]

NON_CURRENT_PATTERNS = [
    r"\bawarded\b", r"\bbid\s+award", r"\baward\s+notice", r"\bbids?\s+received\b", r"\bproof\s+of\s+bids",
    r"\bbid\s+box\s+opening", r"\bopening\s+register", r"\bcancelled\b", r"\bcanceled\b", r"\bdeviation\b",
]


def get_source_dirs(base_root: Path) -> Tuple[Path, Path, Path, Path, Path]:
    """Return standard source directories for the 02_TenderScrapers source layout."""
    raw_dir = base_root / RAW_DIR_NAME
    output_dir = base_root / OUTPUT_DIR_NAME
    log_dir = base_root / LOG_DIR_NAME
    state_dir = base_root / STATE_DIR_NAME
    metadata_dir = base_root / METADATA_DIR_NAME
    return raw_dir, output_dir, log_dir, state_dir, metadata_dir


def ensure_dirs(base_root: Path) -> Tuple[Path, Path, Path, Path, Path]:
    raw_dir, output_dir, log_dir, state_dir, metadata_dir = get_source_dirs(base_root)
    for directory in [raw_dir, output_dir, log_dir, state_dir, metadata_dir]:
        directory.mkdir(parents=True, exist_ok=True)
    return raw_dir, output_dir, log_dir, state_dir, metadata_dir


def setup_logging(log_dir: Path) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "cogta_run.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
        force=True,
    )


def acquire_lock(lock_path: Path) -> None:
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    if lock_path.exists():
        raise RuntimeError(f"Lock file exists; another run may be active: {lock_path}")
    lock_path.write_text(f"pid={os.getpid()}\\n", encoding="utf-8")


def release_lock(lock_path: Path) -> None:
    try:
        if lock_path.exists():
            lock_path.unlink()
    except OSError:
        logging.exception("Failed to remove lock file: %s", lock_path)


def load_seen_ids(path: Path) -> Set[str]:
    try:
        if not path.exists():
            return set()
        data = json.loads(path.read_text(encoding="utf-8"))
        return {clean_text(x).casefold() for x in (data or []) if clean_text(x)}
    except (OSError, ValueError, TypeError) as exc:
        logging.exception("Failed to load seen_ids file %s: %s", path, exc)
        return set()


def save_seen_ids(path: Path, ids: Iterable[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    values = sorted({clean_text(x).casefold() for x in ids if clean_text(x)})
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(values, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def read_text_file(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8").strip() if path.exists() else ""
    except OSError:
        return ""


def write_text_file(path: Path, value: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(value, encoding="utf-8")


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if isinstance(value, float) and pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value)
    text = text.replace("\xa0", " ").replace("\u2013", "-").replace("\u2014", "-").replace("\u2019", "'")
    return re.sub(r"\s+", " ", text).strip()


def normalise_id(value: Any) -> str:
    text = clean_text(value).lower()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")


def save_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", errors="ignore")


def fetch_html(url: str) -> str:
    session = requests.Session()
    session.headers.update(REQUEST_HEADERS)
    try:
        r = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        r.raise_for_status()
        return r.text
    except SSLError:
        logging.warning("SSL verification failed for %s. Retrying without verification.", url)
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        r = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True, verify=False)
        r.raise_for_status()
        return r.text


def parse_datetime(value: Any) -> pd.Timestamp:
    text = clean_text(value)
    if not text or text.lower() in {"n/a", "na", "none", "-", "nil"}:
        return pd.NaT
    text = re.sub(r"(?i)(closing|closing date|closing time|bid closing|deadline|date)\s*: ?", "", text)
    text = re.sub(r"(?i)\bat\b", " ", text)
    text = re.sub(r"(?i)(\d{1,2})h(\d{2})\s*(am|pm)?", lambda m: f"{m.group(1)}:{m.group(2)} {m.group(3) or ''}", text)
    text = re.sub(r"(?i)(\d{1,2})h\s*(am|pm)?", lambda m: f"{m.group(1)}:00 {m.group(2) or ''}", text)
    text = clean_text(text)
    if re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}", text):
        return pd.to_datetime(text, errors="coerce", dayfirst=False)
    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def extract_date_text(text: str, label_hint: str = "") -> str:
    text = clean_text(text)
    # Prefer date close to closing/deadline labels.
    label_pat = r"(?i)(closing\s+date(?:\s*&\s*time)?|closing\s+time|bid\s+closing|deadline)\s*[:\-]?\s*"
    date_pat = r"(\d{1,2}\s+[A-Za-z]+\s+\d{4}(?:\s*(?:at|,)\s*\d{1,2}:\d{2}(?::\d{2})?\s*(?:am|pm)?)?|\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?|\d{1,2}[-/]\d{1,2}[-/]\d{2,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)"
    m = re.search(label_pat + date_pat, text, flags=re.I)
    if m:
        return clean_text(m.group(2))
    for pat in [
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\s*(?:at|,)\s*\d{1,2}:\d{2}(?::\d{2})?\s*(?:am|pm)?\b",
        r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?\b",
        r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\s+\d{1,2}:\d{2}(?::\d{2})?\b",
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\b",
        r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b",
        r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b",
    ]:
        m = re.search(pat, text, flags=re.I)
        if m:
            return clean_text(m.group(0))
    return ""


def extract_reference(*parts: str) -> str:
    text = " ".join(clean_text(p) for p in parts if p)
    patterns = [
        r"\bDCOG[-/A-Z0-9 ]{2,30}\d{4}[/\-]\d{2,4}\b",
        r"\bCOGTA\s*(?:RFP|RFQ|T|BID)?\s*\d{1,4}\s*/\s*\d{4}(?:\s*/\s*\d{2,4})?\b",
        r"\bCOGTA\s*\d{1,4}\s*/\s*\d{4}\b",
        r"\bDTA\s*/\s*[A-Z]+\s*/\s*[A-Z0-9]+\s*/\s*\d{4}\s*/\s*\d{2,4}\b",
        r"\bFQ\s*\d{1,4}\s*/\s*\d{4}\b",
        r"\bRFQ\s*\d{1,4}\s*/\s*\d{4}\b",
        r"\bRFP\s*\d{1,4}\s*/\s*\d{4}\b",
        r"\b[A-Z]{2,}\s*\d{1,4}\s*/\s*\d{4}\s*/\s*\d{2,4}\b",
        r"\b[A-Z]{2,}[-/]\d{1,4}[-/]\d{2,4}\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, flags=re.I)
        if m:
            ref = re.sub(r"\s+", "", m.group(0)).upper()
            ref = re.sub(r"-+", "-", ref)
            return ref.strip("-/")
    return ""


def infer_notice_type(title: str, desc: str) -> str:
    text = f"{title} {desc}".lower()
    if "expression of interest" in text or re.search(r"\beoi\b", text):
        return "EOI"
    if "rfp" in text or "request for proposal" in text:
        return "RFP"
    if "quotation" in text or re.search(r"\brfq\b", text):
        return "RFQ"
    if "rfi" in text or "request for information" in text:
        return "RFI"
    return "Tender"


def is_non_current_text(text: str) -> bool:
    text = clean_text(text).lower()
    return any(re.search(pat, text, flags=re.I) for pat in NON_CURRENT_PATTERNS)


def contains_procurement_signal(text: str, href: str = "") -> bool:
    """Return True only when the visible text/URL actually looks procurement-related.

    Earlier versions allowed generic words such as COGTA/DCOG. Because the site is
    WordPress-based, that caused normal news/category links to be pulled in as
    tender rows. This function deliberately requires procurement language or a
    direct document/download signal.
    """
    joined = f"{text} {href}".lower()
    return any(signal in joined for signal in [
        "tender",
        "bid",
        "rfp",
        "rfq",
        "rfi",
        "quotation",
        "request for proposal",
        "request for quotation",
        "terms of reference",
        "tor",
        "appointment of",
        "service provider",
        "proof of bids",
        "bid award",
        "opening register",
        "download",
    ])


def is_archive_or_category_link(url: str, text: str = "") -> bool:
    """Filter WordPress navigation/category/archive links that are not tender records."""
    low_url = clean_text(url).lower().rstrip("/")
    low_text = clean_text(text).lower()

    if not low_url:
        return True
    if low_url.rstrip("/") in [u.rstrip("/").lower() for u in SOURCE_URLS]:
        return True
    if re.search(r"/category/", low_url):
        return True
    if re.search(r"/tag/|/author/|/page/\d+/?$", low_url):
        return True
    if re.fullmatch(r"20\d{2}", low_text):
        return True
    if low_text in {"tenders", "documents", "cogta", "dcog", "home", "read more", "more"}:
        return True
    return False


def is_old_wordpress_post(url: str, closing_date_text: str = "") -> bool:
    """Treat very old WordPress tender posts with no closing date as archived/non-current."""
    if clean_text(closing_date_text):
        return False
    m = re.search(r"/((?:20)\d{2})/", clean_text(url))
    if not m:
        return False
    try:
        year = int(m.group(1))
    except ValueError:
        return False
    current_year = datetime.now().year
    return year < current_year - 1


def is_doc_link(href: str, text: str = "") -> bool:
    joined = f"{href} {text}".lower()
    if not href or href.startswith("#") or href.lower().startswith("mailto:"):
        return False
    if any(ext in joined for ext in DOC_EXTS):
        return True
    return any(p in joined for p in ["download", "tender", "bid", "rfp", "rfq", "terms of reference", "tor"])


def extract_document_links(soup: BeautifulSoup, page_url: str) -> Tuple[str, str, int]:
    names: List[str] = []
    urls: List[str] = []
    for a in soup.find_all("a", href=True):
        text = clean_text(a.get_text(" "))
        href = clean_text(a.get("href"))
        if not is_doc_link(href, text):
            continue
        url = urljoin(page_url, href)
        # Avoid collecting whole tender directory/category links as document URLs.
        if url.rstrip("/") in [u.rstrip("/") for u in SOURCE_URLS]:
            continue
        if url not in urls:
            urls.append(url)
            names.append(text or Path(urlparse(url).path).name or "Download")
    return "; ".join(names), "; ".join(urls), len(urls)


def table_rows_from_html(html: str, page_url: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    try:
        tables = pd.read_html(StringIO(html))
    except Exception:
        return rows
    for table in tables:
        if table.empty:
            continue
        table.columns = [clean_text(c).lower() for c in table.columns]
        cols_joined = " ".join(table.columns)
        if not any(key in cols_joined for key in ["tender", "bid", "rfq", "rfp", "closing"]):
            continue
        for _, r in table.iterrows():
            vals = {c: clean_text(r.get(c, "")) for c in table.columns}
            title = vals.get("tender title") or vals.get("description") or vals.get("bid description") or vals.get("title") or ""
            ref = vals.get("tender number") or vals.get("bid number") or vals.get("reference") or vals.get("reference number") or extract_reference(" ".join(vals.values()))
            closing = vals.get("closing date") or vals.get("closing date & time") or vals.get("closing") or extract_date_text(" ".join(vals.values()))
            briefing = vals.get("briefing date") or vals.get("briefing") or ""
            if not title and not ref:
                continue
            rows.append({
                "reference_number": ref,
                "title": title or ref,
                "description": title or ref,
                "closing_date_text": closing,
                "briefing_date_text": briefing,
                "document_names": "",
                "document_urls": "",
                "document_count": 0,
                "source_url": page_url,
                "raw_category": "HTML table",
                "raw_status": "Current/Unknown" if not is_non_current_text(" ".join(vals.values())) else "Non-current",
            })
    return rows


def likely_tender_link(text: str, href: str) -> bool:
    """Identify actual tender/procurement links, not generic COGTA news/category links."""
    joined = f"{text} {href}".lower()
    if not text and not href:
        return False
    if any(skip in joined for skip in [
        "facebook", "twitter", "linkedin", "instagram", "youtube", "privacy", "terms", "vacancies",
        "minister", "deputy minister", "budget vote", "summit", "white paper", "state of the province",
    ]):
        # Do not let normal news posts through merely because the URL contains /cogta/.
        if not contains_procurement_signal(text, href):
            return False
    url = urljoin(SOURCE_URLS[0], href)
    if is_archive_or_category_link(url, text):
        return False
    return contains_procurement_signal(text, href)


def extract_listing_links(soup: BeautifulSoup, page_url: str, detail_limit: int) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    seen: set[str] = set()
    content = soup.find("main") or soup.find("article") or soup.body or soup

    for a in content.find_all("a", href=True):
        text = clean_text(a.get_text(" "))
        href = clean_text(a.get("href"))
        url = urljoin(page_url, href)

        if not likely_tender_link(text, href):
            continue
        if is_archive_or_category_link(url, text):
            continue
        if url in seen:
            continue
        seen.add(url)

        title = text or clean_text(urlparse(url).path.strip("/").split("/")[-1].replace("-", " "))
        if not contains_procurement_signal(title, url):
            continue

        ref = extract_reference(title, url)
        date_text = extract_date_text(title)
        raw_status = "Non-current" if is_non_current_text(title) or is_old_wordpress_post(url, date_text) else "Current/Unknown"

        rows.append({
            "reference_number": ref,
            "title": title,
            "description": title,
            "closing_date_text": date_text,
            "briefing_date_text": "",
            "document_names": "",
            "document_urls": "",
            "document_count": 0,
            "source_url": url,
            "raw_category": "Tender/link listing",
            "raw_status": raw_status,
        })

        if len(rows) >= detail_limit:
            break

    return rows


def enrich_from_detail(row: Dict[str, Any]) -> Dict[str, Any]:
    url = clean_text(row.get("source_url"))
    title_seed = clean_text(row.get("title"))

    if not url or not url.startswith("http"):
        return row
    if is_archive_or_category_link(url, title_seed):
        return row
    if not contains_procurement_signal(title_seed, url):
        return row

    # Do not detail-fetch direct documents. Keep the document URL directly.
    if any(url.lower().split("?")[0].endswith(ext) for ext in DOC_EXTS):
        row["document_urls"] = url
        row["document_names"] = title_seed or Path(urlparse(url).path).name
        row["document_count"] = 1
        return row

    try:
        html = fetch_html(url)
    except RequestException as exc:
        logging.warning("Failed to fetch COGTA detail page %s: %s", url, exc)
        return row

    soup = BeautifulSoup(html, "html.parser")
    title_node = soup.find(["h1", "h2"])
    title = clean_text(title_node.get_text(" ")) if title_node else title_seed

    main = soup.find("main") or soup.find("article") or soup.body or soup
    text = clean_text(main.get_text(" "))

    # If the detail page turns out to be a normal news/media post, do not promote it
    # into the master as an open tender row.
    if not contains_procurement_signal(title, text):
        row["raw_status"] = "Non-current"
        return row

    ref = clean_text(row.get("reference_number")) or extract_reference(title, text, url)
    date_text = clean_text(row.get("closing_date_text")) or extract_date_text(text)
    briefing_text = clean_text(row.get("briefing_date_text"))
    doc_names, doc_urls, doc_count = extract_document_links(soup, url)

    raw_status = clean_text(row.get("raw_status")) or "Current/Unknown"
    if is_non_current_text(f"{title} {text}") or is_old_wordpress_post(url, date_text):
        raw_status = "Non-current"

    row.update({
        "reference_number": ref,
        "title": title or title_seed,
        "description": text[:1200] if text else clean_text(row.get("description")),
        "closing_date_text": date_text,
        "briefing_date_text": briefing_text,
        "document_names": doc_names or clean_text(row.get("document_names")),
        "document_urls": doc_urls or clean_text(row.get("document_urls")),
        "document_count": doc_count or int(row.get("document_count") or 0),
        "raw_status": raw_status,
    })
    return row


def build_rows(raw_dir: Path, detail_limit: int) -> List[Dict[str, Any]]:
    all_rows: List[Dict[str, Any]] = []
    combined_html: List[str] = []
    for url in SOURCE_URLS:
        logging.info("Fetching COGTA tender page: %s", url)
        try:
            html = fetch_html(url)
        except RequestException as exc:
            logging.warning("Failed to fetch COGTA page %s: %s", url, exc)
            continue
        combined_html.append(f"\n<!-- SOURCE: {url} -->\n{html}")
        soup = BeautifulSoup(html, "html.parser")
        table_rows = table_rows_from_html(html, url)
        link_rows = extract_listing_links(soup, url, detail_limit=detail_limit)
        logging.info("Parsed COGTA table rows from %s: %s", url, len(table_rows))
        logging.info("Parsed COGTA link rows from %s: %s", url, len(link_rows))
        all_rows.extend(table_rows)
        all_rows.extend(link_rows)

    # Enrich likely detail pages, but cap to avoid very long runs on the slow COGTA site.
    enriched: List[Dict[str, Any]] = []
    for idx, row in enumerate(all_rows[:detail_limit]):
        enriched.append(enrich_from_detail(row))
    enriched.extend(all_rows[detail_limit:])
    save_text(raw_dir / RAW_HTML, "\n".join(combined_html))
    return enriched


def compute_row_hash(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    if df.empty:
        return pd.Series(dtype="object")
    def _hash(row: pd.Series) -> str:
        vals = [clean_text(row.get(c, "")).lower() for c in cols]
        return hashlib.sha256("||".join(vals).encode("utf-8", errors="ignore")).hexdigest()
    return df.apply(_hash, axis=1)


def score_row(row: pd.Series) -> Tuple[int, str]:
    text = f"{row.get('client','')} {row.get('title','')} {row.get('description','')} {row.get('reference_number','')}".lower()
    matches = [kw for kw in KEYWORDS if kw.lower() in text]
    score = min(45, 10 + 4 * len(matches)) if matches else 10
    if row.get("document_count", 0):
        score += 1
    return score, "; ".join(matches)


def prepare_dataframe(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    now = datetime.now()
    prepared: List[Dict[str, Any]] = []
    seen_keys: set[str] = set()
    for row in rows:
        title = clean_text(row.get("title"))
        desc = clean_text(row.get("description")) or title
        source_url = clean_text(row.get("source_url"))
        ref = clean_text(row.get("reference_number")) or extract_reference(title, desc, source_url)
        doc_urls = clean_text(row.get("document_urls"))

        if not contains_procurement_signal(title, desc) and not ref and not doc_urls:
            continue
        if is_archive_or_category_link(source_url, title):
            continue

        key_base = ref or title or source_url or doc_urls
        if not key_base:
            continue
        opportunity_key = f"cogta:{normalise_id(key_base)}"
        if opportunity_key in seen_keys:
            continue
        seen_keys.add(opportunity_key)
        date_text = clean_text(row.get("closing_date_text")) or extract_date_text(f"{title} {desc}")
        dt = parse_datetime(date_text) if date_text else pd.NaT
        briefing_text = clean_text(row.get("briefing_date_text"))
        bdt = parse_datetime(briefing_text) if briefing_text else pd.NaT
        non_current = is_non_current_text(f"{title} {desc} {row.get('raw_status','')}") or is_old_wordpress_post(source_url, date_text)
        is_current = bool((pd.isna(dt) or dt >= pd.Timestamp.now().normalize()) and not non_current)
        raw_status = "Non-current" if non_current else "Open/Unknown Date" if pd.isna(dt) else "Open" if is_current else "Expired"
        record = {
            "source": SOURCE_NAME,
            "opportunity_key": opportunity_key,
            "tender_number": ref,
            "reference_number": ref,
            "notice_type": infer_notice_type(title, desc),
            "title": title,
            "description": desc,
            "client": CLIENT_NAME,
            "sector": SECTOR,
            "category": CATEGORY,
            "country_name": COUNTRY_NAME,
            "province": PROVINCE,
            "region": REGION,
            "date_published": "",
            "issue_date": "",
            "briefing_date": bdt.to_pydatetime().isoformat(sep=" ") if not pd.isna(bdt) else "",
            "closing_date": dt.date().isoformat() if not pd.isna(dt) else "",
            "closing_datetime": dt.to_pydatetime().isoformat(sep=" ") if not pd.isna(dt) else "",
            "status": raw_status,
            "is_current": is_current,
            "is_archived": non_current,
            "is_cancelled": non_current,
            "document_names": clean_text(row.get("document_names")),
            "document_urls": doc_urls,
            "document_count": int(row.get("document_count") or 0),
            "contact_person": "",
            "contact_email": "",
            "contact_telephone": "",
            "submission_method": "See COGTA tender page / tender document / eTenders where applicable",
            "briefing_session_details": briefing_text,
            "conditions": "COGTA page is not a clean structured tender register; missing dates may require eTenders duplicate resolution.",
            "raw_category": clean_text(row.get("raw_category")),
            "raw_status": raw_status,
            "source_url": source_url or SOURCE_URLS[0],
            "run_timestamp": now.isoformat(sep=" ", timespec="seconds"),
            "change_type": "",
            "previous_row_hash": "",
        }
        prepared.append(record)
    df = pd.DataFrame(prepared)
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    if not df.empty:
        scores = df.apply(score_row, axis=1, result_type="expand")
        df["weight_score"] = scores[0]
        df["matched_description_keywords"] = scores[1]
        df["row_hash"] = compute_row_hash(df, HASH_COLS)
    return df[OUTPUT_COLUMNS]


def save_excel(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        sheet = "Sheet1"
        df.to_excel(writer, index=False, sheet_name=sheet)
        workbook = writer.book
        worksheet = writer.sheets[sheet]
        header_fmt = workbook.add_format({"bold": True, "text_wrap": True, "valign": "top", "border": 1})
        for idx, col in enumerate(df.columns):
            worksheet.write(0, idx, col, header_fmt)
            if col in df.columns:
                max_len = min(max([len(str(col))] + [len(str(x)) for x in df[col].head(200).fillna("")]), 80)
            else:
                max_len = len(str(col))
            worksheet.set_column(idx, idx, max_len + 2)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max(len(df), 1), max(len(df.columns) - 1, 0))


def load_existing_master(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    try:
        return pd.read_excel(path, dtype=object)
    except Exception:
        logging.exception("Could not read existing COGTA master: %s", path)
        return pd.DataFrame(columns=OUTPUT_COLUMNS)


def classify_delta_updates(
    new_df: pd.DataFrame,
    old_df: pd.DataFrame,
    seen_ids: Set[str],
    suppress_alerts: bool,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Classify rows using seen_ids as the baseline and row_hash only for true updates."""
    if new_df.empty:
        empty = pd.DataFrame(columns=OUTPUT_COLUMNS)
        return empty.copy(), empty.copy(), empty.copy()

    old_by_key: Dict[str, str] = {}
    if not old_df.empty and "opportunity_key" in old_df.columns:
        for _, row in old_df.iterrows():
            key = clean_text(row.get("opportunity_key")).casefold()
            if key:
                old_by_key[key] = clean_text(row.get("row_hash"))

    rows = []
    delta_rows = []
    update_rows = []

    for _, row in new_df.iterrows():
        record = row.copy()
        key = clean_text(record.get("opportunity_key")).casefold()
        new_hash = clean_text(record.get("row_hash"))
        old_hash = old_by_key.get(key, "")

        if suppress_alerts:
            record["change_type"] = ""
            record["previous_row_hash"] = old_hash
        elif key and key not in seen_ids:
            record["change_type"] = "NEW"
            record["previous_row_hash"] = ""
            delta_rows.append(record.copy())
        elif old_hash and old_hash != new_hash:
            record["change_type"] = "UPDATED"
            record["previous_row_hash"] = old_hash
            update_rows.append(record.copy())
        else:
            record["change_type"] = "UNCHANGED"
            record["previous_row_hash"] = old_hash

        rows.append(record.copy())

    full = pd.DataFrame(rows) if rows else pd.DataFrame(columns=OUTPUT_COLUMNS)
    delta = pd.DataFrame(delta_rows) if delta_rows else pd.DataFrame(columns=OUTPUT_COLUMNS)
    updates = pd.DataFrame(update_rows) if update_rows else pd.DataFrame(columns=OUTPUT_COLUMNS)

    return full[OUTPUT_COLUMNS], delta[OUTPUT_COLUMNS], updates[OUTPUT_COLUMNS]

def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Scrape COGTA public tender pages.")
    parser.add_argument("--base-root", default=BASE_ROOT)
    parser.add_argument("--detail-limit", type=int, default=15)
    parser.add_argument("--top-limit", type=int, default=20)
    args = parser.parse_args(argv)

    base_root = Path(args.base_root)
    raw_dir, output_dir, log_dir, state_dir, metadata_dir = ensure_dirs(base_root)
    setup_logging(log_dir)

    tender_root = base_root.parents[1] if base_root.parent.name == "02_Sources" else base_root.parent
    state_file = state_dir / STATE_FILE
    hash_version_file = state_dir / HASH_VERSION_FILE
    lock_file = state_dir / LOCK_FILE

    logging.info("Using tender scrapers root: %s", tender_root)
    logging.info("Using source root: %s", base_root)
    logging.info("Using raw directory: %s", raw_dir)
    logging.info("Using output directory: %s", output_dir)
    logging.info("Using log directory: %s", log_dir)
    logging.info("Using state directory: %s", state_dir)

    try:
        acquire_lock(lock_file)

        try:
            rows = build_rows(raw_dir, detail_limit=args.detail_limit)
            raw_df = prepare_dataframe(rows)
            logging.info("Parsed COGTA raw rows: %s", len(raw_df))
        except Exception as exc:
            logging.exception("FAILED during COGTA fetch/parse. Creating empty raw outputs and preserving existing master where possible: %s", exc)
            raw_df = pd.DataFrame(columns=OUTPUT_COLUMNS)

        raw_df.to_csv(raw_dir / RAW_CSV, index=False, encoding="utf-8-sig")
        save_excel(raw_df, raw_dir / RAW_XLSX)

        master_candidate = raw_df.copy()
        if not master_candidate.empty and "is_cancelled" in master_candidate.columns:
            master_candidate = master_candidate[
                ~master_candidate["is_cancelled"].astype(str).str.lower().isin(["true", "1", "yes"])
            ].copy()

        old_master = load_existing_master(output_dir / MASTER_XLSX)

        seen_ids = load_seen_ids(state_file)
        recorded_hash_version = read_text_file(hash_version_file)
        suppress_alerts = False

        if not seen_ids and not old_master.empty:
            logging.info("No COGTA seen_ids.json found. Bootstrapping state from existing master and suppressing delta/update alerts for this run.")
            if "opportunity_key" in old_master.columns:
                seen_ids = {clean_text(x).casefold() for x in old_master["opportunity_key"].tolist() if clean_text(x)}
            suppress_alerts = True

        if recorded_hash_version != HASH_VERSION:
            logging.info("COGTA hash/state version changed or was not recorded. Re-baselining and suppressing delta/update alerts for this run.")
            suppress_alerts = True

        if master_candidate.empty and not old_master.empty:
            logging.warning("COGTA produced no master candidates. Preserving previous master and writing empty delta/update outputs.")
            master_df = old_master.copy()
            delta_df = pd.DataFrame(columns=OUTPUT_COLUMNS)
            updates_df = pd.DataFrame(columns=OUTPUT_COLUMNS)
        else:
            master_df, delta_df, updates_df = classify_delta_updates(
                master_candidate,
                old_master,
                seen_ids=seen_ids,
                suppress_alerts=suppress_alerts,
            )

        top_delta_df = (
            delta_df.sort_values("weight_score", ascending=False).head(args.top_limit)
            if not delta_df.empty and "weight_score" in delta_df.columns
            else pd.DataFrame(columns=OUTPUT_COLUMNS)
        )

        save_excel(master_df, output_dir / MASTER_XLSX)
        save_excel(delta_df, output_dir / DELTA_XLSX)
        save_excel(updates_df, output_dir / UPDATES_XLSX)
        save_excel(top_delta_df, output_dir / TOP_DELTA_XLSX)

        if not master_df.empty and "opportunity_key" in master_df.columns:
            save_seen_ids(state_file, master_df["opportunity_key"].tolist())
        write_text_file(hash_version_file, HASH_VERSION)

        logging.info("Saved master Excel: %s | rows: %s", output_dir / MASTER_XLSX, len(master_df))
        logging.info("Saved delta Excel: %s | rows: %s", output_dir / DELTA_XLSX, len(delta_df))
        logging.info("Saved updates Excel: %s | rows: %s", output_dir / UPDATES_XLSX, len(updates_df))
        logging.info("Saved top_delta Excel: %s | rows: %s", output_dir / TOP_DELTA_XLSX, len(top_delta_df))
        return 0

    except Exception as exc:
        logging.exception("FAILED: %s", exc)
        return 1
    finally:
        release_lock(lock_file)

if __name__ == "__main__":
    raise SystemExit(main())
