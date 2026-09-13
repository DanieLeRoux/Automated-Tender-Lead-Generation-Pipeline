r"""City of Ekurhuleni open tenders scraper.

Scrapes open tenders from:
    https://www.ekurhuleni.gov.za/for-my-business/tenders/open-tenders/

The listing page exposes tender references and short descriptions. Each tender
has a detail page with richer fields such as bid number, full description,
closing date/time, contact details, submission instructions and document links.

New architecture:
    Place this script in:
        02_TenderScrapers\02_Sources\CityEkurhuleni\00_Scripts

    It writes raw scrape evidence to:
        02_TenderScrapers\02_Sources\CityEkurhuleni\01_RawData

    It writes operational outputs to:
        02_TenderScrapers\02_Sources\CityEkurhuleni\02_Outputs
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
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from requests.exceptions import RequestException, SSLError
import urllib3
from bs4 import BeautifulSoup

SCRIPT_PATH = Path(__file__).resolve()
# Expected location:
#   02_TenderScrapers/02_Sources/CityEkurhuleni/00_Scripts/city_ekurhuleni_tender_pull_and_clean.py
SOURCE_ROOT = SCRIPT_PATH.parents[1]
TENDER_ROOT = SCRIPT_PATH.parents[3] if len(SCRIPT_PATH.parents) >= 4 else SOURCE_ROOT.parents[1]
BASE_ROOT = SOURCE_ROOT

SOURCE_URL = "https://www.ekurhuleni.gov.za/for-my-business/tenders/open-tenders/"
SOURCE_NAME = "City of Ekurhuleni"
CLIENT_NAME = "City of Ekurhuleni"

RAW_DIR = "01_RawData"
OUTPUT_DIR = "02_Outputs"
LOG_DIR = "03_Logs"
STATE_DIR = "04_State"
METADATA_DIR = "05_Metadata"

RAW_CSV = "city_ekurhuleni_raw.csv"
RAW_XLSX = "city_ekurhuleni_raw_full.xlsx"
RAW_HTML = "city_ekurhuleni_raw_page.html"
MASTER_XLSX = "city_ekurhuleni_master.xlsx"
DELTA_XLSX = "city_ekurhuleni_delta.xlsx"
UPDATES_XLSX = "city_ekurhuleni_updates.xlsx"
TOP_DELTA_XLSX = "city_ekurhuleni_top_delta.xlsx"
LOG_FILE = "city_ekurhuleni_run.log"
STATE_FILE = "seen_ids.json"
LOCK_FILE = "run.lock"

REQUEST_TIMEOUT = 90
MAX_LIST_PAGES = 8
REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

OUTPUT_COLUMNS = [
    "source", "opportunity_key", "tender_number", "reference_number", "notice_type", "title", "description",
    "client", "sector", "category", "country_name", "province", "region", "date_published", "issue_date",
    "briefing_date", "closing_date", "closing_datetime", "status", "is_current", "is_archived", "is_cancelled",
    "document_names", "document_urls", "document_count", "contact_person", "contact_email", "contact_telephone",
    "submission_method", "briefing_session_details", "conditions", "raw_category", "raw_status", "source_url",
    "run_timestamp", "row_hash", "change_type", "previous_row_hash", "weight_score", "matched_description_keywords",
]
HASH_COLS = ["reference_number", "description", "closing_datetime", "document_urls", "source_url", "status"]
DOC_EXTS = (".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".rar")
KEYWORDS = [
    "city of ekurhuleni", "ekurhuleni", "professional services", "professional service provider", "professional sectoral panel",
    "panel", "consultant", "consulting", "consultancy", "engineering", "engineering services", "civil engineering",
    "electrical engineering", "mechanical engineering", "structural engineering", "project management", "programme management",
    "construction monitoring", "contract administration", "technical advisory", "advisory", "feasibility", "design",
    "infrastructure", "water", "water and sanitation", "roads", "stormwater", "transport planning", "property valuation",
    "valuers", "urban planning", "spatial planning", "development planning", "built environment",
]


def setup_logging(base_root: Path) -> None:
    """Create log directory and configure logging for this scraper."""
    (base_root / LOG_DIR).mkdir(parents=True, exist_ok=True)
    log_file = base_root / LOG_DIR / LOG_FILE
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
        force=True,
    )


def ensure_dirs(base_root: Path) -> None:
    """Ensure the new source-level architecture exists."""
    for name in [RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR]:
        (base_root / name).mkdir(parents=True, exist_ok=True)


def acquire_lock(lock_path: Path) -> None:
    """Simple lockfile to prevent overlapping runs."""
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    if lock_path.exists():
        raise RuntimeError(f"Lock file exists (another run may be active): {lock_path}")
    lock_path.write_text(f"pid={os.getpid()}\n", encoding="utf-8")


def release_lock(lock_path: Path) -> None:
    try:
        if lock_path.exists():
            lock_path.unlink()
    except OSError:
        logging.exception("Failed to remove lock file: %s", lock_path)


def save_seen_ids(state_path: Path, keys: Iterable[Any]) -> None:
    """Persist currently known opportunity keys for audit/debugging."""
    try:
        state_path.parent.mkdir(parents=True, exist_ok=True)
        values = sorted({clean_text(k).casefold() for k in keys if clean_text(k)})
        tmp = state_path.with_suffix(state_path.suffix + ".tmp")
        tmp.write_text(json.dumps(values, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(state_path)
    except Exception as exc:  # noqa: BLE001
        logging.warning("Could not save seen IDs file %s: %s", state_path, exc)


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    text = str(value)
    text = text.replace("\xa0", " ").replace("\u2013", "-").replace("\u2014", "-").replace("\u2019", "'")
    return re.sub(r"\s+", " ", text).strip()


def normalise_id(value: Any) -> str:
    text = clean_text(value).lower()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")


def fetch_html(url: str) -> str:
    session = requests.Session()
    session.headers.update(REQUEST_HEADERS)
    try:
        r = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        r.raise_for_status()
        return r.text
    except SSLError:
        logging.warning("SSL verification failed for %s. Retrying without verification for public tender page.", url)
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        r = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True, verify=False)
        r.raise_for_status()
        return r.text


def save_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", errors="ignore")


def parse_datetime(value: str) -> pd.Timestamp:
    text = clean_text(value)
    if not text or text.lower() in {"n/a", "na", "none", "-"}:
        return pd.NaT
    text = re.sub(r"(?i)bid\s*closing\s*date\s*: ?", "", text)
    text = re.sub(r"(?i)closing\s*(date|time)?\s*: ?", "", text)
    text = re.sub(r"(?i)\bat\b", " ", text)
    text = re.sub(r"(?i)(\d{1,2})h(\d{2})\s*(am|pm)?", lambda m: f"{m.group(1)}:{m.group(2)} {m.group(3) or ''}", text)
    text = re.sub(r"(?i)(\d{1,2})h\s*(am|pm)?", lambda m: f"{m.group(1)}:00 {m.group(2) or ''}", text)
    text = clean_text(text)
    if re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}", text):
        ts = pd.to_datetime(text, errors="coerce", dayfirst=False)
    else:
        ts = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if pd.notna(ts) and ts.hour == 0 and ts.minute == 0 and not re.search(r"\b\d{1,2}:\d{2}\b|\b\d{1,2}h", text, flags=re.I):
        ts = ts.replace(hour=10, minute=0)
    return ts


def extract_date_text(text: str) -> str:
    text = clean_text(text)
    patterns = [
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\s*(?:at|,)\s*\d{1,2}:\d{2}(?::\d{2})?\s*(?:am|pm)?\b",
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\b",
        r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?\b",
        r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, flags=re.I)
        if m:
            return clean_text(m.group(0))
    return ""


def field_from_lines(lines: List[str], label_patterns: Iterable[str]) -> str:
    for i, line in enumerate(lines):
        for pat in label_patterns:
            m = re.match(pat, line, flags=re.I)
            if m:
                rest = clean_text(m.group(1) if m.groups() else line.split(":", 1)[-1])
                if rest:
                    return rest
                if i + 1 < len(lines):
                    return clean_text(lines[i + 1])
    return ""


def field_from_text(text: str, label: str, next_labels: Tuple[str, ...]) -> str:
    # Works when the site renders fields on one line, e.g. "Bid Number:... Description:...".
    next_pat = "|".join(re.escape(x) for x in next_labels)
    pat = rf"(?is){re.escape(label)}\s*: ?(.*?)(?=(?:{next_pat})\s*:|$)"
    m = re.search(pat, text)
    if not m:
        return ""
    return clean_text(m.group(1))


def infer_notice_type(reference: str, description: str) -> str:
    txt = f"{reference} {description}".upper()
    if "RFI" in txt or "REQUEST FOR INFORMATION" in txt:
        return "RFI"
    if "RFQ" in txt or "REQUEST FOR QUOTATION" in txt:
        return "RFQ"
    if "RFP" in txt or "REQUEST FOR PROPOSAL" in txt:
        return "RFP"
    return "Tender"


def extract_reference(text: str) -> str:
    text = clean_text(text)
    text = re.sub(r"(?i)^bid\s*number\s*: ?", "", text).strip()
    patterns = [
        r"\b[A-Z]{1,4}\s*-\s*[A-Z]{1,8}\s*\d{1,3}\s*-\s*\d{4}\b",
        r"\b[A-Z]{1,4}\s*-\s*[A-Z]{1,8}\s*\d{1,3}/\d{4}\b",
        r"\b[A-Z]{1,4}\s*[-/]\s*[A-Z]{1,8}\s*[-/]?\s*\d{1,3}\s*[-/]\s*\d{4}\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, flags=re.I)
        if m:
            ref = clean_text(m.group(0))
            ref = re.sub(r"\s*([-\/])\s*", r"\1", ref)
            # Preserve first separator style but remove double spaces after department code.
            return ref.upper()
    # Fallback: return first phrase before Description if it looks useful
    m = re.search(r"(?i)bid\s*number\s*: ?(.+?)(?:description\s*:|$)", text)
    if m:
        return clean_text(m.group(1)).upper()
    return ""


def looks_like_tender_detail(url: str) -> bool:
    low = clean_text(url).lower()
    return "ekurhuleni.gov.za/business/tenders/" in low and not any(x in low for x in ["open-tenders", "awarded", "opening-register", "notices"])


def looks_like_doc(url: str, label: str = "") -> bool:
    low = clean_text(url).lower()
    label_l = clean_text(label).lower()
    path = low.split("?")[0]
    return any(path.endswith(ext) for ext in DOC_EXTS) or any(x in low for x in ["wp-content/uploads", "download", "document"]) or "download" in label_l


def extract_document_links(soup: BeautifulSoup, page_url: str) -> Tuple[str, str]:
    urls: List[str] = []
    names: List[str] = []
    for a in soup.find_all("a", href=True):
        href = clean_text(a.get("href", ""))
        if not href or href.lower().startswith(("javascript:", "mailto:", "tel:", "#")):
            continue
        full = urljoin(page_url, href)
        low = full.lower()
        if any(x in low for x in ["facebook", "twitter", "linkedin", "instagram", "youtube", "/about", "/careers", "/media", "/terms", "/disclaimer"]):
            continue
        label = clean_text(a.get_text(" ")) or os.path.basename(urlparse(full).path) or full
        if looks_like_doc(full, label):
            if full not in urls:
                urls.append(full)
                names.append(label)
    return "; ".join(names), "; ".join(urls)


def discover_listing_pages(start_url: str, first_html: str) -> List[str]:
    pages = [start_url]
    soup = BeautifulSoup(first_html, "html.parser")
    for a in soup.find_all("a", href=True):
        href = urljoin(start_url, a.get("href", ""))
        if "/for-my-business/tenders/open-tenders/" in href and href not in pages:
            # Include numbered pagination pages only.
            if re.search(r"/open-tenders/\d+/?$", href):
                pages.append(href)
    # Also defensively try /2/, /3/ etc. The site currently exposes page 2.
    for n in range(2, MAX_LIST_PAGES + 1):
        candidate = urljoin(start_url, f"/for-my-business/tenders/open-tenders/{n}/")
        if candidate not in pages:
            pages.append(candidate)
    return pages


def discover_detail_links(list_html: str, page_url: str) -> List[Tuple[str, str, str]]:
    soup = BeautifulSoup(list_html, "html.parser")
    links: List[Tuple[str, str, str]] = []
    seen = set()
    for a in soup.find_all("a", href=True):
        full = urljoin(page_url, a.get("href", ""))
        if not looks_like_tender_detail(full):
            continue
        if full in seen:
            continue
        seen.add(full)
        title = clean_text(a.get_text(" "))
        # Attempt list-page description/date from nearby block.
        block_text = ""
        parent = a.find_parent(["article", "div", "li", "h1", "h2", "h3", "h4"])
        if parent:
            texts = [clean_text(parent.get_text(" "))]
            sib = parent.find_next_sibling()
            steps = 0
            while sib is not None and steps < 4:
                if getattr(sib, "name", "") in {"h1", "h2", "h3", "h4"}:
                    break
                texts.append(clean_text(sib.get_text(" ") if hasattr(sib, "get_text") else sib))
                sib = sib.find_next_sibling()
                steps += 1
            block_text = clean_text(" ".join(t for t in texts if t))
        links.append((full, title, block_text))
    return links


def parse_detail_page(detail_url: str, list_title: str = "", list_block: str = "") -> Dict[str, Any]:
    html = fetch_html(detail_url)
    soup = BeautifulSoup(html, "html.parser")
    main = soup.find("main") or soup.find("article") or soup
    text = main.get_text("\n")
    lines = [clean_text(x) for x in text.splitlines() if clean_text(x)]
    collapsed = clean_text(" \n ".join(lines))
    flat = clean_text(" ".join(lines))

    h1 = main.find("h1") or soup.find("h1")
    title = clean_text(h1.get_text(" ")) if h1 else clean_text(list_title)

    labels = (
        "Bid Number", "Description", "Bid closing date", "Information session", "Contact Person (Project manager)",
        "Telephone number", "Official Cell number", "E-mail address", "Bid validity period", "Submit bids to",
        "Non-refundable Document fee",
    )
    reference = field_from_text(flat, "Bid Number", labels) or field_from_lines(lines, [r"Bid\s*Number\s*: ?(.*)"])
    description = field_from_text(flat, "Description", labels) or field_from_lines(lines, [r"Description\s*: ?(.*)"])
    closing_text = field_from_text(flat, "Bid closing date", labels) or field_from_lines(lines, [r"Bid\s*closing\s*date\s*: ?(.*)"])
    briefing_text = field_from_text(flat, "Information session", labels) or field_from_lines(lines, [r"Information\s*session\s*: ?(.*)"])
    contact_person = field_from_text(flat, "Contact Person (Project manager)", labels) or field_from_lines(lines, [r"Contact\s*Person\s*\(Project\s*manager\)\s*: ?(.*)", r"Contact\s*Person\s*: ?(.*)"])
    telephone = field_from_text(flat, "Telephone number", labels) or field_from_lines(lines, [r"Telephone\s*number\s*: ?(.*)"])
    cell = field_from_text(flat, "Official Cell number", labels) or field_from_lines(lines, [r"Official\s*Cell\s*number\s*: ?(.*)"])
    email = field_from_text(flat, "E-mail address", labels) or field_from_lines(lines, [r"E-?mail\s*address\s*: ?(.*)", r"Email\s*: ?(.*)"])
    validity = field_from_text(flat, "Bid validity period", labels) or field_from_lines(lines, [r"Bid\s*validity\s*period\s*: ?(.*)"])
    submit_to = field_from_text(flat, "Submit bids to", labels) or field_from_lines(lines, [r"Submit\s*bids\s*to\s*: ?(.*)"])

    # Fallbacks from listing block if the detail field extraction is imperfect.
    if not reference:
        reference = extract_reference(f"{list_title} {list_block}")
    if not description:
        m = re.search(r"(?is)description\s*: ?(.*?)(?:\b\d{1,2}/\d{1,2}/\d{4}\b|$)", list_block)
        description = clean_text(m.group(1)) if m else clean_text(list_block or title)
    if not closing_text:
        closing_text = extract_date_text(list_block)

    document_names, document_urls = extract_document_links(soup, detail_url)
    return {
        "reference": clean_text(reference),
        "description": clean_text(description),
        "closing_text": clean_text(closing_text),
        "briefing_text": clean_text(briefing_text),
        "contact_person": clean_text(contact_person),
        "contact_email": clean_text(email),
        "contact_telephone": clean_text(telephone or cell),
        "conditions": clean_text("; ".join(x for x in [f"Bid validity period: {validity}" if validity else "", f"Submit bids to: {submit_to}" if submit_to else ""] if x)),
        "document_names": document_names,
        "document_urls": document_urls,
        "title": title,
        "source_url": detail_url,
        "raw_html": html,
    }


def build_row(item: Dict[str, Any]) -> Dict[str, Any]:
    reference = clean_text(item.get("reference"))
    description = clean_text(item.get("description"))
    title = clean_text(item.get("title")) or description or reference
    closing_dt = parse_datetime(clean_text(item.get("closing_text")))
    briefing_dt = parse_datetime(clean_text(item.get("briefing_text")))
    now = pd.Timestamp.now()
    key_seed = reference or f"{description}-{item.get('source_url','')}"
    doc_urls = clean_text(item.get("document_urls"))
    return {
        "source": SOURCE_NAME,
        "opportunity_key": f"city-ekurhuleni:{normalise_id(key_seed)}",
        "tender_number": reference,
        "reference_number": reference,
        "notice_type": infer_notice_type(reference, description),
        "title": title,
        "description": description or title,
        "client": CLIENT_NAME,
        "sector": "Municipal Infrastructure",
        "category": "",
        "country_name": "South Africa",
        "province": "Gauteng",
        "region": "Gauteng",
        "date_published": "",
        "issue_date": "",
        "briefing_date": briefing_dt.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(briefing_dt) else "",
        "closing_date": closing_dt.date().isoformat() if pd.notna(closing_dt) else "",
        "closing_datetime": closing_dt.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(closing_dt) else "",
        "status": "Published",
        "is_current": bool(pd.notna(closing_dt) and closing_dt >= now),
        "is_archived": False,
        "is_cancelled": False,
        "document_names": clean_text(item.get("document_names")),
        "document_urls": doc_urls,
        "document_count": len([u for u in doc_urls.split("; ") if u]) if doc_urls else 0,
        "contact_person": clean_text(item.get("contact_person")),
        "contact_email": clean_text(item.get("contact_email")),
        "contact_telephone": clean_text(item.get("contact_telephone")),
        "submission_method": "Hand delivery to Golden Heights Building tender box / tender instructions",
        "briefing_session_details": clean_text(item.get("briefing_text")),
        "conditions": clean_text(item.get("conditions")),
        "raw_category": "Open Tenders",
        "raw_status": "Published",
        "source_url": clean_text(item.get("source_url")) or SOURCE_URL,
        "run_timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "row_hash": "",
        "change_type": "",
        "previous_row_hash": "",
        "weight_score": 0,
        "matched_description_keywords": "",
    }


def compute_row_hash(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    if df.empty:
        return pd.Series(dtype="object")
    def _hash_row(row: pd.Series) -> str:
        payload = "|".join(clean_text(row.get(c, "")) for c in cols)
        return hashlib.sha256(payload.encode("utf-8", errors="ignore")).hexdigest()
    return df.apply(_hash_row, axis=1)


def score_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        if "weight_score" not in df.columns:
            df["weight_score"] = []
        if "matched_description_keywords" not in df.columns:
            df["matched_description_keywords"] = []
        return df
    scores: List[int] = []
    matches_out: List[str] = []
    for _, row in df.iterrows():
        text = clean_text(" ".join([str(row.get("client", "")), str(row.get("description", "")), str(row.get("reference_number", ""))])).lower()
        matches = [kw for kw in KEYWORDS if kw.lower() in text]
        score = 10 if "city of ekurhuleni" in text or "ekurhuleni" in text else 0
        score += min(len(matches) * 4, 28)
        if row.get("document_count", 0):
            score += 1
        if clean_text(row.get("contact_email")):
            score += 1
        scores.append(score)
        matches_out.append("; ".join(dict.fromkeys(matches)))
    df["weight_score"] = scores
    df["matched_description_keywords"] = matches_out
    return df


def align_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[OUTPUT_COLUMNS]


def read_previous_master(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    try:
        prev = pd.read_excel(path, dtype=str)
        if "opportunity_key" not in prev.columns:
            return pd.DataFrame(columns=OUTPUT_COLUMNS)
        return prev
    except Exception as exc:
        logging.warning("Could not read previous City of Ekurhuleni master: %s", exc)
        return pd.DataFrame(columns=OUTPUT_COLUMNS)


def write_outputs(
    raw_dir: Path,
    out_dir: Path,
    raw_df: pd.DataFrame,
    master_df: pd.DataFrame,
    delta_df: pd.DataFrame,
    updates_df: pd.DataFrame,
    top_delta_df: pd.DataFrame,
) -> None:
    """Write raw scrape evidence and operational Excel outputs."""
    raw_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    raw_df = align_columns(raw_df.copy())
    master_df = align_columns(master_df.copy())
    delta_df = align_columns(delta_df.copy())
    updates_df = align_columns(updates_df.copy())
    top_delta_df = align_columns(top_delta_df.copy())

    raw_df.to_csv(raw_dir / RAW_CSV, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(raw_dir / RAW_XLSX, engine="xlsxwriter") as writer:
        raw_df.to_excel(writer, sheet_name="Raw", index=False)

    with pd.ExcelWriter(out_dir / MASTER_XLSX, engine="xlsxwriter") as writer:
        master_df.to_excel(writer, sheet_name="Master", index=False)
    with pd.ExcelWriter(out_dir / DELTA_XLSX, engine="xlsxwriter") as writer:
        delta_df.to_excel(writer, sheet_name="Delta", index=False)
    with pd.ExcelWriter(out_dir / UPDATES_XLSX, engine="xlsxwriter") as writer:
        updates_df.to_excel(writer, sheet_name="Updates", index=False)
    with pd.ExcelWriter(out_dir / TOP_DELTA_XLSX, engine="xlsxwriter") as writer:
        top_delta_df.to_excel(writer, sheet_name="Top_Delta", index=False)


def fetch_and_parse() -> Tuple[List[Dict[str, Any]], str]:
    combined_html_parts: List[str] = []
    rows: List[Dict[str, Any]] = []
    first_html = fetch_html(SOURCE_URL)
    combined_html_parts.append(f"<!-- {SOURCE_URL} -->\n" + first_html)
    list_pages = discover_listing_pages(SOURCE_URL, first_html)
    logging.info("City of Ekurhuleni listing pages queued: %s", len(list_pages))
    detail_links: List[Tuple[str, str, str]] = []
    seen_detail = set()
    for i, page_url in enumerate(list_pages):
        try:
            html = first_html if i == 0 else fetch_html(page_url)
            combined_html_parts.append(f"\n<!-- {page_url} -->\n" + html)
        except Exception as exc:
            logging.warning("Could not fetch City of Ekurhuleni listing page %s: %s", page_url, exc)
            continue
        links = discover_detail_links(html, page_url)
        logging.info("City of Ekurhuleni detail links from %s: %s", page_url, len(links))
        for link in links:
            if link[0] not in seen_detail:
                seen_detail.add(link[0])
                detail_links.append(link)
        # Stop probing once a numbered page yields no tender detail links and it was not explicitly linked.
        if i > 1 and not links:
            break
    logging.info("City of Ekurhuleni unique tender detail links: %s", len(detail_links))
    for detail_url, title, block_text in detail_links:
        try:
            parsed = parse_detail_page(detail_url, title, block_text)
            combined_html_parts.append(f"\n<!-- {detail_url} -->\n" + parsed.pop("raw_html", ""))
            rows.append(build_row(parsed))
        except Exception as exc:
            logging.warning("Failed to parse City of Ekurhuleni tender detail %s: %s", detail_url, exc)
            # Keep a fallback row from listing page if possible.
            rows.append(build_row({
                "reference": extract_reference(f"{title} {block_text}"),
                "description": block_text or title,
                "closing_text": extract_date_text(block_text),
                "title": title,
                "source_url": detail_url,
            }))
    return rows, "\n".join(combined_html_parts)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Scrape City of Ekurhuleni open tenders")
    parser.add_argument("--base-root", default=str(BASE_ROOT), help="Override source root; normally not needed.")
    args = parser.parse_args(argv)

    base_root = Path(args.base_root).resolve()
    ensure_dirs(base_root)
    setup_logging(base_root)

    raw_dir = base_root / RAW_DIR
    out_dir = base_root / OUTPUT_DIR
    state_dir = base_root / STATE_DIR
    lock_path = state_dir / LOCK_FILE
    seen_ids_path = state_dir / STATE_FILE

    logging.info("Using tender scrapers root: %s", TENDER_ROOT)
    logging.info("Using source root: %s", base_root)
    logging.info("Using raw directory: %s", raw_dir)
    logging.info("Using output directory: %s", out_dir)

    try:
        acquire_lock(lock_path)

        try:
            rows, combined_html = fetch_and_parse()
            save_text(raw_dir / RAW_HTML, combined_html)

            raw_df = pd.DataFrame(rows)
            raw_df = align_columns(raw_df)

            if not raw_df.empty:
                raw_df["row_hash"] = compute_row_hash(raw_df, HASH_COLS)
                raw_df = score_rows(raw_df)

            logging.info("Parsed City of Ekurhuleni raw rows: %s", len(raw_df))

        except Exception as exc:  # noqa: BLE001
            logging.exception(
                "City of Ekurhuleni scraper failed before parsing outputs. "
                "Creating empty outputs so daily run can continue: %s",
                exc,
            )
            raw_df = pd.DataFrame(columns=OUTPUT_COLUMNS)
            save_text(raw_dir / RAW_HTML, "")

        master_df = raw_df.copy()
        if not master_df.empty:
            before = len(master_df)
            master_df = master_df[master_df["is_cancelled"].astype(str).str.lower().ne("true")].copy()
            master_df = master_df[master_df["is_archived"].astype(str).str.lower().ne("true")].copy()
            master_df = master_df[master_df["is_current"].astype(str).str.lower().eq("true")].copy()
            logging.info("Filtered current City of Ekurhuleni tenders: %s -> %s rows", before, len(master_df))

        prev = read_previous_master(out_dir / MASTER_XLSX)
        prev_map: Dict[str, str] = {}
        if not prev.empty and "opportunity_key" in prev.columns:
            prev_hashes = (
                prev["row_hash"].astype(str)
                if "row_hash" in prev.columns
                else pd.Series([""] * len(prev), index=prev.index, dtype="object")
            )
            prev_map = dict(zip(prev["opportunity_key"].astype(str), prev_hashes))

        delta_rows: List[int] = []
        update_rows: List[int] = []

        for idx, row in master_df.iterrows():
            key = str(row.get("opportunity_key", ""))
            current_hash = str(row.get("row_hash", ""))
            prev_hash = prev_map.get(key, "")
            if not prev_hash:
                delta_rows.append(idx)
            elif prev_hash != current_hash:
                update_rows.append(idx)

        if not master_df.empty:
            master_df.loc[:, "previous_row_hash"] = [
                prev_map.get(str(k), "") for k in master_df["opportunity_key"].astype(str)
            ]
            master_df.loc[:, "change_type"] = ""

        delta_df = master_df.loc[delta_rows].copy() if delta_rows else pd.DataFrame(columns=OUTPUT_COLUMNS)
        if not delta_df.empty:
            delta_df.loc[:, "change_type"] = "NEW"

        updates_df = master_df.loc[update_rows].copy() if update_rows else pd.DataFrame(columns=OUTPUT_COLUMNS)
        if not updates_df.empty:
            updates_df.loc[:, "change_type"] = "UPDATED"

        top_delta_df = (
            delta_df.sort_values(["weight_score", "closing_datetime"], ascending=[False, True]).head(20).copy()
            if not delta_df.empty
            else pd.DataFrame(columns=OUTPUT_COLUMNS)
        )

        write_outputs(raw_dir, out_dir, raw_df, master_df, delta_df, updates_df, top_delta_df)
        save_seen_ids(seen_ids_path, master_df.get("opportunity_key", pd.Series(dtype="object")).tolist())

        logging.info(
            "Saved City of Ekurhuleni outputs | raw=%s master=%s delta=%s updates=%s top_delta=%s",
            len(raw_df),
            len(master_df),
            len(delta_df),
            len(updates_df),
            len(top_delta_df),
        )
        return 0

    except (RequestException, RuntimeError, OSError, ValueError) as exc:
        logging.exception("FAILED: %s", exc)
        return 1

    finally:
        release_lock(lock_path)


if __name__ == "__main__":
    raise SystemExit(main())
