"""
ACSA Tender Bulletin Scraper (Raw + Master + Delta + Updates + Top Delta)

Fetch tenders from Airports Company South Africa's public Tender Bulletin page,
clean/normalise them, and maintain:

1) RAW CSV/XLSX/HTML: direct page scrape for audit/debugging
2) MASTER Excel: latest current snapshot per opportunity_key
3) DELTA Excel: brand new current tenders discovered this run only
4) UPDATES Excel: existing current tenders whose key operational fields changed
5) TOP DELTA Excel: highest-scoring new tenders from the current run

Typical daily run:
    python acsa_tender_pull_and_clean.py

Optional:
    python acsa_tender_pull_and_clean.py --include-expired
    python acsa_tender_pull_and_clean.py --top-delta-limit 50
    python acsa_tender_pull_and_clean.py --top-delta-min-score 3
    python acsa_tender_pull_and_clean.py --state-file "path/to/seen_ids.json"
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import re
from pathlib import Path
import sys
import time
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# pylint: disable=abstract-class-instantiated

SOURCE_NAME = "ACSA"
SOURCE_URL = "https://www.airports.co.za/Pages/Tender-Bulletin.aspx"

# === Folder structure (new 02_TenderScrapers architecture) ===
# Expected script location:
#   02_TenderScrapers\02_Sources\ACSA\00_Scripts\acsa_tender_pull_and_clean.py
SCRIPT_DIR = Path(__file__).resolve().parent
SOURCE_ROOT = SCRIPT_DIR.parent
TENDER_SCRAPERS_ROOT = SOURCE_ROOT.parent.parent

RAW_DIR = str(SOURCE_ROOT / "01_RawData")
OUTPUT_DIR = str(SOURCE_ROOT / "02_Outputs")
LOG_DIR = str(SOURCE_ROOT / "03_Logs")
STATE_DIR = str(SOURCE_ROOT / "04_State")
METADATA_DIR = str(SOURCE_ROOT / "05_Metadata")

LOG_PATH = os.path.join(LOG_DIR, "acsa_run.log")
LOCK_FILE = os.path.join(STATE_DIR, "run.lock")

RAW_HTML = os.path.join(RAW_DIR, "acsa_raw_page.html")
RAW_CSV = os.path.join(RAW_DIR, "acsa_raw.csv")
RAW_XLSX = os.path.join(RAW_DIR, "acsa_raw_full.xlsx")

MASTER_XLSX = os.path.join(OUTPUT_DIR, "acsa_master.xlsx")
DELTA_XLSX = os.path.join(OUTPUT_DIR, "acsa_delta.xlsx")
UPDATES_XLSX = os.path.join(OUTPUT_DIR, "acsa_updates.xlsx")
TOP_DELTA_XLSX = os.path.join(OUTPUT_DIR, "acsa_top_delta.xlsx")

STATE_FILE_DEFAULT = os.path.join(STATE_DIR, "seen_ids.json")

TIMEOUT_S = 60
TOP_DELTA_LIMIT_DEFAULT = 20
TOP_DELTA_MIN_SCORE_DEFAULT = 0

# Key fields that define a meaningful operational change.
HASH_COLS = [
    "airport",
    "category",
    "description",
    "briefing_date",
    "closing_date",
    "primary_document_name",
    "primary_document_url",
    "additional_document_names",
    "additional_document_urls",
    "document_count",
    "cancellation_flag",
]

# Scoring weights.
WEIGHT_CATEGORY_MATCH = 5
WEIGHT_DESCRIPTION_KEYWORD = 1

PRIORITY_CATEGORIES = [
    "Professional Services",
    "Provision of services of a consultant",
    "Maintenance and Engineering",
    "Electrical Maintenance",
    "Technical Solution Services",
    "Building and Maintenance",
    "Property Development",
    "Corporate Affairs",
]

DESCRIPTION_KEYWORD_RULES = {
    "Water and Sanitation": ["Water and Sanitation"],
    "Housing": ["Housing"],
    "Education": ["Education"],
    "Transport": ["Transport"],
    "Airport": ["Airport", "Airports"],
    "Runway": ["Runway", "Taxiway", "Taxiways"],
    "Nuclear": ["Nuclear"],
    "Bus Rapid Transit (BRT)": ["Bus Rapid Transit", "BRT"],
    "International Transition Plan Network (ITPN)": [
        "International Transition Plan Network",
        "ITPN",
    ],
    "Design": ["Design"],
    "Refurbishment": ["Refurbishment", "Refurbish"],
    "Engineering, Procurement, and Construction (EPC)": [
        "Engineering, Procurement, and Construction",
        "Engineering Procurement and Construction",
        "EPC",
    ],
    "Fire Engineering": ["Fire Engineering", "Fire Fighting", "Fire"],
    "HVAC": ["HVAC"],
    "Mechanical": ["Mechanical"],
    "Electrical": ["Electrical"],
    "Electronic": ["Electronic"],
    "Wet Services": ["Wet Services"],
    "Project Management": ["Project Management"],
    "Programme Management": ["Programme Management"],
    "Project Management Office (PMO)": ["Project Management Office", "PMO"],
    "Advisory Services": ["Advisory Services", "Advisory"],
    "Structural Engineering": ["Structural Engineering"],
    "Civil Engineering": ["Civil Engineering", "Civil"],
    "Feasibility Study": ["Feasibility Study", "Feasibility"],
    "Construction": ["Construction", "Contractor"],
    "Maintenance": ["Maintenance", "Repairs", "Repair"],
}

OUTPUT_COLUMNS = [
    "source",
    "opportunity_key",
    "tender_number",
    "notice_type",
    "airport",
    "category",
    "description",
    "briefing_date",
    "closing_date",
    "primary_document_name",
    "primary_document_url",
    "additional_document_names",
    "additional_document_urls",
    "document_count",
    "source_url",
    "cancellation_flag",
    "priority_category_match",
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
    )


def build_session() -> requests.Session:
    """Requests session with retry/backoff for flaky endpoints."""
    sess = requests.Session()
    sess.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
            ),
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
    """Simple lockfile to prevent concurrent runs."""
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


def clean_text(value: Any) -> str:
    """Collapse whitespace and remove nuisance non-breaking spaces."""
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").replace("\u200b", " ")
    return re.sub(r"\s+", " ", text).strip()


def normalise_exact_text(value: Any) -> str:
    """Normalise text for exact comparisons, ignoring case/spacing."""
    return clean_text(value).casefold()


def normalise_search_text(value: Any) -> str:
    """Normalise text for phrase matching, ignoring case and punctuation differences."""
    cleaned = re.sub(r"[^a-z0-9]+", " ", clean_text(value).casefold())
    return " ".join(cleaned.split())


def contains_phrase(text_norm: str, phrase: str) -> bool:
    phrase_norm = normalise_search_text(phrase)
    if not phrase_norm:
        return False
    return f" {phrase_norm} " in f" {text_norm} "


def fetch_page_html() -> str:
    """Download the ACSA tender bulletin HTML page."""
    sess = build_session()
    response = sess.get(SOURCE_URL, timeout=TIMEOUT_S)
    response.raise_for_status()
    return response.text


def save_raw_html(html: str) -> None:
    os.makedirs(RAW_DIR, exist_ok=True)
    with open(RAW_HTML, "w", encoding="utf-8") as fh:
        fh.write(html)
    logging.info("Saved raw HTML: %s", RAW_HTML)


def is_likely_tender_table(table) -> bool:
    """Identify the SharePoint list table containing the tender bulletin."""
    table_text = normalise_search_text(table.get_text(" "))
    required = ["airport", "description", "briefing date", "closing date"]
    return all(contains_phrase(table_text, phrase) for phrase in required)


def extract_cell_links(cell) -> List[Tuple[str, str]]:
    """Return (link text, absolute url) pairs from a table cell."""
    links: List[Tuple[str, str]] = []
    for a_tag in cell.find_all("a", href=True):
        name = clean_text(a_tag.get_text(" "))
        href = clean_text(a_tag.get("href"))
        if not href or href.lower().startswith("javascript:"):
            continue
        if not name:
            name = os.path.basename(href).replace("%20", " ")
        links.append((name, urljoin(SOURCE_URL, href)))
    return links


def make_unique_key(fields: Sequence[Any]) -> str:
    """Create stable hash key for records where no clear tender number exists."""
    joined = "||".join(clean_text(x).casefold() for x in fields)
    return hashlib.sha256(joined.encode("utf-8")).hexdigest()[:20]


def extract_tender_number(text: str) -> str:
    """Extract common ACSA tender/reference numbers from descriptions or document names."""
    raw = clean_text(text)
    patterns = [
        r"\b[A-Z]{2,6}\s*[-/]?\s*\d{3,8}\s*[-/]?\s*\d{4}\s*[-/]?\s*(?:RFP|RFQ|RFB|BID)?\b",
        r"\b[A-Z]{2,6}\s*\d{3,8}\s*\d{4}\s*(?:RFP|RFQ|RFB|BID)?\b",
        r"\b(?:COR|CIA|CTIA|KSIA|GRJ|ORTIA|BFN|PLZ|ELS|KIM|UPN)\s*[-/]?\s*\d{3,8}\s*[-/]?\s*\d{4}\s*[-/]?\s*(?:RFP|RFQ|RFB|BID)?\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, raw, flags=re.IGNORECASE)
        if match:
            return re.sub(r"\s+", "", match.group(0).upper()).strip("-/ ")
    return ""


def infer_notice_type(text: str) -> str:
    norm = normalise_search_text(text)
    rules = [
        ("RFP", ["request for proposal", "rfp"]),
        ("RFB", ["request for bids", "request for bid", "rfb"]),
        ("RFQ", ["request for quotation", "rfq"]),
        ("RFI", ["request for information", "rfi"]),
    ]
    for label, phrases in rules:
        if any(contains_phrase(norm, phrase) for phrase in phrases):
            return label
    return "Tender"


def combine_link_names(links: Iterable[Tuple[str, str]]) -> str:
    return "; ".join(clean_text(name) for name, _url in links if clean_text(name))


def combine_link_urls(links: Iterable[Tuple[str, str]]) -> str:
    return "; ".join(clean_text(url) for _name, url in links if clean_text(url))


def parse_date(value: Any) -> pd.Timestamp:
    """Parse ACSA date strings into pandas timestamps without pandas dayfirst warnings."""
    if value is None or pd.isna(value):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return value

    text = clean_text(value)
    if not text or text.casefold() in {"nat", "nan", "none"}:
        return pd.NaT

    # Normalise common separators and SharePoint/export noise.
    cleaned = text.replace("@", " ").replace(".", "/")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()

    explicit_formats = [
        "%d/%m/%Y",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %I:%M %p",
        "%d/%m/%Y %I:%M:%S %p",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d",
        "%Y/%m/%d %H:%M",
        "%Y/%m/%d %H:%M:%S",
        "%m/%d/%Y %I:%M:%S %p",  # Excel/SharePoint occasionally exports this way.
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y",
    ]
    for fmt in explicit_formats:
        try:
            return pd.Timestamp(pd.to_datetime(cleaned, format=fmt, errors="raise"))
        except (ValueError, TypeError):
            continue

    # Final fallback for unusual text dates.
    return pd.to_datetime(cleaned, errors="coerce", dayfirst=True)


def table_to_rows(table) -> List[Dict[str, Any]]:
    """Parse one ACSA tender table into row dictionaries."""
    parsed_rows: List[Dict[str, Any]] = []
    header_map: Dict[int, str] = {}

    for tr in table.find_all("tr"):
        cells = tr.find_all(["th", "td"])
        if not cells:
            continue

        cell_texts = [clean_text(cell.get_text(" ")) for cell in cells]
        header_candidates = [normalise_search_text(x) for x in cell_texts]

        # Detect header row.
        if (
            any(x == "airport" for x in header_candidates)
            and any("description" == x for x in header_candidates)
            and any("closing date" == x for x in header_candidates)
        ):
            header_map = {}
            for idx, header in enumerate(header_candidates):
                if header == "airport":
                    header_map[idx] = "airport"
                elif header in {"tcategory", "category"}:
                    header_map[idx] = "category"
                elif header == "description":
                    header_map[idx] = "description"
                elif header == "briefing date":
                    header_map[idx] = "briefing_date"
                elif header == "closing date":
                    header_map[idx] = "closing_date"
                elif header == "name":
                    header_map[idx] = "primary_document"
                elif header == "additional documents":
                    header_map[idx] = "additional_documents"
            continue

        if not header_map:
            continue

        row: Dict[str, Any] = {
            "source": SOURCE_NAME,
            "airport": "",
            "category": "",
            "description": "",
            "briefing_date": pd.NaT,
            "closing_date": pd.NaT,
            "primary_document_name": "",
            "primary_document_url": "",
            "additional_document_names": "",
            "additional_document_urls": "",
            "document_count": 0,
            "source_url": SOURCE_URL,
        }

        primary_links: List[Tuple[str, str]] = []
        additional_links: List[Tuple[str, str]] = []

        for idx, cell in enumerate(cells):
            field = header_map.get(idx)
            if not field:
                continue
            if field in {"airport", "category", "description"}:
                row[field] = clean_text(cell.get_text(" "))
            elif field == "briefing_date":
                row[field] = parse_date(cell.get_text(" "))
            elif field == "closing_date":
                row[field] = parse_date(cell.get_text(" "))
            elif field == "primary_document":
                primary_links = extract_cell_links(cell)
            elif field == "additional_documents":
                additional_links = extract_cell_links(cell)

        row["primary_document_name"] = combine_link_names(primary_links)
        row["primary_document_url"] = combine_link_urls(primary_links)
        row["additional_document_names"] = combine_link_names(additional_links)
        row["additional_document_urls"] = combine_link_urls(additional_links)
        row["document_count"] = len(primary_links) + len(additional_links)

        # Some SharePoint rendered cells can merge text awkwardly. Only accept records
        # with at least a description or document link and a closing date/document date.
        has_content = bool(row["description"] or row["primary_document_url"] or row["additional_document_urls"])
        if has_content:
            all_text = " ".join(
                [
                    row.get("airport", ""),
                    row.get("category", ""),
                    row.get("description", ""),
                    row.get("primary_document_name", ""),
                    row.get("additional_document_names", ""),
                ]
            )
            tender_number = extract_tender_number(all_text)
            row["tender_number"] = tender_number
            row["notice_type"] = infer_notice_type(all_text)
            row["cancellation_flag"] = bool(
                re.search(r"\b(cancellation|cancelled|withdrawal|withdrawn)\b", all_text, flags=re.IGNORECASE)
            )
            if tender_number:
                row["opportunity_key"] = f"acsa:{tender_number.casefold()}"
            elif row.get("primary_document_url"):
                row["opportunity_key"] = "acsa:doc:" + make_unique_key([row["primary_document_url"]])
            else:
                row["opportunity_key"] = "acsa:" + make_unique_key(
                    [row["airport"], row["category"], row["description"]]
                )
            parsed_rows.append(row)

    return parsed_rows


def parse_with_html_tables(html: str) -> List[Dict[str, Any]]:
    """Parse SharePoint HTML tables from the page."""
    soup = BeautifulSoup(html, "html.parser")
    rows: List[Dict[str, Any]] = []
    tables = soup.find_all("table")
    logging.info("HTML tables found: %d", len(tables))

    for table in tables:
        if is_likely_tender_table(table):
            table_rows = table_to_rows(table)
            logging.info("Parsed tender rows from likely table: %d", len(table_rows))
            rows.extend(table_rows)

    return rows


def parse_with_pandas(html: str) -> List[Dict[str, Any]]:
    """Fallback parser using pandas.read_html if direct table parsing fails."""
    try:
        tables = pd.read_html(html)
    except ValueError:
        return []

    logging.info("pandas.read_html tables found: %d", len(tables))
    out: List[Dict[str, Any]] = []
    for df in tables:
        cols_norm = {normalise_search_text(c): c for c in df.columns}
        required = {"airport", "description", "closing date"}
        if not required.issubset(set(cols_norm)):
            continue

        for _, rec in df.iterrows():
            airport = clean_text(rec.get(cols_norm.get("airport", ""), ""))
            category = clean_text(rec.get(cols_norm.get("tcategory", ""), rec.get(cols_norm.get("category", ""), "")))
            description = clean_text(rec.get(cols_norm.get("description", ""), ""))
            briefing_date = parse_date(rec.get(cols_norm.get("briefing date", ""), ""))
            closing_date = parse_date(rec.get(cols_norm.get("closing date", ""), ""))
            primary_name = clean_text(rec.get(cols_norm.get("name", ""), ""))
            additional_names = clean_text(rec.get(cols_norm.get("additional documents", ""), ""))
            all_text = " ".join([airport, category, description, primary_name, additional_names])
            tender_number = extract_tender_number(all_text)
            key = f"acsa:{tender_number.casefold()}" if tender_number else "acsa:" + make_unique_key(
                [airport, category, description]
            )
            out.append(
                {
                    "source": SOURCE_NAME,
                    "opportunity_key": key,
                    "tender_number": tender_number,
                    "notice_type": infer_notice_type(all_text),
                    "airport": airport,
                    "category": category,
                    "description": description,
                    "briefing_date": briefing_date,
                    "closing_date": closing_date,
                    "primary_document_name": primary_name,
                    "primary_document_url": "",
                    "additional_document_names": additional_names,
                    "additional_document_urls": "",
                    "document_count": 0,
                    "source_url": SOURCE_URL,
                    "cancellation_flag": bool(
                        re.search(r"\b(cancellation|cancelled|withdrawal|withdrawn)\b", all_text, flags=re.IGNORECASE)
                    ),
                }
            )
    return out


def dedupe_acsa_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Dedupe ACSA rows that point to the same tender document/description.

    The ACSA page can expose the same document more than once. When that
    happens, keep the earliest closing date because duplicated rows have shown
    ambiguous day/month parsing differences.
    """
    if df is None or df.empty:
        return df

    out = df.copy()
    if "closing_date" in out.columns:
        out["closing_date"] = out["closing_date"].apply(parse_date)

    def dedupe_key(row):
        doc_url = clean_text(row.get("primary_document_url", "")) or clean_text(row.get("additional_document_urls", ""))
        if doc_url:
            first_url = doc_url.replace("|", ";").split(";")[0].strip().casefold()
            return "doc:" + first_url
        desc = normalise_search_text(row.get("description", ""))
        airport = normalise_search_text(row.get("airport", ""))
        category = normalise_search_text(row.get("category", ""))
        return "text:" + "|".join([airport, category, desc])

    out["_dedupe_key"] = out.apply(dedupe_key, axis=1)
    out = out.sort_values("closing_date", ascending=True, na_position="last", kind="mergesort")
    out = out.drop_duplicates(subset=["_dedupe_key"], keep="first")
    out = out.drop(columns=["_dedupe_key"], errors="ignore")
    return out


def parse_acsa_tenders(html: str) -> pd.DataFrame:
    """Parse ACSA tender rows into a dataframe."""
    rows = parse_with_html_tables(html)
    if not rows:
        logging.warning("Direct HTML parser found 0 rows; trying pandas.read_html fallback.")
        rows = parse_with_pandas(html)

    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame(columns=OUTPUT_COLUMNS)

    # Dedupe obvious duplicate rows by document/description first, then by opportunity key.
    if not df.empty:
        df = dedupe_acsa_rows(df)
    if "opportunity_key" in df.columns and not df.empty:
        df = df.drop_duplicates(subset=["opportunity_key"], keep="first")

    logging.info("Parsed ACSA notice rows: %d | cols: %d", len(df), len(df.columns))
    return df


def filter_current_tenders(df: pd.DataFrame, include_expired: bool, include_cancellations: bool) -> pd.DataFrame:
    """Keep only current/non-cancelled tenders unless overridden by CLI flags."""
    out = df.copy()

    if not include_cancellations and "cancellation_flag" in out.columns:
        before = len(out)
        out = out[~out["cancellation_flag"].fillna(False).astype(bool)].copy()
        logging.info("Filtered cancellation/withdrawal notices: %d -> %d rows", before, len(out))

    if not include_expired and "closing_date" in out.columns:
        before = len(out)
        today = pd.Timestamp.today().normalize()
        out["closing_date"] = out["closing_date"].apply(parse_date)
        out = out[out["closing_date"].notna() & (out["closing_date"] >= today)].copy()
        logging.info("Filtered current tenders by closing date: %d -> %d rows", before, len(out))

    return out


def add_priority_scoring(df: pd.DataFrame) -> pd.DataFrame:
    """Add category/description match details and a weight score."""
    out = df.copy()
    priority_categories_norm = {normalise_exact_text(x) for x in PRIORITY_CATEGORIES}

    if "category" not in out.columns:
        out["priority_category_match"] = False
    else:
        out["priority_category_match"] = out["category"].apply(
            lambda x: normalise_exact_text(x) in priority_categories_norm
        )

    def matched_keywords(record: pd.Series) -> str:
        desc_norm = normalise_search_text(
            " ".join(
                [
                    clean_text(record.get("description", "")),
                    clean_text(record.get("category", "")),
                    clean_text(record.get("airport", "")),
                ]
            )
        )
        matches = []
        for label, phrases in DESCRIPTION_KEYWORD_RULES.items():
            if any(contains_phrase(desc_norm, phrase) for phrase in phrases):
                matches.append(label)
        return "; ".join(matches)

    if out.empty:
        out["matched_description_keywords"] = pd.Series(dtype="object")
    else:
        out["matched_description_keywords"] = out.apply(matched_keywords, axis=1)

    out["description_keyword_match_count"] = out["matched_description_keywords"].apply(
        lambda x: 0 if not x else len(str(x).split("; "))
    )
    out["weight_score"] = (
        out["priority_category_match"].astype(int) * WEIGHT_CATEGORY_MATCH
        + out["description_keyword_match_count"].astype(int) * WEIGHT_DESCRIPTION_KEYWORD
    )

    sort_cols = [c for c in ["weight_score", "closing_date", "briefing_date"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

    return out


def compute_row_hash(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    """Stable SHA256 hash over selected fields (null-safe, order-stable)."""
    if df is None or df.empty:
        return pd.Series([], dtype="object", index=df.index if df is not None else None)

    use_cols = [c for c in cols if c in df.columns]
    if not use_cols:
        return pd.Series([""] * len(df), index=df.index, dtype="object")

    tmp = df[use_cols].copy()
    for c in use_cols:
        if pd.api.types.is_datetime64_any_dtype(tmp[c]):
            tmp[c] = tmp[c].dt.strftime("%Y-%m-%d %H:%M:%S").fillna("")
        else:
            tmp[c] = tmp[c].fillna("").astype(str).str.strip()

    joined = tmp.apply(lambda r: "||".join(r.values.tolist()), axis=1)
    return joined.apply(lambda s: hashlib.sha256(s.encode("utf-8")).hexdigest())


def write_excel_with_table(df: pd.DataFrame, path: str, sheet_name: str, table_name: str) -> None:
    """Write XLSX and create an Excel Table when rows exist. Atomic replace."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp.xlsx"

    out_df = df.copy() if df is not None else pd.DataFrame()
    for col in OUTPUT_COLUMNS:
        if col not in out_df.columns:
            out_df[col] = ""

    # Preserve source-specific ordering for known columns; keep any unexpected columns after.
    ordered_cols = [c for c in OUTPUT_COLUMNS if c in out_df.columns] + [
        c for c in out_df.columns if c not in OUTPUT_COLUMNS
    ]
    out_df = out_df[ordered_cols]

    with pd.ExcelWriter(tmp, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm") as writer:
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


def save_raw_outputs(df: pd.DataFrame) -> None:
    """Save unfiltered raw parsed outputs."""
    os.makedirs(RAW_DIR, exist_ok=True)
    df.to_csv(RAW_CSV, index=False)
    logging.info("Saved raw CSV: %s", RAW_CSV)

    with pd.ExcelWriter(RAW_XLSX, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm") as writer:
        df.to_excel(writer, index=False, sheet_name="Raw_Data")
    logging.info("Saved raw Excel: %s", RAW_XLSX)


def load_master(path: str) -> pd.DataFrame:
    """Load master Excel if it exists; else empty DF."""
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name="Master", engine="openpyxl")
    except Exception as exc:  # noqa: BLE001
        logging.exception("Failed to read master file (%s): %s", path, exc)
        return pd.DataFrame()


def append_and_dedupe_master(master: pd.DataFrame, incoming: pd.DataFrame) -> pd.DataFrame:
    """Append incoming rows to master and dedupe by opportunity_key. Incoming rows always win."""
    if master is None or master.empty:
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
    out = out[out["opportunity_key"].ne("")].copy()

    if "closing_date" in out.columns:
        out["closing_date"] = out["closing_date"].apply(parse_date)

    # Sort by key first, then source priority, so the current run replaces the old row.
    # Closing date is only a secondary display/order field, not the deciding factor.
    sort_cols = ["opportunity_key", "_source_priority"]
    ascending = [True, False]
    if "closing_date" in out.columns:
        sort_cols.append("closing_date")
        ascending.append(True)

    out = out.sort_values(sort_cols, ascending=ascending, na_position="last", kind="mergesort")
    out = out.drop_duplicates(subset=["opportunity_key"], keep="first")
    out = out.drop(columns=["_source_priority"], errors="ignore")
    return out

def build_delta_updates(master_df: pd.DataFrame, mapped: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Build new and changed tender lists."""
    if not master_df.empty and "opportunity_key" in master_df.columns:
        master_df = master_df.copy()
        master_df["opportunity_key"] = master_df["opportunity_key"].fillna("").astype(str).str.strip().str.lower()
        master_df = master_df[master_df["opportunity_key"].ne("")]

        if "row_hash" not in master_df.columns:
            master_df["row_hash"] = compute_row_hash(master_df, HASH_COLS)

        master_ids = set(master_df["opportunity_key"].tolist())
        master_hash_map = dict(
            zip(
                master_df["opportunity_key"].tolist(),
                master_df["row_hash"].fillna("").astype(str).tolist(),
            )
        )

        is_new = ~mapped["opportunity_key"].isin(master_ids)
        delta_df = mapped[is_new].copy()
        delta_df["change_type"] = "NEW"
        delta_df["previous_row_hash"] = ""

        exists = ~is_new
        mapped_prev_hash = mapped["opportunity_key"].map(master_hash_map).fillna("")
        is_updated = exists & (mapped["row_hash"] != mapped_prev_hash)

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

    return delta_df, updates_df


def build_top_delta(df: pd.DataFrame, limit: int, min_score: int) -> pd.DataFrame:
    """Return highest-scoring new tenders from the current run's Delta list."""
    out = df.copy()

    if "weight_score" not in out.columns:
        out["weight_score"] = 0

    out = out[out["weight_score"].fillna(0).astype(int) >= int(min_score)].copy()

    sort_cols = [c for c in ["weight_score", "closing_date", "briefing_date"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

    if limit and int(limit) > 0:
        out = out.head(int(limit)).copy()

    return out


def sort_operational_output(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    sort_cols = [c for c in ["weight_score", "closing_date", "briefing_date"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")
    return out



def _save_seen_ids(path: str, ids: Iterable[str]) -> None:
    """Persist seen opportunity keys for source-level audit/state tracking."""
    try:
        norm = sorted(str(x).strip().casefold() for x in ids if str(x).strip())
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(norm, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except (OSError, TypeError) as exc:
        logging.exception("Failed to save state file: %s -- %s", path, exc)


def main() -> int:
    parser = argparse.ArgumentParser(description="Fetch and clean ACSA Tender Bulletin data")
    parser.add_argument("--include-expired", action="store_true", help="keep expired tenders in clean outputs")
    parser.add_argument("--include-cancellations", action="store_true", help="keep cancellation/withdrawal notices in clean outputs")
    parser.add_argument("--top-delta-limit", type=int, default=TOP_DELTA_LIMIT_DEFAULT,
                        help="maximum rows in acsa_top_delta.xlsx. Use 0 for all qualifying rows.")
    parser.add_argument("--top-delta-min-score", type=int, default=TOP_DELTA_MIN_SCORE_DEFAULT,
                        help="minimum weight_score required for acsa_top_delta.xlsx")
    parser.add_argument("--state-file", default=STATE_FILE_DEFAULT, help="path to JSON file storing seen opportunity keys")
    args = parser.parse_args()

    setup_logging()
    os.makedirs(RAW_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(LOG_DIR, exist_ok=True)
    os.makedirs(STATE_DIR, exist_ok=True)
    os.makedirs(METADATA_DIR, exist_ok=True)
    logging.info("Using tender scrapers root: %s", TENDER_SCRAPERS_ROOT)
    logging.info("Using source root: %s", SOURCE_ROOT)
    logging.info("Using raw directory: %s", RAW_DIR)
    logging.info("Using output directory: %s", OUTPUT_DIR)

    run_ts = pd.Timestamp.now()

    try:
        acquire_lock(LOCK_FILE)

        html = fetch_page_html()
        save_raw_html(html)

        raw_df = parse_acsa_tenders(html)
        save_raw_outputs(raw_df)

        mapped = filter_current_tenders(
            raw_df,
            include_expired=args.include_expired,
            include_cancellations=args.include_cancellations,
        )
        mapped = add_priority_scoring(mapped)
        mapped = mapped.copy()
        mapped["row_hash"] = compute_row_hash(mapped, HASH_COLS)
        mapped["run_timestamp"] = run_ts

        master_df = load_master(MASTER_XLSX)
        if not master_df.empty:
            master_df = filter_current_tenders(
                master_df,
                include_expired=args.include_expired,
                include_cancellations=args.include_cancellations,
            )

        delta_df, updates_df = build_delta_updates(master_df, mapped)
        delta_df = sort_operational_output(delta_df)
        updates_df = sort_operational_output(updates_df)
        top_delta_df = build_top_delta(delta_df, args.top_delta_limit, args.top_delta_min_score)

        logging.info("Delta (NEW) rows this run: %d", len(delta_df))
        logging.info("Updates (CHANGED) rows this run: %d", len(updates_df))
        logging.info("Top delta rows this run: %d", len(top_delta_df))

        write_excel_with_table(delta_df, DELTA_XLSX, sheet_name="Delta", table_name="tbl_delta")
        logging.info("Saved delta Excel: %s | rows: %d", DELTA_XLSX, len(delta_df))

        write_excel_with_table(updates_df, UPDATES_XLSX, sheet_name="Updates", table_name="tbl_updates")
        logging.info("Saved updates Excel: %s | rows: %d", UPDATES_XLSX, len(updates_df))

        write_excel_with_table(top_delta_df, TOP_DELTA_XLSX, sheet_name="Top_Delta", table_name="tbl_top_delta")
        logging.info("Saved top delta Excel: %s | rows: %d", TOP_DELTA_XLSX, len(top_delta_df))

        updated_master = append_and_dedupe_master(master_df, mapped)
        updated_master = sort_operational_output(updated_master)
        write_excel_with_table(updated_master, MASTER_XLSX, sheet_name="Master", table_name="tbl_master")
        logging.info("Saved master Excel: %s | rows: %d", MASTER_XLSX, len(updated_master))

        if "opportunity_key" in updated_master.columns:
            _save_seen_ids(args.state_file, updated_master["opportunity_key"].tolist())
            logging.info("Saved seen IDs state file: %s", args.state_file)

        return 0

    except (requests.RequestException, ValueError, RuntimeError, OSError) as exc:
        logging.exception("FAILED: %s", exc)
        return 1

    finally:
        release_lock(LOCK_FILE)


if __name__ == "__main__":
    raise SystemExit(main())
