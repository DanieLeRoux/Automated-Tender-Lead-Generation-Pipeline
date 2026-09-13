"""
etenders_pull_and_clean (Delta + Updates + Master)

Fetch currently advertised tenders from the South African eTenders API,
clean/normalise them, and maintain:

1) MASTER Excel: latest snapshot per bid_number (deduped by bid_number)
2) DELTA Excel: brand new tenders discovered this run only (new bid_number)
3) UPDATES Excel: existing tenders whose key operational fields changed (hash change)
4) TOP DELTA Excel: highest-scoring new tenders from the current run
5) Document links extracted from eTenders supportDocument metadata where available

Typical daily run:
    python etenders_pull_and_clean.py --since-days 14

Optional:
    python etenders_pull_and_clean.py --since-days 21
    python etenders_pull_and_clean.py --stop-on-seen  (uses state file to early-stop paging)
"""

from __future__ import annotations

import os
import sys
import time
import json
import logging
import argparse
import hashlib
import re
from urllib.parse import quote
from typing import Optional, List, Dict, Any, Set

import requests
import pandas as pd
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# pylint: disable=abstract-class-instantiated

BASE_PAGE = "https://www.etenders.gov.za/Home/opportunities?id=1"
API_URL = "https://www.etenders.gov.za/Home/PaginatedTenderOpportunities"

# ---- tunables ----
PAGE_SIZE = 200
STATUS = 1  # 1 = "Currently Advertised"
ORDER_COL = 2
ORDER_DIR = "desc"
SLEEP_S = 0.4
TIMEOUT_S = 60
MAX_PAGES = 200  # safety brake; PAGE_SIZE*MAX_PAGES = max rows
# -------------------

# === Folder structure ===
# Expected location of this script:
#   02_TenderScrapers\02_Sources\eTenders\00_Scripts\etenders_pull_and_clean.py
#
# Paths are derived from the script location so the scraper remains portable
# inside the reorganised TenderScrapers architecture.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
TENDER_SCRAPERS_ROOT = os.path.abspath(os.path.join(SOURCE_ROOT, "..", ".."))

RAW_DIR = os.path.join(SOURCE_ROOT, "01_RawData")
OUTPUT_DIR = os.path.join(SOURCE_ROOT, "02_Outputs")
LOG_DIR = os.path.join(SOURCE_ROOT, "03_Logs")
STATE_DIR = os.path.join(SOURCE_ROOT, "04_State")
METADATA_DIR = os.path.join(SOURCE_ROOT, "05_Metadata")

LOG_PATH = os.path.join(LOG_DIR, "etenders_run.log")

# Files
MASTER_XLSX = os.path.join(OUTPUT_DIR, "etenders_master.xlsx")
DELTA_XLSX = os.path.join(OUTPUT_DIR, "etenders_delta.xlsx")
UPDATES_XLSX = os.path.join(OUTPUT_DIR, "etenders_updates.xlsx")
TOP_DELTA_XLSX = os.path.join(OUTPUT_DIR, "etenders_top_delta.xlsx")

RAW_CSV = os.path.join(RAW_DIR, "etenders_active_raw.csv")
RAW_XLSX = os.path.join(RAW_DIR, "etenders_active_raw_full.xlsx")

STATE_FILE_DEFAULT = os.path.join(STATE_DIR, "seen_ids.json")
HASH_VERSION_FILE_DEFAULT = os.path.join(STATE_DIR, "hash_version.txt")
LOCK_FILE = os.path.join(STATE_DIR, "run.lock")

# Key operational fields that define "meaningful change"
# (low-noise: excludes description/category/client/date_published etc.)
# Conditions is included because it can materially affect whether/how a tender is pursued.
HASH_COLS = [
    "closing_date",
    "closing_datetime",
    "briefing_session_details",
    "briefing_venue",
    "briefing_compulsory",
    "conditions",
    "contact_person",
    "contact_email",
    "contact_telephone",
    "tender_validity_period",
]

HASH_VERSION = "etenders_hash_v2_structure_2026_07"

DEFAULT_OUTPUT_COLUMNS = [
    "source",
    "opportunity_key",
    "tender_number",
    "reference_number",
    "bid_number",
    "notice_type",
    "title",
    "description",
    "category",
    "client",
    "country_name",
    "province",
    "region",
    "status",
    "type",
    "conditions",
    "document_names",
    "document_urls",
    "document_count",
    "source_url",
    "action_url",
    "date_published",
    "closing_date",
    "closing_datetime",
    "briefing_session_details",
    "briefing_venue",
    "briefing_compulsory",
    "contact_person",
    "contact_email",
    "contact_telephone",
    "tender_validity_period",
    "priority_client_match",
    "matched_description_keywords",
    "description_keyword_match_count",
    "weight_score",
    "row_hash",
    "run_timestamp",
    "change_type",
    "previous_row_hash",
]

# Filter applied to the clean outputs used for Master, Delta and Updates.
# The raw files remain unfiltered for audit/debugging.
TARGET_CATEGORY = "Services: Professional"

# Scoring constants: edit these if you want a more/less aggressive priority score.
WEIGHT_CLIENT_MATCH = 10
WEIGHT_DESCRIPTION_KEYWORD = 1

# Top-delta shortlist defaults. Use the CLI switches below to override these.
TOP_DELTA_LIMIT_DEFAULT = 20
TOP_DELTA_MIN_SCORE_DEFAULT = 0

PRIORITY_CLIENTS = [
    "Umgeni Water",
    "Johannesburg Water",
    "Vaal Central Water",
    "Department of Public Works and Roads",
    "Development Bank of Southern Africa",
    "Transnet SOC Ltd",
    "Transnet National Ports Authority",
    "Health",
    "South African Nuclear Energy Corporation Limited",
    "National Nuclear Regulator",
    "Polokwane Local Municipality",
]

# Labels map to one or more search phrases. Alternative acronyms are included where useful.
DESCRIPTION_KEYWORD_RULES = {
    "Water and Sanitation": ["Water and Sanitation"],
    "Housing": ["Housing"],
    "Education": ["Education"],
    "Transport": ["Transport"],
    "Nuclear": ["Nuclear"],
    "Bus Rapid Transit (BRT)": ["Bus Rapid Transit", "BRT"],
    "International Transition Plan Network (ITPN)": [
        "International Transition Plan Network",
        "ITPN",
    ],
    "Design": ["Design"],
    "Refurbishment": ["Refurbishment"],
    "Engineering, Procurement, and Construction (EPC)": [
        "Engineering, Procurement, and Construction",
        "Engineering Procurement and Construction",
        "EPC",
    ],
    "Fire Engineering": ["Fire Engineering"],
    "HVAC": ["HVAC"],
    "Mechanical": ["Mechanical"],
    "Electrical": ["Electrical"],
    "Electronic": ["Electronic"],
    "Wet Services": ["Wet Services"],
    "Project Management": ["Project Management"],
    "Programme Management": ["Programme Management"],
    "Project Management Office (PMO)": ["Project Management Office", "PMO"],
    "Advisory Services": ["Advisory Services"],
    "Structural Engineering": ["Structural Engineering"],
    "Civil Engineering": ["Civil Engineering"],
    "Feasibility Study": ["Feasibility Study"],
}


def setup_logging():
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
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/json, text/javascript, */*; q=0.01",
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


def _normalize_bid(bid) -> str:
    """Normalize bid_number for stable comparison: strip whitespace + lowercase."""
    if bid is None:
        return ""
    return str(bid).strip().lower()


def _load_seen_ids(path: str) -> Set[str]:
    """Load seen bid_number IDs from a JSON state file."""
    try:
        if not os.path.exists(path):
            return set()
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return set(_normalize_bid(x) for x in (data or []))
    except (OSError, ValueError, TypeError) as e:
        logging.exception("Failed to load state file: %s -- %s", path, e)
        return set()


def _save_seen_ids(path: str, ids: Set[str]):
    """Persist sorted list of seen IDs to path (atomic write)."""
    try:
        norm = sorted(_normalize_bid(x) for x in ids)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(norm, fh, ensure_ascii=False)
        os.replace(tmp, path)
    except (OSError, TypeError) as e:
        logging.exception("Failed to save state file: %s -- %s", path, e)


def _load_hash_version(path: str) -> str:
    try:
        if not os.path.exists(path):
            return ""
        with open(path, "r", encoding="utf-8") as fh:
            return fh.read().strip()
    except OSError as e:
        logging.exception("Failed to load hash version file: %s -- %s", path, e)
        return ""


def _save_hash_version(path: str, version: str) -> None:
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            fh.write(version)
        os.replace(tmp, path)
    except OSError as e:
        logging.exception("Failed to save hash version file: %s -- %s", path, e)


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


def fetch_etenders(
    since_date: Optional[pd.Timestamp] = None,
    seen_ids: Optional[Set[str]] = None,
) -> List[Dict[str, Any]]:
    """Fetch currently advertised tenders.

    If since_date is provided, stop paging when rows are older than since_date
    (assumes API returns results ordered newest->oldest).

    If seen_ids is provided, stop paging when a tender_No already in seen_ids
    is encountered (stop-on-seen behaviour).
    """
    sess = build_session()

    r = sess.get(BASE_PAGE, timeout=TIMEOUT_S)
    r.raise_for_status()

    all_rows: List[Dict[str, Any]] = []
    start = 0
    total = None
    page_count = 0

    base_params = {
        "draw": 1,
        "length": PAGE_SIZE,
        "start": 0,
        "status": STATUS,
        "order[0][column]": ORDER_COL,
        "order[0][dir]": ORDER_DIR,
        "search[value]": "",
        "search[regex]": "false",
    }

    stop_paging = False

    while True:
        page_count += 1
        if page_count > MAX_PAGES:
            logging.warning("MAX_PAGES reached; stopping pagination as a safety brake.")
            break

        params = dict(base_params, start=start)
        r = sess.get(API_URL, params=params, timeout=TIMEOUT_S)
        r.raise_for_status()

        try:
            j = r.json()
        except ValueError:
            j = json.loads(r.text)

        rows = j.get("data") or j.get("aaData") or []
        if total is None:
            total = j.get("recordsFiltered") or j.get("recordsTotal")

        if not rows:
            break

        for row in rows:
            if seen_ids:
                bid = row.get("tender_No")
                if bid is not None and _normalize_bid(bid) in seen_ids:
                    stop_paging = True
                    break

            if since_date is not None:
                raw = row.get("date_Published")
                rd = pd.to_datetime(raw, errors="coerce")
                if pd.notna(rd) and rd < since_date:
                    stop_paging = True
                    break

            all_rows.append(row)

        start += PAGE_SIZE

        if stop_paging:
            logging.info("Stopping early due to since_date/seen_ids condition (page %d)", page_count)
            break

        if total and start >= total:
            break
        if len(rows) < PAGE_SIZE:
            break

        time.sleep(SLEEP_S)

    logging.info("Fetched rows: %d (pages: %d, total: %s)", len(all_rows), page_count, total)
    return all_rows


def drop_duplicate_sources(df: pd.DataFrame) -> pd.DataFrame:
    """Prefer *_name fields where the API includes both."""
    df = df.copy()
    if "categories_name" in df.columns and "category" in df.columns:
        df.drop(columns=["category"], inplace=True)
    if "provinces_name" in df.columns and "province" in df.columns:
        df.drop(columns=["province"], inplace=True)
    return df



def _iter_support_documents(value: Any) -> List[Dict[str, Any]]:
    """Return active eTenders supportDocument records as a list of dictionaries."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []

    docs = value
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped or stripped.lower() in {"nan", "none", "null"}:
            return []
        try:
            docs = json.loads(stripped)
        except json.JSONDecodeError:
            return []

    if isinstance(docs, dict):
        docs = [docs]
    if not isinstance(docs, list):
        return []

    out: List[Dict[str, Any]] = []
    for doc in docs:
        if not isinstance(doc, dict):
            continue
        # Treat missing active flag as active; exclude explicit false.
        if doc.get("active") is False:
            continue
        out.append(doc)
    return out


def _build_etenders_document_url(doc: Dict[str, Any]) -> str:
    """Build public eTenders download URL from supportDocument metadata."""
    support_id = str(doc.get("supportDocumentID") or "").strip()
    filename = str(doc.get("fileName") or "").strip()
    extension = str(doc.get("extension") or "").strip()

    if not support_id or not filename:
        return ""

    if extension and not extension.startswith("."):
        extension = f".{extension}"

    blob_name = support_id
    if extension and not blob_name.lower().endswith(extension.lower()):
        blob_name = f"{blob_name}{extension}"

    return (
        "https://www.etenders.gov.za/home/Download/"
        f"?blobName={quote(blob_name, safe='')}"
        f"&downloadedFileName={quote(filename, safe='')}"
    )


def add_etenders_document_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Extract tender document names/URLs from the eTenders supportDocument API field."""
    out = df.copy()
    if "supportDocument" not in out.columns:
        out["document_names"] = ""
        out["document_urls"] = ""
        out["document_count"] = 0
        return out

    def doc_names(value: Any) -> str:
        names = []
        for doc in _iter_support_documents(value):
            name = str(doc.get("fileName") or "").strip()
            if name and name not in names:
                names.append(name)
        return "; ".join(names)

    def doc_urls(value: Any) -> str:
        urls = []
        for doc in _iter_support_documents(value):
            url = _build_etenders_document_url(doc)
            if url and url not in urls:
                urls.append(url)
        return "; ".join(urls)

    out["document_names"] = out["supportDocument"].apply(doc_names)
    out["document_urls"] = out["supportDocument"].apply(doc_urls)
    out["document_count"] = out["document_urls"].apply(lambda x: 0 if not x else len(str(x).split("; ")))
    return out


def map_and_clean(df: pd.DataFrame) -> pd.DataFrame:
    """Map columns to stable schema, coerce dates, and de-duplicate by bid_number."""
    colmap = {
        "tender_No": "bid_number",
        "description": "description",
        "categories_name": "category",
        "organ_of_State": "client",
        "provinces_name": "province",
        "closing_Date": "closing_date",
        "date_Published": "date_published",
        "status": "status",
        "type": "type",
        "conditions": "conditions",
        "compulsory_briefing_session": "briefing_session_details",
        "briefingVenue": "briefing_venue",
        "contactPerson": "contact_person",
        "email": "contact_email",
        "telephone": "contact_telephone",
        "briefingCompulsory": "briefing_compulsory",
        "validity": "tender_validity_period",
        "document_names": "document_names",
        "document_urls": "document_urls",
        "document_count": "document_count",
    }

    present = {k: v for k, v in colmap.items() if k in df.columns}
    if not present:
        raise RuntimeError("No expected columns found for mapping. eTenders schema likely changed.")

    mapped = df.rename(columns=present)

    desired_order = [
        "bid_number",
        "description",
        "category",
        "client",
        "province",
        "status",
        "type",
        "conditions",
        "document_names",
        "document_urls",
        "document_count",
        "date_published",
        "closing_date",
        "briefing_session_details",
        "briefing_venue",
        "briefing_compulsory",
        "contact_person",
        "contact_email",
        "contact_telephone",
        "tender_validity_period",
    ]
    keep = [c for c in desired_order if c in mapped.columns]
    mapped = mapped[keep].copy()

    for c in ("closing_date", "date_published"):
        if c in mapped.columns:
            mapped[c] = pd.to_datetime(mapped[c], errors="coerce")

    if "bid_number" in mapped.columns:
        # Canonical: strip + lower (keeps Master stable)
        mapped["bid_number"] = mapped["bid_number"].fillna("").astype(str).str.strip().str.lower()
        mapped = mapped[mapped["bid_number"].ne("")]

        # Sort newest first so dedupe keeps newest record
        if "date_published" in mapped.columns:
            mapped = mapped.sort_values("date_published", ascending=False, na_position="last")

        mapped = mapped.drop_duplicates(subset=["bid_number"], keep="first")

    # Common source fields used by the consolidated TenderScrapers layer.
    if "bid_number" in mapped.columns:
        mapped.insert(0, "source", "eTenders")
        mapped.insert(1, "opportunity_key", "etenders:" + mapped["bid_number"].fillna("").astype(str).str.strip().str.lower())
        mapped["tender_number"] = mapped["bid_number"]
        mapped["reference_number"] = mapped["bid_number"]
    else:
        mapped.insert(0, "source", "eTenders")
        mapped.insert(1, "opportunity_key", "")
        mapped["tender_number"] = ""
        mapped["reference_number"] = ""

    mapped["notice_type"] = mapped.get("type", "Tender")
    mapped["title"] = mapped.get("description", "")
    mapped["country_name"] = "South Africa"
    mapped["region"] = mapped.get("province", "")
    mapped["source_url"] = BASE_PAGE

    if "closing_date" in mapped.columns:
        mapped["closing_datetime"] = mapped["closing_date"]
    else:
        mapped["closing_datetime"] = pd.NaT

    def _first_document_url(value: Any) -> str:
        text = str(value or "").strip()
        if not text or text.lower() in {"nan", "none", "null"}:
            return BASE_PAGE
        return text.split(";")[0].strip() or BASE_PAGE

    mapped["action_url"] = mapped["document_urls"].apply(_first_document_url) if "document_urls" in mapped.columns else BASE_PAGE

    preferred = [c for c in DEFAULT_OUTPUT_COLUMNS if c in mapped.columns]
    remaining = [c for c in mapped.columns if c not in preferred]
    return mapped[preferred + remaining]


def _normalise_exact_text(value: Any) -> str:
    """Normalise text for exact comparisons, preserving words but ignoring case/spacing."""
    if value is None or pd.isna(value):
        return ""
    return " ".join(str(value).casefold().strip().split())


def _normalise_search_text(value: Any) -> str:
    """Normalise text for phrase matching, ignoring case and punctuation differences."""
    if value is None or pd.isna(value):
        return ""
    cleaned = re.sub(r"[^a-z0-9]+", " ", str(value).casefold())
    return " ".join(cleaned.split())


def _contains_phrase(text_norm: str, phrase: str) -> bool:
    """Case/punctuation-insensitive phrase contains check with token boundaries."""
    phrase_norm = _normalise_search_text(phrase)
    if not phrase_norm:
        return False
    return f" {phrase_norm} " in f" {text_norm} "


def filter_target_category(df: pd.DataFrame) -> pd.DataFrame:
    """Keep only rows in the target category for clean outputs."""
    if "category" not in df.columns:
        raise RuntimeError("Expected 'category' column missing after mapping.")

    before = len(df)
    target = _normalise_exact_text(TARGET_CATEGORY)
    out = df[df["category"].apply(_normalise_exact_text).eq(target)].copy()
    logging.info(
        "Filtered clean output to category %r: %d -> %d rows",
        TARGET_CATEGORY,
        before,
        len(out),
    )
    return out


def add_priority_scoring(df: pd.DataFrame) -> pd.DataFrame:
    """Add client/description match details and a priority weight score."""
    out = df.copy()
    priority_clients_norm = {_normalise_exact_text(x) for x in PRIORITY_CLIENTS}

    if "client" not in out.columns:
        out["priority_client_match"] = False
    else:
        out["priority_client_match"] = out["client"].apply(
            lambda x: _normalise_exact_text(x) in priority_clients_norm
        )

    def matched_keywords(description: Any) -> str:
        desc_norm = _normalise_search_text(description)
        matches = []
        for label, phrases in DESCRIPTION_KEYWORD_RULES.items():
            if any(_contains_phrase(desc_norm, phrase) for phrase in phrases):
                matches.append(label)
        return "; ".join(matches)

    if "description" not in out.columns:
        out["matched_description_keywords"] = ""
    else:
        out["matched_description_keywords"] = out["description"].apply(matched_keywords)

    out["description_keyword_match_count"] = out["matched_description_keywords"].apply(
        lambda x: 0 if not x else len(str(x).split("; "))
    )
    out["weight_score"] = (
        out["priority_client_match"].astype(int) * WEIGHT_CLIENT_MATCH
        + out["description_keyword_match_count"].astype(int) * WEIGHT_DESCRIPTION_KEYWORD
    )

    # Put the most important records at the top of Delta/Updates/Master.
    sort_cols = [c for c in ["weight_score", "closing_date", "date_published"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

    return out


def build_top_delta(df: pd.DataFrame, limit: int, min_score: int) -> pd.DataFrame:
    """Return the highest-scoring new tenders from the current run's Delta list."""
    out = df.copy()

    if "weight_score" not in out.columns:
        out["weight_score"] = 0

    if min_score is not None:
        out = out[out["weight_score"].fillna(0).astype(int) >= int(min_score)].copy()

    sort_cols = [c for c in ["weight_score", "closing_date", "date_published"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

    if limit and int(limit) > 0:
        out = out.head(int(limit)).copy()

    return out


def compute_row_hash(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    """Stable SHA256 hash over selected fields (null-safe, order-stable)."""
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


def write_excel_with_table(df: pd.DataFrame, path: str, sheet_name: str, table_name: str):
    """Write XLSX and create an Excel Table when rows exist. Atomic replace.

    This function always creates/replaces the workbook, even when the dataframe
    is empty. Empty outputs are useful because they confirm the script ran and
    that there were simply no rows for that output on the current run.
    """
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp.xlsx"

    out_df = df.copy() if df is not None else pd.DataFrame()

    # Avoid creating a corrupt/unfinished workbook if an output has no columns.
    # This should be rare, but it makes the top-delta output robust when Delta is empty.
    if out_df.shape[1] == 0:
        out_df = pd.DataFrame(columns=DEFAULT_OUTPUT_COLUMNS)

    with pd.ExcelWriter(tmp, engine="xlsxwriter", datetime_format="yyyy-mm-dd", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        out_df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]

        nrows, ncols = out_df.shape
        columns = [{"header": col} for col in out_df.columns]

        # Excel tables require at least one data row in xlsxwriter.
        # If there are zero data rows, the workbook is still created with headers.
        if nrows > 0 and ncols > 0:
            last_row = nrows  # inclusive; header row is 0
            worksheet.add_table(
                0,
                0,
                last_row,
                ncols - 1,
                {"name": table_name, "columns": columns},
            )
        elif ncols > 0:
            worksheet.autofilter(0, 0, 0, ncols - 1)

        worksheet.freeze_panes(1, 0)

    os.replace(tmp, path)


def load_master(path: str) -> pd.DataFrame:
    """Load master Excel if it exists; else empty DF."""
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name="Master", engine="openpyxl")
    except Exception as e:
        logging.exception("Failed to read master file (%s): %s", path, e)
        return pd.DataFrame()


def append_and_dedupe_master(master: pd.DataFrame, incoming: pd.DataFrame) -> pd.DataFrame:
    """Append incoming rows to master and dedupe by bid_number.

    Incoming rows win ties, so newly fetched data backfills new columns such as
    conditions and weight_score even when the publication date is unchanged.
    """
    if master.empty:
        out = incoming.copy()
        out["_source_priority"] = 1
    else:
        master_part = master.copy()
        incoming_part = incoming.copy()
        master_part["_source_priority"] = 0
        incoming_part["_source_priority"] = 1
        out = pd.concat([master_part, incoming_part], ignore_index=True)

    if "bid_number" not in out.columns:
        return out.drop(columns=["_source_priority"], errors="ignore")

    out["bid_number"] = out["bid_number"].fillna("").astype(str).str.strip().str.lower()
    out = out[out["bid_number"].ne("")]

    if "date_published" in out.columns:
        out["date_published"] = pd.to_datetime(out["date_published"], errors="coerce")
        out = out.sort_values(
            ["date_published", "_source_priority"],
            ascending=[False, False],
            na_position="last",
            kind="mergesort",
        )
    else:
        out = out.sort_values("_source_priority", ascending=False, kind="mergesort")

    out = out.drop_duplicates(subset=["bid_number"], keep="first")
    out = out.drop(columns=["_source_priority"], errors="ignore")
    return out


def main():
    parser = argparse.ArgumentParser(description="Fetch and clean eTenders data (Delta + Updates + Master)")
    parser.add_argument("--since-date", help="stop paging when rows are older than this date (parseable by pandas)")
    parser.add_argument("--since-days", type=int, help="stop paging when rows are older than N days (shorthand). Recommended: 14.")
    parser.add_argument("--stop-on-seen", action="store_true",
                        help="stop paging early when a previously seen bid_number is encountered (uses --state-file)")
    parser.add_argument("--state-file", default=STATE_FILE_DEFAULT,
                        help="path to JSON file storing seen bid_number IDs (used for stop-on-seen optimisation)")
    parser.add_argument("--hash-version-file", default=HASH_VERSION_FILE_DEFAULT,
                        help="path to text file storing the eTenders hash version")
    parser.add_argument("--top-delta-limit", type=int, default=TOP_DELTA_LIMIT_DEFAULT,
                        help="maximum number of rows to write to etenders_top_delta.xlsx. Use 0 for all qualifying rows.")
    parser.add_argument("--top-delta-min-score", type=int, default=TOP_DELTA_MIN_SCORE_DEFAULT,
                        help="minimum weight_score required for etenders_top_delta.xlsx. Default includes score 0 rows.")
    args = parser.parse_args()

    setup_logging()
    for directory in [RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR]:
        os.makedirs(directory, exist_ok=True)
    logging.info("Using tender scrapers root: %s", TENDER_SCRAPERS_ROOT)
    logging.info("Using source root: %s", SOURCE_ROOT)
    logging.info("Using raw directory: %s", RAW_DIR)
    logging.info("Using output directory: %s", OUTPUT_DIR)
    logging.info("Using log directory: %s", LOG_DIR)
    logging.info("Using state directory: %s", STATE_DIR)

    run_ts = pd.Timestamp.now()

    try:
        acquire_lock(LOCK_FILE)

        # since_date (for paging early-stop)
        if args.since_date:
            since_date = pd.to_datetime(args.since_date)
        elif args.since_days is not None:
            since_date = pd.Timestamp.today().normalize() - pd.Timedelta(days=args.since_days)
        else:
            since_date = None

        # Optional early-stop state
        seen_ids = _load_seen_ids(args.state_file) if args.stop_on_seen else None

        rows = fetch_etenders(
            since_date=since_date,
            seen_ids=seen_ids if args.stop_on_seen else None,
        )
        if not rows:
            logging.warning("No rows returned from eTenders.")
            return 0

        df_full = pd.json_normalize(rows, sep="_")
        logging.info("Raw rows fetched: %d | cols: %d", len(df_full), len(df_full.columns))

        # Raw outputs (optional but useful for debugging)
        df_full.to_csv(RAW_CSV, index=False)
        logging.info("Saved raw CSV: %s", RAW_CSV)

        with pd.ExcelWriter(RAW_XLSX, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
            df_full.to_excel(writer, index=False, sheet_name="Raw_Data")
        logging.info("Saved raw Excel: %s", RAW_XLSX)

        # Clean pipeline
        df_for_clean = drop_duplicate_sources(df_full)
        df_for_clean = add_etenders_document_columns(df_for_clean)
        mapped = map_and_clean(df_for_clean)
        mapped = filter_target_category(mapped)
        mapped = add_priority_scoring(mapped)

        # Compute hash for change detection (operational fields only)
        mapped = mapped.copy()
        mapped["row_hash"] = compute_row_hash(mapped, HASH_COLS)
        mapped["run_timestamp"] = run_ts

        # Load master
        master_df = load_master(MASTER_XLSX)

        # Keep the master aligned to the same Services: Professional scope.
        if not master_df.empty and "category" in master_df.columns:
            master_df = filter_target_category(master_df)

        # Ensure master has canonical bid_number + row_hash (backfill if needed)
        if not master_df.empty and "bid_number" in master_df.columns:
            master_df = master_df.copy()
            master_df["bid_number"] = master_df["bid_number"].fillna("").astype(str).str.strip().str.lower()
            master_df = master_df[master_df["bid_number"].ne("")]

            # Backfill new columns introduced after earlier versions of the script.
            if "conditions" not in master_df.columns:
                master_df["conditions"] = ""
            if "priority_client_match" not in master_df.columns:
                master_df["priority_client_match"] = False
            if "matched_description_keywords" not in master_df.columns:
                master_df["matched_description_keywords"] = ""
            if "description_keyword_match_count" not in master_df.columns:
                master_df["description_keyword_match_count"] = 0
            if "weight_score" not in master_df.columns:
                master_df["weight_score"] = 0
            if "document_names" not in master_df.columns:
                master_df["document_names"] = ""
            if "document_urls" not in master_df.columns:
                master_df["document_urls"] = ""
            if "document_count" not in master_df.columns:
                master_df["document_count"] = 0

            if "row_hash" not in master_df.columns:
                # Backfill using the SAME HASH_COLS definition
                master_df["row_hash"] = compute_row_hash(master_df, HASH_COLS)

        # Migration/state baseline:
        # If this source has an existing master but no seen_ids.json yet, or if
        # the hash version changed, do not flood Delta/Updates on the first run.
        # The master will still be refreshed and the new state files will be saved.
        seen_ids_file_exists = os.path.exists(args.state_file)
        previous_hash_version = _load_hash_version(args.hash_version_file)
        suppress_alerts_for_baseline = (
            not master_df.empty
            and (
                not seen_ids_file_exists
                or previous_hash_version != HASH_VERSION
            )
        )

        # NEW vs UPDATED
        if suppress_alerts_for_baseline:
            logging.info(
                "eTenders hash/state version changed or was not recorded. "
                "Re-baselining and suppressing delta/update alerts for this run."
            )
            delta_df = mapped.iloc[0:0].copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""
            updates_df = mapped.iloc[0:0].copy()
            updates_df["change_type"] = "UPDATED"
            updates_df["previous_row_hash"] = ""
        elif not master_df.empty and "bid_number" in master_df.columns:
            master_ids = set(master_df["bid_number"].tolist())

            # bid -> last known hash
            master_hash_map = dict(
                zip(
                    master_df["bid_number"].tolist(),
                    master_df["row_hash"].fillna("").astype(str).tolist(),
                )
            )

            is_new = ~mapped["bid_number"].isin(master_ids)
            delta_df = mapped[is_new].copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""

            exists = ~is_new
            mapped_prev_hash = mapped["bid_number"].map(master_hash_map).fillna("")
            is_updated = exists & (mapped["row_hash"] != mapped_prev_hash)

            updates_df = mapped[is_updated].copy()
            updates_df["change_type"] = "UPDATED"
            updates_df["previous_row_hash"] = updates_df["bid_number"].map(master_hash_map).fillna("")

        else:
            # First run without a prior master: everything is new; nothing is "updated".
            delta_df = mapped.copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""

            updates_df = mapped.iloc[0:0].copy()
            updates_df["change_type"] = "UPDATED"
            updates_df["previous_row_hash"] = ""

        sort_cols = [c for c in ["weight_score", "closing_date", "date_published"] if c in delta_df.columns]
        if sort_cols:
            ascending = [False if c == "weight_score" else True for c in sort_cols]
            delta_df = delta_df.sort_values(sort_cols, ascending=ascending, na_position="last")
            updates_df = updates_df.sort_values(sort_cols, ascending=ascending, na_position="last")

        top_delta_df = build_top_delta(
            delta_df,
            limit=args.top_delta_limit,
            min_score=args.top_delta_min_score,
        )

        logging.info("Delta (NEW) rows this run after category filter: %d", len(delta_df))
        logging.info("Updates (CHANGED) rows this run after category filter: %d", len(updates_df))
        logging.info(
            "Top delta rows this run: %d | min_score: %d | limit: %d",
            len(top_delta_df),
            args.top_delta_min_score,
            args.top_delta_limit,
        )

        # Write delta (Power Automate ingests this)
        write_excel_with_table(
            delta_df,
            DELTA_XLSX,
            sheet_name="Delta",
            table_name="tbl_delta",
        )
        logging.info("Saved delta Excel: %s | rows: %d", DELTA_XLSX, len(delta_df))

        # Write top-scoring new tenders from Delta for daily review.
        write_excel_with_table(
            top_delta_df,
            TOP_DELTA_XLSX,
            sheet_name="Top_Delta",
            table_name="tbl_top_delta",
        )
        logging.info("Saved top delta Excel: %s | rows: %d", TOP_DELTA_XLSX, len(top_delta_df))

        # Write updates (Power Automate-friendly)
        write_excel_with_table(
            updates_df,
            UPDATES_XLSX,
            sheet_name="Updates",
            table_name="tbl_updates",
        )
        logging.info("Saved updates Excel: %s | rows: %d", UPDATES_XLSX, len(updates_df))

        # Update master with all currently fetched Services: Professional rows.
        # Delta/updates remain only the daily exception lists; master remains the latest snapshot.
        updated_master = append_and_dedupe_master(master_df, mapped)

        write_excel_with_table(
            updated_master,
            MASTER_XLSX,
            sheet_name="Master",
            table_name="tbl_master",
        )
        logging.info("Saved master Excel: %s | rows: %d", MASTER_XLSX, len(updated_master))

        # Update state from the MASTER (keeps state consistent even if delta empty)
        master_seen = set(updated_master["bid_number"].fillna("").astype(str).str.strip().str.lower().tolist())
        _save_seen_ids(args.state_file, master_seen)
        _save_hash_version(args.hash_version_file, HASH_VERSION)

        return 0

    except (requests.RequestException, ValueError, RuntimeError, OSError) as e:
        logging.exception("FAILED: %s", e)
        return 1

    finally:
        release_lock(LOCK_FILE)


if __name__ == "__main__":
    raise SystemExit(main())