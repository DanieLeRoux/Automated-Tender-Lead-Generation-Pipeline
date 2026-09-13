r'''
worldbank_procurement_pull_and_clean_v5 (Delta + Updates + Master + Top Delta)

Fetch World Bank Procurement Notices from the World Bank Finances One API,
clean/normalise them, and maintain:

1) MASTER Excel: latest snapshot per World Bank notice ID / opportunity key
2) DELTA Excel: brand new notices discovered this run only
3) UPDATES Excel: existing notices whose key operational fields changed
4) TOP DELTA Excel: highest-scoring new notices from the current run

Important implementation note:
The World Bank Finances One API documentation lists field filters, but in testing
those filters may be ignored for this dataset. To avoid pulling the full historical
archive by accident, the default mode fetches only the latest tail-window of records
and then filters locally for current/open notices.

Typical daily run:
    python worldbank_procurement_pull_and_clean_v4.py

Default current mode:
    - Fetch latest 20,000 records only
    - Keep notices with deadline_date from today to today + 180 days
    - Exclude expired notices
    - Keep only the default Henry Fagan country_name scope

Optional:
    python worldbank_procurement_pull_and_clean_v4.py --recent-window-records 30000
    python worldbank_procurement_pull_and_clean_v4.py --deadline-horizon-days 365
    python worldbank_procurement_pull_and_clean_v4.py --fetch-mode all   # full archive; use sparingly
'''

from __future__ import annotations

import os
import sys
import time
import json
import logging
import argparse
import hashlib
import re
from typing import Optional, List, Dict, Any, Tuple, Iterable, Set

import requests
import pandas as pd
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# pylint: disable=abstract-class-instantiated

# === World Bank API ===
API_URL = "https://datacatalogapi.worldbank.org/dexapps/fone/api/apiservice"
DATASET_ID = "DS00979"
RESOURCE_ID = "RS00909"

# ---- tunables ----
PAGE_SIZE = 1000  # World Bank API allows up to 1,000 records per request/page
SLEEP_S = 0.25
TIMEOUT_S = 60
DEFAULT_DEADLINE_HORIZON_DAYS = 180
DEFAULT_RECENT_WINDOW_RECORDS = 20000
MAX_ALL_PAGES = 300  # safety brake for full archive mode
# -------------------

# === Folder Structure ===
# This script is intended to live in:
#   02_TenderScrapers\02_Sources\WorldBank\00_Scripts
# Paths are detected dynamically from the script location.
SOURCE_FOLDER_NAME = "WorldBank"
HASH_VERSION = "worldbank_new_structure_v1"


def _detect_source_root() -> str:
    """Return the source root, normally the parent of 00_Scripts."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(script_dir).casefold() == "00_scripts":
        return os.path.abspath(os.path.join(script_dir, os.pardir))
    return os.path.abspath(script_dir)


def _detect_base_root(source_root: str) -> str:
    """Return the TenderScrapers root when the standard structure is used."""
    parent = os.path.dirname(os.path.abspath(source_root))
    if os.path.basename(parent).casefold() == "02_sources":
        return os.path.abspath(os.path.join(parent, os.pardir))
    return os.path.abspath(os.path.join(source_root, os.pardir))


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_ROOT = _detect_source_root()
BASE_ROOT = _detect_base_root(SOURCE_ROOT)

RAW_DIR = os.path.join(SOURCE_ROOT, "01_RawData")
OUTPUT_DIR = os.path.join(SOURCE_ROOT, "02_Outputs")
LOG_DIR = os.path.join(SOURCE_ROOT, "03_Logs")
STATE_DIR = os.path.join(SOURCE_ROOT, "04_State")
METADATA_DIR = os.path.join(SOURCE_ROOT, "05_Metadata")

LOG_PATH = os.path.join(LOG_DIR, "worldbank_run.log")

# Files
MASTER_XLSX = os.path.join(OUTPUT_DIR, "worldbank_master.xlsx")
DELTA_XLSX = os.path.join(OUTPUT_DIR, "worldbank_delta.xlsx")
UPDATES_XLSX = os.path.join(OUTPUT_DIR, "worldbank_updates.xlsx")
TOP_DELTA_XLSX = os.path.join(OUTPUT_DIR, "worldbank_top_delta.xlsx")

RAW_CSV = os.path.join(RAW_DIR, "worldbank_raw.csv")
RAW_XLSX = os.path.join(RAW_DIR, "worldbank_raw_full.xlsx")

STATE_FILE_DEFAULT = os.path.join(STATE_DIR, "seen_ids.json")
HASH_VERSION_FILE = os.path.join(STATE_DIR, "hash_version.txt")
TRANSLATION_CACHE_FILE = os.path.join(STATE_DIR, "translation_cache.json")
LOCK_FILE = os.path.join(STATE_DIR, "run.lock")


def recompute_paths(
    *,
    source_root: Optional[str] = None,
    base_root: Optional[str] = None,
    portal_subfolder: str = "",
) -> None:
    """Recompute global paths when CLI overrides are used.

    Normal use should not need overrides. The default is dynamic detection from
    the script location. --base-root and --portal-subfolder are retained only
    for backward compatibility with older launchers.
    """
    global SCRIPT_DIR, SOURCE_ROOT, BASE_ROOT, RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR
    global LOG_PATH, MASTER_XLSX, DELTA_XLSX, UPDATES_XLSX, TOP_DELTA_XLSX, RAW_CSV, RAW_XLSX
    global STATE_FILE_DEFAULT, HASH_VERSION_FILE, TRANSLATION_CACHE_FILE, LOCK_FILE

    if source_root:
        SOURCE_ROOT = os.path.abspath(source_root)
    elif base_root:
        SOURCE_ROOT = os.path.abspath(os.path.join(base_root, portal_subfolder)) if portal_subfolder else os.path.abspath(base_root)
    else:
        SOURCE_ROOT = _detect_source_root()

    BASE_ROOT = _detect_base_root(SOURCE_ROOT)
    SCRIPT_DIR = os.path.join(SOURCE_ROOT, "00_Scripts")

    RAW_DIR = os.path.join(SOURCE_ROOT, "01_RawData")
    OUTPUT_DIR = os.path.join(SOURCE_ROOT, "02_Outputs")
    LOG_DIR = os.path.join(SOURCE_ROOT, "03_Logs")
    STATE_DIR = os.path.join(SOURCE_ROOT, "04_State")
    METADATA_DIR = os.path.join(SOURCE_ROOT, "05_Metadata")

    LOG_PATH = os.path.join(LOG_DIR, "worldbank_run.log")

    MASTER_XLSX = os.path.join(OUTPUT_DIR, "worldbank_master.xlsx")
    DELTA_XLSX = os.path.join(OUTPUT_DIR, "worldbank_delta.xlsx")
    UPDATES_XLSX = os.path.join(OUTPUT_DIR, "worldbank_updates.xlsx")
    TOP_DELTA_XLSX = os.path.join(OUTPUT_DIR, "worldbank_top_delta.xlsx")

    RAW_CSV = os.path.join(RAW_DIR, "worldbank_raw.csv")
    RAW_XLSX = os.path.join(RAW_DIR, "worldbank_raw_full.xlsx")

    STATE_FILE_DEFAULT = os.path.join(STATE_DIR, "seen_ids.json")
    HASH_VERSION_FILE = os.path.join(STATE_DIR, "hash_version.txt")
    TRANSLATION_CACHE_FILE = os.path.join(STATE_DIR, "translation_cache.json")
    LOCK_FILE = os.path.join(STATE_DIR, "run.lock")

# Key operational fields that define "meaningful change" for updates.
HASH_COLS = [
    "bid_description",
    "deadline_date",
    "notice_type",
    "procurement_category",
    "procurement_method",
    "sector",
    "url",
]

# Weighting settings. Adjust here as you refine what Henry Fagan wants to prioritise.
PRIORITY_PROCUREMENT_CATEGORIES = [
    "consultant services",
    "consulting services",
]

DESCRIPTION_KEYWORDS = [
    "Water and Sanitation",
    "Housing",
    "Education",
    "Transport",
    "Nuclear",
    "Bus Rapid Transit",
    "BRT",
    "International Transition Plan Network",
    "ITPN",
    "Design",
    "Refurbishment",
    "Engineering",
    "Procurement",
    "Construction",
    "Engineering, Procurement, and Construction",
    "EPC",
    "Fire Engineering",
    "HVAC",
    "Mechanical",
    "Electrical",
    "Electronic",
    "Wet Services",
    "Project Management",
    "Programme Management",
    "Project Management Office",
    "PMO",
    "Advisory Services",
    "Structural Engineering",
    "Civil Engineering",
    "Feasibility Study",
]

CATEGORY_SCORE = 5
KEYWORD_SCORE = 1

# Default local country/region-style filter applied to the clean outputs.
# Raw files remain unfiltered for audit/debugging.
DEFAULT_TARGET_COUNTRY_NAMES = [
    "Eastern and Southern Africa",
    "Africa",
    "South Africa",
    "Western and Central Africa",
    "Central Africa",
    "Western Africa",
    "Eastern Africa",
    "United Kingdom",
    "Saudi Arabia",
    "Netherlands",
]
DEFAULT_COUNTRY_NAME_FILTER = "|".join(DEFAULT_TARGET_COUNTRY_NAMES)


def ensure_directories() -> None:
    """Create the standard source folders used by the new structure."""
    for folder in [RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR]:
        os.makedirs(folder, exist_ok=True)


def setup_logging():
    """Create log directory and configure root logger with file + stdout handlers."""
    ensure_directories()

    root = logging.getLogger()
    for handler in list(root.handlers):
        root.removeHandler(handler)

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
        allowed_methods=frozenset(["GET", "POST"]),
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


def _clean_col_name(col: str) -> str:
    """Normalise API column names defensively."""
    return re.sub(r"[^a-z0-9]+", "_", str(col).strip().lower()).strip("_")


def _safe_str(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _normalise_for_match(value: Any) -> str:
    """Lowercase and remove punctuation noise while keeping word boundaries."""
    text = _safe_str(value).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _parse_api_payload(payload: Any) -> Tuple[List[Dict[str, Any]], Optional[int]]:
    """Extract rows and count from World Bank API responses."""
    if isinstance(payload, dict):
        rows = payload.get("data") or payload.get("Data") or []
        count = payload.get("count") or payload.get("Count")
    elif isinstance(payload, list):
        rows = payload
        count = None
    else:
        rows = []
        count = None

    if isinstance(rows, dict):
        rows = [rows]
    if rows is None:
        rows = []

    return rows, count



def _load_json_dict(path: str) -> Dict[str, Any]:
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError, TypeError) as e:
        logging.warning("Could not read JSON state/cache file %s: %s", path, e)
        return {}


def _save_json_dict(path: str, data: Dict[str, Any]) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2, sort_keys=True)
    os.replace(tmp, path)


def _load_seen_ids(path: str) -> Set[str]:
    if not os.path.exists(path):
        return set()
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, list):
            return {str(x).strip().lower() for x in data if str(x).strip()}
        if isinstance(data, dict):
            return {str(x).strip().lower() for x in data.get("seen_ids", []) if str(x).strip()}
    except (OSError, ValueError, TypeError) as e:
        logging.warning("Could not read seen IDs file %s: %s", path, e)
    return set()


def _save_seen_ids(path: str, ids: Iterable[str]) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    norm = sorted({str(x).strip().lower() for x in ids if str(x).strip()})
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(norm, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _read_hash_version(path: str) -> str:
    try:
        if not os.path.exists(path):
            return ""
        with open(path, "r", encoding="utf-8") as fh:
            return fh.read().strip()
    except OSError:
        return ""


def _write_hash_version(path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        fh.write(HASH_VERSION + "\n")
    os.replace(tmp, path)


def is_migration_baseline_run() -> bool:
    """Suppress deltas/updates on the first successful run after restructuring."""
    return _read_hash_version(HASH_VERSION_FILE) != HASH_VERSION


def _translation_cache_key(provider: str, text: str) -> str:
    payload = f"{provider}|en|{text}".encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def _translate_libretranslate(
    sess: requests.Session,
    text: str,
    *,
    endpoint: str,
    api_key: Optional[str],
    timeout_s: int,
) -> str:
    url = endpoint.rstrip("/") + "/translate"
    payload: Dict[str, Any] = {
        "q": text,
        "source": "auto",
        "target": "en",
        "format": "text",
    }
    if api_key:
        payload["api_key"] = api_key

    response = sess.post(url, data=payload, timeout=timeout_s)
    response.raise_for_status()
    data = response.json()
    translated = data.get("translatedText") or data.get("translated_text") or ""
    return _safe_str(translated)


def apply_optional_translation(
    df: pd.DataFrame,
    *,
    enabled: bool,
    provider: str,
    endpoint: Optional[str],
    api_key: Optional[str],
    cache_path: str,
    timeout_s: int = 30,
) -> pd.DataFrame:
    """Optionally translate World Bank descriptions to English.

    Translation is deliberately opt-in so the normal scraper has no extra external
    dependency. Original source text is always preserved in bid_description.
    """
    out = df.copy()
    if "bid_description" not in out.columns:
        out["bid_description"] = ""

    out["bid_description_original"] = out["bid_description"].apply(_safe_str)
    out["bid_description_en"] = out["bid_description_original"]
    out["translation_status"] = "not_requested"
    out["translation_provider"] = ""

    if out.empty or not enabled:
        return out

    provider = (provider or "libretranslate").strip().lower()
    if provider != "libretranslate":
        logging.warning("Unsupported translation provider '%s'. Skipping translation.", provider)
        out["translation_status"] = "skipped_unsupported_provider"
        out["translation_provider"] = provider
        return out

    if not endpoint:
        logging.warning("Translation requested but no LibreTranslate endpoint was supplied. Skipping translation.")
        out["translation_status"] = "skipped_no_endpoint"
        out["translation_provider"] = provider
        return out

    sess = build_session()
    cache = _load_json_dict(cache_path)
    cache_changed = False

    translated_values: List[str] = []
    statuses: List[str] = []

    for text in out["bid_description_original"].tolist():
        text = _safe_str(text)
        if not text:
            translated_values.append("")
            statuses.append("blank")
            continue

        key = _translation_cache_key(provider, text)
        cached = cache.get(key)
        if isinstance(cached, dict) and cached.get("translated_text"):
            translated_values.append(_safe_str(cached.get("translated_text")))
            statuses.append(_safe_str(cached.get("status") or "cached"))
            continue

        try:
            translated = _translate_libretranslate(
                sess,
                text,
                endpoint=endpoint,
                api_key=api_key,
                timeout_s=timeout_s,
            )
            if not translated:
                translated = text
                status = "empty_response_used_original"
            elif translated.strip() == text.strip():
                status = "same_or_english"
            else:
                status = "translated"

            cache[key] = {"translated_text": translated, "status": status}
            cache_changed = True
            translated_values.append(translated)
            statuses.append(status)
            time.sleep(SLEEP_S)
        except (requests.RequestException, ValueError, RuntimeError) as e:
            logging.warning("Translation failed for one notice; using original text. Error: %s", e)
            translated_values.append(text)
            statuses.append("failed_used_original")

    out["bid_description_en"] = translated_values
    out["translation_status"] = statuses
    out["translation_provider"] = provider

    if cache_changed:
        _save_json_dict(cache_path, cache)

    return out


def fetch_worldbank_page(sess: requests.Session, *, skip: int, top: int) -> Tuple[List[Dict[str, Any]], Optional[int]]:
    """Fetch one unfiltered API page.

    The Finances One API appears to ignore some field filters for this dataset.
    We therefore fetch a controlled tail-window and apply filters locally.
    """
    params: Dict[str, Any] = {
        "datasetId": DATASET_ID,
        "resourceId": RESOURCE_ID,
        "top": top,
        "skip": skip,
        "type": "json",
    }

    r = sess.get(API_URL, params=params, timeout=TIMEOUT_S)
    r.raise_for_status()

    try:
        payload = r.json()
    except ValueError:
        payload = json.loads(r.text)

    return _parse_api_payload(payload)


def _dedupe_raw_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Deduplicate raw API rows using ID first, then a stable composite fallback."""
    seen = set()
    out: List[Dict[str, Any]] = []
    for row in rows:
        key = str(row.get("id") or row.get("ID") or "").strip()
        if not key:
            key = "|".join(
                [
                    str(row.get("project_id") or row.get("Project ID") or "").strip(),
                    str(row.get("bid_description") or row.get("Bid Description") or "").strip(),
                    str(row.get("publication_date") or row.get("Publication Date") or "").strip(),
                    str(row.get("deadline_date") or row.get("Deadline Date") or "").strip(),
                ]
            ).lower()
        if key and key not in seen:
            seen.add(key)
            out.append(row)
    return out


def fetch_tail_window(sess: requests.Session, *, recent_window_records: int) -> List[Dict[str, Any]]:
    """Fetch only the latest N records from the dataset.

    The API pages from older to newer rows. The current/open notices are near the end of
    the dataset, so this avoids pulling the full historical archive every day.
    """
    first_rows, total = fetch_worldbank_page(sess, skip=0, top=1)
    if total is None:
        raise RuntimeError("World Bank API did not return a total count; cannot safely calculate tail-window.")

    total_count = int(total)
    window = max(int(recent_window_records), PAGE_SIZE)
    start_skip = max(total_count - window, 0)
    # Align to page boundary for cleaner paging.
    start_skip = (start_skip // PAGE_SIZE) * PAGE_SIZE

    logging.info(
        "World Bank dataset count: %d | fetching latest window from skip=%d to end (window requested: %d records)",
        total_count,
        start_skip,
        window,
    )

    all_rows: List[Dict[str, Any]] = []
    skip = start_skip
    page = 0
    while skip < total_count:
        page += 1
        rows, _ = fetch_worldbank_page(sess, skip=skip, top=PAGE_SIZE)
        if not rows:
            break
        all_rows.extend(rows)
        logging.info("Fetched tail page %d | skip=%d | rows=%d | total so far=%d", page, skip, len(rows), len(all_rows))
        if len(rows) < PAGE_SIZE:
            break
        skip += PAGE_SIZE
        time.sleep(SLEEP_S)

    return _dedupe_raw_rows(all_rows)


def fetch_all_worldbank(sess: requests.Session) -> List[Dict[str, Any]]:
    """Fetch the full archive. Use sparingly."""
    all_rows: List[Dict[str, Any]] = []
    skip = 0
    page = 0
    total = None

    while True:
        page += 1
        if page > MAX_ALL_PAGES:
            logging.warning("Full archive safety brake reached at %d pages; stopping.", MAX_ALL_PAGES)
            break

        rows, count = fetch_worldbank_page(sess, skip=skip, top=PAGE_SIZE)
        if total is None:
            total = int(count) if count is not None else None
        if not rows:
            break

        all_rows.extend(rows)
        logging.info("Fetched archive page %d | skip=%d | rows=%d | total so far=%d", page, skip, len(rows), len(all_rows))

        if len(rows) < PAGE_SIZE:
            break
        skip += PAGE_SIZE
        if total and skip >= total:
            break
        time.sleep(SLEEP_S)

    return _dedupe_raw_rows(all_rows)


def fetch_worldbank(*, fetch_mode: str, recent_window_records: int) -> List[Dict[str, Any]]:
    """Fetch World Bank procurement notices.

    current/recent mode: controlled tail-window only.
    all mode: full archive.
    """
    sess = build_session()
    fetch_mode = (fetch_mode or "current").strip().lower()
    if fetch_mode not in {"current", "recent", "all"}:
        raise ValueError("fetch_mode must be one of: current, recent, all")

    if fetch_mode == "all":
        logging.info("Fetching full World Bank archive. This can be slow; use sparingly.")
        rows = fetch_all_worldbank(sess)
    else:
        logging.info("Fetching World Bank latest tail-window only; local filters will determine %s notices.", fetch_mode)
        rows = fetch_tail_window(sess, recent_window_records=recent_window_records)

    logging.info("Fetched World Bank rows after raw de-duplication: %d", len(rows))
    return rows


def empty_mapped_frame() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "source",
            "opportunity_key",
            "notice_id",
            "project_id",
            "bid_description",
            "bid_description_original",
            "bid_description_en",
            "translation_status",
            "translation_provider",
            "country_code",
            "country_name",
            "region",
            "sector",
            "notice_type",
            "procurement_category",
            "procurement_method",
            "publication_date",
            "deadline_date",
            "url",
        ]
    )


def _apply_text_filter(df: pd.DataFrame, column: str, filter_value: Optional[str]) -> pd.DataFrame:
    """Apply case-insensitive exact/multiple-value filter to a mapped dataframe."""
    if not filter_value or column not in df.columns or df.empty:
        return df
    allowed = {_normalise_for_match(x) for x in str(filter_value).split("|") if str(x).strip()}
    if not allowed:
        return df
    return df[df[column].apply(_normalise_for_match).isin(allowed)].copy()


def map_and_clean(
    df: pd.DataFrame,
    *,
    fetch_mode: str,
    since_date: Optional[pd.Timestamp],
    deadline_horizon_days: int,
    include_expired: bool,
    country_name: Optional[str],
    region: Optional[str],
    procurement_category: Optional[str],
) -> pd.DataFrame:
    """Map World Bank API fields to a stable schema and apply local filters."""
    if df.empty:
        return empty_mapped_frame()

    mapped = df.copy()
    mapped.columns = [_clean_col_name(c) for c in mapped.columns]

    # Defensive aliases in case the API returns slightly different labels.
    aliases = {
        "bid_description": ["bid_description", "description"],
        "country_code": ["country_code"],
        "country_name": ["country_name", "country"],
        "deadline_date": ["deadline_date", "closing_date"],
        "id": ["id", "notice_id"],
        "notice_type": ["notice_type"],
        "procurement_category": ["procurement_category"],
        "procurement_method": ["procurement_method"],
        "project_id": ["project_id"],
        "publication_date": ["publication_date", "published_date"],
        "region": ["region"],
        "sector": ["sector"],
        "url": ["url"],
    }

    out = pd.DataFrame(index=mapped.index)
    for target, options in aliases.items():
        src = next((c for c in options if c in mapped.columns), None)
        out[target] = mapped[src] if src else ""

    out["source"] = "World Bank"
    out["notice_id"] = out["id"].apply(_safe_str)
    out.drop(columns=["id"], inplace=True)

    for c in ["publication_date", "deadline_date"]:
        out[c] = pd.to_datetime(out[c], errors="coerce", dayfirst=True)

    # Opportunity key: prefer official ID; fallback to a stable composite.
    fallback = (
        out["project_id"].apply(_safe_str)
        + "|"
        + out["bid_description"].apply(_safe_str)
        + "|"
        + out["publication_date"].dt.strftime("%Y-%m-%d").fillna("")
        + "|"
        + out["deadline_date"].dt.strftime("%Y-%m-%d").fillna("")
    )
    out["opportunity_key"] = out["notice_id"].where(out["notice_id"].ne(""), fallback)
    out["opportunity_key"] = "worldbank:" + out["opportunity_key"].astype(str).str.strip().str.lower()

    # Local filters because API field filters can be ignored.
    out = _apply_text_filter(out, "country_name", country_name)
    out = _apply_text_filter(out, "region", region)
    out = _apply_text_filter(out, "procurement_category", procurement_category)

    fetch_mode = (fetch_mode or "current").strip().lower()
    today = pd.Timestamp.today().normalize()

    if fetch_mode == "current":
        horizon = today + pd.Timedelta(days=max(int(deadline_horizon_days), 0))
        # Current opportunities require a deadline that has not passed.
        out = out[out["deadline_date"].notna()].copy()
        out = out[(out["deadline_date"] >= today) & (out["deadline_date"] <= horizon)].copy()
    elif fetch_mode == "recent" and since_date is not None:
        out = out[out["publication_date"].notna() & (out["publication_date"] >= pd.Timestamp(since_date).normalize())].copy()
        if not include_expired:
            out = out[out["deadline_date"].notna() & (out["deadline_date"] >= today)].copy()
    elif not include_expired:
        out = out[out["deadline_date"].notna() & (out["deadline_date"] >= today)].copy()

    # Stable ordering and de-duplication.
    if not out.empty:
        out = out.sort_values(["publication_date", "deadline_date"], ascending=[False, True], na_position="last")
        out = out.drop_duplicates(subset=["opportunity_key"], keep="first")

    desired_order = [
        "source",
        "opportunity_key",
        "notice_id",
        "project_id",
        "bid_description",
        "country_code",
        "country_name",
        "region",
        "sector",
        "notice_type",
        "procurement_category",
        "procurement_method",
        "publication_date",
        "deadline_date",
        "url",
    ]
    return out[desired_order].copy()


def _keyword_matches(text: str, keywords: List[str]) -> List[str]:
    norm = _normalise_for_match(text)
    matches = []
    for kw in keywords:
        kw_norm = _normalise_for_match(kw)
        if not kw_norm:
            continue
        if re.search(rf"\b{re.escape(kw_norm)}\b", norm):
            matches.append(kw)
    return matches


def apply_weight_score(df: pd.DataFrame) -> pd.DataFrame:
    """Apply weight score to prioritise relevant World Bank notices."""
    df = df.copy()
    if df.empty:
        df["priority_procurement_category_match"] = []
        df["matched_keywords"] = []
        df["keyword_match_count"] = []
        df["weight_score"] = []
        return df

    category_matches = []
    matched_keywords_col = []
    keyword_counts = []
    scores = []

    priority_categories_norm = {_normalise_for_match(x) for x in PRIORITY_PROCUREMENT_CATEGORIES}

    for _, row in df.iterrows():
        score = 0
        proc_cat = _normalise_for_match(row.get("procurement_category", ""))
        category_match = proc_cat in priority_categories_norm or any(x in proc_cat for x in priority_categories_norm)
        if category_match:
            score += CATEGORY_SCORE

        combined_text = " ".join(
            [
                _safe_str(row.get("bid_description_en", "") or row.get("bid_description", "")),
                _safe_str(row.get("sector", "")),
                _safe_str(row.get("procurement_method", "")),
                _safe_str(row.get("procurement_category", "")),
            ]
        )
        matches = _keyword_matches(combined_text, DESCRIPTION_KEYWORDS)
        score += len(matches) * KEYWORD_SCORE

        category_matches.append("Yes" if category_match else "No")
        matched_keywords_col.append(", ".join(matches))
        keyword_counts.append(len(matches))
        scores.append(score)

    df["priority_procurement_category_match"] = category_matches
    df["matched_keywords"] = matched_keywords_col
    df["keyword_match_count"] = keyword_counts
    df["weight_score"] = scores
    return df


def compute_row_hash(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    """Stable SHA256 hash over selected fields."""
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
    """Write XLSX and create an Excel Table. Always creates a workbook, even if empty."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp.xlsx"

    out_df = df.copy()
    if out_df.shape[1] == 0:
        out_df = pd.DataFrame({"message": []})

    with pd.ExcelWriter(
        tmp,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        out_df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]

        nrows, ncols = out_df.shape
        columns = [{"header": str(col)} for col in out_df.columns]

        # Excel table range is inclusive. If no data rows, create one blank row beneath the header.
        last_row = max(nrows, 1)
        worksheet.add_table(
            0,
            0,
            last_row,
            ncols - 1,
            {"name": table_name, "columns": columns},
        )
        worksheet.freeze_panes(1, 0)

        for idx, col in enumerate(out_df.columns):
            width = min(max(len(str(col)) + 2, 14), 45)
            worksheet.set_column(idx, idx, width)

    os.replace(tmp, path)


def load_master(path: str) -> pd.DataFrame:
    """Load master Excel if it exists; refuse to overwrite if it cannot be read."""
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name="Master", engine="openpyxl")
    except Exception as e:  # pylint: disable=broad-exception-caught
        raise RuntimeError(f"Existing master could not be read; refusing to overwrite: {path}") from e


def append_and_dedupe_master(master: pd.DataFrame, incoming: pd.DataFrame) -> pd.DataFrame:
    """Append incoming rows to master and dedupe by opportunity_key.

    Current-run rows deliberately win over older master rows when keys collide.
    """
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

    out["opportunity_key"] = out["opportunity_key"].fillna("").astype(str).str.strip().str.lower()
    out = out[out["opportunity_key"].ne("")]

    if "publication_date" in out.columns:
        out["publication_date"] = pd.to_datetime(out["publication_date"], errors="coerce", dayfirst=True)
    if "deadline_date" in out.columns:
        out["deadline_date"] = pd.to_datetime(out["deadline_date"], errors="coerce", dayfirst=True)

    sort_cols = ["_source_priority"]
    ascending = [False]
    if "publication_date" in out.columns:
        sort_cols.append("publication_date")
        ascending.append(False)
    if "deadline_date" in out.columns:
        sort_cols.append("deadline_date")
        ascending.append(True)

    out = out.sort_values(sort_cols, ascending=ascending, na_position="last", kind="mergesort")
    out = out.drop_duplicates(subset=["opportunity_key"], keep="first")
    out = out.drop(columns=["_source_priority"], errors="ignore")

    display_cols = [c for c in ["publication_date", "deadline_date"] if c in out.columns]
    if display_cols:
        display_ascending = [False if c == "publication_date" else True for c in display_cols]
        out = out.sort_values(display_cols, ascending=display_ascending, na_position="last")

    return out


def build_top_delta(delta_df: pd.DataFrame, *, limit: int, min_score: int) -> pd.DataFrame:
    """Create top-scoring delta list."""
    if delta_df.empty:
        return delta_df.copy()

    top = delta_df.copy()
    if "weight_score" in top.columns:
        top = top[top["weight_score"].fillna(0).astype(int) >= min_score]
        sort_cols = [c for c in ["weight_score", "deadline_date", "publication_date"] if c in top.columns]
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        if sort_cols:
            top = top.sort_values(sort_cols, ascending=ascending, na_position="last")

    if limit and limit > 0:
        top = top.head(limit)

    return top


def main():
    parser = argparse.ArgumentParser(description="Fetch and clean World Bank Procurement Notices")
    parser.add_argument("--since-date", help="for --fetch-mode recent: keep notices published from this date onwards")
    parser.add_argument("--since-days", type=int, default=14,
                        help="for --fetch-mode recent: keep notices published in the last N days. Default: 14")
    parser.add_argument("--fetch-mode", choices=["current", "recent", "all"], default="current",
                        help="current = open notices by deadline_date; recent = recent publications; all = full archive. Default: current")
    parser.add_argument("--deadline-horizon-days", type=int, default=DEFAULT_DEADLINE_HORIZON_DAYS,
                        help="for --fetch-mode current: keep deadlines from today to today + N days. Default: 180")
    parser.add_argument("--recent-window-records", type=int, default=DEFAULT_RECENT_WINDOW_RECORDS,
                        help="for current/recent modes: fetch latest N records before local filtering. Default: 20000")
    parser.add_argument("--fetch-all", action="store_true",
                        help="legacy shortcut for --fetch-mode all. Use sparingly.")
    parser.add_argument("--include-expired", action="store_true",
                        help="include notices whose deadline_date has passed; mainly useful with --fetch-mode all")
    parser.add_argument(
        "--country-name",
        default=DEFAULT_COUNTRY_NAME_FILTER,
        help=(
            "local exact World Bank country_name filter; use | for multiple values. "
            "Default keeps Henry Fagan target countries/regions only."
        ),
    )
    parser.add_argument("--region", help="optional local exact World Bank region filter; use | for multiple values")
    parser.add_argument("--procurement-category", help="optional local exact procurement category filter; use | for multiple values")
    parser.add_argument("--top-delta-limit", type=int, default=20,
                        help="number of rows to keep in top delta file. Use 0 for all. Default: 20")
    parser.add_argument("--top-delta-min-score", type=int, default=0,
                        help="minimum weight_score for top delta file. Default: 0")
    parser.add_argument("--source-root", help="optional override for the WorldBank source root folder")
    parser.add_argument("--base-root", help="deprecated compatibility override; treated as source root unless --portal-subfolder is used")
    parser.add_argument("--portal-subfolder", default="", help="deprecated compatibility subfolder used with --base-root")
    parser.add_argument("--state-file", default=None, help="optional override for seen_ids.json path")
    parser.add_argument("--translate-non-english", action="store_true", help="translate bid descriptions to English using the selected provider")
    parser.add_argument("--translation-provider", default="libretranslate", choices=["libretranslate"], help="translation provider used when --translate-non-english is set")
    parser.add_argument("--translation-endpoint", default=os.getenv("LIBRETRANSLATE_URL"), help="LibreTranslate endpoint, e.g. http://localhost:5000")
    parser.add_argument("--translation-api-key", default=os.getenv("LIBRETRANSLATE_API_KEY"), help="optional LibreTranslate API key")
    args = parser.parse_args()

    recompute_paths(source_root=args.source_root, base_root=args.base_root, portal_subfolder=args.portal_subfolder)
    setup_logging()
    state_file = args.state_file or STATE_FILE_DEFAULT
    logging.info("Using base root: %s", BASE_ROOT)
    logging.info("Using source root: %s", SOURCE_ROOT)
    logging.info("Using raw directory: %s", RAW_DIR)
    logging.info("Using output directory: %s", OUTPUT_DIR)
    logging.info("Using country_name filter: %s", args.country_name or "<none>")

    run_ts = pd.Timestamp.now()

    try:
        acquire_lock(LOCK_FILE)

        if args.since_date:
            since_date = pd.to_datetime(args.since_date, dayfirst=True).normalize()
        elif args.since_days is not None:
            since_date = pd.Timestamp.today().normalize() - pd.Timedelta(days=args.since_days)
        else:
            since_date = None

        fetch_mode = "all" if args.fetch_all else args.fetch_mode

        rows = fetch_worldbank(
            fetch_mode=fetch_mode,
            recent_window_records=args.recent_window_records,
        )

        df_full = pd.json_normalize(rows, sep="_") if rows else pd.DataFrame()
        logging.info("Raw rows fetched: %d | cols: %d", len(df_full), len(df_full.columns))

        if df_full.empty:
            raise RuntimeError("World Bank API returned zero raw rows; refusing to overwrite existing outputs/master.")

        # Raw outputs before active filter/scoring. Useful for audit/debugging.
        df_full.to_csv(RAW_CSV, index=False)
        logging.info("Saved raw CSV: %s", RAW_CSV)

        with pd.ExcelWriter(
            RAW_XLSX,
            engine="xlsxwriter",
            engine_kwargs={"options": {"strings_to_urls": False}},
        ) as writer:
            df_full.to_excel(writer, index=False, sheet_name="Raw_Data")
        logging.info("Saved raw Excel: %s", RAW_XLSX)

        mapped = map_and_clean(
            df_full,
            fetch_mode=fetch_mode,
            since_date=since_date,
            deadline_horizon_days=args.deadline_horizon_days,
            include_expired=args.include_expired,
            country_name=args.country_name,
            region=args.region,
            procurement_category=args.procurement_category,
        )
        logging.info("Rows after local filters: %d", len(mapped))

        mapped = apply_optional_translation(
            mapped,
            enabled=args.translate_non_english,
            provider=args.translation_provider,
            endpoint=args.translation_endpoint,
            api_key=args.translation_api_key,
            cache_path=TRANSLATION_CACHE_FILE,
        )

        mapped = apply_weight_score(mapped)
        mapped["row_hash"] = compute_row_hash(mapped, HASH_COLS)
        mapped["run_timestamp"] = run_ts

        master_df = load_master(MASTER_XLSX)

        # Keep existing master aligned to the same country_name scope.
        # This prevents older out-of-scope records from being retained after the filter changes.
        if not master_df.empty:
            master_df = _apply_text_filter(master_df, "country_name", args.country_name)

        if not master_df.empty and "opportunity_key" in master_df.columns:
            master_df = master_df.copy()
            master_df["opportunity_key"] = master_df["opportunity_key"].fillna("").astype(str).str.strip().str.lower()
            master_df = master_df[master_df["opportunity_key"].ne("")]
            if "row_hash" not in master_df.columns or master_df["row_hash"].fillna("").eq("").all():
                master_df["row_hash"] = compute_row_hash(master_df, HASH_COLS)

        migration_baseline = is_migration_baseline_run()
        if migration_baseline:
            logging.info(
                "Migration baseline active: first run for hash version %s. Delta/update/top-delta alerts will be suppressed.",
                HASH_VERSION,
            )

        if migration_baseline:
            delta_df = mapped.iloc[0:0].copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""

            updates_df = mapped.iloc[0:0].copy()
            updates_df["change_type"] = "UPDATED"
            updates_df["previous_row_hash"] = ""
        elif not master_df.empty and "opportunity_key" in master_df.columns:
            master_keys = set(master_df["opportunity_key"].tolist())
            master_hash_map = dict(
                zip(
                    master_df["opportunity_key"].tolist(),
                    master_df["row_hash"].fillna("").astype(str).tolist(),
                )
            )

            is_new = ~mapped["opportunity_key"].isin(master_keys) if not mapped.empty else pd.Series([], dtype=bool)
            delta_df = mapped[is_new].copy() if not mapped.empty else mapped.copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""

            exists = ~is_new if not mapped.empty else pd.Series([], dtype=bool)
            mapped_prev_hash = mapped["opportunity_key"].map(master_hash_map).fillna("") if not mapped.empty else pd.Series([], dtype="object")
            is_updated = exists & (mapped["row_hash"] != mapped_prev_hash) if not mapped.empty else pd.Series([], dtype=bool)

            updates_df = mapped[is_updated].copy() if not mapped.empty else mapped.copy()
            updates_df["change_type"] = "UPDATED"
            if not updates_df.empty:
                updates_df["previous_row_hash"] = updates_df["opportunity_key"].map(master_hash_map).fillna("")
            else:
                updates_df["previous_row_hash"] = ""
        else:
            delta_df = mapped.copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""

            updates_df = mapped.iloc[0:0].copy()
            updates_df["change_type"] = "UPDATED"
            updates_df["previous_row_hash"] = ""

        top_delta_df = build_top_delta(
            delta_df,
            limit=args.top_delta_limit,
            min_score=args.top_delta_min_score,
        )

        logging.info("Delta (NEW) rows this run: %d", len(delta_df))
        logging.info("Updates (CHANGED) rows this run: %d", len(updates_df))
        logging.info("Top delta rows this run: %d", len(top_delta_df))

        write_excel_with_table(delta_df, DELTA_XLSX, sheet_name="Delta", table_name="tbl_wb_delta")
        logging.info("Saved delta Excel: %s | rows: %d", DELTA_XLSX, len(delta_df))

        write_excel_with_table(updates_df, UPDATES_XLSX, sheet_name="Updates", table_name="tbl_wb_updates")
        logging.info("Saved updates Excel: %s | rows: %d", UPDATES_XLSX, len(updates_df))

        write_excel_with_table(top_delta_df, TOP_DELTA_XLSX, sheet_name="Top_Delta", table_name="tbl_wb_top_delta")
        logging.info("Saved top delta Excel: %s | rows: %d", TOP_DELTA_XLSX, len(top_delta_df))

        # Update master with all currently fetched/cleaned rows so new calculated columns backfill.
        # If filters return no mapped rows while an existing master is present, do not replace it with an empty workbook.
        if mapped.empty and not master_df.empty:
            updated_master = master_df.copy()
            logging.warning("Mapped World Bank output is empty; preserving existing master instead of overwriting it.")
        else:
            updated_master = append_and_dedupe_master(master_df, mapped)
            write_excel_with_table(updated_master, MASTER_XLSX, sheet_name="Master", table_name="tbl_wb_master")
            logging.info("Saved master Excel: %s | rows: %d", MASTER_XLSX, len(updated_master))

        if "opportunity_key" in updated_master.columns:
            _save_seen_ids(state_file, updated_master["opportunity_key"].tolist())
        _write_hash_version(HASH_VERSION_FILE)

        return 0

    except (requests.RequestException, ValueError, RuntimeError, OSError) as e:
        logging.exception("FAILED: %s", e)
        return 1

    finally:
        release_lock(LOCK_FILE)


if __name__ == "__main__":
    raise SystemExit(main())
