"""
coct_tender_pull_and_clean (Delta + Updates + Master)

Fetch current tender notices from the City of Cape Town Tender Portal,
clean/normalise them, score them, and maintain:

1) MASTER Excel: latest snapshot per tender_number / opportunity_key
2) DELTA Excel: brand new tenders discovered this run only
3) UPDATES Excel: existing tenders whose key fields changed
4) TOP DELTA Excel: highest-scoring new tenders from the current run

The City of Cape Town public portal exposes the tender notice list in HTML.
Additional detail requires supplier registration/login, so this scraper only
uses the public notice table and captures the detail URL where available.

Typical daily run:
    python coct_tender_pull_and_clean.py

Optional:
    python coct_tender_pull_and_clean.py --include-expired
    python coct_tender_pull_and_clean.py --top-delta-limit 50
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import logging
import os
import re
import sys
from pathlib import Path
from html.parser import HTMLParser
from typing import Any, Dict, List, Optional, Set
from urllib.parse import urljoin

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# pylint: disable=abstract-class-instantiated

SOURCE_NAME = "City of Cape Town"
CLIENT_NAME = "City of Cape Town"
PROVINCE_NAME = "Western Cape"
PORTAL_URL = "https://web1.capetown.gov.za/web1/tenderportal/Tender"
BASE_URL = "https://web1.capetown.gov.za"

# === Folder structure ===
# Expected location:
#   02_TenderScrapers\02_Sources\CityOfCapeTown\00_Scripts\coct_tender_pull_and_clean.py
SCRIPT_PATH = Path(__file__).resolve()
SOURCE_ROOT = SCRIPT_PATH.parents[1]
TENDER_SCRAPERS_ROOT = SCRIPT_PATH.parents[3]

RAW_DIR = os.fspath(SOURCE_ROOT / "01_RawData")
OUTPUT_DIR = os.fspath(SOURCE_ROOT / "02_Outputs")
LOG_DIR = os.fspath(SOURCE_ROOT / "03_Logs")
STATE_DIR = os.fspath(SOURCE_ROOT / "04_State")
METADATA_DIR = os.fspath(SOURCE_ROOT / "05_Metadata")

LOG_PATH = os.path.join(LOG_DIR, "coct_run.log")

MASTER_XLSX = os.path.join(OUTPUT_DIR, "coct_master.xlsx")
DELTA_XLSX = os.path.join(OUTPUT_DIR, "coct_delta.xlsx")
UPDATES_XLSX = os.path.join(OUTPUT_DIR, "coct_updates.xlsx")
TOP_DELTA_XLSX = os.path.join(OUTPUT_DIR, "coct_top_delta.xlsx")

RAW_CSV = os.path.join(RAW_DIR, "coct_raw.csv")
RAW_XLSX = os.path.join(RAW_DIR, "coct_raw_full.xlsx")
RAW_HTML = os.path.join(RAW_DIR, "coct_raw_page.html")

STATE_FILE_DEFAULT = os.path.join(STATE_DIR, "seen_ids.json")
HASH_VERSION_FILE = os.path.join(STATE_DIR, "hash_version.txt")
LOCK_FILE = os.path.join(STATE_DIR, "run.lock")
HASH_VERSION = "coct_v2_structure_paths_stable_hash_2026_07"

TIMEOUT_S = 60

# Hash fields that define a meaningful change in a public tender notice.
HASH_COLS = [
    "description",
    "directorate",
    "department",
    "closing_date",
    "closing_datetime",
    "posted_date",
    "posted_datetime",
    "details_url",
]

# Top-delta shortlist defaults.
TOP_DELTA_LIMIT_DEFAULT = 20
TOP_DELTA_MIN_SCORE_DEFAULT = 0

# City of Cape Town directorates aligned to Henry Fagan-type opportunities.
WEIGHT_DIRECTORATE_MATCH = 5
WEIGHT_DESCRIPTION_KEYWORD = 1

PRIORITY_DIRECTORATES = [
    "Water and Sanitation",
    "Human Settlements",
    "Urban Mobility",
    "Energy",
    "Future Planning and Resilience",
    "Corporate Services",
]

DESCRIPTION_KEYWORD_RULES = {
    "Water and Sanitation": ["Water and Sanitation"],
    "Housing": ["Housing"],
    "Education": ["Education"],
    "Transport": ["Transport"],
    "Nuclear": ["Nuclear"],
    "Bus Rapid Transit (BRT)": ["Bus Rapid Transit", "BRT", "IRT"],
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
    "Professional Services": ["Professional Services"],
    "Consulting": ["Consulting", "Consultant", "Consultancy"],
    "Infrastructure": ["Infrastructure"],
    "Wastewater": ["Wastewater", "Waste Water"],
    "Sewer": ["Sewer", "Sewerage"],
    "Stormwater": ["Stormwater", "Storm Water"],
}

EXPECTED_OUTPUT_COLUMNS = [
    "source",
    "opportunity_key",
    "tender_number",
    "client",
    "province",
    "description",
    "directorate",
    "department",
    "closing_date",
    "closing_datetime",
    "posted_date",
    "posted_datetime",
    "details_url",
    "detail_id",
    "priority_directorate_match",
    "matched_description_keywords",
    "description_keyword_match_count",
    "weight_score",
    "row_hash",
    "run_timestamp",
    "change_type",
    "previous_row_hash",
]


class LinkTableParser(HTMLParser):
    """Small standard-library parser to extract tender detail links from tables."""

    def __init__(self):
        super().__init__()
        self.in_table = False
        self.in_tr = False
        self.in_td_or_th = False
        self.table_index = -1
        self.current_cell_text: List[str] = []
        self.current_row_cells: List[str] = []
        self.current_row_links: List[str] = []
        self.tables: List[List[Dict[str, Any]]] = []

    def handle_starttag(self, tag: str, attrs: List[tuple[str, Optional[str]]]):
        attrs_dict = dict(attrs)
        if tag.lower() == "table":
            self.in_table = True
            self.table_index += 1
            self.tables.append([])
        elif self.in_table and tag.lower() == "tr":
            self.in_tr = True
            self.current_row_cells = []
            self.current_row_links = []
        elif self.in_table and self.in_tr and tag.lower() in {"td", "th"}:
            self.in_td_or_th = True
            self.current_cell_text = []
        elif self.in_table and self.in_tr and tag.lower() == "a":
            href = attrs_dict.get("href")
            if href:
                self.current_row_links.append(urljoin(BASE_URL, href))

    def handle_data(self, data: str):
        if self.in_table and self.in_tr and self.in_td_or_th:
            self.current_cell_text.append(data)

    def handle_endtag(self, tag: str):
        tag = tag.lower()
        if self.in_table and self.in_tr and tag in {"td", "th"}:
            text = " ".join(" ".join(self.current_cell_text).split())
            self.current_row_cells.append(text)
            self.current_cell_text = []
            self.in_td_or_th = False
        elif self.in_table and tag == "tr":
            if self.current_row_cells or self.current_row_links:
                self.tables[self.table_index].append(
                    {
                        "cells": self.current_row_cells,
                        "links": self.current_row_links,
                    }
                )
            self.in_tr = False
        elif tag == "table":
            self.in_table = False


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
        force=True,
    )


def build_session() -> requests.Session:
    """Requests session with retry/backoff for flaky endpoints."""
    sess = requests.Session()
    sess.headers.update(
        {
            "User-Agent": "Mozilla/5.0",
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


def _normalise_exact_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return " ".join(str(value).casefold().strip().split())


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


def _normalise_key(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return " ".join(str(value).strip().split()).casefold()


def _make_unique_columns(columns: List[Any]) -> List[str]:
    seen: Dict[str, int] = {}
    out = []
    for col in columns:
        base = " ".join(str(col).strip().split())
        if base in seen:
            seen[base] += 1
            out.append(f"{base}.{seen[base]}")
        else:
            seen[base] = 0
            out.append(base)
    return out


def _coerce_datetime(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def _load_seen_ids(path: str) -> Set[str]:
    try:
        if not os.path.exists(path):
            return set()
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return set(_normalise_key(x) for x in (data or []))
    except (OSError, ValueError, TypeError) as e:
        logging.exception("Failed to load state file: %s -- %s", path, e)
        return set()


def _save_seen_ids(path: str, ids: Set[str]):
    try:
        norm = sorted(_normalise_key(x) for x in ids)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(norm, fh, ensure_ascii=False)
        os.replace(tmp, path)
    except (OSError, TypeError) as e:
        logging.exception("Failed to save state file: %s -- %s", path, e)


def _read_hash_version(path: str) -> str:
    try:
        if not os.path.exists(path):
            return ""
        with open(path, "r", encoding="utf-8") as fh:
            return fh.read().strip()
    except OSError as exc:
        logging.warning("Could not read hash version file %s: %s", path, exc)
        return ""


def _write_hash_version(path: str, version: str) -> None:
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            fh.write(version)
        os.replace(tmp, path)
    except OSError as exc:
        logging.warning("Could not write hash version file %s: %s", path, exc)


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


def fetch_portal_html() -> str:
    sess = build_session()
    response = sess.get(PORTAL_URL, timeout=TIMEOUT_S)
    response.raise_for_status()
    return response.text


def extract_detail_links_by_tender(html: str) -> Dict[str, Dict[str, str]]:
    """Extract public row-level detail links, even though opening details requires login."""
    parser = LinkTableParser()
    parser.feed(html)

    result: Dict[str, Dict[str, str]] = {}
    tender_no_regex = re.compile(r"^(?:RFI\d+|\d+[A-Z]?)/\d{4}/\d{2}$", re.IGNORECASE)

    for table in parser.tables:
        # Identify the tender table by checking its header row.
        if not table:
            continue
        header_cells = [_normalise_search_text(c) for c in table[0].get("cells", [])]
        header_joined = " ".join(header_cells)
        if "tender number" not in header_joined or "description" not in header_joined:
            continue

        for row in table[1:]:
            cells = row.get("cells", [])
            if not cells:
                continue
            tender_number = cells[0].strip()
            if not tender_no_regex.match(tender_number):
                continue

            links = row.get("links", [])
            detail_urls = [u for u in links if "/Tender/Details/" in u or "/tenderportal/Tender/Details/" in u]
            add_notice_urls = [u for u in links if "/Tender/AddNotice/" in u or "/tenderportal/Tender/AddNotice/" in u]

            details_url = detail_urls[0] if detail_urls else ""
            add_notice_url = add_notice_urls[0] if add_notice_urls else ""
            detail_id_match = re.search(r"/Details/(\d+)", details_url, re.IGNORECASE)
            detail_id = detail_id_match.group(1) if detail_id_match else ""

            result.setdefault(
                _normalise_key(tender_number),
                {
                    "details_url": details_url,
                    "add_notice_url": add_notice_url,
                    "detail_id": detail_id,
                },
            )

    return result


def parse_tender_table(html: str) -> pd.DataFrame:
    """Parse the public tender table from the City portal HTML."""
    try:
        tables = pd.read_html(io.StringIO(html))
    except ImportError as exc:
        raise RuntimeError(
            "pandas.read_html requires an HTML parser package. Install lxml with: pip install lxml"
        ) from exc
    except ValueError as exc:
        raise RuntimeError("No HTML tables found on the City of Cape Town tender page.") from exc

    selected = None
    for table in tables:
        cols = _make_unique_columns(list(table.columns))
        table.columns = cols
        cols_norm = [_normalise_search_text(c) for c in cols]
        if "tender number" in cols_norm and "description" in cols_norm:
            selected = table.copy()
            break

    if selected is None:
        raise RuntimeError("Could not identify the tender notice table on the City of Cape Town page.")

    selected.columns = _make_unique_columns(list(selected.columns))
    return selected


def map_and_clean(raw_table: pd.DataFrame, detail_link_map: Dict[str, Dict[str, str]]) -> pd.DataFrame:
    """Map City portal columns to stable schema and de-duplicate by tender_number."""
    df = raw_table.copy()

    # Rename the columns we expect from the portal. Duplicate labels are named .1 by _make_unique_columns.
    colmap = {
        "Tender Number": "tender_number",
        "Description": "description",
        "Directorate": "directorate",
        "Department": "department",
        "Closing Date": "closing_date",
        "Closing Date.1": "closing_datetime",
        "Posted Date": "posted_date",
        "Posted Date.1": "posted_datetime",
    }
    present = {k: v for k, v in colmap.items() if k in df.columns}
    if "Tender Number" not in present or "Description" not in present:
        raise RuntimeError("Expected City tender columns are missing. Portal structure may have changed.")

    df = df.rename(columns=present)

    keep = [
        "tender_number",
        "description",
        "directorate",
        "department",
        "closing_date",
        "closing_datetime",
        "posted_date",
        "posted_datetime",
    ]
    keep = [c for c in keep if c in df.columns]
    df = df[keep].copy()

    for c in ["tender_number", "description", "directorate", "department"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).apply(lambda x: " ".join(x.split()))

    # Filter out malformed/header rows.
    df["tender_number"] = df["tender_number"].fillna("").astype(str).str.strip()
    df = df[df["tender_number"].ne("")].copy()
    df = df[~df["tender_number"].str.casefold().eq("tender number")].copy()

    # Coerce dates.
    for c in ["closing_date", "closing_datetime", "posted_date", "posted_datetime"]:
        if c in df.columns:
            df[c] = _coerce_datetime(df[c])

    # Public source identifier.
    df.insert(0, "source", SOURCE_NAME)
    df.insert(1, "opportunity_key", "coct:" + df["tender_number"].fillna("").astype(str).str.strip().str.casefold())
    df["client"] = CLIENT_NAME
    df["province"] = PROVINCE_NAME

    # Detail links from HTML rows.
    df["_tender_key"] = df["tender_number"].apply(_normalise_key)
    df["details_url"] = df["_tender_key"].map(lambda k: detail_link_map.get(k, {}).get("details_url", ""))
    df["detail_id"] = df["_tender_key"].map(lambda k: detail_link_map.get(k, {}).get("detail_id", ""))
    df = df.drop(columns=["_tender_key"], errors="ignore")

    # Keep latest row per tender number. The portal can show duplicate tender numbers after notices/updates.
    if "posted_datetime" in df.columns:
        df = df.sort_values("posted_datetime", ascending=False, na_position="last", kind="mergesort")
    elif "posted_date" in df.columns:
        df = df.sort_values("posted_date", ascending=False, na_position="last", kind="mergesort")

    df = df.drop_duplicates(subset=["opportunity_key"], keep="first")

    return df


def filter_current_tenders(df: pd.DataFrame) -> pd.DataFrame:
    """Keep only rows where the closing date/datetime has not passed."""
    if "closing_datetime" in df.columns and df["closing_datetime"].notna().any():
        close = df["closing_datetime"]
    elif "closing_date" in df.columns:
        close = df["closing_date"]
    else:
        raise RuntimeError("No closing date field available for current tender filtering.")

    before = len(df)
    today = pd.Timestamp.today().normalize()
    out = df[close.notna() & (close >= today)].copy()
    logging.info("Filtered current tenders by closing date: %d -> %d rows", before, len(out))
    return out


def add_priority_scoring(df: pd.DataFrame) -> pd.DataFrame:
    """Add directorate/description match details and a priority weight score."""
    out = df.copy()
    priority_directorates_norm = {_normalise_exact_text(x) for x in PRIORITY_DIRECTORATES}

    if "directorate" not in out.columns:
        out["priority_directorate_match"] = False
    else:
        out["priority_directorate_match"] = out["directorate"].apply(
            lambda x: _normalise_exact_text(x) in priority_directorates_norm
        )

    def matched_keywords(description: Any) -> str:
        desc_norm = _normalise_search_text(description)
        matches = []
        for label, phrases in DESCRIPTION_KEYWORD_RULES.items():
            if any(_contains_phrase(desc_norm, phrase) for phrase in phrases):
                matches.append(label)
        return "; ".join(matches)

    out["matched_description_keywords"] = out["description"].apply(matched_keywords) if "description" in out.columns else ""
    out["description_keyword_match_count"] = out["matched_description_keywords"].apply(
        lambda x: 0 if not x else len(str(x).split("; "))
    )
    out["weight_score"] = (
        out["priority_directorate_match"].astype(int) * WEIGHT_DIRECTORATE_MATCH
        + out["description_keyword_match_count"].astype(int) * WEIGHT_DESCRIPTION_KEYWORD
    )

    return sort_for_review(out)


def sort_for_review(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    sort_cols = [c for c in ["weight_score", "closing_datetime", "closing_date", "posted_datetime"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")
    return out


def build_top_delta(df: pd.DataFrame, limit: int, min_score: int) -> pd.DataFrame:
    out = df.copy()
    if "weight_score" not in out.columns:
        out["weight_score"] = 0
    out = out[out["weight_score"].fillna(0).astype(int) >= int(min_score)].copy()
    out = sort_for_review(out)
    if limit and int(limit) > 0:
        out = out.head(int(limit)).copy()
    return out


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
    """Write XLSX and create an Excel Table when rows exist. Atomic replace."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp.xlsx"

    out_df = df.copy() if df is not None else pd.DataFrame()
    if out_df.shape[1] == 0:
        out_df = pd.DataFrame(columns=EXPECTED_OUTPUT_COLUMNS)

    with pd.ExcelWriter(
        tmp,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd hh:mm",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
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
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name="Master", engine="openpyxl")
    except Exception as e:  # pylint: disable=broad-except
        logging.exception("Failed to read master file (%s): %s", path, e)
        return pd.DataFrame()


def append_and_dedupe_master(master: pd.DataFrame, incoming: pd.DataFrame) -> pd.DataFrame:
    """Append incoming rows to master and dedupe by opportunity_key. Incoming/current rows win."""
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

    out["opportunity_key"] = out["opportunity_key"].fillna("").astype(str).str.strip().str.casefold()
    out = out[out["opportunity_key"].ne("")].copy()

    for c in ["closing_date", "closing_datetime", "posted_date", "posted_datetime", "run_timestamp"]:
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce", dayfirst=True)

    # Current run rows must win over older master rows for the same opportunity key.
    sort_cols = [c for c in ["_source_priority", "posted_datetime", "posted_date"] if c in out.columns]
    if sort_cols:
        out = out.sort_values(sort_cols, ascending=[False] * len(sort_cols), na_position="last", kind="mergesort")

    out = out.drop_duplicates(subset=["opportunity_key"], keep="first")
    out = out.drop(columns=["_source_priority"], errors="ignore")
    return sort_for_review(out)


def _prepare_master_for_comparison(master_df: pd.DataFrame) -> pd.DataFrame:
    """Normalise a loaded master so it can be compared safely against the current run."""
    if master_df is None or master_df.empty or "opportunity_key" not in master_df.columns:
        return pd.DataFrame()

    out = master_df.copy()
    out["opportunity_key"] = out["opportunity_key"].fillna("").astype(str).str.strip().str.casefold()
    out = out[out["opportunity_key"].ne("")].copy()

    for c in ["closing_date", "closing_datetime", "posted_date", "posted_datetime", "run_timestamp"]:
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce", dayfirst=True)

    for c in [
        "row_hash",
        "weight_score",
        "priority_directorate_match",
        "matched_description_keywords",
        "description_keyword_match_count",
    ]:
        if c not in out.columns:
            if c == "row_hash":
                out[c] = compute_row_hash(out, HASH_COLS)
            elif c == "priority_directorate_match":
                out[c] = False
            elif c == "matched_description_keywords":
                out[c] = ""
            else:
                out[c] = 0

    return out


def build_delta_updates(
    master_df: pd.DataFrame,
    mapped: pd.DataFrame,
    seen_ids: Set[str],
    suppress_alerts: bool,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Build delta and updates, with optional one-run suppression during state/bootstrap migrations."""
    if suppress_alerts:
        delta_df = mapped.iloc[0:0].copy()
        delta_df["change_type"] = "NEW"
        delta_df["previous_row_hash"] = ""

        updates_df = mapped.iloc[0:0].copy()
        updates_df["change_type"] = "UPDATED"
        updates_df["previous_row_hash"] = ""
        return delta_df, updates_df

    if not master_df.empty and "opportunity_key" in master_df.columns:
        master_ids = set(master_df["opportunity_key"].tolist())
        master_hash_map = dict(
            zip(
                master_df["opportunity_key"].tolist(),
                master_df["row_hash"].fillna("").astype(str).tolist(),
            )
        )

        # Prefer the master for delta checks, but use seen_ids when available as an extra guard.
        baseline_ids = master_ids | set(seen_ids)
        is_new = ~mapped["opportunity_key"].isin(baseline_ids)

        delta_df = mapped[is_new].copy()
        delta_df["change_type"] = "NEW"
        delta_df["previous_row_hash"] = ""

        exists = ~is_new
        mapped_prev_hash = mapped["opportunity_key"].map(master_hash_map).fillna("")
        is_updated = exists & mapped["opportunity_key"].isin(master_ids) & (mapped["row_hash"] != mapped_prev_hash)

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


def main() -> int:
    parser = argparse.ArgumentParser(description="Fetch and clean City of Cape Town tender notices")
    parser.add_argument("--include-expired", action="store_true", help="do not filter out closed tenders")
    parser.add_argument("--state-file", default=STATE_FILE_DEFAULT, help="path to JSON file storing seen opportunity keys")
    parser.add_argument("--top-delta-limit", type=int, default=TOP_DELTA_LIMIT_DEFAULT,
                        help="maximum rows to write to coct_top_delta.xlsx. Use 0 for all qualifying rows.")
    parser.add_argument("--top-delta-min-score", type=int, default=TOP_DELTA_MIN_SCORE_DEFAULT,
                        help="minimum weight_score required for coct_top_delta.xlsx.")
    args = parser.parse_args()

    setup_logging()
    for folder in [RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR]:
        os.makedirs(folder, exist_ok=True)

    logging.info("Using tender scrapers root: %s", TENDER_SCRAPERS_ROOT)
    logging.info("Using source root: %s", SOURCE_ROOT)
    logging.info("Using raw directory: %s", RAW_DIR)
    logging.info("Using output directory: %s", OUTPUT_DIR)
    logging.info("Using log directory: %s", LOG_DIR)
    logging.info("Using state directory: %s", STATE_DIR)

    run_ts = pd.Timestamp.now()

    try:
        acquire_lock(LOCK_FILE)

        html = fetch_portal_html()
        with open(RAW_HTML, "w", encoding="utf-8") as fh:
            fh.write(html)
        logging.info("Saved raw HTML: %s", RAW_HTML)

        detail_link_map = extract_detail_links_by_tender(html)
        raw_table = parse_tender_table(html)
        logging.info("Raw tender table rows: %d | cols: %d", len(raw_table), len(raw_table.columns))

        raw_table.to_csv(RAW_CSV, index=False)
        logging.info("Saved raw CSV: %s", RAW_CSV)

        with pd.ExcelWriter(
            RAW_XLSX,
            engine="xlsxwriter",
            datetime_format="yyyy-mm-dd hh:mm",
            engine_kwargs={"options": {"strings_to_urls": False}},
        ) as writer:
            raw_table.to_excel(writer, index=False, sheet_name="Raw_Data")
        logging.info("Saved raw Excel: %s", RAW_XLSX)

        mapped = map_and_clean(raw_table, detail_link_map)
        if not args.include_expired:
            mapped = filter_current_tenders(mapped)

        mapped = add_priority_scoring(mapped)
        mapped["row_hash"] = compute_row_hash(mapped, HASH_COLS)
        mapped["run_timestamp"] = run_ts
        mapped["opportunity_key"] = mapped["opportunity_key"].fillna("").astype(str).str.strip().str.casefold()

        master_df = _prepare_master_for_comparison(load_master(MASTER_XLSX))
        seen_ids = _load_seen_ids(args.state_file)
        recorded_hash_version = _read_hash_version(HASH_VERSION_FILE)

        suppress_alerts = False
        if not seen_ids and not master_df.empty:
            suppress_alerts = True
            logging.info(
                "No seen_ids.json found. Bootstrapping City of Cape Town state from existing master "
                "and suppressing delta/update alerts for this run."
            )

        if recorded_hash_version != HASH_VERSION and not master_df.empty:
            suppress_alerts = True
            logging.info(
                "City of Cape Town hash/state version changed or was not recorded. "
                "Re-baselining and suppressing delta/update alerts for this run."
            )

        delta_df, updates_df = build_delta_updates(master_df, mapped, seen_ids, suppress_alerts)

        delta_df = sort_for_review(delta_df)
        updates_df = sort_for_review(updates_df)
        top_delta_df = build_top_delta(delta_df, args.top_delta_limit, args.top_delta_min_score)

        logging.info("Delta (NEW) rows this run: %d", len(delta_df))
        logging.info("Updates (CHANGED) rows this run: %d", len(updates_df))
        logging.info("Top delta rows this run: %d", len(top_delta_df))

        write_excel_with_table(delta_df, DELTA_XLSX, "Delta", "tbl_delta")
        logging.info("Saved delta Excel: %s | rows: %d", DELTA_XLSX, len(delta_df))

        write_excel_with_table(updates_df, UPDATES_XLSX, "Updates", "tbl_updates")
        logging.info("Saved updates Excel: %s | rows: %d", UPDATES_XLSX, len(updates_df))

        write_excel_with_table(top_delta_df, TOP_DELTA_XLSX, "Top_Delta", "tbl_top_delta")
        logging.info("Saved top delta Excel: %s | rows: %d", TOP_DELTA_XLSX, len(top_delta_df))

        updated_master = append_and_dedupe_master(master_df, mapped)
        write_excel_with_table(updated_master, MASTER_XLSX, "Master", "tbl_master")
        logging.info("Saved master Excel: %s | rows: %d", MASTER_XLSX, len(updated_master))

        master_seen = set(updated_master["opportunity_key"].fillna("").astype(str).str.strip().str.casefold().tolist())
        _save_seen_ids(args.state_file, master_seen)
        _write_hash_version(HASH_VERSION_FILE, HASH_VERSION)

        return 0

    except (requests.RequestException, ValueError, RuntimeError, OSError) as e:
        logging.exception("FAILED: %s", e)
        return 1
    finally:
        release_lock(LOCK_FILE)


if __name__ == "__main__":
    raise SystemExit(main())
