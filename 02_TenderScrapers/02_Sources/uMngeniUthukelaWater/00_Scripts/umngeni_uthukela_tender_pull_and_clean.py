"""
uMngeni-uThukela Water Tender Scraper (Raw + Delta + Updates + Master + Top Delta)

Scrapes advertised tenders from the uMngeni-uThukela Water tender archive page,
cleans/normalises them, and maintains:

1) MASTER Excel: latest snapshot per reference_number
2) DELTA Excel: brand new tenders discovered this run only
3) UPDATES Excel: existing tenders whose key operational fields changed
4) TOP DELTA Excel: highest-scoring new tenders from the current run
5) RAW outputs: unfiltered scrape for audit/debugging

Default behaviour:
    - Scrapes the public tender archive page.
    - Saves raw rows unfiltered.
    - Keeps only current tenders in the clean outputs unless --include-expired is used.
    - Excludes archived/cancelled/withdrawn notices unless explicitly included.

Typical daily run:
    python umngeni_uthukela_tender_pull_and_clean.py

Optional:
    python umngeni_uthukela_tender_pull_and_clean.py --include-expired
    python umngeni_uthukela_tender_pull_and_clean.py --include-archived
    python umngeni_uthukela_tender_pull_and_clean.py --top-delta-limit 50
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import re
import sys
from io import StringIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Set
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# -----------------------------
# Source
# -----------------------------
SOURCE_NAME = "uMngeni-uThukela Water"
SOURCE_URL = "https://www.umngeni-uthukela.co.za/tender/"

# -----------------------------
# Dynamic Folder Detection
# -----------------------------
SOURCE_FOLDER_NAME = "UmngeniUthukelaWater"
SCRIPTS_FOLDER_NAME = "00_Scripts"


def resolve_source_root() -> Path:
    """Resolve 02_Sources/[Source] from this script's location.

    Expected placement:
        02_TenderScrapers/02_Sources/UmngeniUthukelaWater/00_Scripts/this_script.py

    The fallback keeps the script usable during testing if it is run from the
    source folder itself.
    """
    script_dir = Path(__file__).resolve().parent

    if script_dir.name == SCRIPTS_FOLDER_NAME:
        return script_dir.parent

    if script_dir.name == SOURCE_FOLDER_NAME:
        return script_dir

    # Fallback for ad-hoc testing only. In production, place the script in
    # 02_Sources/[Source]/00_Scripts so all paths resolve to that source folder.
    return script_dir


SOURCE_ROOT = resolve_source_root()
BASE_ROOT = str(SOURCE_ROOT)

SCRIPT_DIR = str(SOURCE_ROOT / "00_Scripts")
RAW_DIR = str(SOURCE_ROOT / "01_RawData")
OUTPUT_DIR = str(SOURCE_ROOT / "02_Outputs")
LOG_DIR = str(SOURCE_ROOT / "03_Logs")
STATE_DIR = str(SOURCE_ROOT / "04_State")
METADATA_DIR = str(SOURCE_ROOT / "05_Metadata")

LOG_PATH = os.path.join(LOG_DIR, "umngeni_uthukela_run.log")

RAW_HTML = os.path.join(RAW_DIR, "umngeni_uthukela_raw_page.html")
RAW_CSV = os.path.join(RAW_DIR, "umngeni_uthukela_raw.csv")
RAW_XLSX = os.path.join(RAW_DIR, "umngeni_uthukela_raw_full.xlsx")

MASTER_XLSX = os.path.join(OUTPUT_DIR, "umngeni_uthukela_master.xlsx")
DELTA_XLSX = os.path.join(OUTPUT_DIR, "umngeni_uthukela_delta.xlsx")
UPDATES_XLSX = os.path.join(OUTPUT_DIR, "umngeni_uthukela_updates.xlsx")
TOP_DELTA_XLSX = os.path.join(OUTPUT_DIR, "umngeni_uthukela_top_delta.xlsx")

STATE_FILE_DEFAULT = os.path.join(STATE_DIR, "seen_ids.json")
HASH_VERSION_FILE = os.path.join(STATE_DIR, "hash_version.txt")
LOCK_FILE = os.path.join(STATE_DIR, "run.lock")

# -----------------------------
# Tunables
# -----------------------------
TIMEOUT_S = 60
TOP_DELTA_LIMIT_DEFAULT = 20
TOP_DELTA_MIN_SCORE_DEFAULT = 0

# Change-detection fields. If one of these changes for an existing tender,
# the tender is written to Updates.
HASH_COLS = [
    "description",
    "issue_date",
    "closing_datetime",
    "status",
    "view_url",
    "document_urls",
]

# Dedicated source score. uMngeni-uThukela Water is a priority water-board source,
# so all valid notices from this portal receive a priority-client/source match.
WEIGHT_PRIORITY_SOURCE = 10
WEIGHT_DESCRIPTION_KEYWORD = 1

DESCRIPTION_KEYWORD_RULES = {
    "Water and Sanitation": ["Water and Sanitation", "Water", "Sanitation", "Wastewater", "Waterworks", "Water Treatment"],
    "Housing": ["Housing"],
    "Education": ["Education"],
    "Transport": ["Transport"],
    "Nuclear": ["Nuclear"],
    "Bus Rapid Transit (BRT)": ["Bus Rapid Transit", "BRT"],
    "International Transition Plan Network (ITPN)": ["International Transition Plan Network", "ITPN"],
    "Design": ["Design", "Detailed Design"],
    "Refurbishment": ["Refurbishment", "Refurbish", "Rehabilitation", "Upgrade", "Maintenance", "Repairs"],
    "Engineering, Procurement, and Construction (EPC)": [
        "Engineering, Procurement, and Construction",
        "Engineering Procurement and Construction",
        "EPC",
        "Turnkey",
    ],
    "Construction": ["Construction", "Construct"],
    "Fire Engineering": ["Fire Engineering"],
    "HVAC": ["HVAC", "Heating Ventilation and Air Conditioning"],
    "Mechanical": ["Mechanical"],
    "Electrical": ["Electrical"],
    "Electronic": ["Electronic"],
    "Wet Services": ["Wet Services"],
    "Project Management": ["Project Management", "Contract Administration", "Close Out"],
    "Programme Management": ["Programme Management"],
    "Project Management Office (PMO)": ["Project Management Office", "PMO"],
    "Advisory Services": ["Advisory Services", "Advisory"],
    "Structural Engineering": ["Structural Engineering"],
    "Civil Engineering": ["Civil Engineering", "Civil"],
    "Feasibility Study": ["Feasibility Study", "Feasibility Studies"],
    "Environmental": ["Environmental", "Water Use License", "Water Use Licence", "Impact Assessment", "EIA"],
}


def setup_logging() -> None:
    for folder in [SCRIPT_DIR, RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR]:
        os.makedirs(folder, exist_ok=True)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(LOG_PATH, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def build_session() -> requests.Session:
    sess = requests.Session()
    sess.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
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


def _normalise_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return " ".join(str(value).strip().split())


def _normalise_key(value: Any) -> str:
    text = _normalise_text(value).casefold()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")


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


def _clean_url(url: str) -> str:
    return urljoin(SOURCE_URL, str(url).strip())


def _parse_dt(value: Any) -> pd.Timestamp:
    if value is None or pd.isna(value):
        return pd.NaT
    text = _normalise_text(value)
    if not text:
        return pd.NaT

    # Remove common labels if present.
    text = re.sub(r"^(issue date|closing date|date)\s*:?\s*", "", text, flags=re.I)

    # Site uses dd/mm/yyyy h:mm am/pm.
    parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
    return parsed


def _format_links(links: List[str]) -> str:
    cleaned = []
    for link in links:
        if link and link not in cleaned:
            cleaned.append(link)
    return "; ".join(cleaned)


def fetch_html() -> str:
    sess = build_session()
    response = sess.get(SOURCE_URL, timeout=TIMEOUT_S)
    response.raise_for_status()
    html = response.text

    os.makedirs(RAW_DIR, exist_ok=True)
    with open(RAW_HTML, "w", encoding="utf-8") as fh:
        fh.write(html)
    logging.info("Saved raw HTML: %s", RAW_HTML)

    return html


def _row_links(td) -> List[str]:
    links = []
    for a in td.find_all("a", href=True):
        links.append(_clean_url(a["href"]))
    return links


def _row_link_texts(td) -> List[str]:
    texts = []
    for a in td.find_all("a", href=True):
        txt = _normalise_text(a.get_text(" ", strip=True))
        if txt:
            texts.append(txt)
    return texts


def parse_from_html_tables(soup: BeautifulSoup) -> List[Dict[str, Any]]:
    """Parse the normal tender archive table structure."""
    rows: List[Dict[str, Any]] = []

    for table in soup.find_all("table"):
        trs = table.find_all("tr")
        if not trs:
            continue

        # Find header row.
        header_cells = trs[0].find_all(["th", "td"])
        headers = [_normalise_text(c.get_text(" ", strip=True)).casefold() for c in header_cells]
        header_blob = " | ".join(headers)

        if not ("title" in header_blob and "reference" in header_blob and "closing" in header_blob):
            continue

        logging.info("Detected uMngeni-uThukela tender table with headers: %s", header_blob)

        for tr in trs[1:]:
            cells = tr.find_all("td")
            if len(cells) < 5:
                continue

            title_cell = cells[0]
            ref_cell = cells[1]
            view_cell = cells[2] if len(cells) > 2 else None
            issue_cell = cells[3] if len(cells) > 3 else None
            closing_cell = cells[4] if len(cells) > 4 else None
            status_cell = cells[5] if len(cells) > 5 else None

            description = _normalise_text(title_cell.get_text(" ", strip=True))
            reference_number = _normalise_text(ref_cell.get_text(" ", strip=True))

            if not description and not reference_number:
                continue

            title_links = _row_links(title_cell)
            view_links = _row_links(view_cell) if view_cell else []
            all_links = []
            all_link_texts = []
            for c in cells:
                all_links.extend(_row_links(c))
                all_link_texts.extend(_row_link_texts(c))

            document_links = []
            document_names = []
            for link, text in zip(all_links, all_link_texts + [""] * max(0, len(all_links) - len(all_link_texts))):
                lower_link = link.lower()
                lower_text = text.casefold()
                if "download" in lower_text or any(ext in lower_link for ext in [".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip"]):
                    document_links.append(link)
                    document_names.append(text or os.path.basename(link))

            # Some rows have a Download anchor in a final cell; if the text extraction pairing
            # above misses it, scan all anchors directly.
            for a in tr.find_all("a", href=True):
                text = _normalise_text(a.get_text(" ", strip=True))
                link = _clean_url(a["href"])
                if text.casefold() == "download" and link not in document_links:
                    document_links.append(link)
                    document_names.append(text)

            issue_raw = _normalise_text(issue_cell.get_text(" ", strip=True)) if issue_cell else ""
            closing_raw = _normalise_text(closing_cell.get_text(" ", strip=True)) if closing_cell else ""
            status = _normalise_text(status_cell.get_text(" ", strip=True)) if status_cell else ""

            issue_dt = _parse_dt(issue_raw)
            closing_dt = _parse_dt(closing_raw)

            opportunity_key = f"umngeni_uthukela:{_normalise_key(reference_number or description)}"

            rows.append(
                {
                    "source": SOURCE_NAME,
                    "opportunity_key": opportunity_key,
                    "reference_number": reference_number,
                    "notice_type": "Tender",
                    "description": description,
                    "client": SOURCE_NAME,
                    "issue_date": issue_dt,
                    "closing_date": closing_dt.date() if pd.notna(closing_dt) else pd.NaT,
                    "closing_datetime": closing_dt,
                    "status": status,
                    "title_url": title_links[0] if title_links else "",
                    "view_url": view_links[0] if view_links else "",
                    "document_names": "; ".join(dict.fromkeys(document_names)),
                    "document_urls": _format_links(document_links),
                    "document_count": len(set(document_links)),
                    "source_url": (view_links[0] if view_links else (title_links[0] if title_links else SOURCE_URL)),
                    "archived_flag": "archived" in status.casefold(),
                    "cancellation_flag": bool(re.search(r"\b(cancelled|canceled|withdrawn|cancellation)\b", f"{description} {status}", re.I)),
                }
            )

    return rows


def parse_from_pandas_tables(html: str) -> List[Dict[str, Any]]:
    """Fallback parser using pandas if the direct HTML parser does not find rows."""
    rows: List[Dict[str, Any]] = []

    try:
        tables = pd.read_html(StringIO(html))
    except (ValueError, ImportError):
        return rows

    for df in tables:
        if df.empty:
            continue

        columns = [str(c).strip().casefold() for c in df.columns]
        col_blob = " | ".join(columns)

        if not ("title" in col_blob and "reference" in col_blob and "closing" in col_blob):
            continue

        colmap = {}
        for col in df.columns:
            norm = str(col).strip().casefold()
            if "title" in norm:
                colmap[col] = "description"
            elif "reference" in norm:
                colmap[col] = "reference_number"
            elif "issue" in norm:
                colmap[col] = "issue_date"
            elif "closing" in norm:
                colmap[col] = "closing_datetime"
            elif "open" in norm or "archived" in norm or "result" in norm:
                colmap[col] = "status"

        temp = df.rename(columns=colmap)

        for _, r in temp.iterrows():
            description = _normalise_text(r.get("description", ""))
            reference_number = _normalise_text(r.get("reference_number", ""))
            if not description and not reference_number:
                continue

            issue_dt = _parse_dt(r.get("issue_date", ""))
            closing_dt = _parse_dt(r.get("closing_datetime", ""))
            status = _normalise_text(r.get("status", ""))

            rows.append(
                {
                    "source": SOURCE_NAME,
                    "opportunity_key": f"umngeni_uthukela:{_normalise_key(reference_number or description)}",
                    "reference_number": reference_number,
                    "notice_type": "Tender",
                    "description": description,
                    "client": SOURCE_NAME,
                    "issue_date": issue_dt,
                    "closing_date": closing_dt.date() if pd.notna(closing_dt) else pd.NaT,
                    "closing_datetime": closing_dt,
                    "status": status,
                    "title_url": "",
                    "view_url": "",
                    "document_names": "",
                    "document_urls": "",
                    "document_count": 0,
                    "source_url": SOURCE_URL,
                    "archived_flag": "archived" in status.casefold(),
                    "cancellation_flag": bool(re.search(r"\b(cancelled|canceled|withdrawn|cancellation)\b", f"{description} {status}", re.I)),
                }
            )

    return rows


def parse_umngeni_html(html: str) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "html.parser")

    rows = parse_from_html_tables(soup)
    if rows:
        logging.info("Parsed rows using direct table parser: %d", len(rows))
        return rows

    rows = parse_from_pandas_tables(html)
    if rows:
        logging.info("Parsed rows using pandas table fallback: %d", len(rows))
        return rows

    # Diagnostic only; main treats zero parsed rows as a parse failure so existing
    # master files are not wiped by an unexpected site/layout issue.
    text_preview = " | ".join(
        [_normalise_text(line) for line in soup.get_text("\n").splitlines() if _normalise_text(line)][:30]
    )
    logging.warning("No uMngeni-uThukela rows parsed. First extracted text lines: %s", text_preview)
    return []


def map_and_clean(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(
            columns=[
                "source",
                "opportunity_key",
                "reference_number",
                "notice_type",
                "description",
                "client",
                "issue_date",
                "closing_date",
                "closing_datetime",
                "status",
                "title_url",
                "view_url",
                "document_names",
                "document_urls",
                "document_count",
                "source_url",
                "archived_flag",
                "cancellation_flag",
            ]
        )

    df = pd.DataFrame(rows)

    for c in ["issue_date", "closing_date", "closing_datetime"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)

    for c in ["reference_number", "description", "status", "document_urls", "document_names"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).str.strip()

    if "opportunity_key" in df.columns:
        df["opportunity_key"] = df["opportunity_key"].fillna("").astype(str).str.strip().str.lower()
        df = df[df["opportunity_key"].ne("")]
        df = df.drop_duplicates(subset=["opportunity_key"], keep="first")

    sort_cols = [c for c in ["closing_datetime", "issue_date"] if c in df.columns]
    if sort_cols:
        df = df.sort_values(sort_cols, ascending=[True] * len(sort_cols), na_position="last")

    return df


def filter_current(df: pd.DataFrame, include_expired: bool, include_archived: bool, include_cancellations: bool) -> pd.DataFrame:
    out = df.copy()
    before = len(out)

    if not include_cancellations and "cancellation_flag" in out.columns:
        out = out[~out["cancellation_flag"].fillna(False).astype(bool)].copy()
        logging.info("Filtered cancellation/withdrawal notices: %d -> %d rows", before, len(out))
        before = len(out)

    if not include_archived and "archived_flag" in out.columns:
        out = out[~out["archived_flag"].fillna(False).astype(bool)].copy()
        logging.info("Filtered archived notices: %d -> %d rows", before, len(out))
        before = len(out)

    if not include_expired and "closing_datetime" in out.columns:
        now = pd.Timestamp.now()
        out["closing_datetime"] = pd.to_datetime(out["closing_datetime"], errors="coerce")
        out = out[pd.notna(out["closing_datetime"]) & (out["closing_datetime"] >= now)].copy()
        logging.info("Filtered current tenders by closing_datetime >= now: %d -> %d rows", before, len(out))

    return out


def add_priority_scoring(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    out["priority_client_match"] = True

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
        out["priority_client_match"].astype(int) * WEIGHT_PRIORITY_SOURCE
        + out["description_keyword_match_count"].astype(int) * WEIGHT_DESCRIPTION_KEYWORD
    )

    sort_cols = [c for c in ["weight_score", "closing_datetime", "issue_date"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

    return out


def compute_row_hash(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    if df is None or df.empty:
        return pd.Series([], dtype=str, index=df.index if df is not None else None)

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


def build_top_delta(df: pd.DataFrame, limit: int, min_score: int) -> pd.DataFrame:
    out = df.copy()

    if "weight_score" not in out.columns:
        out["weight_score"] = 0

    if min_score is not None:
        out = out[out["weight_score"].fillna(0).astype(int) >= int(min_score)].copy()

    sort_cols = [c for c in ["weight_score", "closing_datetime", "issue_date"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

    if limit and int(limit) > 0:
        out = out.head(int(limit)).copy()

    return out


def write_excel_with_table(df: pd.DataFrame, path: str, sheet_name: str, table_name: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp.xlsx"

    out_df = df.copy() if df is not None else pd.DataFrame()

    if out_df.shape[1] == 0:
        out_df = pd.DataFrame(
            columns=[
                "source",
                "opportunity_key",
                "reference_number",
                "notice_type",
                "description",
                "client",
                "issue_date",
                "closing_date",
                "closing_datetime",
                "status",
                "title_url",
                "view_url",
                "document_names",
                "document_urls",
                "document_count",
                "source_url",
                "archived_flag",
                "cancellation_flag",
                "priority_client_match",
                "matched_description_keywords",
                "description_keyword_match_count",
                "weight_score",
                "row_hash",
                "run_timestamp",
                "change_type",
                "previous_row_hash",
            ]
        )

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
    except Exception as e:
        logging.exception("Failed to read master file (%s): %s", path, e)
        raise RuntimeError(f"Existing master could not be read safely; not overwriting it: {path}") from e


def append_and_dedupe_master(master: pd.DataFrame, incoming: pd.DataFrame) -> pd.DataFrame:
    if master.empty:
        out = incoming.copy()
        out["_source_priority"] = 1
        out["_source_order"] = range(len(out))
    else:
        master_part = master.copy()
        incoming_part = incoming.copy()
        master_part["_source_priority"] = 0
        incoming_part["_source_priority"] = 1
        master_part["_source_order"] = range(len(master_part))
        incoming_part["_source_order"] = range(len(incoming_part))
        out = pd.concat([master_part, incoming_part], ignore_index=True)

    if "opportunity_key" not in out.columns:
        return out.drop(columns=["_source_priority", "_source_order"], errors="ignore")

    out["opportunity_key"] = out["opportunity_key"].fillna("").astype(str).str.strip().str.lower()
    out = out[out["opportunity_key"].ne("")].copy()

    # For duplicate keys, the current run must win over older master rows.
    out = out.sort_values(
        ["opportunity_key", "_source_priority", "_source_order"],
        ascending=[True, False, True],
        kind="mergesort",
    )
    out = out.drop_duplicates(subset=["opportunity_key"], keep="first")
    out = out.drop(columns=["_source_priority", "_source_order"], errors="ignore")

    sort_cols = [c for c in ["weight_score", "closing_datetime", "issue_date"] if c in out.columns]
    if sort_cols:
        if "closing_datetime" in out.columns:
            out["closing_datetime"] = pd.to_datetime(out["closing_datetime"], errors="coerce")
        if "issue_date" in out.columns:
            out["issue_date"] = pd.to_datetime(out["issue_date"], errors="coerce")
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last", kind="mergesort")

    return out


def _load_seen_ids(path: str) -> Set[str]:
    try:
        if not os.path.exists(path):
            return set()
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return set(str(x).strip().lower() for x in (data or []))
    except (OSError, ValueError, TypeError) as e:
        logging.exception("Failed to load state file: %s -- %s", path, e)
        return set()


def _save_seen_ids(path: str, ids: Set[str]) -> None:
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(sorted(ids), fh, ensure_ascii=False)
        os.replace(tmp, path)
    except (OSError, TypeError) as e:
        logging.exception("Failed to save state file: %s -- %s", path, e)


def _is_migration_baseline_run(hash_version_path: str) -> bool:
    """First migrated run suppresses delta/update alerts to avoid false positives."""
    return not os.path.exists(hash_version_path)


def _write_hash_version(path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        fh.write("umngeni_uthukela_hash_v1\n")
    os.replace(tmp, path)


def _empty_like(df: pd.DataFrame, change_type: str) -> pd.DataFrame:
    out = df.iloc[0:0].copy()
    out["change_type"] = change_type
    out["previous_row_hash"] = ""
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description="Scrape and clean uMngeni-uThukela Water tenders")
    parser.add_argument("--include-expired", action="store_true", help="include expired tenders in clean outputs")
    parser.add_argument("--include-archived", action="store_true", help="include rows marked Archived in clean outputs")
    parser.add_argument("--include-cancellations", action="store_true", help="include cancelled/withdrawn notices in clean outputs")
    parser.add_argument("--top-delta-limit", type=int, default=TOP_DELTA_LIMIT_DEFAULT,
                        help="maximum number of rows to write to top-delta. Use 0 for all qualifying rows.")
    parser.add_argument("--top-delta-min-score", type=int, default=TOP_DELTA_MIN_SCORE_DEFAULT,
                        help="minimum weight_score required for top-delta.")
    parser.add_argument("--state-file", default=STATE_FILE_DEFAULT,
                        help="path to JSON file storing seen opportunity IDs")
    args = parser.parse_args()

    setup_logging()
    os.makedirs(RAW_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(STATE_DIR, exist_ok=True)
    logging.info("Using source root: %s", BASE_ROOT)
    logging.info("Using raw directory: %s", RAW_DIR)
    logging.info("Using output directory: %s", OUTPUT_DIR)
    logging.info("Using state directory: %s", STATE_DIR)

    run_ts = pd.Timestamp.now()

    try:
        acquire_lock(LOCK_FILE)

        html = fetch_html()
        rows = parse_umngeni_html(html)
        df_raw = map_and_clean(rows)

        logging.info("Raw parsed rows: %d | cols: %d", len(df_raw), len(df_raw.columns))

        df_raw.to_csv(RAW_CSV, index=False)
        logging.info("Saved raw CSV: %s", RAW_CSV)

        write_excel_with_table(df_raw, RAW_XLSX, sheet_name="Raw_Data", table_name="tbl_raw")
        logging.info("Saved raw Excel: %s", RAW_XLSX)

        if df_raw.empty:
            raise ValueError(
                "No rows were parsed from the uMngeni-uThukela tender page. "
                "Raw diagnostics were saved, but clean outputs/master/state were not overwritten."
            )

        mapped = filter_current(
            df_raw,
            include_expired=args.include_expired,
            include_archived=args.include_archived,
            include_cancellations=args.include_cancellations,
        )

        mapped = add_priority_scoring(mapped)
        mapped = mapped.copy()
        mapped["row_hash"] = compute_row_hash(mapped, HASH_COLS)
        mapped["run_timestamp"] = run_ts

        master_df = load_master(MASTER_XLSX)

        if not master_df.empty and "opportunity_key" in master_df.columns:
            master_df = master_df.copy()
            master_df["opportunity_key"] = master_df["opportunity_key"].fillna("").astype(str).str.strip().str.lower()
            master_df = master_df[master_df["opportunity_key"].ne("")]

            for col, default in [
                ("priority_client_match", False),
                ("matched_description_keywords", ""),
                ("description_keyword_match_count", 0),
                ("weight_score", 0),
                ("row_hash", ""),
            ]:
                if col not in master_df.columns:
                    master_df[col] = default

            if not args.include_expired:
                master_df = filter_current(
                    master_df,
                    include_expired=False,
                    include_archived=args.include_archived,
                    include_cancellations=args.include_cancellations,
                )

            if "row_hash" not in master_df.columns or master_df["row_hash"].fillna("").eq("").all():
                master_df["row_hash"] = compute_row_hash(master_df, HASH_COLS)

        migration_baseline = _is_migration_baseline_run(HASH_VERSION_FILE)
        if migration_baseline:
            logging.info(
                "Migration baseline mode active: suppressing delta/update/top-delta alerts for this first migrated run."
            )
            delta_df = _empty_like(mapped, "NEW")
            updates_df = _empty_like(mapped, "UPDATED")
        elif not master_df.empty and "opportunity_key" in master_df.columns:
            master_ids = set(master_df["opportunity_key"].astype(str).str.lower().tolist())
            master_hash_map = dict(
                zip(
                    master_df["opportunity_key"].astype(str).str.lower().tolist(),
                    master_df["row_hash"].fillna("").astype(str).tolist(),
                )
            )

            is_new = ~mapped["opportunity_key"].isin(master_ids)
            delta_df = mapped[is_new].copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""

            exists = ~is_new
            prev_hash = mapped["opportunity_key"].map(master_hash_map).fillna("")
            is_updated = exists & (mapped["row_hash"] != prev_hash)

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

        sort_cols = [c for c in ["weight_score", "closing_datetime", "issue_date"] if c in delta_df.columns]
        if sort_cols:
            ascending = [False if c == "weight_score" else True for c in sort_cols]
            delta_df = delta_df.sort_values(sort_cols, ascending=ascending, na_position="last")
            updates_df = updates_df.sort_values(sort_cols, ascending=ascending, na_position="last")

        top_delta_df = build_top_delta(
            delta_df,
            limit=args.top_delta_limit,
            min_score=args.top_delta_min_score,
        )

        logging.info("Delta rows this run: %d", len(delta_df))
        logging.info("Updates rows this run: %d", len(updates_df))
        logging.info("Top delta rows this run: %d", len(top_delta_df))

        write_excel_with_table(delta_df, DELTA_XLSX, sheet_name="Delta", table_name="tbl_delta")
        logging.info("Saved delta Excel: %s | rows: %d", DELTA_XLSX, len(delta_df))

        write_excel_with_table(updates_df, UPDATES_XLSX, sheet_name="Updates", table_name="tbl_updates")
        logging.info("Saved updates Excel: %s | rows: %d", UPDATES_XLSX, len(updates_df))

        write_excel_with_table(top_delta_df, TOP_DELTA_XLSX, sheet_name="Top_Delta", table_name="tbl_top_delta")
        logging.info("Saved top delta Excel: %s | rows: %d", TOP_DELTA_XLSX, len(top_delta_df))

        updated_master = append_and_dedupe_master(master_df, mapped)
        write_excel_with_table(updated_master, MASTER_XLSX, sheet_name="Master", table_name="tbl_master")
        logging.info("Saved master Excel: %s | rows: %d", MASTER_XLSX, len(updated_master))

        if "opportunity_key" in updated_master.columns:
            seen = set(updated_master["opportunity_key"].fillna("").astype(str).str.strip().str.lower().tolist())
            _save_seen_ids(args.state_file, seen)

        _write_hash_version(HASH_VERSION_FILE)

        return 0

    except (requests.RequestException, ValueError, RuntimeError, OSError) as e:
        logging.exception("FAILED: %s", e)
        return 1

    finally:
        release_lock(LOCK_FILE)


if __name__ == "__main__":
    raise SystemExit(main())
