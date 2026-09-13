r"""NECSA current tenders scraper.

Scrapes current tenders from:
    https://www.necsa.co.za/tenders/

Outputs:
    necsa_raw.csv
    necsa_raw_full.xlsx
    necsa_raw_page.html
    necsa_master.xlsx
    necsa_delta.xlsx
    necsa_updates.xlsx
    necsa_top_delta.xlsx

This source is parsed as a static WordPress/HTML page. Tender blocks are label-based
(BID NUMBER, BID DESCRIPTION, CLOSING DATE, CLOSING TIME, etc.) and related document
ZIP links are captured where the page exposes "Click here for Download all Related Documents and Information" links.
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
from urllib.parse import urljoin, urlparse, unquote

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from requests.exceptions import RequestException, SSLError
from urllib3.util.retry import Retry
import urllib3

# -----------------------------------------------------------------------------
# Source configuration
# -----------------------------------------------------------------------------
SOURCE_URL = "https://www.necsa.co.za/tenders/"
SOURCE_NAME = "NECSA"
CLIENT_NAME = "South African Nuclear Energy Corporation"

# Dynamic TenderScrapers folder detection.
# Expected location:
#   ...\02_TenderScrapers\02_Sources\NECSA\00_Scripts\necsa_tender_pull_and_clean.py
SCRIPT_PATH = Path(__file__).resolve()
try:
    SOURCE_ROOT = SCRIPT_PATH.parents[1]
    TENDER_SCRAPERS_ROOT = SCRIPT_PATH.parents[3]
except IndexError:
    SOURCE_ROOT = SCRIPT_PATH.parent
    TENDER_SCRAPERS_ROOT = SOURCE_ROOT.parent

RAW_DIR_NAME = "01_RawData"
OUTPUT_DIR_NAME = "02_Outputs"
LOG_DIR_NAME = "03_Logs"
STATE_DIR_NAME = "04_State"
METADATA_DIR_NAME = "05_Metadata"

RAW_CSV = "necsa_raw.csv"
RAW_XLSX = "necsa_raw_full.xlsx"
RAW_HTML = "necsa_raw_page.html"
MASTER_XLSX = "necsa_master.xlsx"
DELTA_XLSX = "necsa_delta.xlsx"
UPDATES_XLSX = "necsa_updates.xlsx"
TOP_DELTA_XLSX = "necsa_top_delta.xlsx"
LOG_FILE = "necsa_run.log"
SEEN_IDS_FILE = "seen_ids.json"
HASH_VERSION_FILE = "hash_version.txt"
LOCK_FILE = "run.lock"

# Increment this only when row_hash/change-detection logic changes.
HASH_VERSION = "necsa_v2_stable_hash_no_document_urls"

REQUEST_TIMEOUT = 45
REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# Keep the hash stable. Full document_urls can occasionally change because of WordPress
# upload/CMS paths while the underlying tender has not materially changed.
HASH_COLS = [
    "reference_number",
    "description",
    "closing_datetime",
    "briefing_session_details",
    "document_names",
    "document_count",
    "source_url",
    "contact_email",
    "status",
]

OUTPUT_COLUMNS = [
    "source",
    "opportunity_key",
    "tender_number",
    "reference_number",
    "notice_type",
    "title",
    "description",
    "client",
    "sector",
    "category",
    "country_name",
    "province",
    "region",
    "date_published",
    "issue_date",
    "briefing_date",
    "closing_date",
    "closing_datetime",
    "status",
    "is_current",
    "is_archived",
    "is_cancelled",
    "document_names",
    "document_urls",
    "document_count",
    "contact_person",
    "contact_email",
    "contact_telephone",
    "submission_method",
    "briefing_session_details",
    "conditions",
    "raw_category",
    "raw_status",
    "source_url",
    "run_timestamp",
    "row_hash",
    "change_type",
    "previous_row_hash",
    "weight_score",
    "matched_description_keywords",
]

KEYWORDS = [
    "nuclear", "professional engineering", "professional services", "engineering services",
    "design", "upgrade", "hvac", "facility", "contract advisory", "project management",
    "programme management", "consultant", "consulting", "technical advisory", "feasibility",
    "master plan", "infrastructure", "construction monitoring", "contract administration",
]


# -----------------------------------------------------------------------------
# Folders, logging, state and lock handling
# -----------------------------------------------------------------------------
def resolve_paths(base_root_override: str | None = None) -> Dict[str, Path]:
    source_root = Path(base_root_override).resolve() if base_root_override else SOURCE_ROOT
    return {
        "source_root": source_root,
        "tender_root": TENDER_SCRAPERS_ROOT,
        "raw_dir": source_root / RAW_DIR_NAME,
        "output_dir": source_root / OUTPUT_DIR_NAME,
        "log_dir": source_root / LOG_DIR_NAME,
        "state_dir": source_root / STATE_DIR_NAME,
        "metadata_dir": source_root / METADATA_DIR_NAME,
    }


def ensure_dirs(paths: Dict[str, Path]) -> None:
    for key in ["raw_dir", "output_dir", "log_dir", "state_dir", "metadata_dir"]:
        paths[key].mkdir(parents=True, exist_ok=True)


def setup_logging(log_dir: Path) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / LOG_FILE
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
        force=True,
    )


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
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalise_id(value: Any) -> str:
    text = clean_text(value).lower()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")


def normalise_key(value: Any) -> str:
    return clean_text(value).casefold()


def load_seen_ids(path: Path) -> Set[str]:
    try:
        if not path.exists():
            return set()
        data = json.loads(path.read_text(encoding="utf-8"))
        return {normalise_key(x) for x in (data or []) if normalise_key(x)}
    except Exception as exc:
        logging.warning("Could not load seen IDs from %s: %s", path, exc)
        return set()


def save_seen_ids(path: Path, ids: Iterable[Any]) -> None:
    try:
        values = sorted({normalise_key(x) for x in ids if normalise_key(x)})
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(json.dumps(values, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(path)
    except Exception as exc:
        logging.warning("Could not save seen IDs to %s: %s", path, exc)


def load_hash_version(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8").strip() if path.exists() else ""
    except Exception:
        return ""


def save_hash_version(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(HASH_VERSION, encoding="utf-8")


def acquire_lock(lock_path: Path) -> None:
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


def save_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", errors="ignore")


# -----------------------------------------------------------------------------
# Fetching and parsing
# -----------------------------------------------------------------------------
def build_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(REQUEST_HEADERS)
    retry = Retry(
        total=4,
        connect=4,
        read=4,
        status=4,
        backoff_factor=0.8,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


def fetch_html(url: str) -> str:
    session = build_session()
    try:
        response = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        response.raise_for_status()
        return response.text
    except SSLError:
        logging.warning("SSL verification failed for %s. Retrying without verification.", url)
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        response = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True, verify=False)
        response.raise_for_status()
        return response.text


def markerise_download_links(soup: BeautifulSoup) -> Tuple[BeautifulSoup, Dict[str, Dict[str, str]]]:
    """Replace relevant download anchor text with markers so line-based parsing can keep link order."""
    link_map: Dict[str, Dict[str, str]] = {}
    idx = 0
    for a in soup.find_all("a", href=True):
        text = clean_text(a.get_text(" "))
        href = a.get("href", "")
        combined = f"{text} {href}".lower()
        if "download all related documents" in combined or re.search(r"/wp-content/uploads/.+\.(zip|pdf|docx?|xlsx?)", combined):
            idx += 1
            marker = f"__NECSA_DOC_LINK_{idx}__"
            full_url = urljoin(SOURCE_URL, href)
            basename = Path(urlparse(full_url).path).name
            name = unquote(basename) if basename else text or "Tender document"
            link_map[marker] = {"name": name, "url": full_url, "anchor_text": text}
            a.string = f"{marker} {text}"
    return soup, link_map


def visible_lines_from_soup(soup: BeautifulSoup) -> List[str]:
    for element in soup(["script", "style", "noscript", "svg"]):
        element.decompose()
    text = soup.get_text("\n")
    lines = [clean_text(line) for line in text.splitlines()]
    return [line for line in lines if line]


def is_major_label(text: str) -> bool:
    upper = clean_text(text).upper().rstrip(":")
    labels = {
        "BID NUMBER", "BID DESCRIPTION", "CLOSING DATE", "CLOSING TIME", "BID VALIDITY PERIOD",
        "DELIVERY ADDRESS", "ENQUIRES", "ENQUIRIES", "DATE FOR SITE MEETING", "NON-COMPLUSORY SITE MEEING",
        "NON-COMPLUSORY SITE MEETING", "NON-COMPULSORY SITE MEETING", "COMPULSORY SITE MEETING",
    }
    return upper in labels


def label_value(block: List[str], label: str) -> str:
    label_upper = label.upper().rstrip(":")
    for i, line in enumerate(block):
        cleaned = clean_text(line)
        upper = cleaned.upper().rstrip(":")
        if upper == label_upper or upper.startswith(label_upper + ":"):
            inline = re.sub(rf"(?i)^{re.escape(label)}\s*:?\s*", "", cleaned).strip()
            if inline and inline.upper().rstrip(":") != label_upper:
                return inline
            for j in range(i + 1, min(i + 4, len(block))):
                candidate = clean_text(block[j])
                if candidate and not is_major_label(candidate):
                    return candidate
    return ""


def collect_description(block: List[str]) -> str:
    for i, line in enumerate(block):
        if clean_text(line).upper().startswith("BID DESCRIPTION"):
            parts: List[str] = []
            inline = re.sub(r"(?i)^BID DESCRIPTION\s*:?\s*", "", clean_text(line)).strip()
            if inline and inline.upper() != "BID DESCRIPTION":
                parts.append(inline)
            for j in range(i + 1, len(block)):
                candidate = clean_text(block[j])
                if not candidate:
                    continue
                if candidate.upper().startswith("CLOSING DATE") or is_major_label(candidate):
                    break
                if candidate.startswith("__NECSA_DOC_LINK_"):
                    break
                parts.append(candidate)
            return clean_text(" ".join(parts))
    return ""


def normalise_time(value: str) -> str:
    text = clean_text(value)
    if not text:
        return "11:00"
    text = text.replace("@", " ")
    text = re.sub(r"(?i)(\d{1,2})h(\d{2})(am|pm)", r"\1:\2 \3", text)
    text = re.sub(r"(?i)(\d{1,2})h(\d{2})\s*(am|pm)?", lambda m: f"{m.group(1)}:{m.group(2)} {m.group(3) or ''}", text)
    text = re.sub(r"(?i)(\d{1,2})h\s*(am|pm)?", lambda m: f"{m.group(1)}:00 {m.group(2) or ''}", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or "11:00"


def parse_necsa_datetime(date_text: str, time_text: str = "") -> pd.Timestamp:
    date_text = clean_text(date_text)
    time_text = normalise_time(time_text)
    if not date_text:
        return pd.NaT
    date_text = re.sub(r"(?i)(\d+)(st|nd|rd|th)", r"\1", date_text)
    text = clean_text(f"{date_text} {time_text}")
    if re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}", text):
        return pd.to_datetime(text, errors="coerce", dayfirst=False)
    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def extract_contact(block: List[str]) -> Tuple[str, str, str]:
    email = ""
    tel = ""
    person = ""
    for i, line in enumerate(block):
        text = clean_text(line)
        if not email:
            match = re.search(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", text, flags=re.I)
            if match:
                email = match.group(0)
        if not tel:
            match = re.search(r"(?:\+27|0)\s?\(?\d{2,3}\)?[\s-]?\d{3}[\s-]?\d{4}", text)
            if match:
                tel = match.group(0)
        if text.upper() in {"ENQUIRES:", "ENQUIRES", "ENQUIRIES:", "ENQUIRIES"} and i + 1 < len(block):
            candidate = clean_text(block[i + 1])
            if candidate and not candidate.lower().startswith("email") and "@" not in candidate:
                person = candidate
    return person, email, tel


def extract_briefing(block: List[str]) -> str:
    parts = []
    capture_terms = ["SITE MEE", "SITE MEETING", "CLARRIFICATION", "CLARIFICATION", "VIRTUAL"]
    for i, line in enumerate(block):
        upper = clean_text(line).upper()
        if any(term in upper for term in capture_terms):
            snippet = [clean_text(line)]
            for j in range(i + 1, min(i + 5, len(block))):
                nxt = clean_text(block[j])
                if not nxt or nxt.startswith("__NECSA_DOC_LINK_"):
                    continue
                if nxt.upper().startswith("BID VALIDITY") or nxt.upper().startswith("DELIVERY ADDRESS"):
                    break
                snippet.append(nxt)
            parts.append(" ".join(snippet))
    return clean_text(" | ".join(parts))


def parse_necsa_html(html: str) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "html.parser")
    soup, link_map = markerise_download_links(soup)
    lines = visible_lines_from_soup(soup)
    bid_indices = [idx for idx, line in enumerate(lines) if clean_text(line).upper().rstrip(":") == "BID NUMBER"]
    rows: List[Dict[str, Any]] = []
    run_ts = datetime.now().isoformat(timespec="seconds")

    for idx, start in enumerate(bid_indices):
        end = bid_indices[idx + 1] if idx + 1 < len(bid_indices) else len(lines)
        block = lines[start:end]
        bid_number = ""
        for j in range(start + 1, min(start + 5, len(lines))):
            candidate = clean_text(lines[j])
            if candidate and not is_major_label(candidate):
                bid_number = candidate
                break
        if not bid_number:
            continue

        description = collect_description(block)
        closing_date_text = label_value(block, "CLOSING DATE")
        closing_time_text = label_value(block, "CLOSING TIME")
        closing_dt = parse_necsa_datetime(closing_date_text, closing_time_text)
        contact_person, contact_email, contact_tel = extract_contact(block)
        briefing = extract_briefing(block)

        doc_names: List[str] = []
        doc_urls: List[str] = []
        for line in block:
            for marker, payload in link_map.items():
                if marker in line:
                    doc_names.append(payload["name"])
                    doc_urls.append(payload["url"])

        # Some pages put the document link immediately before a block. If there is no link
        # in the block, match by bid number in URL/file name.
        if not doc_urls:
            bid_norm = re.sub(r"[^a-z0-9]", "", bid_number.lower())
            for payload in link_map.values():
                url_norm = re.sub(r"[^a-z0-9]", "", (payload.get("url") or "").lower())
                if bid_norm and bid_norm in url_norm:
                    doc_names.append(payload["name"])
                    doc_urls.append(payload["url"])
                    break

        notice_type = "Expression of Interest" if re.search(r"\bEOI\b|expression of interest", f"{bid_number} {description}", re.I) else "Tender"
        is_current = bool(pd.notna(closing_dt) and closing_dt >= pd.Timestamp.now().normalize())
        is_cancelled = bool(re.search(r"cancelled|withdrawn", f"{bid_number} {description}", re.I))
        matched = [kw for kw in KEYWORDS if kw.lower() in description.lower()]
        score = len(set(matched))
        if doc_urls:
            score += 1
        if contact_email:
            score += 1
        if "professional" in description.lower() or "engineering" in description.lower():
            score += 5
        if "nuclear" in description.lower() or "necsa" in description.lower():
            score += 5

        unique_doc_names = list(dict.fromkeys([clean_text(name) for name in doc_names if clean_text(name)]))
        unique_doc_urls = list(dict.fromkeys([clean_text(url) for url in doc_urls if clean_text(url)]))

        rows.append({
            "source": SOURCE_NAME,
            "opportunity_key": f"necsa:{normalise_id(bid_number)}",
            "tender_number": bid_number,
            "reference_number": bid_number,
            "notice_type": notice_type,
            "title": description[:180],
            "description": description,
            "client": CLIENT_NAME,
            "sector": "Nuclear",
            "category": "Services: Professional",
            "country_name": "South Africa",
            "province": "North West",
            "region": "Pelindaba",
            "date_published": "",
            "issue_date": "",
            "briefing_date": "",
            "closing_date": closing_dt.date().isoformat() if pd.notna(closing_dt) else "",
            "closing_datetime": closing_dt.isoformat(sep=" ") if pd.notna(closing_dt) else "",
            "status": "Published",
            "is_current": is_current,
            "is_archived": False,
            "is_cancelled": is_cancelled,
            "document_names": "; ".join(unique_doc_names),
            "document_urls": "; ".join(unique_doc_urls),
            "document_count": len(unique_doc_urls),
            "contact_person": contact_person,
            "contact_email": contact_email,
            "contact_telephone": contact_tel,
            "submission_method": "Physical bid box / email for international suppliers",
            "briefing_session_details": briefing,
            "conditions": "THIS BID IS SUBJECT TO THE PREFERENTIAL PROCUREMENT POLICY FRAMEWORK ACT AND THE PREFERENTIAL PROCUREMENT REGULATIONS, 2022, THE GENERAL CONDITIONS OF CONTRACT (GCC) AND, IF APPLICABLE, ANY OTHER SPECIAL CONDITIONS OF CONTRACT.",
            "raw_category": "Tenders",
            "raw_status": "Published",
            "source_url": SOURCE_URL,
            "run_timestamp": run_ts,
            "row_hash": "",
            "change_type": "",
            "previous_row_hash": "",
            "weight_score": score,
            "matched_description_keywords": "; ".join(sorted(set(matched))),
        })
    return rows


def fetch_and_parse(raw_dir: Path) -> List[Dict[str, Any]]:
    logging.info("Fetching NECSA tender page: %s", SOURCE_URL)
    html = fetch_html(SOURCE_URL)
    save_text(raw_dir / RAW_HTML, html)
    logging.info("Saved raw HTML: %s", raw_dir / RAW_HTML)
    rows = parse_necsa_html(html)
    logging.info("Parsed NECSA tender rows: %s", len(rows))
    return rows


# -----------------------------------------------------------------------------
# Dataframe, change detection and output helpers
# -----------------------------------------------------------------------------
def compute_row_hash(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    if df.empty:
        return pd.Series(dtype="object", index=df.index)
    safe = df.copy()
    for col in cols:
        if col not in safe.columns:
            safe[col] = ""

    def _hash(row: pd.Series) -> str:
        payload = "||".join(clean_text(row.get(col, "")).casefold() for col in cols)
        return hashlib.sha256(payload.encode("utf-8", errors="ignore")).hexdigest()

    return safe.apply(_hash, axis=1)


def ensure_output_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in OUTPUT_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    return out[OUTPUT_COLUMNS]


def normalise_df(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    df = ensure_output_columns(df)
    run_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if not df.empty:
        df["run_timestamp"] = run_ts
        df["row_hash"] = compute_row_hash(df, HASH_COLS)
    return ensure_output_columns(df)


def read_excel_if_exists(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    try:
        return ensure_output_columns(pd.read_excel(path, dtype=object))
    except Exception as exc:
        logging.warning("Could not read existing Excel %s: %s", path, exc)
        return pd.DataFrame(columns=OUTPUT_COLUMNS)


def filter_current(df: pd.DataFrame, include_expired: bool = False) -> pd.DataFrame:
    if include_expired or df.empty:
        return df.copy()
    out = df.copy()
    dt = pd.to_datetime(out.get("closing_datetime", ""), errors="coerce")
    mask = dt.notna() & (dt >= pd.Timestamp.now())
    return out.loc[mask].copy()


def key_for_row(row: pd.Series) -> str:
    return normalise_key(row.get("opportunity_key")) or normalise_key(row.get("reference_number"))


def classify_changes(
    current_df: pd.DataFrame,
    old_master: pd.DataFrame,
    seen_ids: Set[str],
    suppress_alerts: bool,
    top_delta_limit: int,
    top_delta_min_score: int,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    current = ensure_output_columns(current_df)
    if current.empty:
        empty = pd.DataFrame(columns=OUTPUT_COLUMNS)
        return current, empty, empty, empty

    if suppress_alerts:
        logging.info("NECSA hash/state version changed or was not recorded. Re-baselining and suppressing delta/update alerts for this run.")
        current = current.copy()
        current["change_type"] = ""
        current["previous_row_hash"] = ""
        empty = pd.DataFrame(columns=OUTPUT_COLUMNS)
        return current[OUTPUT_COLUMNS], empty, empty, empty

    previous = ensure_output_columns(old_master)
    prev_hash_by_key: Dict[str, str] = {}
    prev_keys: Set[str] = set()
    if not previous.empty:
        for _, row in previous.iterrows():
            key = key_for_row(row)
            if key:
                prev_keys.add(key)
                prev_hash_by_key[key] = clean_text(row.get("row_hash"))

    baseline_keys = set(seen_ids) | prev_keys
    change_types: List[str] = []
    previous_hashes: List[str] = []

    for _, row in current.iterrows():
        key = key_for_row(row)
        old_hash = prev_hash_by_key.get(key, "")
        new_hash = clean_text(row.get("row_hash"))
        previous_hashes.append(old_hash)
        if key and key not in baseline_keys:
            change_types.append("NEW")
        elif old_hash and old_hash != new_hash:
            change_types.append("UPDATED")
        else:
            change_types.append("")

    current = current.copy()
    current["change_type"] = change_types
    current["previous_row_hash"] = previous_hashes

    delta = current[current["change_type"] == "NEW"].copy()
    updates = current[current["change_type"] == "UPDATED"].copy()

    top_delta = delta.copy()
    if not top_delta.empty:
        top_delta["weight_score"] = pd.to_numeric(top_delta["weight_score"], errors="coerce").fillna(0)
        top_delta = top_delta[top_delta["weight_score"] >= int(top_delta_min_score)]
        top_delta = top_delta.sort_values(["weight_score", "closing_datetime"], ascending=[False, True], na_position="last")
        if top_delta_limit and top_delta_limit > 0:
            top_delta = top_delta.head(top_delta_limit)

    return current[OUTPUT_COLUMNS], delta[OUTPUT_COLUMNS], updates[OUTPUT_COLUMNS], ensure_output_columns(top_delta)


def write_excel(path: Path, df: pd.DataFrame, sheet_name: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    out_df = ensure_output_columns(df)
    tmp = path.with_suffix(path.suffix + ".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        out_df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        for idx, col in enumerate(out_df.columns):
            width = min(max(len(str(col)) + 2, 12), 60)
            if col in {"description", "document_urls", "document_names", "briefing_session_details", "source_url", "conditions"}:
                width = 50
            worksheet.set_column(idx, idx, width)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max(len(out_df), 1), max(len(out_df.columns) - 1, 0))
    tmp.replace(path)


def save_outputs(
    raw_dir: Path,
    output_dir: Path,
    raw_df: pd.DataFrame,
    master: pd.DataFrame | None = None,
    delta: pd.DataFrame | None = None,
    updates: pd.DataFrame | None = None,
    top_delta: pd.DataFrame | None = None,
    *,
    write_master: bool = True,
) -> None:
    raw_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    raw_df = ensure_output_columns(raw_df)
    raw_df.to_csv(raw_dir / RAW_CSV, index=False, encoding="utf-8-sig")
    write_excel(raw_dir / RAW_XLSX, raw_df, "Raw")

    empty = pd.DataFrame(columns=OUTPUT_COLUMNS)
    master = ensure_output_columns(master if master is not None else empty)
    delta = ensure_output_columns(delta if delta is not None else empty)
    updates = ensure_output_columns(updates if updates is not None else empty)
    top_delta = ensure_output_columns(top_delta if top_delta is not None else empty)

    if write_master:
        write_excel(output_dir / MASTER_XLSX, master, "Master")
    write_excel(output_dir / DELTA_XLSX, delta, "Delta")
    write_excel(output_dir / UPDATES_XLSX, updates, "Updates")
    write_excel(output_dir / TOP_DELTA_XLSX, top_delta, "Top_Delta")
    logging.info(
        "Saved NECSA outputs | raw=%s master=%s delta=%s updates=%s top_delta=%s",
        len(raw_df), len(master), len(delta), len(updates), len(top_delta),
    )


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Pull and clean NECSA current tenders.")
    parser.add_argument("--base-root", default=None, help="Override source root. Usually not needed in 02_TenderScrapers.")
    parser.add_argument("--include-expired", action="store_true")
    parser.add_argument("--top-delta-limit", type=int, default=20)
    parser.add_argument("--top-delta-min-score", type=int, default=0)
    args = parser.parse_args(argv)

    paths = resolve_paths(args.base_root)
    ensure_dirs(paths)
    setup_logging(paths["log_dir"])

    logging.info("Using tender scrapers root: %s", paths["tender_root"])
    logging.info("Using source root: %s", paths["source_root"])
    logging.info("Using raw directory: %s", paths["raw_dir"])
    logging.info("Using output directory: %s", paths["output_dir"])
    logging.info("Using log directory: %s", paths["log_dir"])
    logging.info("Using state directory: %s", paths["state_dir"])

    lock_path = paths["state_dir"] / LOCK_FILE
    seen_path = paths["state_dir"] / SEEN_IDS_FILE
    hash_path = paths["state_dir"] / HASH_VERSION_FILE
    master_path = paths["output_dir"] / MASTER_XLSX

    try:
        acquire_lock(lock_path)

        rows = fetch_and_parse(paths["raw_dir"])
        raw_df = normalise_df(rows)
        logging.info("Raw NECSA rows: %s | cols: %s", len(raw_df), len(raw_df.columns))

        clean_df = ensure_output_columns(raw_df)
        before = len(clean_df)
        if not clean_df.empty:
            clean_df = clean_df[clean_df["is_cancelled"].astype(str).str.lower() != "true"].copy()
            clean_df = filter_current(clean_df, include_expired=args.include_expired)
        logging.info("Filtered current NECSA tenders: %s -> %s rows", before, len(clean_df))

        previous = read_excel_if_exists(master_path)
        seen_ids = load_seen_ids(seen_path)
        suppress_alerts = load_hash_version(hash_path) != HASH_VERSION

        if raw_df.empty:
            logging.warning("No NECSA rows parsed. Leaving existing master untouched and writing empty delta/update outputs.")
            empty = pd.DataFrame(columns=OUTPUT_COLUMNS)
            save_outputs(paths["raw_dir"], paths["output_dir"], raw_df, delta=empty, updates=empty, top_delta=empty, write_master=False)
            return 0

        master_df, delta_df, updates_df, top_delta_df = classify_changes(
            clean_df,
            previous,
            seen_ids,
            suppress_alerts,
            top_delta_limit=args.top_delta_limit,
            top_delta_min_score=args.top_delta_min_score,
        )

        save_outputs(paths["raw_dir"], paths["output_dir"], raw_df, master_df, delta_df, updates_df, top_delta_df, write_master=True)
        save_seen_ids(seen_path, master_df["opportunity_key"].fillna("").astype(str).tolist())
        save_hash_version(hash_path)

        logging.info("Saved delta Excel: %s | rows: %s", paths["output_dir"] / DELTA_XLSX, len(delta_df))
        logging.info("Saved updates Excel: %s | rows: %s", paths["output_dir"] / UPDATES_XLSX, len(updates_df))
        logging.info("Saved top delta Excel: %s | rows: %s", paths["output_dir"] / TOP_DELTA_XLSX, len(top_delta_df))
        logging.info("Saved master Excel: %s | rows: %s", paths["output_dir"] / MASTER_XLSX, len(master_df))
        return 0

    except RequestException as exc:
        logging.exception("FAILED: %s", exc)
    except Exception as exc:  # deliberately broad so daily orchestration can continue
        logging.exception("FAILED: %s", exc)
    finally:
        release_lock(lock_path)

    # On failure, do not overwrite the existing master. Empty delta/update files confirm the run failed cleanly.
    empty = pd.DataFrame(columns=OUTPUT_COLUMNS)
    save_outputs(paths["raw_dir"], paths["output_dir"], empty, delta=empty, updates=empty, top_delta=empty, write_master=False)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
