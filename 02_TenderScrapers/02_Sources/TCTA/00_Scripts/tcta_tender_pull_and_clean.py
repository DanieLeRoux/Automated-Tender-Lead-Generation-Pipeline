# -*- coding: utf-8 -*-
"""
TCTA Tender Scraper
Source: https://www.tcta.co.za/tenders/

Purpose:
- Fetch public TCTA active tenders page.
- Parse tender blocks from the page heading structure.
- Save raw/master/delta/updates/top-delta outputs using the standard TenderScrapers pattern.

Notes:
- This scraper intentionally captures all parsed active-page rows in raw outputs.
- Master/delta/update outputs are filtered to current tenders by closing_datetime.
- Central TenderScrapers consolidation applies final scoring and relevance filtering.
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
from typing import Dict, Iterable, List, Optional, Tuple
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
from dateutil import parser as dtparser
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


SOURCE_URL = "https://www.tcta.co.za/tenders/"
SOURCE_NAME = "TCTA"
CLIENT_NAME = "Trans-Caledon Tunnel Authority"

# Default location when the script is not being run from the intended 00_Scripts folder.
DEFAULT_SOURCE_ROOT = (
    r"C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics"
    r"\02_TenderScrapers\02_Sources\TCTA"
)

RAW_CSV_NAME = "tcta_raw.csv"
RAW_XLSX_NAME = "tcta_raw_full.xlsx"
RAW_HTML_NAME = "tcta_raw_page.html"
MASTER_XLSX_NAME = "tcta_master.xlsx"
DELTA_XLSX_NAME = "tcta_delta.xlsx"
UPDATES_XLSX_NAME = "tcta_updates.xlsx"
TOP_DELTA_XLSX_NAME = "tcta_top_delta.xlsx"

STATE_FILE_NAME = "row_hash_state.json"
SEEN_IDS_FILE_NAME = "seen_ids.json"
HASH_VERSION_FILE_NAME = "hash_version.txt"
LOCK_FILE_NAME = "run.lock"
HASH_VERSION = "tcta_hash_v2_2026_07_new_structure"

REQUEST_TIMEOUT = 45
REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-ZA,en;q=0.9",
}

STANDARD_COLUMNS = [
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
    "source_original_weight_score",
    "weight_score",
    "matched_description_keywords",
    "run_timestamp",
    "row_hash",
    "change_type",
    "previous_row_hash",
]

# Keep document_urls out of the operational hash because URL encoding/order can create false update noise.
HASH_COLS = [
    "tender_number",
    "reference_number",
    "description",
    "closing_datetime",
    "conditions",
    "status",
]

KEYWORDS = [
    "water",
    "masterplan",
    "master plan",
    "implementation readiness",
    "project management",
    "programme management",
    "construction",
    "advisory",
    "consulting",
    "professional",
    "engineering",
    "environmental",
    "infrastructure",
    "design",
    "feasibility",
    "study",
]

SOURCE_ROOT: Path
TENDER_ROOT: Path
RAW_DIR: Path
OUTPUT_DIR: Path
LOG_DIR: Path
STATE_DIR: Path
METADATA_DIR: Path

RAW_CSV: Path
RAW_XLSX: Path
RAW_HTML: Path
MASTER_XLSX: Path
DELTA_XLSX: Path
UPDATES_XLSX: Path
TOP_DELTA_XLSX: Path
STATE_FILE: Path
SEEN_IDS_FILE: Path
HASH_VERSION_FILE: Path
LOCK_FILE: Path


def clean_text(value: object) -> str:
    if value is None:
        return ""
    try:
        if isinstance(value, float) and pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).replace("\xa0", " ").replace("\u2013", "-").replace("\u2014", "-").replace("\u2019", "'")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalise_id(value: object) -> str:
    text = clean_text(value).lower()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")


def resolve_source_root(cli_source_root: Optional[str] = None, cli_base_root: Optional[str] = None) -> Path:
    if cli_source_root:
        return Path(cli_source_root)
    if cli_base_root:
        return Path(cli_base_root)

    script_path = Path(__file__).resolve()
    if script_path.parent.name == "00_Scripts":
        return script_path.parent.parent

    return Path(DEFAULT_SOURCE_ROOT)


def configure_paths(source_root: Path) -> None:
    global SOURCE_ROOT, TENDER_ROOT, RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR
    global RAW_CSV, RAW_XLSX, RAW_HTML, MASTER_XLSX, DELTA_XLSX, UPDATES_XLSX, TOP_DELTA_XLSX
    global STATE_FILE, SEEN_IDS_FILE, HASH_VERSION_FILE, LOCK_FILE

    SOURCE_ROOT = source_root
    TENDER_ROOT = SOURCE_ROOT.parents[1] if len(SOURCE_ROOT.parents) > 1 and SOURCE_ROOT.parent.name == "02_Sources" else SOURCE_ROOT.parent

    RAW_DIR = SOURCE_ROOT / "01_RawData"
    OUTPUT_DIR = SOURCE_ROOT / "02_Outputs"
    LOG_DIR = SOURCE_ROOT / "03_Logs"
    STATE_DIR = SOURCE_ROOT / "04_State"
    METADATA_DIR = SOURCE_ROOT / "05_Metadata"

    RAW_CSV = RAW_DIR / RAW_CSV_NAME
    RAW_XLSX = RAW_DIR / RAW_XLSX_NAME
    RAW_HTML = RAW_DIR / RAW_HTML_NAME
    MASTER_XLSX = OUTPUT_DIR / MASTER_XLSX_NAME
    DELTA_XLSX = OUTPUT_DIR / DELTA_XLSX_NAME
    UPDATES_XLSX = OUTPUT_DIR / UPDATES_XLSX_NAME
    TOP_DELTA_XLSX = OUTPUT_DIR / TOP_DELTA_XLSX_NAME

    STATE_FILE = STATE_DIR / STATE_FILE_NAME
    SEEN_IDS_FILE = STATE_DIR / SEEN_IDS_FILE_NAME
    HASH_VERSION_FILE = STATE_DIR / HASH_VERSION_FILE_NAME
    LOCK_FILE = STATE_DIR / LOCK_FILE_NAME


def ensure_dirs() -> None:
    for path in [RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR]:
        path.mkdir(parents=True, exist_ok=True)


def setup_logging() -> None:
    ensure_dirs()
    log_path = LOG_DIR / "tcta_run.log"

    root_logger = logging.getLogger()
    root_logger.handlers.clear()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
    )


def build_session() -> requests.Session:
    sess = requests.Session()
    sess.headers.update(REQUEST_HEADERS)

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


def fetch_html(url: str) -> str:
    logging.info("Fetching TCTA tenders: %s", url)
    sess = build_session()
    response = sess.get(url, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    response.encoding = response.encoding or "utf-8"
    return response.text


def save_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", errors="ignore")


def acquire_lock() -> None:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    if LOCK_FILE.exists():
        raise RuntimeError(f"Lock file exists (another run may be active): {LOCK_FILE}")
    LOCK_FILE.write_text(f"pid={os.getpid()}\n", encoding="utf-8")


def release_lock() -> None:
    try:
        if LOCK_FILE.exists():
            LOCK_FILE.unlink()
    except OSError:
        logging.exception("Failed to remove lock file: %s", LOCK_FILE)


def parse_datetime(text: str) -> Optional[pd.Timestamp]:
    text = clean_text(text)
    if not text:
        return None
    text = text.replace(" at ", " ")
    try:
        dt = dtparser.parse(text, dayfirst=True, fuzzy=True)
        return pd.Timestamp(dt)
    except Exception:
        return None


def next_tag_sibling(node: Tag) -> Optional[Tag]:
    sib = node.next_sibling
    while sib is not None:
        if isinstance(sib, Tag):
            return sib
        sib = sib.next_sibling
    return None


def is_tender_title_heading(tag: Tag) -> bool:
    if tag.name is None or tag.name.lower() not in {"h5", "h4"}:
        return False
    text = clean_text(tag.get_text(" ", strip=True))
    if len(text) < 15:
        return False
    upper = text.upper()
    excluded = {
        "TENDERS",
        "CONNECT WITH US",
        "TCTA",
    }
    if upper in excluded:
        return False
    if "DID YOU FIND" in upper or "WERE YOU SATISFIED" in upper:
        return False

    return bool(
        re.search(
            r"\b(APPOINTMENT|CONSULTING|REQUEST|INFORMATION|INVITATION|NFORMATION)\b",
            upper,
        )
    )


def extract_reference(block_text: str) -> str:
    """Extract TCTA tender/reference number from a tender block."""
    text = block_text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    patterns = [
        r"Tender\s*N\^?\{?o\}?\s*[:\-]?\s*([^|]+?)(?=\s+(?:The\s+RFB|The\s+RFQ|The\s+RFI|\d{1,2}\s+[A-Za-z]+\s+\d{4}|Brief\s+Document|Additional\s+Doc|Addendum|Appendix|Cost\s+Sheet|$))",
        r"Tender\s+No\.?\s*[:\-]?\s*([^|]+?)(?=\s+(?:The\s+RFB|The\s+RFQ|The\s+RFI|\d{1,2}\s+[A-Za-z]+\s+\d{4}|Brief\s+Document|Additional\s+Doc|Addendum|Appendix|Cost\s+Sheet|$))",
        r"Tender\s+N\s*o\s*[:\-]?\s*([^|]+?)(?=\s+(?:The\s+RFB|The\s+RFQ|The\s+RFI|\d{1,2}\s+[A-Za-z]+\s+\d{4}|Brief\s+Document|Additional\s+Doc|Addendum|Appendix|Cost\s+Sheet|$))",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            ref = clean_text(match.group(1))
            ref = re.sub(r"(?i)^(Tender\s*)?[:\-\s]*", "", ref).strip()
            return ref
    return ""


def extract_closing(block_text: str) -> Tuple[str, Optional[pd.Timestamp]]:
    patterns = [
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\s+at\s+\d{1,2}:\d{2}:\d{2}\b",
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\s+at\s+\d{1,2}:\d{2}\b",
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, block_text, flags=re.I)
        if match:
            raw = clean_text(match.group(0))
            return raw, parse_datetime(raw)
    return "", None


def infer_notice_type(reference: str, description: str) -> str:
    text = f"{reference} {description}".upper()
    if "RFI" in text or "REQUEST FOR INFORMATION" in text:
        return "Request for Information"
    if "RFQ" in text:
        return "Request for Quotation"
    if "RFP" in text:
        return "Request for Proposal"
    if "RFB" in text or "TENDER" in text:
        return "Request for Bid"
    return "Tender"


def matched_keywords(description: str) -> str:
    text = clean_text(description).lower()
    matches = [kw for kw in KEYWORDS if kw in text]
    return "; ".join(dict.fromkeys(matches))


def source_score(description: str) -> int:
    text = clean_text(description).upper()
    score = 0
    if "TCTA" in text or "TRANS-CALEDON" in text:
        score += 10
    for kw in [
        "WATER",
        "MASTERPLAN",
        "MASTER PLAN",
        "IMPLEMENTATION READINESS",
        "PROJECT MANAGEMENT",
        "CONSTRUCTION",
        "ADVISORY",
        "CONSULTING",
        "PROFESSIONAL",
        "ENGINEERING",
    ]:
        if kw in text:
            score += 3
    return score


def collect_document_order_block(heading: Tag) -> Tuple[str, List[str], List[str]]:
    """Collect visible text and links from this tender heading until the next tender heading."""
    parts: List[str] = []
    links: List[str] = []
    link_names: List[str] = []

    for node in heading.next_elements:
        if node is heading:
            continue
        if isinstance(node, Tag):
            if node is not heading and is_tender_title_heading(node):
                break

            node_text = clean_text(node.get_text(" ", strip=True))
            if node.name and node.name.lower() in {"h1", "h2", "h3"} and node_text.lower() in {"let us know", "connect with us"}:
                break

            if node.name and node.name.lower() == "a" and node.get("href"):
                href = urljoin(SOURCE_URL, node.get("href", ""))
                name = clean_text(node.get_text(" ", strip=True)) or "Document"
                href_l = href.lower()

                generic_etenders = "etenders.gov.za" in href_l and (
                    href_l.rstrip("/").endswith("etenders.gov.za")
                    or href_l.endswith("/#")
                    or href_l.endswith("#")
                )
                if href and not href.endswith("#") and not generic_etenders:
                    if "tcta.co.za" in href_l or any(ext in href_l for ext in [".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip"]):
                        if href not in links:
                            links.append(href)
                            link_names.append(name)
        else:
            txt = clean_text(node)
            if txt:
                parts.append(txt)

    text = "\n".join(parts)
    text = re.sub(r"\n{2,}", "\n", text).strip()
    return text, link_names, links


def make_tcta_row(description: str, block_text: str, link_names: List[str], links: List[str], run_ts: str) -> Dict[str, object]:
    reference = extract_reference(block_text)
    _, closing_dt = extract_closing(block_text)

    key_src = reference or description
    opportunity_key = f"tcta:{normalise_id(key_src)[:120]}"

    source_original_weight_score = source_score(description)
    keyword_text = matched_keywords(description)

    return {
        "source": SOURCE_NAME,
        "opportunity_key": opportunity_key,
        "tender_number": reference,
        "reference_number": reference,
        "notice_type": infer_notice_type(reference, description),
        "title": description,
        "description": description,
        "client": CLIENT_NAME,
        "sector": "Water",
        "category": "Tenders",
        "country_name": "South Africa",
        "province": "National",
        "region": "South Africa",
        "date_published": "",
        "issue_date": "",
        "briefing_date": "",
        "closing_date": closing_dt.date().isoformat() if closing_dt is not None and pd.notna(closing_dt) else "",
        "closing_datetime": closing_dt.isoformat(sep=" ") if closing_dt is not None and pd.notna(closing_dt) else "",
        "status": "Published",
        "is_current": bool(closing_dt is not None and pd.notna(closing_dt) and closing_dt >= pd.Timestamp.now()),
        "is_archived": False,
        "is_cancelled": False,
        "document_names": "; ".join(dict.fromkeys(link_names)),
        "document_urls": "; ".join(dict.fromkeys(links)),
        "document_count": len(dict.fromkeys(links)),
        "contact_person": "",
        "contact_email": "",
        "contact_telephone": "",
        "submission_method": "See tender documents / source page",
        "briefing_session_details": "",
        "conditions": "",
        "raw_category": "",
        "raw_status": "Published",
        "source_url": SOURCE_URL,
        "source_original_weight_score": source_original_weight_score,
        "weight_score": source_original_weight_score,
        "matched_description_keywords": keyword_text,
        "run_timestamp": run_ts,
        "row_hash": "",
        "change_type": "",
        "previous_row_hash": "",
    }


def parse_tcta_rows(html: str) -> List[Dict[str, object]]:
    soup = BeautifulSoup(html, "html.parser")
    rows: List[Dict[str, object]] = []
    run_ts = datetime.now().replace(microsecond=0).isoformat(sep=" ")

    headings = [h for h in soup.find_all(["h5", "h4"]) if is_tender_title_heading(h)]
    logging.info("Candidate TCTA tender title headings found: %s", len(headings))

    for heading in headings:
        description = clean_text(heading.get_text(" ", strip=True))
        block_text, link_names, links = collect_document_order_block(heading)
        block_text = "\n".join([description, block_text]).strip()

        if not description:
            continue

        rows.append(make_tcta_row(description, block_text, link_names, links, run_ts))

    if not rows:
        logging.warning("No rows parsed from headings. Trying text fallback parser.")
        visible = soup.get_text("\n", strip=True)
        chunks = re.split(r"\n\s*\*\s*\*\s*\*\s*\n|\* \* \*", visible)
        for chunk in chunks:
            if not re.search(r"Tender\s+N", chunk, flags=re.I):
                continue
            lines = [clean_text(x) for x in chunk.split("\n") if clean_text(x)]
            if not lines:
                continue
            description = lines[0]
            rows.append(make_tcta_row(description, chunk, [], [], run_ts))

    deduped: List[Dict[str, object]] = []
    seen = set()
    for row in rows:
        key = clean_text(row.get("reference_number", "")) or clean_text(row.get("description", ""))
        key = key.lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(row)

    with_dates = sum(1 for row in deduped if clean_text(row.get("closing_datetime", "")))
    logging.info("Parsed TCTA tender rows: %s | rows with closing dates: %s", len(deduped), with_dates)
    return deduped


def ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in STANDARD_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    return out[STANDARD_COLUMNS]


def compute_row_hash(row: pd.Series) -> str:
    parts = [clean_text(row.get(col, "")) for col in HASH_COLS]
    payload = "||".join(parts)
    return hashlib.sha256(payload.encode("utf-8", errors="ignore")).hexdigest()


def add_hashes(df: pd.DataFrame) -> pd.DataFrame:
    out = ensure_columns(df)
    if not out.empty:
        out["row_hash"] = out.apply(compute_row_hash, axis=1)
    return out


def load_state() -> Dict[str, str]:
    if not STATE_FILE.exists():
        return {}
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return {str(k): str(v) for k, v in data.items()}
    except Exception:
        logging.warning("Could not load TCTA row hash state file; starting fresh.")
    return {}


def save_state(state: Dict[str, str]) -> None:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    tmp = STATE_FILE.with_suffix(STATE_FILE.suffix + ".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(STATE_FILE)


def save_seen_ids(ids: Iterable[str]) -> None:
    seen = sorted({clean_text(x) for x in ids if clean_text(x)})
    tmp = SEEN_IDS_FILE.with_suffix(SEEN_IDS_FILE.suffix + ".tmp")
    tmp.write_text(json.dumps(seen, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(SEEN_IDS_FILE)


def hash_version_changed() -> bool:
    if not HASH_VERSION_FILE.exists():
        return True
    try:
        return clean_text(HASH_VERSION_FILE.read_text(encoding="utf-8")) != HASH_VERSION
    except OSError:
        return True


def save_hash_version() -> None:
    HASH_VERSION_FILE.write_text(HASH_VERSION + "\n", encoding="utf-8")


def load_existing_master() -> pd.DataFrame:
    if not MASTER_XLSX.exists():
        return ensure_columns(pd.DataFrame())
    try:
        return ensure_columns(pd.read_excel(MASTER_XLSX, dtype=object))
    except Exception as exc:
        logging.warning("Could not read existing master %s: %s", MASTER_XLSX, exc)
        return ensure_columns(pd.DataFrame())


def write_excel(df: pd.DataFrame, path: Path, sheet_name: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    out_df = ensure_columns(df) if path.name != RAW_XLSX_NAME else df.copy()
    if out_df.shape[1] == 0:
        out_df = pd.DataFrame(columns=STANDARD_COLUMNS)

    tmp = path.with_suffix(path.suffix + ".tmp.xlsx")
    with pd.ExcelWriter(
        tmp,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd hh:mm:ss",
        date_format="yyyy-mm-dd",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        out_df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max(len(out_df), 1), max(len(out_df.columns) - 1, 0))

        for col_num, col in enumerate(out_df.columns):
            width = min(max(len(str(col)) + 2, 12), 55)
            if col in {"description", "title", "document_urls", "document_names", "source_url"}:
                width = 50
            worksheet.set_column(col_num, col_num, width)

    tmp.replace(path)


def save_outputs(
    raw_df: pd.DataFrame,
    current_df: pd.DataFrame,
    delta_df: pd.DataFrame,
    updates_df: pd.DataFrame,
    top_delta_df: pd.DataFrame,
) -> None:
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    raw_df.to_csv(RAW_CSV, index=False, encoding="utf-8-sig")
    write_excel(raw_df, RAW_XLSX, "Raw")
    write_excel(current_df, MASTER_XLSX, "Master")
    write_excel(delta_df, DELTA_XLSX, "Delta")
    write_excel(updates_df, UPDATES_XLSX, "Updates")
    write_excel(top_delta_df, TOP_DELTA_XLSX, "Top_Delta")


def empty_df() -> pd.DataFrame:
    return ensure_columns(pd.DataFrame())


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Scrape TCTA active tenders.")
    parser.add_argument("--source-root", default=None, help="Override the TCTA source root folder.")
    parser.add_argument("--base-root", default=None, help="Backward-compatible alias for --source-root.")
    parser.add_argument("--include-expired", action="store_true", help="Keep expired tenders in source master.")
    parser.add_argument("--top-delta-limit", type=int, default=20)
    parser.add_argument("--top-delta-min-score", type=int, default=0)
    args = parser.parse_args(argv)

    configure_paths(resolve_source_root(args.source_root, args.base_root))
    setup_logging()

    logging.info("Using tender scrapers root: %s", TENDER_ROOT)
    logging.info("Using source root: %s", SOURCE_ROOT)
    logging.info("Using raw directory: %s", RAW_DIR)
    logging.info("Using output directory: %s", OUTPUT_DIR)
    logging.info("Using log directory: %s", LOG_DIR)
    logging.info("Using state directory: %s", STATE_DIR)

    try:
        acquire_lock()

        try:
            html = fetch_html(SOURCE_URL)
            save_text(RAW_HTML, html)
            logging.info("Saved raw HTML: %s", RAW_HTML)
            rows = parse_tcta_rows(html)
        except Exception as exc:
            logging.exception("TCTA scrape failed; preserving existing master and writing empty daily exception outputs: %s", exc)
            existing_master = load_existing_master()
            empty = empty_df()
            save_outputs(empty, existing_master, empty, empty, empty)
            return 0

        raw_df = add_hashes(pd.DataFrame(rows))
        logging.info("Parsed TCTA raw rows: %s | cols: %s", len(raw_df), len(raw_df.columns))

        if raw_df.empty:
            current_df = empty_df()
        else:
            current_df = raw_df.copy()
            if not args.include_expired:
                current_df["_closing_datetime_parsed"] = pd.to_datetime(current_df["closing_datetime"], errors="coerce")
                before = len(current_df)
                current_df = current_df[
                    current_df["_closing_datetime_parsed"].notna()
                    & (current_df["_closing_datetime_parsed"] >= pd.Timestamp.now())
                ].copy()
                current_df = current_df.drop(columns=["_closing_datetime_parsed"])
                logging.info("Filtered current TCTA tenders: %s -> %s rows", before, len(current_df))
            current_df = add_hashes(current_df)

        old_state = load_state()
        new_state: Dict[str, str] = {}
        delta_rows = []
        update_rows = []

        suppress_alerts = hash_version_changed()
        if suppress_alerts:
            logging.info("TCTA hash/state version changed or was not recorded. Re-baselining and suppressing delta/update alerts for this run.")

        if not current_df.empty:
            for _, row in current_df.iterrows():
                key = clean_text(row.get("opportunity_key", ""))
                row_hash = clean_text(row.get("row_hash", "")) or compute_row_hash(row)
                new_state[key] = row_hash

                old_hash = old_state.get(key)
                if old_hash is None:
                    r = row.copy()
                    r["change_type"] = "NEW"
                    delta_rows.append(r)
                elif old_hash != row_hash:
                    r = row.copy()
                    r["change_type"] = "UPDATED"
                    r["previous_row_hash"] = old_hash
                    update_rows.append(r)

        if suppress_alerts:
            delta_df = empty_df()
            updates_df = empty_df()
        else:
            delta_df = ensure_columns(pd.DataFrame(delta_rows))
            updates_df = ensure_columns(pd.DataFrame(update_rows))

        if not delta_df.empty:
            delta_df["weight_score"] = pd.to_numeric(delta_df["weight_score"], errors="coerce").fillna(0)
            delta_df = delta_df.sort_values(["weight_score", "closing_datetime"], ascending=[False, True], na_position="last")
            top_delta_df = delta_df[delta_df["weight_score"] >= int(args.top_delta_min_score)].head(int(args.top_delta_limit)).copy()
        else:
            top_delta_df = empty_df()

        current_df = ensure_columns(current_df)
        save_outputs(raw_df, current_df, delta_df, updates_df, top_delta_df)
        save_state(new_state)
        save_seen_ids(new_state.keys())
        save_hash_version()

        logging.info("Saved raw CSV: %s | rows: %s", RAW_CSV, len(raw_df))
        logging.info("Saved master Excel: %s | rows: %s", MASTER_XLSX, len(current_df))
        logging.info("Saved delta Excel: %s | rows: %s", DELTA_XLSX, len(delta_df))
        logging.info("Saved updates Excel: %s | rows: %s", UPDATES_XLSX, len(updates_df))
        logging.info("Saved top delta Excel: %s | rows: %s", TOP_DELTA_XLSX, len(top_delta_df))
        return 0

    finally:
        release_lock()


if __name__ == "__main__":
    raise SystemExit(main())
