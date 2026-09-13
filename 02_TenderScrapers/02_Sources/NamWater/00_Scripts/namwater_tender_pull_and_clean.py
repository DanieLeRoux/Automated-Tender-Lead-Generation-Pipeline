r"""NamWater public bids-list scraper.

Scrapes the public NamWater bids-list page:
    https://www.namwater.com.na/index.php/bids-list

Important limitation:
The visible bids-list page exposes bid numbers, descriptions and document links,
but it does not currently expose closing dates in the listing HTML. This scraper
therefore preserves records with missing closing dates instead of dropping them.
The central TenderScrapers consolidation layer can then keep them in the master
and route relevant unresolved items to Needs Manual Date Check.
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
from requests.exceptions import SSLError, RequestException
import urllib3
from bs4 import BeautifulSoup

# -----------------------------------------------------------------------------
# Source configuration
# -----------------------------------------------------------------------------
SOURCE_NAME = "NamWater"
CLIENT_NAME = "NamWater"
COUNTRY_NAME = "Namibia"
PROVINCE = ""
REGION = "Namibia"
SECTOR = "Water"
CATEGORY = "Services: Professional"
SOURCE_URL = "https://www.namwater.com.na/index.php/bids-list"

# Dynamic TenderScrapers folder detection.
# Expected location:
#   ...\02_TenderScrapers\02_Sources\NamWater\00_Scripts\namwater_tender_pull_and_clean.py
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

RAW_CSV = "namwater_raw.csv"
RAW_XLSX = "namwater_raw_full.xlsx"
RAW_HTML = "namwater_raw_page.html"
MASTER_XLSX = "namwater_master.xlsx"
DELTA_XLSX = "namwater_delta.xlsx"
UPDATES_XLSX = "namwater_updates.xlsx"
TOP_DELTA_XLSX = "namwater_top_delta.xlsx"
LOG_FILE = "namwater_run.log"
SEEN_IDS_FILE = "seen_ids.json"
HASH_VERSION_FILE = "hash_version.txt"
LOCK_FILE = "run.lock"

# Increment this only when the definition of row_hash/change detection changes.
HASH_VERSION = "namwater_v2_stable_hash_no_document_urls"

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

# Keep the hash stable. NamWater document URLs can include CMS paths or filenames that may change
# without the underlying notice changing, so full document_urls are deliberately excluded.
HASH_COLS = ["reference_number", "description", "closing_datetime", "document_names", "document_count", "source_url", "status"]

KEYWORDS = [
    "professional services", "professional service provider", "psp", "consultant", "consulting", "consultancy",
    "engineering", "engineering services", "civil engineering", "mechanical engineering", "electrical engineering",
    "water", "bulk water", "water supply", "water quality", "master plan", "integrated bulk water master plan",
    "rehabilitation", "boreholes", "sludge", "treatment", "scheme", "feasibility", "design", "project management",
]

BID_REF_RE = re.compile(r"\b[A-Z]{1,5}(?:/[A-Z]{1,5})+/NW[-–]\d{3}/\d{4}\b", re.I)


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
    return re.sub(r"\s+", " ", text).strip()


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
def fetch_html(url: str) -> str:
    session = requests.Session()
    session.headers.update(REQUEST_HEADERS)
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


def parse_datetime(value: Any) -> pd.Timestamp:
    text = clean_text(value)
    if not text or text.lower() in {"n/a", "na", "none", "-"}:
        return pd.NaT
    text = re.sub(r"(?i)(closing|bid closing|closing date|closing time|date)\s*: ?", "", text)
    text = re.sub(r"(?i)\bat\b", " ", text)
    text = re.sub(r"(?i)(\d{1,2})h(\d{2})\s*(am|pm)?", lambda m: f"{m.group(1)}:{m.group(2)} {m.group(3) or ''}", text)
    text = clean_text(text)
    if re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}", text):
        return pd.to_datetime(text, errors="coerce", dayfirst=False)
    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def infer_notice_type(ref: str, desc: str) -> str:
    text = f"{ref} {desc}".lower()
    if "rfq" in text or "quotation" in text:
        return "RFQ"
    if "rfp" in text or "request for proposal" in text or "selection of consultant" in text:
        return "RFP"
    return "Tender"


def is_doc_link(text: str, href: str) -> bool:
    joined = f"{text} {href}".lower()
    if not href or href.startswith("#") or href.lower().startswith("mailto:"):
        return False
    if "print" in joined or "email" in joined:
        return False
    if "download" in joined:
        return True
    return any(ext in joined for ext in DOC_EXTS)


def filename_from_url(url: str, fallback: str) -> str:
    path = unquote(urlparse(url).path)
    name = Path(path).name
    return name or fallback or "Download"


def split_bid_line(line: str) -> Tuple[str, str]:
    match = BID_REF_RE.search(line)
    if not match:
        return "", line
    ref = match.group(0).replace("–", "-").upper()
    rest = clean_text(line[:match.start()] + " " + line[match.end():])
    return ref, rest


def parse_from_text_lines(soup: BeautifulSoup, page_url: str) -> List[Dict[str, Any]]:
    full_text = soup.get_text("\n")
    lines = [clean_text(x) for x in full_text.splitlines()]
    lines = [x for x in lines if x]

    start = 0
    for idx, line in enumerate(lines):
        if line.lower().startswith("bid number") and "description" in line.lower():
            start = idx + 1
            break
    end = len(lines)
    for idx in range(start, len(lines)):
        if lines[idx].lower().startswith("who are we"):
            end = idx
            break
    relevant = lines[start:end]

    records: List[Dict[str, Any]] = []
    current: Dict[str, Any] | None = None
    noise = {"download", "documents", "description", "details", "print", "email"}

    for line in relevant:
        ref, remainder = split_bid_line(line)
        if ref:
            if current:
                records.append(current)
            current = {"reference_number": ref, "description_parts": []}
            if remainder and remainder.lower() not in noise:
                current["description_parts"].append(remainder)
            continue
        if current:
            low = line.lower()
            if low in noise or low.startswith("bid number"):
                continue
            if low.startswith("home") or low.startswith("about us"):
                continue
            current["description_parts"].append(line)
    if current:
        records.append(current)

    download_links: List[Tuple[str, str]] = []
    for a in soup.find_all("a", href=True):
        text = clean_text(a.get_text(" "))
        href = clean_text(a.get("href"))
        if not is_doc_link(text, href):
            continue
        url = urljoin(page_url, href)
        download_links.append((text or filename_from_url(url, "Download"), url))

    rows: List[Dict[str, Any]] = []
    for idx, rec in enumerate(records):
        ref = clean_text(rec.get("reference_number"))
        desc = clean_text(" ".join(rec.get("description_parts", [])))
        desc = clean_text(BID_REF_RE.sub("", desc))
        desc = re.sub(r"(?i)\bdownload\b", "", desc)
        desc = clean_text(desc)
        doc_names: List[str] = []
        doc_urls: List[str] = []
        if idx < len(download_links):
            text, url = download_links[idx]
            doc_urls = [url]
            doc_names = [filename_from_url(url, text or f"{ref} Download")]
        rows.append(make_row(ref, desc, doc_names, doc_urls, page_url))
    return rows


def parse_table_rows(soup: BeautifulSoup, page_url: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for tr in soup.find_all("tr"):
        text = clean_text(tr.get_text(" "))
        if not BID_REF_RE.search(text):
            continue
        ref, remainder = split_bid_line(text)
        if not ref:
            continue
        desc = clean_text(re.sub(r"(?i)\b(bid number|description|documents|download)\b", " ", remainder))
        doc_names: List[str] = []
        doc_urls: List[str] = []
        for a in tr.find_all("a", href=True):
            a_text = clean_text(a.get_text(" "))
            href = clean_text(a.get("href"))
            if not is_doc_link(a_text, href):
                continue
            url = urljoin(page_url, href)
            doc_urls.append(url)
            doc_names.append(filename_from_url(url, a_text or f"{ref} Download"))
        rows.append(make_row(ref, desc, doc_names, doc_urls, page_url))
    return rows


def make_row(ref: str, desc: str, doc_names: List[str], doc_urls: List[str], page_url: str) -> Dict[str, Any]:
    ref = clean_text(ref).upper()
    desc = clean_text(desc)
    title = f"{ref} - {desc}" if desc else ref
    key = f"namwater:{normalise_id(ref or desc)}"
    matched = [kw for kw in KEYWORDS if kw.lower() in f"{ref} {desc}".lower()]
    weight_score = min(20, len(matched) * 2 + 6) if matched else 0
    unique_names = list(dict.fromkeys([clean_text(x) for x in doc_names if clean_text(x)]))
    unique_urls = list(dict.fromkeys([clean_text(x) for x in doc_urls if clean_text(x)]))
    return {
        "source": SOURCE_NAME,
        "opportunity_key": key,
        "tender_number": ref,
        "reference_number": ref,
        "notice_type": infer_notice_type(ref, desc),
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
        "briefing_date": "",
        "closing_date": "",
        "closing_datetime": "",
        "status": "Open / closing date not shown on source page",
        "is_current": True,
        "is_archived": False,
        "is_cancelled": False,
        "document_names": "; ".join(unique_names),
        "document_urls": "; ".join(unique_urls),
        "document_count": len(unique_urls),
        "contact_person": "",
        "contact_email": "",
        "contact_telephone": "+264 61 71 0000",
        "submission_method": "See NamWater bid document / source page",
        "briefing_session_details": "",
        "conditions": "Closing date is not exposed on the NamWater bids-list page; check bid document/source page manually.",
        "raw_category": "Bids List",
        "raw_status": "Closing date not shown",
        "source_url": page_url,
        "run_timestamp": "",
        "row_hash": "",
        "change_type": "",
        "previous_row_hash": "",
        "weight_score": weight_score,
        "matched_description_keywords": "; ".join(matched),
    }


def fetch_and_parse(raw_dir: Path) -> List[Dict[str, Any]]:
    logging.info("Fetching NamWater bids list: %s", SOURCE_URL)
    html = fetch_html(SOURCE_URL)
    save_text(raw_dir / RAW_HTML, html)
    soup = BeautifulSoup(html, "html.parser")

    rows = parse_table_rows(soup, SOURCE_URL)
    if not rows:
        rows = parse_from_text_lines(soup, SOURCE_URL)

    dedup: Dict[str, Dict[str, Any]] = {}
    for row in rows:
        key = normalise_id(row.get("reference_number") or row.get("description"))
        if not key:
            continue
        if key not in dedup:
            dedup[key] = row
            continue
        old = dedup[key]
        if len(str(row.get("description", ""))) > len(str(old.get("description", ""))):
            old["description"] = row.get("description", old.get("description"))
            old["title"] = row.get("title", old.get("title"))
        urls = [u.strip() for u in f"{old.get('document_urls','')}; {row.get('document_urls','')}".split(";") if u.strip()]
        names = [n.strip() for n in f"{old.get('document_names','')}; {row.get('document_names','')}".split(";") if n.strip()]
        old["document_urls"] = "; ".join(dict.fromkeys(urls))
        old["document_names"] = "; ".join(dict.fromkeys(names))
        old["document_count"] = len(dict.fromkeys(urls))
    out = list(dedup.values())
    logging.info("Parsed NamWater rows: %s", len(out))
    return out


# -----------------------------------------------------------------------------
# Dataframe, change detection and output helpers
# -----------------------------------------------------------------------------
def compute_row_hash(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    if df.empty:
        return pd.Series(dtype="object", index=df.index)

    use_cols = [col for col in cols if col in df.columns]

    def _hash(row: pd.Series) -> str:
        payload = "||".join(clean_text(row.get(col, "")).casefold() for col in use_cols)
        return hashlib.sha256(payload.encode("utf-8", errors="ignore")).hexdigest()

    return df.apply(_hash, axis=1)


def add_missing_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in OUTPUT_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    return out[OUTPUT_COLUMNS]


def normalise_df(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    df = add_missing_columns(df)
    run_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if not df.empty:
        df["run_timestamp"] = run_ts
        df["row_hash"] = compute_row_hash(df, HASH_COLS)
    return add_missing_columns(df)


def load_previous(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    try:
        return add_missing_columns(pd.read_excel(path, dtype=object))
    except Exception as exc:
        logging.warning("Could not read previous master %s: %s", path, exc)
        return pd.DataFrame(columns=OUTPUT_COLUMNS)


def key_for_row(row: pd.Series) -> str:
    return normalise_key(row.get("opportunity_key")) or normalise_key(row.get("reference_number"))


def classify_changes(
    current: pd.DataFrame,
    previous: pd.DataFrame,
    seen_ids: Set[str],
    suppress_alerts: bool,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    current = add_missing_columns(current)
    if current.empty:
        empty = pd.DataFrame(columns=OUTPUT_COLUMNS)
        return current, empty, empty, empty

    if suppress_alerts:
        logging.info("NamWater hash/state version changed or was not recorded. Re-baselining and suppressing delta/update alerts for this run.")
        current = current.copy()
        current["change_type"] = ""
        current["previous_row_hash"] = ""
        empty = pd.DataFrame(columns=OUTPUT_COLUMNS)
        return current[OUTPUT_COLUMNS], empty, empty, empty

    previous = add_missing_columns(previous)
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
    prev_hashes: List[str] = []

    for _, row in current.iterrows():
        key = key_for_row(row)
        old_hash = prev_hash_by_key.get(key, "")
        new_hash = clean_text(row.get("row_hash"))
        prev_hashes.append(old_hash)
        if key and key not in baseline_keys:
            change_types.append("NEW")
        elif old_hash and old_hash != new_hash:
            change_types.append("UPDATED")
        else:
            change_types.append("")

    current = current.copy()
    current["change_type"] = change_types
    current["previous_row_hash"] = prev_hashes
    delta = current[current["change_type"] == "NEW"].copy()
    updates = current[current["change_type"] == "UPDATED"].copy()
    top_delta = delta.sort_values(["weight_score", "reference_number"], ascending=[False, True]).head(20).copy() if not delta.empty else delta.copy()
    return current[OUTPUT_COLUMNS], delta[OUTPUT_COLUMNS], updates[OUTPUT_COLUMNS], top_delta[OUTPUT_COLUMNS]


def write_excel(df: pd.DataFrame, path: Path, sheet_name: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    out_df = add_missing_columns(df)
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

    raw_df = add_missing_columns(raw_df)
    raw_df.to_csv(raw_dir / RAW_CSV, index=False, encoding="utf-8-sig")
    write_excel(raw_df, raw_dir / RAW_XLSX, "Raw")

    empty = pd.DataFrame(columns=OUTPUT_COLUMNS)
    master = add_missing_columns(master if master is not None else empty)
    delta = add_missing_columns(delta if delta is not None else empty)
    updates = add_missing_columns(updates if updates is not None else empty)
    top_delta = add_missing_columns(top_delta if top_delta is not None else empty)

    if write_master:
        write_excel(master, output_dir / MASTER_XLSX, "Master")
    write_excel(delta, output_dir / DELTA_XLSX, "Delta")
    write_excel(updates, output_dir / UPDATES_XLSX, "Updates")
    write_excel(top_delta, output_dir / TOP_DELTA_XLSX, "Top_Delta")
    logging.info(
        "Saved NamWater outputs | raw=%s master=%s delta=%s updates=%s top_delta=%s",
        len(raw_df), len(master), len(delta), len(updates), len(top_delta),
    )


def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Scrape NamWater public bids-list tenders.")
    parser.add_argument("--base-root", default=None, help="Override source root. Usually not needed in 02_TenderScrapers.")
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
        logging.info("Raw NamWater rows: %s | cols: %s", len(raw_df), len(raw_df.columns))

        previous = load_previous(master_path)
        seen_ids = load_seen_ids(seen_path)
        suppress_alerts = load_hash_version(hash_path) != HASH_VERSION

        if raw_df.empty:
            logging.warning("No NamWater rows parsed. Leaving existing master untouched and writing empty delta/update outputs.")
            save_outputs(paths["raw_dir"], paths["output_dir"], raw_df, delta=pd.DataFrame(columns=OUTPUT_COLUMNS), updates=pd.DataFrame(columns=OUTPUT_COLUMNS), top_delta=pd.DataFrame(columns=OUTPUT_COLUMNS), write_master=False)
            return 0

        master, delta, updates, top_delta = classify_changes(raw_df, previous, seen_ids, suppress_alerts)
        save_outputs(paths["raw_dir"], paths["output_dir"], raw_df, master, delta, updates, top_delta, write_master=True)

        save_seen_ids(seen_path, master["opportunity_key"].fillna("").astype(str).tolist())
        save_hash_version(hash_path)
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
