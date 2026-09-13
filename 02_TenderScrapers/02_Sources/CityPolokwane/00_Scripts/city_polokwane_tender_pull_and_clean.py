r"""City of Polokwane tender documents scraper.

Scrapes public tender document and bid advertisement pages from:
    https://www.polokwane.gov.za/tender-documents/
    https://www.polokwane.gov.za/bids-advertisements/

This version is aligned to the 02_TenderScrapers architecture:

    02_Sources/CityPolokwane/00_Scripts
    02_Sources/CityPolokwane/01_RawData
    02_Sources/CityPolokwane/02_Outputs
    02_Sources/CityPolokwane/03_Logs
    02_Sources/CityPolokwane/04_State
    02_Sources/CityPolokwane/05_Metadata

The Polokwane site exposes tender document links publicly, but closing dates are
not consistently available in the visible HTML. This scraper therefore preserves
records with missing dates so the central consolidation layer can resolve dates
from matching eTenders records where possible, or isolate unresolved relevant
items in manual-date review.
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
from urllib.parse import urljoin

import pandas as pd
import requests
import urllib3
from bs4 import BeautifulSoup
from requests.exceptions import RequestException, SSLError

SOURCE_NAME = "City of Polokwane"
SOURCE_SLUG = "city_polokwane"
CLIENT_NAME = "City of Polokwane"
COUNTRY_NAME = "South Africa"
PROVINCE = "Limpopo"
REGION = "Polokwane"
SECTOR = "Municipal Infrastructure"
CATEGORY = "Services: Professional"

TENDER_DOCS_URL = "https://www.polokwane.gov.za/tender-documents/"
BIDS_ADVERTS_URL = "https://www.polokwane.gov.za/bids-advertisements/"
SOURCE_URLS = [TENDER_DOCS_URL, BIDS_ADVERTS_URL]

RAW_DIR_NAME = "01_RawData"
OUTPUT_DIR_NAME = "02_Outputs"
LOG_DIR_NAME = "03_Logs"
STATE_DIR_NAME = "04_State"
METADATA_DIR_NAME = "05_Metadata"

RAW_CSV = "city_polokwane_raw.csv"
RAW_XLSX = "city_polokwane_raw_full.xlsx"
RAW_HTML = "city_polokwane_raw_page.html"
MASTER_XLSX = "city_polokwane_master.xlsx"
DELTA_XLSX = "city_polokwane_delta.xlsx"
UPDATES_XLSX = "city_polokwane_updates.xlsx"
TOP_DELTA_XLSX = "city_polokwane_top_delta.xlsx"

SEEN_IDS_JSON = "seen_ids.json"
HASH_VERSION_TXT = "hash_version.txt"
RUN_LOCK = "run.lock"
HASH_VERSION = "city_polokwane_v2_20260704_stable_hash"

REQUEST_TIMEOUT = 90
REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
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

# Keep the row hash stable. Polokwane document URLs can be long and are not always a reliable signal
# for meaningful operational updates, so the hash uses document_names/document_count instead.
HASH_COLS = [
    "reference_number",
    "description",
    "closing_datetime",
    "document_names",
    "document_count",
    "source_url",
    "status",
]

KEYWORDS = [
    "polokwane", "professional services", "professional service provider", "psp", "panel", "consultant",
    "consulting", "consultancy", "engineering", "engineering services", "civil engineering", "structural engineering",
    "mechanical engineering", "electrical engineering", "water", "water and wastewater", "wastewater", "roads",
    "stormwater", "infrastructure", "project management", "programme management", "pmo", "contract administration",
    "construction monitoring", "feasibility", "design", "built environment", "town planning", "spatial planning",
]


def default_source_root() -> Path:
    """Resolve source root from 00_Scripts/script.py -> source folder."""
    return Path(__file__).resolve().parents[1]


def resolve_dirs(source_root: Path) -> Tuple[Path, Path, Path, Path, Path]:
    raw_dir = source_root / RAW_DIR_NAME
    output_dir = source_root / OUTPUT_DIR_NAME
    log_dir = source_root / LOG_DIR_NAME
    state_dir = source_root / STATE_DIR_NAME
    metadata_dir = source_root / METADATA_DIR_NAME
    for path in [raw_dir, output_dir, log_dir, state_dir, metadata_dir]:
        path.mkdir(parents=True, exist_ok=True)
    return raw_dir, output_dir, log_dir, state_dir, metadata_dir


def setup_logging(log_dir: Path) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "city_polokwane_run.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
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


def save_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", errors="ignore")


def acquire_lock(lock_path: Path) -> None:
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    if lock_path.exists():
        raise RuntimeError(f"Lock file exists; another run may be active: {lock_path}")
    lock_path.write_text(f"pid={os.getpid()}\n", encoding="utf-8")


def release_lock(lock_path: Path) -> None:
    try:
        if lock_path.exists():
            lock_path.unlink()
    except OSError:
        logging.exception("Failed to remove lock file: %s", lock_path)


def load_seen_ids(path: Path) -> Set[str]:
    if not path.exists():
        return set()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return {normalise_key(x) for x in data if clean_text(x)}
    except Exception as exc:
        logging.warning("Could not read seen_ids state file %s: %s", path, exc)
        return set()


def save_seen_ids(path: Path, ids: Iterable[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    cleaned = sorted({normalise_key(x) for x in ids if clean_text(x)})
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(cleaned, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def load_hash_version(path: Path) -> str:
    try:
        return clean_text(path.read_text(encoding="utf-8")) if path.exists() else ""
    except Exception:
        return ""


def save_hash_version(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(HASH_VERSION, encoding="utf-8")


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
    if not text or text.lower() in {"n/a", "na", "none", "-"}:
        return pd.NaT
    text = re.sub(r"(?i)(closing|bid closing|closing date|closing time|date)\s*: ?", "", text)
    text = re.sub(r"(?i)\bat\b", " ", text)
    text = re.sub(r"(?i)(\d{1,2})h(\d{2})\s*(am|pm)?", lambda m: f"{m.group(1)}:{m.group(2)} {m.group(3) or ''}", text)
    text = re.sub(r"(?i)(\d{1,2})h\s*(am|pm)?", lambda m: f"{m.group(1)}:00 {m.group(2) or ''}", text)
    text = clean_text(text)
    if re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}", text):
        return pd.to_datetime(text, errors="coerce", dayfirst=False)
    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def extract_date_text(text: str) -> str:
    text = clean_text(text)
    patterns = [
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\s*(?:at|,)\s*\d{1,2}:\d{2}(?::\d{2})?\s*(?:am|pm)?\b",
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\b",
        r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?\b",
        r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            return clean_text(match.group(0))
    return ""


def extract_reference(*parts: str) -> str:
    text = " ".join(clean_text(p) for p in parts if p)
    patterns = [
        r"\bPM\s*\d{1,3}\s*[-/]\s*\d{2,4}\s*[-/]\s*\d{2,4}\b",
        r"\bPM\s*\d{1,3}\s*[-/]\s*\d{2,4}\b",
        r"\b[A-Z]{2,}\s*\d{1,4}\s*/\s*\d{4}\b",
        r"\b[A-Z]{2,}\d{1,4}\s*/\s*\d{4}\b",
        r"\b[A-Z]{2,}[-/]\d{1,4}[-/]\d{2,4}\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            ref = re.sub(r"\s+", "", match.group(0)).upper()
            ref = re.sub(r"[-/]+", "-", ref)
            return ref.strip("-")
    return ""


def infer_notice_type(title: str, desc: str) -> str:
    text = f"{title} {desc}".lower()
    if "expression of interest" in text or "eoi" in text:
        return "EOI"
    if "rfp" in text or "request for proposal" in text:
        return "RFP"
    if "quotation" in text or "rfq" in text:
        return "RFQ"
    return "Tender"


def is_relevant_download_link(text: str, href: str) -> bool:
    joined = f"{text} {href}".lower()
    if not href or href.startswith("#") or href.lower().startswith(("mailto:", "tel:", "javascript:")):
        return False
    if any(ext in joined for ext in DOC_EXTS):
        return True
    phrases = ["tender document", "advert for projects", "erratum", "drawings", "boq", "bid", "pm"]
    return any(phrase in joined for phrase in phrases)


def group_links_from_tender_documents(soup: BeautifulSoup, page_url: str) -> List[Dict[str, Any]]:
    """Group document links by reference where possible."""
    groups: Dict[str, Dict[str, Any]] = {}
    ungrouped_idx = 0

    for a in soup.find_all("a", href=True):
        text = clean_text(a.get_text(" "))
        href = clean_text(a.get("href"))
        if not is_relevant_download_link(text, href):
            continue

        url = urljoin(page_url, href)
        ref = extract_reference(text, href)
        if not ref:
            ungrouped_idx += 1
            ref = f"POLOKWANE-DOC-{ungrouped_idx:03d}"

        group = groups.setdefault(
            ref,
            {"reference_number": ref, "title_parts": [], "document_names": [], "document_urls": []},
        )
        if text and text not in group["title_parts"]:
            group["title_parts"].append(text)
        if url and url not in group["document_urls"]:
            group["document_urls"].append(url)
            group["document_names"].append(text or Path(url).name or "Download")

    rows: List[Dict[str, Any]] = []
    for ref, group in groups.items():
        title = clean_text("; ".join(group["title_parts"][:3]))
        desc_candidates = [part for part in group["title_parts"] if len(part) > 20]
        desc = clean_text(desc_candidates[0] if desc_candidates else title)
        rows.append(
            {
                "reference_number": ref,
                "title": title or ref,
                "description": desc or title or ref,
                "document_names": "; ".join(group["document_names"]),
                "document_urls": "; ".join(group["document_urls"]),
                "document_count": len(group["document_urls"]),
                "source_url": page_url,
                "raw_category": "Tender Documents",
            }
        )
    return rows


def parse_adverts_page(soup: BeautifulSoup, page_url: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for a in soup.find_all("a", href=True):
        text = clean_text(a.get_text(" "))
        href = clean_text(a.get("href"))
        if not text or "advert for projects" not in text.lower():
            continue
        url = urljoin(page_url, href)
        ref = extract_reference(text) or normalise_id(text).upper()
        rows.append(
            {
                "reference_number": ref,
                "title": text,
                "description": text,
                "document_names": text,
                "document_urls": url,
                "document_count": 1,
                "source_url": page_url,
                "raw_category": "Bids Advertisements",
            }
        )
    return rows


def build_rows(raw_dir: Path) -> List[Dict[str, Any]]:
    all_rows: List[Dict[str, Any]] = []
    combined_html: List[str] = []

    for url in SOURCE_URLS:
        logging.info("Fetching Polokwane tender page: %s", url)
        try:
            html = fetch_html(url)
        except RequestException as exc:
            logging.warning("Failed to fetch Polokwane page %s: %s", url, exc)
            continue
        combined_html.append(f"\n<!-- SOURCE: {url} -->\n{html}")
        soup = BeautifulSoup(html, "html.parser")
        if "bids-advertisements" in url:
            rows = parse_adverts_page(soup, url)
        else:
            rows = group_links_from_tender_documents(soup, url)
        logging.info("Parsed Polokwane rows from %s: %s", url, len(rows))
        all_rows.extend(rows)

    save_text(raw_dir / RAW_HTML, "\n".join(combined_html))
    return all_rows


def compute_row_hash(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    if df.empty:
        return pd.Series(dtype="object", index=df.index)

    def _hash(row: pd.Series) -> str:
        vals = [clean_text(row.get(col, "")).lower() for col in cols]
        return hashlib.sha256("||".join(vals).encode("utf-8", errors="ignore")).hexdigest()

    return df.apply(_hash, axis=1)


def score_row(row: pd.Series) -> Tuple[int, str]:
    text = f"{row.get('client','')} {row.get('title','')} {row.get('description','')} {row.get('reference_number','')}".lower()
    matches = [keyword for keyword in KEYWORDS if keyword.lower() in text]
    score = min(40, 10 + 4 * len(matches)) if matches else 10
    if row.get("document_count", 0):
        score += 1
    return score, "; ".join(dict.fromkeys(matches))


def prepare_dataframe(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    now = datetime.now()
    prepared: List[Dict[str, Any]] = []
    seen_keys: set[str] = set()

    for row in rows:
        ref = clean_text(row.get("reference_number"))
        title = clean_text(row.get("title"))
        desc = clean_text(row.get("description")) or title
        doc_urls = clean_text(row.get("document_urls"))
        key_base = ref or title or doc_urls
        opportunity_key = f"city_polokwane:{normalise_id(key_base)}"
        if opportunity_key in seen_keys:
            continue
        seen_keys.add(opportunity_key)

        date_text = extract_date_text(f"{title} {desc}")
        parsed_dt = parse_datetime(date_text) if date_text else pd.NaT
        is_current = bool(pd.isna(parsed_dt) or parsed_dt >= pd.Timestamp.now().normalize())

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
            "briefing_date": "",
            "closing_date": parsed_dt.date().isoformat() if not pd.isna(parsed_dt) else "",
            "closing_datetime": parsed_dt.to_pydatetime().isoformat(sep=" ") if not pd.isna(parsed_dt) else "",
            "status": "Open/Unknown Date" if pd.isna(parsed_dt) else "Open" if is_current else "Expired",
            "is_current": is_current,
            "is_archived": False,
            "is_cancelled": False,
            "document_names": clean_text(row.get("document_names")),
            "document_urls": doc_urls,
            "document_count": int(row.get("document_count") or 0),
            "contact_person": "",
            "contact_email": "",
            "contact_telephone": "",
            "submission_method": "See tender document / municipality instructions",
            "briefing_session_details": "",
            "conditions": "Closing dates are not consistently exposed on the Polokwane tender-document page; eTenders duplicate date resolution may apply.",
            "raw_category": clean_text(row.get("raw_category")),
            "raw_status": "Open/Unknown Date" if pd.isna(parsed_dt) else "Open" if is_current else "Expired",
            "source_url": clean_text(row.get("source_url")) or TENDER_DOCS_URL,
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


def align_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in OUTPUT_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    return out[OUTPUT_COLUMNS]


def save_excel(df: pd.DataFrame, path: Path, sheet_name: str = "Sheet1") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    out = align_columns(df)
    tmp = path.with_suffix(path.suffix + ".tmp.xlsx")
    with pd.ExcelWriter(
        tmp,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
        datetime_format="yyyy-mm-dd hh:mm",
    ) as writer:
        out.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        header_fmt = workbook.add_format({"bold": True, "text_wrap": True, "valign": "top", "border": 1})
        for idx, col in enumerate(out.columns):
            worksheet.write(0, idx, col, header_fmt)
            values = [len(str(col))] + [len(str(x)) for x in out[col].head(200).fillna("")]
            max_len = min(max(values), 80)
            worksheet.set_column(idx, idx, max_len + 2)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max(len(out), 1), max(len(out.columns) - 1, 0))
    tmp.replace(path)


def load_existing_master(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    try:
        df = pd.read_excel(path, dtype=object)
        for col in OUTPUT_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        return df[OUTPUT_COLUMNS]
    except Exception as exc:
        logging.warning("Could not read existing Polokwane master %s: %s", path, exc)
        return pd.DataFrame(columns=OUTPUT_COLUMNS)


def build_delta_updates(
    master_candidate: pd.DataFrame,
    old_master: pd.DataFrame,
    seen_ids: Set[str],
    suppress_alerts: bool,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    master_df = align_columns(master_candidate)
    if master_df.empty:
        return master_df, master_df.copy(), master_df.copy()

    if suppress_alerts:
        master_df["change_type"] = ""
        master_df["previous_row_hash"] = ""
        return master_df, pd.DataFrame(columns=OUTPUT_COLUMNS), pd.DataFrame(columns=OUTPUT_COLUMNS)

    old_hash_by_key: Dict[str, str] = {}
    if not old_master.empty and "opportunity_key" in old_master.columns:
        for _, row in old_master.iterrows():
            key = normalise_key(row.get("opportunity_key"))
            if key:
                old_hash_by_key[key] = clean_text(row.get("row_hash"))

    baseline_ids = set(seen_ids) if seen_ids else set(old_hash_by_key.keys())
    delta_rows = []
    update_rows = []
    full_rows = []

    for _, row in master_df.iterrows():
        record = row.copy()
        key = normalise_key(record.get("opportunity_key"))
        current_hash = clean_text(record.get("row_hash"))
        previous_hash = old_hash_by_key.get(key, "")

        if key not in baseline_ids:
            record["change_type"] = "NEW"
            record["previous_row_hash"] = ""
            delta_rows.append(record.copy())
        elif previous_hash and previous_hash != current_hash:
            record["change_type"] = "UPDATED"
            record["previous_row_hash"] = previous_hash
            update_rows.append(record.copy())
        else:
            record["change_type"] = ""
            record["previous_row_hash"] = previous_hash
        full_rows.append(record.copy())

    full = pd.DataFrame(full_rows) if full_rows else pd.DataFrame(columns=OUTPUT_COLUMNS)
    delta = pd.DataFrame(delta_rows) if delta_rows else pd.DataFrame(columns=OUTPUT_COLUMNS)
    updates = pd.DataFrame(update_rows) if update_rows else pd.DataFrame(columns=OUTPUT_COLUMNS)
    return align_columns(full), align_columns(delta), align_columns(updates)


def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Scrape City of Polokwane public tender documents.")
    parser.add_argument("--base-root", default=None, help="Optional source root override. Defaults to parent of 00_Scripts.")
    parser.add_argument("--top-limit", type=int, default=20)
    args = parser.parse_args(argv)

    source_root = Path(args.base_root).expanduser() if args.base_root else default_source_root()
    raw_dir, output_dir, log_dir, state_dir, metadata_dir = resolve_dirs(source_root)
    setup_logging(log_dir)

    logging.info("Using source root: %s", source_root)
    logging.info("Using raw directory: %s", raw_dir)
    logging.info("Using output directory: %s", output_dir)
    logging.info("Using log directory: %s", log_dir)
    logging.info("Using state directory: %s", state_dir)
    logging.info("Using metadata directory: %s", metadata_dir)

    lock_file = state_dir / RUN_LOCK
    seen_file = state_dir / SEEN_IDS_JSON
    hash_version_file = state_dir / HASH_VERSION_TXT

    try:
        acquire_lock(lock_file)

        try:
            rows = build_rows(raw_dir)
            raw_df = prepare_dataframe(rows)
            logging.info("Parsed City of Polokwane raw rows: %s", len(raw_df))
        except Exception as exc:
            logging.exception("FAILED before output classification: %s", exc)
            raw_df = pd.DataFrame(columns=OUTPUT_COLUMNS)
            save_text(raw_dir / RAW_HTML, "")

        raw_df.to_csv(raw_dir / RAW_CSV, index=False, encoding="utf-8-sig")
        save_excel(raw_df, raw_dir / RAW_XLSX, sheet_name="Raw_Data")

        master_candidate = raw_df.copy()
        if "is_cancelled" in master_candidate.columns:
            master_candidate = master_candidate[
                ~master_candidate["is_cancelled"].astype(str).str.lower().isin(["true", "1", "yes"])
            ].copy()

        old_master = load_existing_master(output_dir / MASTER_XLSX)
        seen_ids = load_seen_ids(seen_file)
        stored_hash_version = load_hash_version(hash_version_file)

        suppress_alerts = False
        if old_master.empty and not seen_ids:
            suppress_alerts = False
        elif not seen_ids:
            logging.info("No seen_ids.json found. Bootstrapping Polokwane state from existing master and suppressing alerts for this run.")
            suppress_alerts = True
        elif stored_hash_version != HASH_VERSION:
            logging.info("Polokwane hash/state version changed or was not recorded. Re-baselining and suppressing delta/update alerts for this run.")
            suppress_alerts = True

        master_df, delta_df, updates_df = build_delta_updates(master_candidate, old_master, seen_ids, suppress_alerts)
        top_delta_df = (
            delta_df.sort_values("weight_score", ascending=False).head(args.top_limit).copy()
            if not delta_df.empty
            else pd.DataFrame(columns=OUTPUT_COLUMNS)
        )

        save_excel(master_df, output_dir / MASTER_XLSX, sheet_name="Master")
        save_excel(delta_df, output_dir / DELTA_XLSX, sheet_name="Delta")
        save_excel(updates_df, output_dir / UPDATES_XLSX, sheet_name="Updates")
        save_excel(top_delta_df, output_dir / TOP_DELTA_XLSX, sheet_name="Top_Delta")

        save_seen_ids(seen_file, master_df.get("opportunity_key", pd.Series(dtype=str)).tolist())
        save_hash_version(hash_version_file)

        logging.info("Saved master Excel: %s | rows: %s", output_dir / MASTER_XLSX, len(master_df))
        logging.info("Saved delta Excel: %s | rows: %s", output_dir / DELTA_XLSX, len(delta_df))
        logging.info("Saved updates Excel: %s | rows: %s", output_dir / UPDATES_XLSX, len(updates_df))
        logging.info("Saved top_delta Excel: %s | rows: %s", output_dir / TOP_DELTA_XLSX, len(top_delta_df))
        return 0

    except (RequestException, ValueError, RuntimeError, OSError) as exc:
        logging.exception("FAILED: %s", exc)
        return 1
    finally:
        release_lock(lock_file)


if __name__ == "__main__":
    raise SystemExit(main())
