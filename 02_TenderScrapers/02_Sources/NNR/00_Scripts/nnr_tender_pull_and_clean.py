r"""NNR advertised tenders scraper.

Scrapes open tenders from:
    https://nnr.co.za/tenders/

Output folder:
    C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\NNR\Outputs

Outputs:
    nnr_raw.csv
    nnr_raw_full.xlsx
    nnr_raw_page.html
    nnr_master.xlsx
    nnr_delta.xlsx
    nnr_updates.xlsx
    nnr_top_delta.xlsx

The NNR page contains several sections: Open Tenders, Closed Tenders, Received Tenders,
Awarded Tenders and Cancelled Tenders. This scraper only treats records found in the
Open Tenders section as operational tender records. Historical/awarded/cancelled text is
left out of the cleaned source outputs.
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
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple
from urllib.parse import urljoin, urlparse, unquote

import pandas as pd
import requests
from bs4 import BeautifulSoup

# === Source and folder configuration ===
SOURCE_URL = "https://nnr.co.za/tenders/"
SOURCE_NAME = "NNR"
CLIENT_NAME = "National Nuclear Regulator"

SCRIPT_PATH = Path(__file__).resolve()
SOURCE_ROOT = SCRIPT_PATH.parents[1] if len(SCRIPT_PATH.parents) >= 2 else Path.cwd()
TENDER_ROOT = SCRIPT_PATH.parents[3] if len(SCRIPT_PATH.parents) >= 4 else SOURCE_ROOT.parents[1]
BASE_ROOT = str(SOURCE_ROOT)

RAW_DIR = "01_RawData"
OUTPUT_DIR = "02_Outputs"
LOG_DIR = "03_Logs"
STATE_DIR = "04_State"
METADATA_DIR = "05_Metadata"

RAW_CSV = "nnr_raw.csv"
RAW_XLSX = "nnr_raw_full.xlsx"
RAW_HTML = "nnr_raw_page.html"
MASTER_XLSX = "nnr_master.xlsx"
DELTA_XLSX = "nnr_delta.xlsx"
UPDATES_XLSX = "nnr_updates.xlsx"
TOP_DELTA_XLSX = "nnr_top_delta.xlsx"

LOG_FILE = "nnr_run.log"
LOCK_FILE = "run.lock"
STATE_FILE = "seen_ids.json"
HASH_VERSION_FILE = "hash_version.txt"
HASH_VERSION = "nnr_hash_v2_20260705"

REQUEST_TIMEOUT = 45
REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

HASH_COLS = [
    "reference_number",
    "description",
    "closing_datetime",
    "document_names",
    "document_count",
    "source_url",
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
    "nuclear", "national nuclear regulator", "technical support", "engineering", "engineering services",
    "professional services", "professional service provider", "consultant", "consulting", "design",
    "construction management", "project management", "programme management", "principal agent", "architect",
    "structural engineer", "electrical engineer", "mechanical engineer", "quantity surveyor", "scientific services",
    "technical advisory", "infrastructure", "facility", "upgrade", "hvac", "master plan", "feasibility",
]


def setup_logging(base_root: Path) -> None:
    log_dir = base_root / LOG_DIR
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / LOG_FILE
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
        force=True,
    )


def ensure_dirs(base_root: Path) -> None:
    for name in [RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR]:
        (base_root / name).mkdir(parents=True, exist_ok=True)


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


def load_seen_ids(path: Path) -> Set[str]:
    try:
        if not path.exists():
            return set()
        data = json.loads(path.read_text(encoding="utf-8"))
        return {clean_text(x).lower() for x in (data or []) if clean_text(x)}
    except (OSError, ValueError, TypeError) as exc:
        logging.warning("Could not load seen IDs from %s: %s", path, exc)
        return set()


def save_seen_ids(path: Path, ids: Iterable[str]) -> None:
    try:
        values = sorted({clean_text(x).lower() for x in ids if clean_text(x)})
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(json.dumps(values, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(path)
    except (OSError, TypeError) as exc:
        logging.warning("Could not save seen IDs to %s: %s", path, exc)


def read_hash_version(path: Path) -> str:
    try:
        return clean_text(path.read_text(encoding="utf-8")) if path.exists() else ""
    except OSError:
        return ""


def write_hash_version(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(HASH_VERSION, encoding="utf-8")


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalise_id(value: Any) -> str:
    text = clean_text(value).lower()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")


def fetch_html(url: str) -> str:
    response = requests.get(url, headers=REQUEST_HEADERS, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    return response.text


def save_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", errors="ignore")


def normalise_time_text(text: str) -> str:
    text = clean_text(text)
    text = re.sub(r"(?i)\bat\b", " ", text)
    text = re.sub(r"(?i)(\d{1,2})h(\d{2})\s*(am|pm)?", lambda m: f"{m.group(1)}:{m.group(2)} {m.group(3) or ''}", text)
    text = re.sub(r"(?i)(\d{1,2})h\s*(am|pm)?", lambda m: f"{m.group(1)}:00 {m.group(2) or ''}", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def parse_datetime(value: str) -> pd.Timestamp:
    text = clean_text(value)
    if not text:
        return pd.NaT
    text = text.replace("~", " ")
    text = re.sub(r"(?i)closing\s*date\s*:?", "", text)
    text = re.sub(r"(?i)(\d+)(st|nd|rd|th)", r"\1", text)
    text = normalise_time_text(text)
    # NNR sometimes uses plain date + 11h00am. If only a date is found, assume 11:00.
    ts = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        date_match = re.search(r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\b", text)
        if date_match:
            ts = pd.to_datetime(date_match.group(0) + " 11:00", errors="coerce", dayfirst=True)
    if pd.notna(ts) and ts.hour == 0 and ts.minute == 0 and not re.search(r"\b\d{1,2}:\d{2}\b|\b\d{1,2}h", text, flags=re.I):
        ts = ts.replace(hour=11, minute=0)
    return ts


def is_tender_ref(text: str) -> bool:
    return bool(re.search(r"\bNNR\s*SCM|\bNNRSCM", clean_text(text), flags=re.I))


def extract_reference(text: str) -> str:
    text = clean_text(text)
    # Accept NNRSCM-01-2026, NNRSCM 02/2025, NNRSCM01/2016, etc.
    match = re.search(r"\bNNR\s*SCM\s*[-:]?\s*\d{1,2}\s*[-/]\s*\d{4}\b|\bNNRSCM\s*[-:]?\s*\d{1,2}\s*[-/]\s*\d{4}\b", text, flags=re.I)
    if match:
        ref = clean_text(match.group(0))
        ref = re.sub(r"\s+", " ", ref)
        return ref.upper().replace("NNR SCM", "NNRSCM")
    # Fallback: take first token if it starts with NNRSCM.
    match = re.search(r"\bNNRSCM[^\s:]+", text, flags=re.I)
    return clean_text(match.group(0)).upper() if match else ""


def strip_reference_from_description(text: str, ref: str) -> str:
    desc = clean_text(text)
    if ref:
        desc = re.sub(re.escape(ref), "", desc, flags=re.I).strip(" :-")
        # also strip flexible version
        flexible = re.escape(ref).replace(r"\ ", r"\s+")
        desc = re.sub(flexible, "", desc, flags=re.I).strip(" :-")
    return clean_text(desc)


def split_description_and_closing(text: str) -> Tuple[str, str]:
    text = clean_text(text)
    # Prefer date formats like 17 June 2026 at 11h00am at end of line.
    date_pat = r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}(?:\s*(?:at|,)\s*\d{1,2}(?::\d{2}|h\d{2}|h)?\s*(?:am|pm)?)?\b"
    matches = list(re.finditer(date_pat, text, flags=re.I))
    if matches:
        m = matches[-1]
        desc = clean_text(text[:m.start()])
        closing = clean_text(text[m.start():m.end()])
        return desc, closing
    # Other compact patterns, if NNR changes format.
    date_pat2 = r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}(?:\s*\d{1,2}(?::\d{2}|h\d{2}|h)?\s*(?:am|pm)?)?\b"
    matches = list(re.finditer(date_pat2, text, flags=re.I))
    if matches:
        m = matches[-1]
        return clean_text(text[:m.start()]), clean_text(text[m.start():m.end()])
    return text, ""


def extract_section_lines(soup: BeautifulSoup, start_label: str, stop_labels: Iterable[str]) -> List[str]:
    """Fallback text extraction from Open Tenders section."""
    for element in soup(["script", "style", "noscript", "svg"]):
        element.decompose()
    text = soup.get_text("\n")
    lines = [clean_text(line) for line in text.splitlines()]
    lines = [line for line in lines if line]
    start_idx = None
    for i, line in enumerate(lines):
        if clean_text(line).lower() == start_label.lower():
            start_idx = i + 1
            break
    if start_idx is None:
        return []
    stop_lower = [s.lower() for s in stop_labels]
    out: List[str] = []
    for line in lines[start_idx:]:
        if clean_text(line).lower() in stop_lower:
            break
        out.append(line)
    return out


def parse_open_tenders_from_tables(soup: BeautifulSoup) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    tables = soup.find_all("table")
    for table in tables:
        table_text = clean_text(table.get_text(" ")).lower()
        if "reference" not in table_text or "closing date" not in table_text:
            continue
        # Try to ensure this is the open tender table and not closed tenders.
        previous_text = ""
        prev = table.find_previous(string=True)
        if prev:
            previous_text = clean_text(prev)
        for tr in table.find_all("tr"):
            cells = tr.find_all(["td", "th"])
            if len(cells) < 2:
                continue
            cell_texts = [clean_text(c.get_text(" ")) for c in cells]
            combined = clean_text(" ".join(cell_texts))
            if not is_tender_ref(combined):
                continue
            ref = extract_reference(cell_texts[0]) or extract_reference(combined)
            if not ref:
                continue
            link = ""
            doc_name = ""
            a = cells[0].find("a", href=True) or tr.find("a", href=True)
            if a:
                link = urljoin(SOURCE_URL, a.get("href", ""))
                doc_name = clean_text(a.get_text(" ")) or ref
            if len(cells) >= 3:
                desc = cell_texts[1]
                closing_text = cell_texts[2]
            else:
                desc_with_date, closing_text = split_description_and_closing(strip_reference_from_description(combined, ref))
                desc = desc_with_date
            rows.append(build_row(ref, desc, closing_text, link, doc_name, raw_section="Open Tenders"))
    return rows


def parse_open_tenders_from_text(soup: BeautifulSoup) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    # Map reference text to anchor URLs.
    link_by_ref: Dict[str, Tuple[str, str]] = {}
    for a in soup.find_all("a", href=True):
        text = clean_text(a.get_text(" "))
        if is_tender_ref(text):
            ref = extract_reference(text)
            if ref:
                link_by_ref[ref] = (urljoin(SOURCE_URL, a.get("href", "")), text)

    section = extract_section_lines(soup, "Open Tenders", ["Closed Tenders", "Received Tenders", "Awarded Tenders", "Cancelled Tenders", "Tender History"])
    if not section:
        return []
    for line in section:
        if not is_tender_ref(line):
            continue
        ref = extract_reference(line)
        if not ref:
            continue
        remainder = strip_reference_from_description(line, ref)
        desc, closing_text = split_description_and_closing(remainder)
        link, doc_name = link_by_ref.get(ref, ("", ""))
        rows.append(build_row(ref, desc, closing_text, link, doc_name, raw_section="Open Tenders"))
    return rows


def build_row(ref: str, description: str, closing_text: str, link: str, doc_name: str, raw_section: str) -> Dict[str, Any]:
    closing_dt = parse_datetime(closing_text)
    now = pd.Timestamp.now()
    document_urls = link if link else ""
    document_names = doc_name if link else ""
    return {
        "source": SOURCE_NAME,
        "opportunity_key": f"nnr:{normalise_id(ref)}",
        "tender_number": ref,
        "reference_number": ref,
        "notice_type": "Tender",
        "title": clean_text(description),
        "description": clean_text(description),
        "client": CLIENT_NAME,
        "sector": "Nuclear",
        "category": "",
        "country_name": "South Africa",
        "province": "National",
        "region": "National",
        "date_published": "",
        "issue_date": "",
        "briefing_date": "",
        "closing_date": closing_dt.date().isoformat() if pd.notna(closing_dt) else "",
        "closing_datetime": closing_dt.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(closing_dt) else "",
        "status": "Published",
        "is_current": bool(pd.notna(closing_dt) and closing_dt >= now),
        "is_archived": False,
        "is_cancelled": False,
        "document_names": document_names,
        "document_urls": document_urls,
        "document_count": 1 if document_urls else 0,
        "contact_person": "",
        "contact_email": "",
        "contact_telephone": "+27 (12) 674 7100",
        "submission_method": "",
        "briefing_session_details": "",
        "conditions": "",
        "raw_category": raw_section,
        "raw_status": "Open Tenders",
        "source_url": link or SOURCE_URL,
    }


def parse_nnr(html: str) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "html.parser")
    rows = parse_open_tenders_from_tables(soup)
    if not rows:
        rows = parse_open_tenders_from_text(soup)
    # de-duplicate by reference
    seen = set()
    unique: List[Dict[str, Any]] = []
    for row in rows:
        key = row.get("opportunity_key") or normalise_id(row.get("reference_number"))
        if key in seen:
            continue
        seen.add(key)
        unique.append(row)
    return unique


def compute_row_hash(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    if df.empty:
        return pd.Series(dtype="object")
    def row_hash(row: pd.Series) -> str:
        parts = [clean_text(row.get(col, "")) for col in cols]
        return hashlib.sha256("||".join(parts).encode("utf-8", errors="ignore")).hexdigest()
    return df.apply(row_hash, axis=1)


def match_keywords(description: str) -> str:
    desc = clean_text(description).lower()
    matches = []
    for kw in KEYWORDS:
        if re.search(r"\b" + re.escape(kw.lower()) + r"\b", desc):
            matches.append(kw)
    return "; ".join(dict.fromkeys(matches))


def basic_weight_score(row: pd.Series) -> int:
    score = 10  # NNR priority source
    matches = [m for m in clean_text(row.get("matched_description_keywords", "")).split("; ") if m]
    score += len(matches) * 2
    if clean_text(row.get("document_urls", "")):
        score += 1
    return score


def save_excel(path: Path, df: pd.DataFrame) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    out_df = standardise_columns(df) if df is not None else pd.DataFrame(columns=OUTPUT_COLUMNS)
    with pd.ExcelWriter(
        path,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd hh:mm:ss",
        date_format="yyyy-mm-dd",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        out_df.to_excel(writer, index=False, sheet_name="Data")
        workbook = writer.book
        worksheet = writer.sheets["Data"]
        header_fmt = workbook.add_format({"bold": True, "text_wrap": True})
        for col_num, value in enumerate(out_df.columns):
            worksheet.write(0, col_num, value, header_fmt)
            width = min(max(len(str(value)) + 2, 12), 45)
            if len(out_df) and value in out_df.columns:
                try:
                    width = min(max(float(out_df[value].astype(str).str.len().quantile(0.95)) + 2, width), 60)
                except Exception:
                    pass
            worksheet.set_column(col_num, col_num, width)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max(len(out_df), 1), max(len(out_df.columns) - 1, 0))


def standardise_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[OUTPUT_COLUMNS]


def filter_current(df: pd.DataFrame, include_expired: bool = False) -> pd.DataFrame:
    if include_expired or df.empty:
        return df.copy()
    out = df.copy()
    dt = pd.to_datetime(out.get("closing_datetime"), errors="coerce")
    before = len(out)
    out = out[dt >= pd.Timestamp.now()].copy()
    logging.info("Filtered current NNR tenders: %s -> %s rows", before, len(out))
    return out


def compare_with_master(
    current_df: pd.DataFrame,
    master_path: Path,
    suppress_alerts: bool = False,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    current = standardise_columns(current_df.copy())
    if not current.empty:
        current["change_type"] = ""
        current["previous_row_hash"] = ""

    if master_path.exists():
        try:
            master = standardise_columns(pd.read_excel(master_path))
        except Exception as exc:
            logging.warning("Could not read existing master, starting fresh: %s", exc)
            master = pd.DataFrame(columns=OUTPUT_COLUMNS)
    else:
        master = pd.DataFrame(columns=OUTPUT_COLUMNS)

    if current.empty:
        logging.warning("Current NNR parse returned no open/current rows. Preserving existing master.")
        empty = standardise_columns(pd.DataFrame(columns=OUTPUT_COLUMNS))
        return empty, empty, standardise_columns(master)

    if suppress_alerts:
        logging.info("NNR hash/state version changed or was not recorded. Re-baselining and suppressing delta/update alerts for this run.")
        current["change_type"] = "UNCHANGED"
        empty = standardise_columns(pd.DataFrame(columns=OUTPUT_COLUMNS))
        return empty, empty, standardise_columns(current)

    if master.empty or "opportunity_key" not in master.columns:
        delta = current.copy()
        delta["change_type"] = "NEW"
        updates = current.iloc[0:0].copy()
        return standardise_columns(delta), standardise_columns(updates), standardise_columns(current)

    master = master.copy()
    master["opportunity_key"] = master["opportunity_key"].fillna("").astype(str).str.strip()
    if "row_hash" not in master.columns:
        master["row_hash"] = compute_row_hash(master, HASH_COLS)

    master_lookup = master.drop_duplicates("opportunity_key").set_index("opportunity_key")
    delta_rows = []
    update_rows = []

    for _, row in current.iterrows():
        key = clean_text(row.get("opportunity_key"))
        if not key or key not in master_lookup.index:
            r = row.copy()
            r["change_type"] = "NEW"
            delta_rows.append(r)
            continue

        old_row = master_lookup.loc[key]
        if isinstance(old_row, pd.DataFrame):
            old_row = old_row.iloc[0]
        prev_hash = clean_text(old_row.get("row_hash", ""))
        if clean_text(row.get("row_hash", "")) != prev_hash:
            r = row.copy()
            r["change_type"] = "UPDATED"
            r["previous_row_hash"] = prev_hash
            update_rows.append(r)

    delta = pd.DataFrame(delta_rows) if delta_rows else current.iloc[0:0].copy()
    updates = pd.DataFrame(update_rows) if update_rows else current.iloc[0:0].copy()

    # Source master is the latest open/current NNR snapshot; current run wins.
    new_master = current.copy()
    return standardise_columns(delta), standardise_columns(updates), standardise_columns(new_master)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Scrape NNR open tenders.")
    parser.add_argument("--base-root", default=BASE_ROOT)
    parser.add_argument("--include-expired", action="store_true")
    parser.add_argument("--top-delta-limit", type=int, default=50)
    parser.add_argument("--top-delta-min-score", type=int, default=0)
    args = parser.parse_args(argv)

    base_root = Path(args.base_root)
    ensure_dirs(base_root)
    setup_logging(base_root)

    raw_dir = base_root / RAW_DIR
    out_dir = base_root / OUTPUT_DIR
    state_dir = base_root / STATE_DIR
    lock_path = state_dir / LOCK_FILE
    hash_version_path = state_dir / HASH_VERSION_FILE
    seen_ids_path = state_dir / STATE_FILE

    logging.info("Using tender scrapers root: %s", TENDER_ROOT)
    logging.info("Using source root: %s", base_root)
    logging.info("Using raw directory: %s", raw_dir)
    logging.info("Using output directory: %s", out_dir)
    logging.info("Using log directory: %s", base_root / LOG_DIR)
    logging.info("Using state directory: %s", state_dir)

    try:
        acquire_lock(lock_path)
        logging.info("Fetching NNR tenders: %s", SOURCE_URL)

        html = fetch_html(SOURCE_URL)
        save_text(raw_dir / RAW_HTML, html)
        logging.info("Saved raw HTML: %s", raw_dir / RAW_HTML)

        rows = parse_nnr(html)
        logging.info("Parsed NNR open tender rows: %s", len(rows))

        raw_df = pd.DataFrame(rows)
        if raw_df.empty:
            raw_df = pd.DataFrame(columns=[c for c in OUTPUT_COLUMNS if c not in {"row_hash", "change_type", "previous_row_hash", "weight_score", "matched_description_keywords"}])

        raw_df["run_timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        raw_df["row_hash"] = compute_row_hash(raw_df, HASH_COLS)
        raw_df["matched_description_keywords"] = raw_df.get("description", pd.Series(dtype=str)).apply(match_keywords) if not raw_df.empty else ""
        raw_df["weight_score"] = raw_df.apply(basic_weight_score, axis=1) if not raw_df.empty else 0
        raw_df["change_type"] = ""
        raw_df["previous_row_hash"] = ""
        raw_df = standardise_columns(raw_df)

        raw_df.to_csv(raw_dir / RAW_CSV, index=False, encoding="utf-8-sig")
        save_excel(raw_dir / RAW_XLSX, raw_df)
        logging.info("Saved raw outputs")

        clean_df = filter_current(raw_df, include_expired=args.include_expired)
        clean_df = standardise_columns(clean_df)
        clean_df["row_hash"] = compute_row_hash(clean_df, HASH_COLS)

        suppress_alerts = read_hash_version(hash_version_path) != HASH_VERSION
        delta, updates, master = compare_with_master(
            clean_df,
            out_dir / MASTER_XLSX,
            suppress_alerts=suppress_alerts,
        )

        if not delta.empty:
            delta["weight_score"] = pd.to_numeric(delta["weight_score"], errors="coerce").fillna(0)
            delta = delta.sort_values(["weight_score", "closing_datetime"], ascending=[False, True], na_position="last")

        if not updates.empty:
            updates["weight_score"] = pd.to_numeric(updates["weight_score"], errors="coerce").fillna(0)
            updates = updates.sort_values(["weight_score", "closing_datetime"], ascending=[False, True], na_position="last")

        top_delta = delta.copy()
        if not top_delta.empty:
            top_delta = top_delta[top_delta["weight_score"].fillna(0).astype(float) >= args.top_delta_min_score]
            if args.top_delta_limit and args.top_delta_limit > 0:
                top_delta = top_delta.head(args.top_delta_limit)

        save_excel(out_dir / DELTA_XLSX, delta)
        save_excel(out_dir / UPDATES_XLSX, updates)
        save_excel(out_dir / TOP_DELTA_XLSX, top_delta)
        save_excel(out_dir / MASTER_XLSX, standardise_columns(master))

        if "opportunity_key" in master.columns:
            save_seen_ids(seen_ids_path, master["opportunity_key"].fillna("").astype(str).tolist())
        write_hash_version(hash_version_path)

        logging.info("Saved delta Excel: %s | rows: %s", out_dir / DELTA_XLSX, len(delta))
        logging.info("Saved updates Excel: %s | rows: %s", out_dir / UPDATES_XLSX, len(updates))
        logging.info("Saved top delta Excel: %s | rows: %s", out_dir / TOP_DELTA_XLSX, len(top_delta))
        logging.info("Saved master Excel: %s | rows: %s", out_dir / MASTER_XLSX, len(master))
        return 0

    except Exception as exc:
        logging.exception("FAILED: %s", exc)
        try:
            empty = standardise_columns(pd.DataFrame(columns=OUTPUT_COLUMNS))
            save_excel(out_dir / DELTA_XLSX, empty)
            save_excel(out_dir / UPDATES_XLSX, empty)
            save_excel(out_dir / TOP_DELTA_XLSX, empty)
            if not (out_dir / MASTER_XLSX).exists():
                save_excel(out_dir / MASTER_XLSX, empty)
        except Exception:
            logging.exception("Failed to write safe empty outputs after failure.")
        return 1
    finally:
        release_lock(lock_path)


if __name__ == "__main__":
    raise SystemExit(main())
