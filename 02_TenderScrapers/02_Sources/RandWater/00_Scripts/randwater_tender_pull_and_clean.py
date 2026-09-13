r"""Rand Water available tenders scraper.

Scrapes advertised tenders from:
    https://www.randwater.co.za/availabletenders.php

Output folders under the new TenderScrapers source structure:
    01_RawData, 02_Outputs, 03_Logs, 04_State, 05_Metadata
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
from io import StringIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from requests.exceptions import SSLError, RequestException
import urllib3
from bs4 import BeautifulSoup

SCRIPT_PATH = Path(__file__).resolve()
SOURCE_ROOT = SCRIPT_PATH.parents[1] if len(SCRIPT_PATH.parents) >= 2 else Path.cwd()
TENDER_ROOT = SCRIPT_PATH.parents[3] if len(SCRIPT_PATH.parents) >= 4 else SOURCE_ROOT.parents[1]
BASE_ROOT = str(SOURCE_ROOT)

SOURCE_URL = "https://www.randwater.co.za/availabletenders.php"
SOURCE_NAME = "Rand Water"
CLIENT_NAME = "Rand Water"

RAW_DIR = SOURCE_ROOT / "01_RawData"
OUTPUT_DIR = SOURCE_ROOT / "02_Outputs"
LOG_DIR = SOURCE_ROOT / "03_Logs"
STATE_DIR = SOURCE_ROOT / "04_State"
METADATA_DIR = SOURCE_ROOT / "05_Metadata"

RAW_CSV = RAW_DIR / "randwater_raw.csv"
RAW_XLSX = RAW_DIR / "randwater_raw_full.xlsx"
RAW_HTML = RAW_DIR / "randwater_raw_page.html"
MASTER_XLSX = OUTPUT_DIR / "randwater_master.xlsx"
DELTA_XLSX = OUTPUT_DIR / "randwater_delta.xlsx"
UPDATES_XLSX = OUTPUT_DIR / "randwater_updates.xlsx"
TOP_DELTA_XLSX = OUTPUT_DIR / "randwater_top_delta.xlsx"

LOG_PATH = LOG_DIR / "randwater_run.log"
LOCK_FILE = STATE_DIR / "run.lock"
SEEN_IDS_FILE = STATE_DIR / "seen_ids.json"
HASH_VERSION_FILE = STATE_DIR / "hash_version.txt"
HASH_VERSION = "randwater_hash_v2_no_full_document_urls"

REQUEST_TIMEOUT = 60
REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

OUTPUT_COLUMNS = [
    "source", "opportunity_key", "tender_number", "reference_number", "notice_type", "title", "description",
    "client", "sector", "category", "country_name", "province", "region", "date_published", "issue_date",
    "briefing_date", "closing_date", "closing_datetime", "status", "is_current", "is_archived", "is_cancelled",
    "document_names", "document_urls", "document_count", "contact_person", "contact_email", "contact_telephone",
    "submission_method", "briefing_session_details", "conditions", "raw_category", "raw_status", "source_url",
    "run_timestamp", "row_hash", "change_type", "previous_row_hash", "weight_score", "matched_description_keywords",
]
HASH_COLS = ["reference_number", "description", "closing_datetime", "document_names", "source_url", "status"]
KEYWORDS = [
    "rand water", "water", "water and sanitation", "infrastructure", "professional services", "professional service provider",
    "consultant", "consulting", "engineering", "engineering services", "design", "project management", "programme management",
    "construction monitoring", "contract administration", "technical advisory", "feasibility", "master plan",
    "environmental risk", "environmental assessment", "erla", "hazardous installation", "safety report", "pump station",
    "water treatment", "wastewater", "civil engineering", "mechanical engineering", "electrical engineering",
]
DOC_EXTS = (".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".rar")


def setup_logging(base_root: Path) -> None:
    del base_root  # paths are derived from the script location
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(LOG_PATH, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def ensure_dirs(base_root: Path) -> None:
    del base_root  # paths are derived from the script location
    for directory in [RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR]:
        directory.mkdir(parents=True, exist_ok=True)


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    text = str(value).replace("\xa0", " ").replace("\u2013", "-").replace("\u2014", "-")
    return re.sub(r"\s+", " ", text).strip()


def normalise_id(value: Any) -> str:
    text = clean_text(value).lower()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")


def fetch_html(url: str) -> str:
    """Fetch Rand Water HTML with SSL fallback.

    Rand Water's public website can present a certificate-chain issue on some
    Windows/Python installations. We first try normal certificate verification.
    If that fails with an SSL certificate verification error, we retry with
    verification disabled for this public tender page only.
    """
    session = requests.Session()
    session.headers.update(REQUEST_HEADERS)

    candidates = [url]
    if url.lower().startswith("https://"):
        candidates.append("http://" + url[len("https://"):])

    last_exc = None
    for candidate in candidates:
        try:
            r = session.get(candidate, timeout=REQUEST_TIMEOUT, allow_redirects=True)
            r.raise_for_status()
            return r.text
        except SSLError as exc:
            last_exc = exc
            logging.warning("SSL verification failed for %s. Retrying without certificate verification for this public tender page.", candidate)
            try:
                urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
                r = session.get(candidate, timeout=REQUEST_TIMEOUT, allow_redirects=True, verify=False)
                r.raise_for_status()
                return r.text
            except RequestException as retry_exc:
                last_exc = retry_exc
                logging.warning("Rand Water SSL fallback failed for %s: %s", candidate, retry_exc)
        except RequestException as exc:
            last_exc = exc
            logging.warning("Rand Water request failed for %s: %s", candidate, exc)

    raise RuntimeError(f"Could not fetch Rand Water tender page after HTTPS/HTTP attempts: {last_exc}")


def save_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", errors="ignore")


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


def load_hash_version(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8").strip() if path.exists() else ""
    except OSError:
        return ""


def save_hash_version(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(HASH_VERSION, encoding="utf-8")
    tmp.replace(path)


def save_seen_ids(path: Path, ids: List[str]) -> None:
    try:
        clean_ids = sorted({clean_text(x).lower() for x in ids if clean_text(x)})
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(json.dumps(clean_ids, ensure_ascii=False), encoding="utf-8")
        tmp.replace(path)
    except (OSError, TypeError) as exc:
        logging.exception("Failed to save state file %s: %s", path, exc)


def should_suppress_alerts() -> bool:
    hash_changed = load_hash_version(HASH_VERSION_FILE) != HASH_VERSION
    state_missing = not SEEN_IDS_FILE.exists()
    if hash_changed or state_missing:
        logging.info(
            "Rand Water hash/state version changed or was not recorded. "
            "Re-baselining and suppressing delta/update alerts for this run."
        )
        return True
    return False


def normalise_time(text: str) -> str:
    text = clean_text(text)
    text = re.sub(r"(?i)\bat\b", " ", text)
    text = re.sub(r"(?i)(\d{1,2})h(\d{2})\s*(am|pm)?", lambda m: f"{m.group(1)}:{m.group(2)} {m.group(3) or ''}", text)
    text = re.sub(r"(?i)(\d{1,2})h\s*(am|pm)?", lambda m: f"{m.group(1)}:00 {m.group(2) or ''}", text)
    return clean_text(text)


def extract_date_text(text: str) -> str:
    text = clean_text(text)
    patterns = [
        r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:\s+\d{1,2}(?::\d{2}|h\d{2}|h)?\s*(?:am|pm)?)?\b",
        r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}(?:\s+\d{1,2}(?::\d{2}|h\d{2}|h)?\s*(?:am|pm)?)?\b",
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}(?:\s*(?:at|,)\s*\d{1,2}(?::\d{2}|h\d{2}|h)?\s*(?:am|pm)?)?\b",
    ]
    for pat in patterns:
        matches = list(re.finditer(pat, text, flags=re.I))
        if matches:
            return clean_text(matches[-1].group(0))
    return ""


def parse_datetime(value: str) -> pd.Timestamp:
    text = clean_text(value)
    if not text or text in {"-", "N/A", "n/a"}:
        return pd.NaT
    text = re.sub(r"(?i)closing\s*(date|time)?\s*:?", "", text)
    text = normalise_time(text)
    if re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}", text):
        ts = pd.to_datetime(text, errors="coerce", dayfirst=False)
    else:
        ts = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if pd.notna(ts) and ts.hour == 0 and ts.minute == 0 and not re.search(r"\b\d{1,2}:\d{2}\b|\b\d{1,2}h", text, flags=re.I):
        ts = ts.replace(hour=11, minute=0)
    return ts


def extract_reference(text: str) -> str:
    text = clean_text(text)
    patterns = [
        r"\bRW\s*\d{6,}[-/]?\d{0,4}\s*[A-Z]*\b",
        r"\bRFQ\s*\d{6,}\s*[A-Z]*\b",
        r"\bRFP\s*\d{3,}[-/]?\d{0,4}\s*[A-Z]*\b",
        r"\b\d{6,}\s*[A-Z]{0,4}\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, flags=re.I)
        if m:
            return clean_text(m.group(0)).upper()
    return ""


def looks_like_doc(url: str) -> bool:
    u = clean_text(url).lower()
    path = u.split("?")[0]
    return any(path.endswith(ext) for ext in DOC_EXTS) or any(t in u for t in ["download", "media/", "sc_pdf", "document"])


def extract_links(element: Any) -> Tuple[str, str, str]:
    urls: List[str] = []
    names: List[str] = []
    detail = ""
    if element is None:
        return "", "", ""
    for a in element.find_all("a", href=True):
        href = clean_text(a.get("href", ""))
        if not href or href.lower().startswith(("javascript:", "mailto:", "tel:")):
            continue
        full = urljoin(SOURCE_URL, href)
        label = clean_text(a.get_text(" ")) or os.path.basename(urlparse(full).path) or full
        if looks_like_doc(full):
            if full not in urls:
                urls.append(full)
                names.append(label)
        elif not detail:
            detail = full
    return "; ".join(names), "; ".join(urls), detail


def build_row(reference: str, description: str, closing_text: str, source_url: str, document_names: str = "", document_urls: str = "", detail_url: str = "", issue_text: str = "", briefing_text: str = "", raw_category: str = "Available Tenders") -> Dict[str, Any]:
    description = clean_text(description)
    reference = clean_text(reference) or extract_reference(description)
    if reference:
        desc2 = re.sub(re.escape(reference), "", description, flags=re.I).strip(" -:|")
        if desc2:
            description = desc2
    closing_dt = parse_datetime(closing_text)
    issue_dt = parse_datetime(issue_text)
    briefing_dt = parse_datetime(briefing_text)
    now = pd.Timestamp.now()
    key_seed = reference or f"{description}-{closing_text}"
    return {
        "source": SOURCE_NAME,
        "opportunity_key": f"randwater:{normalise_id(key_seed)}",
        "tender_number": reference,
        "reference_number": reference,
        "notice_type": "Tender",
        "title": description,
        "description": description,
        "client": CLIENT_NAME,
        "sector": "Water",
        "category": "",
        "country_name": "South Africa",
        "province": "Gauteng",
        "region": "Gauteng",
        "date_published": issue_dt.date().isoformat() if pd.notna(issue_dt) else "",
        "issue_date": issue_dt.date().isoformat() if pd.notna(issue_dt) else "",
        "briefing_date": briefing_dt.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(briefing_dt) else "",
        "closing_date": closing_dt.date().isoformat() if pd.notna(closing_dt) else "",
        "closing_datetime": closing_dt.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(closing_dt) else "",
        "status": "Published",
        "is_current": bool(pd.notna(closing_dt) and closing_dt >= now),
        "is_archived": False,
        "is_cancelled": False,
        "document_names": clean_text(document_names),
        "document_urls": clean_text(document_urls),
        "document_count": len([u for u in clean_text(document_urls).split("; ") if u]) if document_urls else 0,
        "contact_person": "",
        "contact_email": "",
        "contact_telephone": "+27 11 682 0911",
        "submission_method": "Rand Water tender portal / tender instructions",
        "briefing_session_details": clean_text(briefing_text),
        "conditions": "",
        "raw_category": raw_category,
        "raw_status": "Available Tenders",
        "source_url": detail_url or source_url or SOURCE_URL,
    }


def infer_description(cells: List[str]) -> str:
    candidates = []
    for c in cells:
        c = clean_text(c)
        if not c or c in {"-", "N/A"}:
            continue
        if extract_date_text(c):
            continue
        if extract_reference(c) and len(c) < 40:
            continue
        if c.lower() in {"description", "tender title", "tender description", "closing date", "briefing date", "reference", "tender number"}:
            continue
        candidates.append(c)
    return max(candidates, key=len) if candidates else clean_text(" ".join(cells))


def parse_tables(soup: BeautifulSoup) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    tables = soup.find_all("table")
    logging.info("HTML tables found: %s", len(tables))
    for table in tables:
        txt = clean_text(table.get_text(" ")).lower()
        if not any(k in txt for k in ["closing", "tender", "bid", "rfq", "rfp"]):
            continue
        headers = [clean_text(h.get_text(" ")).lower() for h in table.find_all("th")]
        for tr in table.find_all("tr"):
            cells_el = tr.find_all(["td", "th"])
            if not cells_el:
                continue
            cells = [clean_text(c.get_text(" ")) for c in cells_el]
            combined = clean_text(" ".join(cells))
            if not combined or len(cells) < 2 or not extract_date_text(combined):
                continue
            if sum(1 for c in cells if c.lower() in {"description", "closing date", "tender number", "reference", "briefing date"}) >= 2:
                continue
            doc_names, doc_urls, detail_url = extract_links(tr)
            ref = extract_reference(combined)
            desc = infer_description(cells)
            dates = [extract_date_text(c) for c in cells if extract_date_text(c)]
            closing = dates[-1] if dates else ""
            briefing = dates[-2] if len(dates) >= 2 else ""
            if desc and closing:
                rows.append(build_row(ref, desc, closing, detail_url or SOURCE_URL, doc_names, doc_urls, detail_url, "", briefing, "HTML table"))
    return rows


def parse_visible_text(soup: BeautifulSoup) -> List[Dict[str, Any]]:
    for element in soup(["script", "style", "noscript", "svg"]):
        element.decompose()
    text = soup.get_text("\n")
    lines = [clean_text(x) for x in text.splitlines()]
    lines = [x for x in lines if x]
    rows: List[Dict[str, Any]] = []
    for line in lines:
        if len(line) < 40 or len(line) > 800:
            continue
        if not extract_date_text(line):
            continue
        low = line.lower()
        if not any(t in low for t in ["tender", "rfq", "rfp", "bid", "rand water", "closing"]):
            continue
        closing = extract_date_text(line)
        desc = line.replace(closing, " ")
        desc = re.sub(r"(?i)closing\s*(date|date\s*&\s*time|time)?\s*:?", " ", desc)
        desc = clean_text(desc)
        ref = extract_reference(desc)
        if desc and closing:
            rows.append(build_row(ref, desc, closing, SOURCE_URL, raw_category="Visible text"))
    return rows


def parse_randwater(html: str) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "html.parser")
    rows = parse_tables(soup)
    if not rows:
        rows = parse_visible_text(soup)
    seen = set()
    unique: List[Dict[str, Any]] = []
    for row in rows:
        key = row.get("opportunity_key") or normalise_id(row.get("description", "") + row.get("closing_datetime", ""))
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
    score = 10  # Rand Water priority source
    matches = [m for m in clean_text(row.get("matched_description_keywords", "")).split("; ") if m]
    score += len(matches) * 2
    if clean_text(row.get("document_urls", "")):
        score += 1
    return score


def save_excel(path: Path, df: pd.DataFrame) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm:ss", date_format="yyyy-mm-dd", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
        workbook = writer.book
        worksheet = writer.sheets["Data"]
        header_fmt = workbook.add_format({"bold": True, "text_wrap": True})
        for col_num, value in enumerate(df.columns):
            worksheet.write(0, col_num, value, header_fmt)
            width = min(max(len(str(value)) + 2, 12), 45)
            if len(df) and value in df.columns:
                try:
                    width = min(max(df[value].astype(str).str.len().quantile(0.95) + 2, width), 60)
                except Exception:
                    pass
            worksheet.set_column(col_num, col_num, width)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max(len(df), 1), max(len(df.columns) - 1, 0))


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
    logging.info("Filtered current Rand Water tenders: %s -> %s rows", before, len(out))
    return out


def compare_with_master(current_df: pd.DataFrame, master_path: Path) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if current_df.empty:
        if master_path.exists():
            try:
                master_existing = pd.read_excel(master_path)
            except Exception:
                master_existing = pd.DataFrame(columns=OUTPUT_COLUMNS)
        else:
            master_existing = pd.DataFrame(columns=OUTPUT_COLUMNS)
        return current_df.copy(), current_df.copy(), standardise_columns(master_existing)

    current = current_df.copy()
    current["change_type"] = ""
    current["previous_row_hash"] = ""

    if master_path.exists():
        try:
            master = pd.read_excel(master_path)
        except Exception as exc:
            logging.warning("Could not read existing master, starting fresh: %s", exc)
            master = pd.DataFrame(columns=OUTPUT_COLUMNS)
    else:
        master = pd.DataFrame(columns=OUTPUT_COLUMNS)

    if master.empty or "opportunity_key" not in master.columns:
        delta = current.copy()
        delta["change_type"] = "NEW"
        updates = current.iloc[0:0].copy()
        return standardise_columns(delta), standardise_columns(updates), standardise_columns(current.copy())

    master_lookup = master.drop_duplicates("opportunity_key").set_index("opportunity_key")
    delta_rows = []
    update_rows = []
    for _, row in current.iterrows():
        key = row.get("opportunity_key")
        if key not in master_lookup.index:
            r = row.copy()
            r["change_type"] = "NEW"
            delta_rows.append(r)
        else:
            prev_hash = clean_text(master_lookup.loc[key].get("row_hash", ""))
            if clean_text(row.get("row_hash", "")) != prev_hash:
                r = row.copy()
                r["change_type"] = "UPDATED"
                r["previous_row_hash"] = prev_hash
                update_rows.append(r)
    delta = pd.DataFrame(delta_rows) if delta_rows else current.iloc[0:0].copy()
    updates = pd.DataFrame(update_rows) if update_rows else current.iloc[0:0].copy()
    old_other = master[~master.get("opportunity_key", pd.Series(dtype=str)).isin(current["opportunity_key"])].copy()
    new_master = pd.concat([old_other, current], ignore_index=True, sort=False)
    return standardise_columns(delta), standardise_columns(updates), standardise_columns(new_master)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Scrape Rand Water available tenders.")
    parser.add_argument("--base-root", default=BASE_ROOT, help="Kept for backwards compatibility; paths are derived from the script location.")
    parser.add_argument("--include-expired", action="store_true")
    parser.add_argument("--top-delta-limit", type=int, default=50)
    parser.add_argument("--top-delta-min-score", type=int, default=0)
    args = parser.parse_args(argv)

    base_root = SOURCE_ROOT
    ensure_dirs(base_root)
    setup_logging(base_root)

    logging.info("Using tender scrapers root: %s", TENDER_ROOT)
    logging.info("Using source root: %s", SOURCE_ROOT)
    logging.info("Using raw directory: %s", RAW_DIR)
    logging.info("Using output directory: %s", OUTPUT_DIR)
    logging.info("Using log directory: %s", LOG_DIR)
    logging.info("Using state directory: %s", STATE_DIR)
    logging.info("Fetching Rand Water tenders: %s", SOURCE_URL)

    lock_acquired = False
    try:
        acquire_lock(LOCK_FILE)
        lock_acquired = True

        try:
            html = fetch_html(SOURCE_URL)
        except Exception as exc:
            logging.exception("Rand Water fetch failed; preserving existing master where possible: %s", exc)
            html = ""

        save_text(RAW_HTML, html)
        logging.info("Saved raw HTML: %s", RAW_HTML)

        rows = parse_randwater(html) if html else []
        logging.info("Parsed Rand Water tender rows: %s", len(rows))

        raw_df = pd.DataFrame(rows)
        if raw_df.empty:
            raw_df = pd.DataFrame(columns=[
                c for c in OUTPUT_COLUMNS
                if c not in {"row_hash", "change_type", "previous_row_hash", "weight_score", "matched_description_keywords"}
            ])

        raw_df["run_timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        raw_df["row_hash"] = compute_row_hash(raw_df, HASH_COLS)
        raw_df["matched_description_keywords"] = raw_df.get("description", pd.Series(dtype=str)).apply(match_keywords) if not raw_df.empty else ""
        raw_df["weight_score"] = raw_df.apply(basic_weight_score, axis=1) if not raw_df.empty else 0
        raw_df["change_type"] = ""
        raw_df["previous_row_hash"] = ""
        raw_df = standardise_columns(raw_df)

        raw_df.to_csv(RAW_CSV, index=False, encoding="utf-8-sig")
        save_excel(RAW_XLSX, raw_df)
        logging.info("Saved raw outputs")

        clean_df = filter_current(raw_df, include_expired=args.include_expired)
        clean_df = standardise_columns(clean_df)
        clean_df["row_hash"] = compute_row_hash(clean_df, HASH_COLS)

        delta, updates, master = compare_with_master(clean_df, MASTER_XLSX)

        if should_suppress_alerts():
            delta = standardise_columns(pd.DataFrame(columns=OUTPUT_COLUMNS))
            updates = standardise_columns(pd.DataFrame(columns=OUTPUT_COLUMNS))

        if not delta.empty:
            delta = delta.sort_values(["weight_score", "closing_datetime"], ascending=[False, True], na_position="last")
        if not updates.empty:
            updates = updates.sort_values(["weight_score", "closing_datetime"], ascending=[False, True], na_position="last")

        top_delta = delta.copy()
        if not top_delta.empty:
            top_delta = top_delta[top_delta["weight_score"].fillna(0).astype(float) >= args.top_delta_min_score].head(args.top_delta_limit)

        save_excel(DELTA_XLSX, delta)
        save_excel(UPDATES_XLSX, updates)
        save_excel(TOP_DELTA_XLSX, top_delta)
        master = standardise_columns(master)
        save_excel(MASTER_XLSX, master)

        if "opportunity_key" in master.columns:
            save_seen_ids(SEEN_IDS_FILE, master["opportunity_key"].fillna("").astype(str).tolist())
        save_hash_version(HASH_VERSION_FILE)

        logging.info("Saved delta Excel: %s | rows: %s", DELTA_XLSX, len(delta))
        logging.info("Saved updates Excel: %s | rows: %s", UPDATES_XLSX, len(updates))
        logging.info("Saved top delta Excel: %s | rows: %s", TOP_DELTA_XLSX, len(top_delta))
        logging.info("Saved master Excel: %s | rows: %s", MASTER_XLSX, len(master))
        return 0

    finally:
        if lock_acquired:
            release_lock(LOCK_FILE)


if __name__ == "__main__":
    raise SystemExit(main())
