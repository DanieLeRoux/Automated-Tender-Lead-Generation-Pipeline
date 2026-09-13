r"""Johannesburg Water advertised tenders / RFPs scraper.

Scrapes public procurement pages from Johannesburg Water SCM:
    - Advertised Tenders: https://scm.johannesburgwater.co.za/supply-chain/tenders/all-open-tenders/
    - Advertised RFPs:    https://scm.johannesburgwater.co.za/supply-chain/request-for-proposal/all-advertised-rfps/

Updated for the consolidated TenderScrapers folder structure:
    02_Sources\JohannesburgWater\01_RawData
    02_Sources\JohannesburgWater\02_Outputs
    02_Sources\JohannesburgWater\03_Logs
    02_Sources\JohannesburgWater\04_State
    02_Sources\JohannesburgWater\05_Metadata
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
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.exceptions import RequestException, SSLError
import urllib3

SOURCE_SLUG = "johannesburg_water"
SOURCE_NAME = "Johannesburg Water"
CLIENT_NAME = "Johannesburg Water"
COUNTRY_NAME = "South Africa"
PROVINCE = "Gauteng"
REGION = "Johannesburg"
SECTOR = "Water"
CATEGORY = "Services: Professional"

SOURCE_PAGES = [
    {
        "url": "https://scm.johannesburgwater.co.za/supply-chain/tenders/all-open-tenders/",
        "page_type": "Tender",
        "raw_html": "johannesburg_water_raw_tenders_page.html",
    },
    {
        "url": "https://scm.johannesburgwater.co.za/supply-chain/request-for-proposal/all-advertised-rfps/",
        "page_type": "RFP",
        "raw_html": "johannesburg_water_raw_rfps_page.html",
    },
]

RAW_DIR_NAME = "01_RawData"
OUTPUT_DIR_NAME = "02_Outputs"
LOG_DIR_NAME = "03_Logs"
STATE_DIR_NAME = "04_State"
METADATA_DIR_NAME = "05_Metadata"

RAW_CSV = "johannesburg_water_raw.csv"
RAW_XLSX = "johannesburg_water_raw_full.xlsx"
MASTER_XLSX = "johannesburg_water_master.xlsx"
DELTA_XLSX = "johannesburg_water_delta.xlsx"
UPDATES_XLSX = "johannesburg_water_updates.xlsx"
TOP_DELTA_XLSX = "johannesburg_water_top_delta.xlsx"

REQUEST_TIMEOUT = 90
REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}
DOC_EXTS = (".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".rar")

# Increment this only when the meaningful-change hash logic changes.
HASH_VERSION = "johannesburg_water_hash_v2_2026_07_04"

OUTPUT_COLUMNS = [
    "source", "opportunity_key", "tender_number", "reference_number", "notice_type", "title", "description",
    "client", "sector", "category", "country_name", "province", "region", "date_published", "issue_date",
    "briefing_date", "closing_date", "closing_datetime", "status", "is_current", "is_archived", "is_cancelled",
    "document_names", "document_urls", "document_count", "contact_person", "contact_email", "contact_telephone",
    "submission_method", "briefing_session_details", "conditions", "raw_category", "raw_status", "source_url",
    "run_timestamp", "row_hash", "change_type", "previous_row_hash", "weight_score", "matched_description_keywords",
]

# Deliberately excludes full document_urls. Some public sites vary link/query formatting, causing false updates.
HASH_COLS = [
    "reference_number",
    "description",
    "closing_datetime",
    "document_names",
    "document_count",
    "source_url",
    "notice_type",
]

KEYWORDS = [
    "johannesburg water", "joburg water", "water", "water and sanitation", "wastewater", "sewer", "pipeline",
    "professional services", "professional service provider", "psp", "consultant", "consulting", "consultancy",
    "engineering", "engineering services", "civil engineering", "mechanical engineering", "electrical engineering",
    "structural engineering", "project management", "programme management", "construction monitoring", "contract administration",
    "feasibility", "design", "master plan", "infrastructure", "advisory", "technical advisory", "built environment",
]


def default_source_root() -> Path:
    """Return the source root when the script sits in 00_Scripts."""
    script_path = Path(__file__).resolve()
    try:
        if script_path.parent.name == "00_Scripts":
            return script_path.parents[1]
        return script_path.parent
    except IndexError:
        return script_path.parent


def tender_root_from_source(source_root: Path) -> Path:
    try:
        if source_root.parent.name == "02_Sources":
            return source_root.parents[1]
    except IndexError:
        pass
    return source_root.parent


def ensure_dirs(source_root: Path) -> Tuple[Path, Path, Path, Path, Path]:
    raw_dir = source_root / RAW_DIR_NAME
    output_dir = source_root / OUTPUT_DIR_NAME
    log_dir = source_root / LOG_DIR_NAME
    state_dir = source_root / STATE_DIR_NAME
    metadata_dir = source_root / METADATA_DIR_NAME
    for path in (raw_dir, output_dir, log_dir, state_dir, metadata_dir):
        path.mkdir(parents=True, exist_ok=True)
    return raw_dir, output_dir, log_dir, state_dir, metadata_dir


def setup_logging(log_dir: Path) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "johannesburg_water_run.log"
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


def normalise_state_key(value: Any) -> str:
    return clean_text(value).casefold()


def save_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", errors="ignore")


def load_seen_ids(path: Path) -> set[str]:
    try:
        if not path.exists():
            return set()
        data = json.loads(path.read_text(encoding="utf-8"))
        return {normalise_state_key(x) for x in (data or []) if normalise_state_key(x)}
    except Exception as exc:
        logging.warning("Could not read seen IDs file %s: %s", path, exc)
        return set()


def save_seen_ids(path: Path, keys: Iterable[Any]) -> None:
    try:
        values = sorted({normalise_state_key(k) for k in keys if normalise_state_key(k)})
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(json.dumps(values, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(path)
    except Exception as exc:
        logging.warning("Could not save seen IDs file %s: %s", path, exc)


def read_hash_version(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8").strip() if path.exists() else ""
    except Exception:
        return ""


def write_hash_version(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(HASH_VERSION + "\n", encoding="utf-8")


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


def fetch_html(url: str) -> str:
    session = requests.Session()
    session.headers.update(REQUEST_HEADERS)
    try:
        response = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        response.raise_for_status()
        return response.text
    except SSLError:
        logging.warning("SSL verification failed for %s. Retrying without verification for public tender page.", url)
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        response = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True, verify=False)
        response.raise_for_status()
        return response.text


def parse_jw_datetime(value: str, is_closing: bool = False) -> pd.Timestamp:
    text = clean_text(value)
    if not text or text.lower() in {"n/a", "na", "none", "-", "not attached"}:
        return pd.NaT
    text = re.sub(r"(?i)closing\s*(date|time)?\s*: ?", "", text)
    text = re.sub(r"(?i)date\s*of\s*advert\s*: ?", "", text)
    text = re.sub(r"(?i)\bat\b", " ", text)
    text = re.sub(r"(?i)(\d{1,2})h(\d{2})\s*(am|pm)?", lambda m: f"{m.group(1)}:{m.group(2)} {m.group(3) or ''}", text)
    text = re.sub(r"(?i)(\d{1,2})h\s*(am|pm)?", lambda m: f"{m.group(1)}:00 {m.group(2) or ''}", text)
    text = clean_text(text)

    # Johannesburg Water currently uses mm/dd/yyyy in the public tables, e.g. 04/22/2026.
    timestamp = pd.to_datetime(text, errors="coerce", dayfirst=False)
    if pd.notna(timestamp) and is_closing and timestamp.hour == 0 and timestamp.minute == 0:
        # The table commonly closes at 10:30am; if the time was not captured, use 10:30 as a practical default.
        if not re.search(r"\b\d{1,2}:\d{2}\b|\b\d{1,2}h", text, flags=re.I):
            timestamp = timestamp.replace(hour=10, minute=30)
    return timestamp


def extract_links(cell: Any, page_url: str) -> Tuple[str, str, int]:
    names: List[str] = []
    urls: List[str] = []
    if cell is None:
        return "", "", 0
    for a in cell.find_all("a", href=True):
        href = clean_text(a.get("href"))
        text = clean_text(a.get_text(" "))
        if not href or href.startswith("#") or href.lower().startswith("mailto:"):
            continue
        full_url = urljoin(page_url, href)
        joined = f"{text} {full_url}".lower()
        if any(ext in full_url.lower() for ext in DOC_EXTS) or any(
            word in joined for word in ["tender", "document", "download", "advert", "addendum", "minutes", "boq"]
        ):
            if full_url not in urls:
                urls.append(full_url)
                names.append(text or Path(full_url).name)
    return "; ".join(names), "; ".join(urls), len(urls)


def normalise_headers(cells: List[str]) -> List[str]:
    out: List[str] = []
    for header in cells:
        value = clean_text(header).lower()
        value = re.sub(r"[^a-z0-9]+", "_", value).strip("_")
        out.append(value)
    return out


def infer_notice_type(reference: str, page_type: str, description: str) -> str:
    text = f"{reference} {page_type} {description}".upper()
    if "RFP" in text or "REQUEST FOR PROPOSAL" in text:
        return "RFP"
    if "RFQ" in text or "REQUEST FOR QUOTATION" in text:
        return "RFQ"
    return "Tender"


def make_row(
    reference: str,
    description: str,
    page_type: str,
    closing_dt: pd.Timestamp,
    advert_dt: pd.Timestamp,
    briefing: str,
    doc_names: str,
    doc_urls: str,
    doc_count: int,
    source_url: str,
    raw_category: str,
    raw_status: str,
) -> Dict[str, Any]:
    reference = clean_text(reference).upper()
    description = clean_text(description)
    title = description
    notice_type = infer_notice_type(reference, page_type, description)
    closing_date = closing_dt.date().isoformat() if pd.notna(closing_dt) else ""
    issue_date = advert_dt.date().isoformat() if pd.notna(advert_dt) else ""
    is_current = bool(pd.notna(closing_dt) and closing_dt >= pd.Timestamp.now())
    opportunity_key = f"johannesburg_water:{normalise_id(reference or description)}"
    return {
        "source": SOURCE_NAME,
        "opportunity_key": opportunity_key,
        "tender_number": reference,
        "reference_number": reference,
        "notice_type": notice_type,
        "title": title,
        "description": description,
        "client": CLIENT_NAME,
        "sector": SECTOR,
        "category": CATEGORY,
        "country_name": COUNTRY_NAME,
        "province": PROVINCE,
        "region": REGION,
        "date_published": issue_date,
        "issue_date": issue_date,
        "briefing_date": "",
        "closing_date": closing_date,
        "closing_datetime": closing_dt.isoformat(sep=" ") if pd.notna(closing_dt) else "",
        "status": raw_status,
        "is_current": is_current,
        "is_archived": False,
        "is_cancelled": False,
        "document_names": clean_text(doc_names),
        "document_urls": clean_text(doc_urls),
        "document_count": int(doc_count or 0),
        "contact_person": "",
        "contact_email": "",
        "contact_telephone": "",
        "submission_method": "As per Johannesburg Water tender document / SCM instructions",
        "briefing_session_details": clean_text(briefing),
        "conditions": "",
        "raw_category": raw_category,
        "raw_status": raw_status,
        "source_url": source_url,
        "run_timestamp": datetime.now().isoformat(timespec="seconds"),
        "row_hash": "",
        "change_type": "",
        "previous_row_hash": "",
        "weight_score": 0,
        "matched_description_keywords": "",
    }


def parse_tables(html: str, page_url: str, page_type: str) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "html.parser")
    rows: List[Dict[str, Any]] = []

    for table in soup.find_all("table"):
        tr_list = table.find_all("tr")
        if not tr_list:
            continue

        header_cells = tr_list[0].find_all(["th", "td"])
        headers = normalise_headers([c.get_text(" ") for c in header_cells])
        header_text = " ".join(headers)
        if not any(x in header_text for x in ["tender_no", "rfp_no", "description"]):
            continue
        if "closing_date" not in header_text and "closing" not in header_text:
            continue

        for tr in tr_list[1:]:
            cells = tr.find_all("td")
            if len(cells) < 3:
                continue
            values = [clean_text(c.get_text(" ")) for c in cells]
            row_map = {headers[i] if i < len(headers) else f"col_{i}": values[i] for i in range(len(values))}

            reference = ""
            for key in ["tender_no", "tender_number", "rfp_no", "rfp_number", "bid_no", "bid_number"]:
                if row_map.get(key):
                    reference = clean_text(row_map[key])
                    break
            description = row_map.get("description", "") or row_map.get("desc", "")
            closing_text = row_map.get("closing_date", "") or row_map.get("closing", "")
            advert_text = row_map.get("date_of_advert", "") or row_map.get("advert_date", "")
            briefing = row_map.get("briefing_session", "") or row_map.get("briefing", "")

            if not reference and not description:
                continue
            if clean_text(reference).lower() in {"tender no", "rfp no"}:
                continue

            doc_names_parts: List[str] = []
            doc_urls_parts: List[str] = []
            doc_count = 0
            for idx, cell in enumerate(cells):
                header = headers[idx] if idx < len(headers) else ""
                if header in {"download", "addendum", "document", "documents", "tender_document"} or cell.find("a", href=True):
                    names, urls, count = extract_links(cell, page_url)
                    if urls:
                        doc_names_parts.append(names)
                        doc_urls_parts.append(urls)
                        doc_count += count

            doc_names = "; ".join([x for x in doc_names_parts if x])
            doc_urls = "; ".join([x for x in doc_urls_parts if x])
            closing_dt = parse_jw_datetime(closing_text, is_closing=True)
            advert_dt = parse_jw_datetime(advert_text, is_closing=False)

            rows.append(
                make_row(
                    reference=reference,
                    description=description,
                    page_type=page_type,
                    closing_dt=closing_dt,
                    advert_dt=advert_dt,
                    briefing=briefing,
                    doc_names=doc_names,
                    doc_urls=doc_urls,
                    doc_count=doc_count,
                    source_url=page_url,
                    raw_category=page_type,
                    raw_status="Open",
                )
            )
    return rows


def parse_text_fallback(html: str, page_url: str, page_type: str) -> List[Dict[str, Any]]:
    """Fallback parser for WordPress pages where tables are flattened by the renderer."""
    soup = BeautifulSoup(html, "html.parser")
    full_text = clean_text(soup.get_text(" \n "))
    if not full_text:
        return []
    ref_pat = re.compile(
        r"\b(?:JW\s*[A-Z]{2,5}\s*\d{3}/\d{2}|JW\s*[A-Z]{2,5}\s*\d{3}\s*/\s*\d{2}|JW\s*\d{4,6}|JW\s*[A-Z]{2,5}\d{3}/\d{2})\b",
        re.I,
    )
    matches = list(ref_pat.finditer(full_text))
    rows: List[Dict[str, Any]] = []
    for i, match in enumerate(matches):
        start = match.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else min(len(full_text), start + 1500)
        block = full_text[start:end]
        reference = clean_text(match.group(0)).upper().replace(" ", "")
        date_matches = list(
            re.finditer(r"\b\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}\s*(?:am|pm)?\b", block, flags=re.I)
        )
        advert_dt = pd.NaT
        closing_dt = pd.NaT
        if len(date_matches) >= 1:
            advert_dt = parse_jw_datetime(date_matches[0].group(0), is_closing=False)
        if len(date_matches) >= 2:
            closing_dt = parse_jw_datetime(date_matches[1].group(0), is_closing=True)
        desc_end = date_matches[0].start() if date_matches else min(len(block), 300)
        description = clean_text(block[len(match.group(0)):desc_end]) or reference
        rows.append(
            make_row(
                reference=reference,
                description=description,
                page_type=page_type,
                closing_dt=closing_dt,
                advert_dt=advert_dt,
                briefing="",
                doc_names="",
                doc_urls="",
                doc_count=0,
                source_url=page_url,
                raw_category=page_type,
                raw_status="Open",
            )
        )
    return rows


def score_row(row: Dict[str, Any]) -> Tuple[int, str]:
    text = f"{row.get('client','')} {row.get('description','')} {row.get('reference_number','')}".lower()
    matches = [keyword for keyword in KEYWORDS if keyword.lower() in text]
    score = 0
    if "johannesburg water" in text or "joburg water" in text:
        score += 10
    score += min(len(matches) * 3, 24)
    if row.get("document_count", 0):
        score += 1
    return score, "; ".join(sorted(set(matches)))


def compute_row_hash(df: pd.DataFrame, cols: Iterable[str]) -> pd.Series:
    if df.empty:
        return pd.Series(dtype="object", index=df.index)
    use_cols = [c for c in cols if c in df.columns]

    def _hash(row: pd.Series) -> str:
        payload = "||".join(clean_text(row.get(c, "")) for c in use_cols)
        return hashlib.sha256(payload.encode("utf-8", errors="ignore")).hexdigest()

    return df.apply(_hash, axis=1)


def add_missing_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in OUTPUT_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    return out[OUTPUT_COLUMNS]


def load_previous_master(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    try:
        df = pd.read_excel(path)
        return add_missing_columns(df)
    except Exception as exc:
        logging.warning("Could not read previous master %s: %s", path, exc)
        return pd.DataFrame(columns=OUTPUT_COLUMNS)


def write_excel(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        sheet = "Data"
        df.to_excel(writer, index=False, sheet_name=sheet)
        worksheet = writer.sheets[sheet]
        for i, col in enumerate(df.columns):
            width = min(max(len(str(col)) + 2, 12), 60)
            if col in {"description", "document_urls", "document_names", "briefing_session_details", "source_url"}:
                width = 50
            worksheet.set_column(i, i, width)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max(len(df), 1), max(len(df.columns) - 1, 0))


def fetch_and_parse(raw_dir: Path) -> pd.DataFrame:
    all_rows: List[Dict[str, Any]] = []
    for page in SOURCE_PAGES:
        url = page["url"]
        page_type = page["page_type"]
        logging.info("Fetching Johannesburg Water %s page: %s", page_type, url)
        try:
            html = fetch_html(url)
            save_text(raw_dir / page["raw_html"], html)
        except RequestException as exc:
            logging.error("Failed to fetch Johannesburg Water page %s: %s", url, exc)
            continue

        rows = parse_tables(html, url, page_type)
        if not rows:
            logging.info("No table rows parsed for %s. Trying text fallback.", page_type)
            rows = parse_text_fallback(html, url, page_type)
        logging.info("Parsed Johannesburg Water %s rows: %s", page_type, len(rows))
        all_rows.extend(rows)

    df = pd.DataFrame(all_rows)
    df = add_missing_columns(df)
    if df.empty:
        return df
    df["_dedupe_key"] = (
        df["reference_number"].astype(str).str.lower().str.strip()
        + "|"
        + df["description"].astype(str).str.lower().str.strip()
    )
    df = df.drop_duplicates("_dedupe_key").drop(columns=["_dedupe_key"])
    scores = df.apply(lambda r: score_row(r.to_dict()), axis=1)
    df["weight_score"] = [s[0] for s in scores]
    df["matched_description_keywords"] = [s[1] for s in scores]
    df["row_hash"] = compute_row_hash(df, HASH_COLS)
    return add_missing_columns(df)


def build_current_master_candidate(raw_df: pd.DataFrame) -> pd.DataFrame:
    master_df = raw_df.copy()
    if master_df.empty:
        return add_missing_columns(master_df)
    master_df["_closing_dt"] = pd.to_datetime(master_df["closing_datetime"], errors="coerce")
    before = len(master_df)
    master_df = master_df[(master_df["_closing_dt"].isna()) | (master_df["_closing_dt"] >= pd.Timestamp.now())].copy()
    master_df = master_df.drop(columns=["_closing_dt"])
    logging.info("Filtered current Johannesburg Water rows: %s -> %s", before, len(master_df))
    master_df = add_missing_columns(master_df)
    if not master_df.empty:
        master_df["row_hash"] = compute_row_hash(master_df, HASH_COLS)
    return master_df


def classify_delta_updates(master_df: pd.DataFrame, previous_master: pd.DataFrame, suppress_alerts: bool) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if master_df.empty:
        empty = pd.DataFrame(columns=OUTPUT_COLUMNS)
        return master_df.copy(), empty.copy(), empty.copy()

    full = master_df.copy()
    if suppress_alerts:
        full["change_type"] = "BASELINE"
        full["previous_row_hash"] = ""
        empty = pd.DataFrame(columns=OUTPUT_COLUMNS)
        return add_missing_columns(full), empty.copy(), empty.copy()

    previous_hash_by_key: Dict[str, str] = {}
    if not previous_master.empty and "opportunity_key" in previous_master.columns:
        previous_hash_by_key = dict(
            zip(
                previous_master["opportunity_key"].fillna("").astype(str),
                previous_master["row_hash"].fillna("").astype(str),
            )
        )

    delta_rows: List[int] = []
    update_rows: List[int] = []
    for idx, row in full.iterrows():
        key = str(row.get("opportunity_key", ""))
        prev_hash = previous_hash_by_key.get(key)
        new_hash = str(row.get("row_hash", ""))
        if not prev_hash:
            delta_rows.append(idx)
            full.at[idx, "change_type"] = "NEW"
            full.at[idx, "previous_row_hash"] = ""
        elif prev_hash != new_hash:
            update_rows.append(idx)
            full.at[idx, "change_type"] = "UPDATED"
            full.at[idx, "previous_row_hash"] = prev_hash
        else:
            full.at[idx, "change_type"] = "UNCHANGED"
            full.at[idx, "previous_row_hash"] = ""

    delta_df = add_missing_columns(full.loc[delta_rows].copy() if delta_rows else pd.DataFrame(columns=OUTPUT_COLUMNS))
    updates_df = add_missing_columns(full.loc[update_rows].copy() if update_rows else pd.DataFrame(columns=OUTPUT_COLUMNS))
    return add_missing_columns(full), delta_df, updates_df


def should_suppress_alerts(state_dir: Path) -> bool:
    hash_version_file = state_dir / "hash_version.txt"
    previous_version = read_hash_version(hash_version_file)
    if previous_version != HASH_VERSION:
        logging.info("Johannesburg Water hash/state version changed or was not recorded. Re-baselining and suppressing delta/update alerts for this run.")
        return True
    return False


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Johannesburg Water tender scraper")
    parser.add_argument("--base-root", default=str(default_source_root()))
    args = parser.parse_args(argv)

    source_root = Path(args.base_root).resolve()
    tender_root = tender_root_from_source(source_root)
    raw_dir, output_dir, log_dir, state_dir, _metadata_dir = ensure_dirs(source_root)
    setup_logging(log_dir)

    state_file = state_dir / "seen_ids.json"
    hash_version_file = state_dir / "hash_version.txt"
    lock_file = state_dir / "run.lock"

    logging.info("Using tender scrapers root: %s", tender_root)
    logging.info("Using source root: %s", source_root)
    logging.info("Using raw directory: %s", raw_dir)
    logging.info("Using output directory: %s", output_dir)
    logging.info("Using log directory: %s", log_dir)
    logging.info("Using state directory: %s", state_dir)

    try:
        acquire_lock(lock_file)

        try:
            raw_df = fetch_and_parse(raw_dir)
            logging.info("Raw Johannesburg Water rows: %s | cols: %s", len(raw_df), len(raw_df.columns))
        except Exception as exc:
            logging.exception("FAILED to parse Johannesburg Water. Writing empty raw outputs to avoid breaking daily run: %s", exc)
            raw_df = add_missing_columns(pd.DataFrame())

        raw_df.to_csv(raw_dir / RAW_CSV, index=False, encoding="utf-8-sig")
        write_excel(raw_df, raw_dir / RAW_XLSX)

        previous_master = load_previous_master(output_dir / MASTER_XLSX)

        if raw_df.empty and not previous_master.empty:
            logging.warning("No Johannesburg Water rows parsed. Preserving existing master and writing empty delta/update/top-delta outputs.")
            empty = pd.DataFrame(columns=OUTPUT_COLUMNS)
            write_excel(empty, output_dir / DELTA_XLSX)
            write_excel(empty, output_dir / UPDATES_XLSX)
            write_excel(empty, output_dir / TOP_DELTA_XLSX)
            save_seen_ids(state_file, previous_master.get("opportunity_key", pd.Series(dtype=str)).tolist())
            write_hash_version(hash_version_file)
            logging.info("Saved delta Excel: %s | rows: 0", output_dir / DELTA_XLSX)
            logging.info("Saved updates Excel: %s | rows: 0", output_dir / UPDATES_XLSX)
            logging.info("Saved top delta Excel: %s | rows: 0", output_dir / TOP_DELTA_XLSX)
            logging.info("Preserved master Excel: %s | rows: %s", output_dir / MASTER_XLSX, len(previous_master))
            return 0

        master_candidate = build_current_master_candidate(raw_df)
        suppress_alerts = should_suppress_alerts(state_dir)
        master_df, delta_df, updates_df = classify_delta_updates(master_candidate, previous_master, suppress_alerts)
        top_delta_df = delta_df.sort_values("weight_score", ascending=False).head(20).copy() if not delta_df.empty else delta_df.copy()

        write_excel(delta_df, output_dir / DELTA_XLSX)
        write_excel(updates_df, output_dir / UPDATES_XLSX)
        write_excel(top_delta_df, output_dir / TOP_DELTA_XLSX)
        write_excel(master_df, output_dir / MASTER_XLSX)

        save_seen_ids(state_file, master_df.get("opportunity_key", pd.Series(dtype=str)).tolist())
        write_hash_version(hash_version_file)

        logging.info("Saved delta Excel: %s | rows: %s", output_dir / DELTA_XLSX, len(delta_df))
        logging.info("Saved updates Excel: %s | rows: %s", output_dir / UPDATES_XLSX, len(updates_df))
        logging.info("Saved top delta Excel: %s | rows: %s", output_dir / TOP_DELTA_XLSX, len(top_delta_df))
        logging.info("Saved master Excel: %s | rows: %s", output_dir / MASTER_XLSX, len(master_df))
        return 0

    except (RequestException, ValueError, RuntimeError, OSError) as exc:
        logging.exception("FAILED: %s", exc)
        return 1
    finally:
        release_lock(lock_file)


if __name__ == "__main__":
    raise SystemExit(main())
