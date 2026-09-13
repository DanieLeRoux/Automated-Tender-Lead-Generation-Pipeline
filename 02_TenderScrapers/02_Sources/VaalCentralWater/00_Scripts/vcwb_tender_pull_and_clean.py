"""
vcwb_tender_pull_and_clean (Vaal Central Water Board / Advertised Tenders)

Scrapes advertised tenders from the public Vaal Central Water advertised tenders page,
cleans/normalises the notice rows, filters to currently open tenders, scores entries,
and maintains:

1) MASTER Excel: latest current snapshot per opportunity_key
2) DELTA Excel: brand new tenders discovered this run only
3) UPDATES Excel: existing tenders whose key operational fields changed
4) TOP DELTA Excel: highest-scoring new tenders from the current run
5) RAW CSV/XLSX/HTML: unfiltered source output for audit/debugging

Typical daily run:
    python vcwb_tender_pull_and_clean.py

Optional:
    python vcwb_tender_pull_and_clean.py --include-expired
    python vcwb_tender_pull_and_clean.py --top-delta-limit 50
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import re
import sys
import time
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# pylint: disable=abstract-class-instantiated

SOURCE_NAME = "Vaal Central Water Board"
SOURCE_URL = "https://vaalcentralwater.co.za/advertised-tenders/"

# === Dynamic Folder Detection ===
SOURCE_FOLDER_NAME = "VaalCentralWater"
SCRIPTS_FOLDER_NAME = "00_Scripts"


def resolve_source_root() -> Path:
    """Resolve 02_Sources/[Source] from this script's location.

    Expected placement:
        02_TenderScrapers/02_Sources/VaalCentralWater/00_Scripts/this_script.py

    The fallback keeps the script usable during ad-hoc testing if it is run
    directly from the source folder or another temporary folder.
    """
    script_dir = Path(__file__).resolve().parent

    if script_dir.name == SCRIPTS_FOLDER_NAME:
        return script_dir.parent

    if script_dir.name == SOURCE_FOLDER_NAME:
        return script_dir

    return script_dir


SOURCE_ROOT = resolve_source_root()
BASE_ROOT = str(SOURCE_ROOT)

SCRIPT_DIR = str(SOURCE_ROOT / "00_Scripts")
RAW_DIR = str(SOURCE_ROOT / "01_RawData")
OUTPUT_DIR = str(SOURCE_ROOT / "02_Outputs")
LOG_DIR = str(SOURCE_ROOT / "03_Logs")
STATE_DIR = str(SOURCE_ROOT / "04_State")
METADATA_DIR = str(SOURCE_ROOT / "05_Metadata")

LOG_PATH = os.path.join(LOG_DIR, "vcwb_run.log")

MASTER_XLSX = os.path.join(OUTPUT_DIR, "vcwb_master.xlsx")
DELTA_XLSX = os.path.join(OUTPUT_DIR, "vcwb_delta.xlsx")
UPDATES_XLSX = os.path.join(OUTPUT_DIR, "vcwb_updates.xlsx")
TOP_DELTA_XLSX = os.path.join(OUTPUT_DIR, "vcwb_top_delta.xlsx")

RAW_CSV = os.path.join(RAW_DIR, "vcwb_raw.csv")
RAW_XLSX = os.path.join(RAW_DIR, "vcwb_raw_full.xlsx")
RAW_HTML = os.path.join(RAW_DIR, "vcwb_raw_page.html")

STATE_FILE_DEFAULT = os.path.join(STATE_DIR, "seen_ids.json")
HASH_VERSION_FILE = os.path.join(STATE_DIR, "hash_version.txt")
LOCK_FILE = os.path.join(STATE_DIR, "run.lock")

TIMEOUT_S = 60
SLEEP_S = 0.2

TOP_DELTA_LIMIT_DEFAULT = 20
TOP_DELTA_MIN_SCORE_DEFAULT = 0

# Fields that define a meaningful operational update.
# Document names/URLs are included because VCWB often posts addenda, errata and tender extensions.
HASH_COLS = [
    "description",
    "date_uploaded",
    "closing_date",
    "document_names",
    "document_urls",
    "notice_type",
]

# Source-specific: all notices come from VCWB, so ranking is based on description keywords.
WEIGHT_DESCRIPTION_KEYWORD = 1

DESCRIPTION_KEYWORD_RULES = {
    "Water and Sanitation": ["Water and Sanitation", "Water"],
    "Bulk Infrastructure": ["Bulk Infrastructure"],
    "Infrastructure Masterplan": ["Infrastructure Masterplan", "Masterplan", "Master Plan"],
    "Design": ["Design"],
    "Implementation": ["Implementation"],
    "Refurbishment": ["Refurbishment", "Refurbish"],
    "Construction": ["Construction"],
    "Maintenance": ["Maintenance"],
    "Mechanical": ["Mechanical"],
    "Electrical": ["Electrical"],
    "Civil Engineering": ["Civil Engineering", "Civil"],
    "Structural Engineering": ["Structural Engineering"],
    "Professional Service Provider": ["Professional Service Provider", "PSP"],
    "Engineering": ["Engineering"],
    "Project Management": ["Project Management"],
    "Programme Management": ["Programme Management"],
    "Feasibility Study": ["Feasibility Study"],
    "Pipelines": ["Pipeline", "Pipelines"],
    "Pump Stations": ["Pump Station", "Pump Stations"],
    "Reservoir": ["Reservoir", "Reservoirs"],
    "Wastewater Treatment Works": ["Wastewater Treatment Works", "WWTW"],
    "Water Treatment Works": ["Water Treatment Works", "WTW"],
    "Fire Engineering": ["Fire", "Fire Preventative", "Fire Engineering"],
    "HVAC": ["HVAC"],
    "Advisory Services": ["Advisory Services"],
}

DEFAULT_COLUMNS = [
    "source",
    "opportunity_key",
    "reference_number",
    "notice_type",
    "description",
    "client",
    "date_uploaded",
    "closing_date",
    "document_names",
    "document_urls",
    "document_count",
    "source_url",
    "priority_client_match",
    "matched_description_keywords",
    "description_keyword_match_count",
    "weight_score",
    "row_hash",
    "run_timestamp",
    "change_type",
    "previous_row_hash",
]


def setup_logging() -> None:
    """Create the source folder tree and configure file + stdout logging."""
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
    """Requests session with retry/backoff for flaky CMS endpoints."""
    sess = requests.Session()
    sess.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-ZA,en;q=0.9",
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
        raise RuntimeError(f"Lock file exists (another run may be active): {lock_path}")
    with open(lock_path, "w", encoding="utf-8") as fh:
        fh.write(f"pid={os.getpid()}\n")


def release_lock(lock_path: str) -> None:
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
    except OSError:
        logging.exception("Failed to remove lock file: %s", lock_path)


def _normalise_key(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).casefold().strip()
    text = re.sub(r"\s+", " ", text)
    text = text.replace("\\", "/")
    return text


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


def _save_seen_ids(path: str, ids: Iterable[str]) -> None:
    try:
        norm = sorted(_normalise_key(x) for x in ids if _normalise_key(x))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(norm, fh, ensure_ascii=False)
        os.replace(tmp, path)
    except (OSError, TypeError) as e:
        logging.exception("Failed to save state file: %s -- %s", path, e)


def fetch_html() -> str:
    sess = build_session()
    response = sess.get(SOURCE_URL, timeout=TIMEOUT_S)
    response.raise_for_status()
    return response.text


def save_raw_html(html: str) -> None:
    os.makedirs(RAW_DIR, exist_ok=True)
    with open(RAW_HTML, "w", encoding="utf-8") as fh:
        fh.write(html)
    logging.info("Saved raw HTML: %s", RAW_HTML)


def _clean_text(text: Any) -> str:
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text).replace("\xa0", " ")).strip()


def _tokenise_content(soup: BeautifulSoup) -> List[Dict[str, str]]:
    """Create a linear stream of visible text and links from the main page content."""
    for bad in soup(["script", "style", "noscript", "svg"]):
        bad.decompose()

    root: Tag = soup.find("main") or soup.find("article") or soup.find("body") or soup  # type: ignore[assignment]

    tokens: List[Dict[str, str]] = []
    block_tags = {"p", "div", "span", "li", "td", "th", "h1", "h2", "h3", "h4", "strong", "em"}

    def append_text(value: str) -> None:
        text = _clean_text(value)
        if not text:
            return
        # Split obvious label/value lines without destroying dates.
        parts = [p.strip() for p in re.split(r"[\r\n]+", text) if p.strip()]
        for part in parts:
            tokens.append({"kind": "text", "text": part, "url": ""})

    for elem in root.descendants:
        if isinstance(elem, Tag):
            if elem.name == "a":
                link_text = _clean_text(elem.get_text(" ", strip=True))
                href = elem.get("href")
                if link_text or href:
                    tokens.append(
                        {
                            "kind": "link",
                            "text": link_text or "Document",
                            "url": urljoin(SOURCE_URL, str(href or "")),
                        }
                    )
            elif elem.name in {"br", "hr"}:
                continue
        else:
            parent = elem.parent if hasattr(elem, "parent") else None
            if isinstance(parent, Tag) and parent.name == "a":
                continue
            if isinstance(parent, Tag) and parent.name in block_tags:
                append_text(str(elem))

    # Compact duplicates created by nested tags.
    compacted: List[Dict[str, str]] = []
    for tok in tokens:
        if not tok["text"]:
            continue
        if compacted and compacted[-1] == tok:
            continue
        compacted.append(tok)

    return compacted


def _is_reference_number(text: str) -> bool:
    text = _clean_text(text)
    if not text:
        return False
    if text.lower().startswith("alert:"):
        return False
    if len(text) > 80:
        return False
    # Known VCWB/Bloem Water reference prefixes seen on the tender page.
    return bool(re.match(r"^(?:VCW|VCWB|BW)\s*[A-Z0-9/\\\- ]+\d{2,4}$", text, flags=re.IGNORECASE))


def _find_content_start(tokens: Sequence[Dict[str, str]]) -> int:
    for i, tok in enumerate(tokens):
        if tok["text"].casefold() == "reference no.":
            return i + 1
    for i, tok in enumerate(tokens):
        if _is_reference_number(tok["text"]):
            return i
    return 0


def _group_notice_blocks(tokens: Sequence[Dict[str, str]]) -> List[List[Dict[str, str]]]:
    start = _find_content_start(tokens)
    blocks: List[List[Dict[str, str]]] = []
    current: List[Dict[str, str]] = []

    for tok in tokens[start:]:
        text_cf = tok["text"].casefold()
        if text_cf in {"contact info", "copyright - vaal central water 2023"} or text_cf.startswith("contact info"):
            break

        if tok["kind"] == "text" and _is_reference_number(tok["text"]):
            if current:
                blocks.append(current)
            current = [tok]
        else:
            if current:
                current.append(tok)

    if current:
        blocks.append(current)

    return blocks


def _parse_date_loose(value: Any) -> pd.Timestamp:
    """Parse the inconsistent date formats used on the VCWB page."""
    text = _clean_text(value)
    if not text:
        return pd.NaT
    if text in {"–", "-", "—"}:
        return pd.NaT

    text = re.sub(r"(?i)date uploaded\s*:", "", text)
    text = re.sub(r"(?i)closing date\s*:", "", text)
    text = re.sub(r"(?i)advert uploaded", "", text)
    text = re.sub(r"(?i)erratum uploaded", "", text)
    text = text.strip(" :-–—")

    # Common patterns: 15 May 2026, 17 APRIL 2026, 10.10.2025, 26.03.2024.
    patterns = [
        r"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\b",
        r"\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b",
    ]
    match_text = text
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            match_text = m.group(0)
            break

    return pd.to_datetime(match_text, errors="coerce", dayfirst=True)


def _first_date_after_label(texts: Sequence[str], label_regex: str) -> pd.Timestamp:
    for idx, text in enumerate(texts):
        if re.search(label_regex, text, flags=re.IGNORECASE):
            # Date can be in same token or next few tokens.
            candidate = _parse_date_loose(text)
            if pd.notna(candidate):
                return candidate
            for nxt in texts[idx + 1 : idx + 4]:
                candidate = _parse_date_loose(nxt)
                if pd.notna(candidate):
                    return candidate
    return pd.NaT


def _extract_description(texts: Sequence[str]) -> str:
    description_parts: List[str] = []
    for text in texts[1:]:
        if re.search(r"(?i)date uploaded\s*:|advert uploaded|closing date\s*:", text):
            break
        if text.casefold() in {"date uploaded:", "closing date:"}:
            break
        description_parts.append(text)
    return _clean_text(" ".join(description_parts))


def _infer_notice_type(reference_number: str, description: str) -> str:
    combined = f"{reference_number} {description}".casefold()
    if "request for quotation" in combined or "rfq" in combined:
        return "RFQ"
    if "request for proposal" in combined or "rfp" in combined:
        return "RFP"
    if "expression of interest" in combined or "eoi" in combined:
        return "EOI"
    if "tender" in combined:
        return "Tender"
    return "Tender"


def _parse_notice_block(block: Sequence[Dict[str, str]]) -> Optional[Dict[str, Any]]:
    if not block:
        return None

    reference_number = _clean_text(block[0]["text"])
    if not _is_reference_number(reference_number):
        return None

    texts = [_clean_text(tok["text"]) for tok in block if tok["kind"] == "text" and _clean_text(tok["text"])]
    links = [tok for tok in block if tok["kind"] == "link" and tok.get("url")]

    description = _extract_description(texts)
    date_uploaded = _first_date_after_label(texts, r"date uploaded\s*:|advert uploaded")
    closing_date = _first_date_after_label(texts, r"closing date\s*:")

    document_names = "; ".join(_clean_text(link["text"]) for link in links if _clean_text(link["text"]))
    document_urls = "; ".join(link["url"] for link in links if link.get("url"))

    opportunity_key = "vcwb:" + _normalise_key(reference_number)

    return {
        "source": SOURCE_NAME,
        "opportunity_key": opportunity_key,
        "reference_number": reference_number,
        "notice_type": _infer_notice_type(reference_number, description),
        "description": description,
        "client": SOURCE_NAME,
        "date_uploaded": date_uploaded,
        "closing_date": closing_date,
        "document_names": document_names,
        "document_urls": document_urls,
        "document_count": len(links),
        "source_url": SOURCE_URL,
    }


def parse_vcwb_page(html: str) -> pd.DataFrame:
    soup = BeautifulSoup(html, "html.parser")
    tokens = _tokenise_content(soup)
    blocks = _group_notice_blocks(tokens)

    logging.info("VCWB parser tokens found: %d", len(tokens))
    logging.info("VCWB notice blocks found: %d", len(blocks))

    rows: List[Dict[str, Any]] = []
    for block in blocks:
        row = _parse_notice_block(block)
        if row:
            rows.append(row)

    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame(columns=[c for c in DEFAULT_COLUMNS if c not in {"row_hash", "run_timestamp", "change_type", "previous_row_hash", "priority_client_match", "matched_description_keywords", "description_keyword_match_count", "weight_score"}])

    logging.info("Parsed VCWB notice rows: %d | cols: %d", len(df), len(df.columns))
    return df


def filter_current(df: pd.DataFrame, include_expired: bool = False) -> pd.DataFrame:
    out = df.copy()
    if include_expired:
        logging.info("Expired tenders included by CLI switch: %d rows", len(out))
        return out

    before = len(out)
    if "closing_date" not in out.columns:
        logging.warning("No closing_date column found; treating clean output as empty.")
        return out.iloc[0:0].copy()

    out["closing_date"] = pd.to_datetime(out["closing_date"], errors="coerce")
    today = pd.Timestamp.today().normalize()
    out = out[pd.notna(out["closing_date"]) & (out["closing_date"].dt.normalize() >= today)].copy()
    logging.info("Filtered current tenders by closing date: %d -> %d rows", before, len(out))
    return out


def add_priority_scoring(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["priority_client_match"] = True  # all records are from the target client/source

    def matched_keywords(description: Any) -> str:
        desc_norm = _normalise_search_text(description)
        matches: List[str] = []
        for label, phrases in DESCRIPTION_KEYWORD_RULES.items():
            if any(_contains_phrase(desc_norm, phrase) for phrase in phrases):
                matches.append(label)
        return "; ".join(matches)

    if "description" not in out.columns:
        out["matched_description_keywords"] = ""
    else:
        out["matched_description_keywords"] = out["description"].apply(matched_keywords)

    out["description_keyword_match_count"] = out["matched_description_keywords"].apply(
        lambda x: 0 if not x else len(str(x).split("; "))
    )
    out["weight_score"] = out["description_keyword_match_count"].astype(int) * WEIGHT_DESCRIPTION_KEYWORD

    sort_cols = [c for c in ["weight_score", "closing_date", "date_uploaded"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

    return out


def build_top_delta(df: pd.DataFrame, limit: int, min_score: int) -> pd.DataFrame:
    out = df.copy()
    if "weight_score" not in out.columns:
        out["weight_score"] = 0

    out = out[out["weight_score"].fillna(0).astype(int) >= int(min_score)].copy()

    sort_cols = [c for c in ["weight_score", "closing_date", "date_uploaded"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

    if limit and int(limit) > 0:
        out = out.head(int(limit)).copy()

    return out


def compute_row_hash(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    if df is None or df.empty:
        return pd.Series([], dtype="object", index=df.index if df is not None else None)

    use_cols = [c for c in cols if c in df.columns]
    if not use_cols:
        return pd.Series([""] * len(df), index=df.index, dtype="object")

    tmp = df[use_cols].copy()
    for c in use_cols:
        if pd.api.types.is_datetime64_any_dtype(tmp[c]):
            tmp[c] = tmp[c].dt.strftime("%Y-%m-%d %H:%M:%S").fillna("")
        else:
            tmp[c] = tmp[c].fillna("").astype(str).str.strip()

    joined = tmp.apply(lambda r: "||".join(r.values.tolist()), axis=1)
    return joined.apply(lambda s: hashlib.sha256(s.encode("utf-8")).hexdigest())


def write_excel_with_table(df: pd.DataFrame, path: str, sheet_name: str, table_name: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp.xlsx"

    out_df = df.copy() if df is not None else pd.DataFrame()
    if out_df.shape[1] == 0:
        out_df = pd.DataFrame(columns=DEFAULT_COLUMNS)

    with pd.ExcelWriter(
        tmp,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        out_df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]

        nrows, ncols = out_df.shape
        columns = [{"header": col} for col in out_df.columns]

        if nrows > 0 and ncols > 0:
            worksheet.add_table(0, 0, nrows, ncols - 1, {"name": table_name, "columns": columns})
        elif ncols > 0:
            worksheet.autofilter(0, 0, 0, ncols - 1)

        worksheet.freeze_panes(1, 0)

    os.replace(tmp, path)


def load_master(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name="Master", engine="openpyxl")
    except Exception as e:  # pylint: disable=broad-exception-caught
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
        out = pd.concat([master_part, incoming_part], ignore_index=True, sort=False)

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

    sort_cols = [c for c in ["weight_score", "closing_date", "date_uploaded"] if c in out.columns]
    if sort_cols:
        for col in ["closing_date", "date_uploaded"]:
            if col in out.columns:
                out[col] = pd.to_datetime(out[col], errors="coerce")
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last", kind="mergesort")

    return out

def _normalise_master(master_df: pd.DataFrame, include_expired: bool) -> pd.DataFrame:
    if master_df.empty:
        return master_df
    out = master_df.copy()

    if "opportunity_key" in out.columns:
        out["opportunity_key"] = out["opportunity_key"].fillna("").astype(str).str.strip().str.lower()
        out = out[out["opportunity_key"].ne("")]

    for col in ["date_uploaded", "closing_date"]:
        if col in out.columns:
            out[col] = pd.to_datetime(out[col], errors="coerce")

    out = filter_current(out, include_expired=include_expired)

    # Backfill columns introduced after earlier versions.
    for col in DEFAULT_COLUMNS:
        if col not in out.columns:
            if col in {"priority_client_match"}:
                out[col] = True
            elif col in {"description_keyword_match_count", "weight_score", "document_count"}:
                out[col] = 0
            else:
                out[col] = ""

    if "row_hash" not in out.columns or out["row_hash"].fillna("").eq("").all():
        out["row_hash"] = compute_row_hash(out, HASH_COLS)

    return out




def _is_migration_baseline_run(hash_version_path: str) -> bool:
    """First migrated run suppresses delta/update alerts to avoid false positives."""
    return not os.path.exists(hash_version_path)


def _write_hash_version(path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        fh.write("vcwb_hash_v1\n")
    os.replace(tmp, path)


def _empty_like(df: pd.DataFrame, change_type: str) -> pd.DataFrame:
    out = df.iloc[0:0].copy()
    out["change_type"] = change_type
    out["previous_row_hash"] = ""
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description="Scrape and clean VCWB advertised tenders")
    parser.add_argument("--include-expired", action="store_true", help="include tenders where closing_date has passed")
    parser.add_argument("--state-file", default=STATE_FILE_DEFAULT, help="path to JSON file storing seen opportunity keys")
    parser.add_argument("--top-delta-limit", type=int, default=TOP_DELTA_LIMIT_DEFAULT, help="maximum rows in vcwb_top_delta.xlsx; use 0 for all")
    parser.add_argument("--top-delta-min-score", type=int, default=TOP_DELTA_MIN_SCORE_DEFAULT, help="minimum weight_score for top delta output")
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
        save_raw_html(html)

        raw_df = parse_vcwb_page(html)

        raw_df.to_csv(RAW_CSV, index=False)
        logging.info("Saved raw CSV: %s", RAW_CSV)

        write_excel_with_table(raw_df, RAW_XLSX, sheet_name="Raw_Data", table_name="tbl_raw")
        logging.info("Saved raw Excel: %s", RAW_XLSX)

        if raw_df.empty:
            raise ValueError(
                "No rows were parsed from the VCWB advertised tenders page. "
                "Raw diagnostics were saved, but clean outputs/master/state were not overwritten."
            )

        mapped = raw_df.copy()
        for col in ["date_uploaded", "closing_date"]:
            if col in mapped.columns:
                mapped[col] = pd.to_datetime(mapped[col], errors="coerce")

        mapped = filter_current(mapped, include_expired=args.include_expired)
        mapped = add_priority_scoring(mapped)
        mapped["row_hash"] = compute_row_hash(mapped, HASH_COLS)
        mapped["run_timestamp"] = run_ts

        master_df = load_master(MASTER_XLSX)
        master_df = _normalise_master(master_df, include_expired=args.include_expired)

        migration_baseline = _is_migration_baseline_run(HASH_VERSION_FILE)
        if migration_baseline:
            logging.info(
                "Migration baseline mode active: suppressing delta/update/top-delta alerts for this first migrated run."
            )
            delta_df = _empty_like(mapped, "NEW")
            updates_df = _empty_like(mapped, "UPDATED")
        elif not master_df.empty and "opportunity_key" in master_df.columns:
            master_ids = set(master_df["opportunity_key"].fillna("").astype(str).str.lower().tolist())
            master_hash_map = dict(
                zip(
                    master_df["opportunity_key"].fillna("").astype(str).str.lower().tolist(),
                    master_df["row_hash"].fillna("").astype(str).tolist(),
                )
            )

            is_new = ~mapped["opportunity_key"].isin(master_ids) if not mapped.empty else pd.Series([], dtype=bool)
            delta_df = mapped[is_new].copy() if not mapped.empty else mapped.copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""

            mapped_prev_hash = mapped["opportunity_key"].map(master_hash_map).fillna("") if not mapped.empty else pd.Series([], dtype="object")
            is_updated = (~is_new) & (mapped["row_hash"] != mapped_prev_hash) if not mapped.empty else pd.Series([], dtype=bool)
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

        sort_cols = [c for c in ["weight_score", "closing_date", "date_uploaded"] if c in delta_df.columns]
        if sort_cols:
            ascending = [False if c == "weight_score" else True for c in sort_cols]
            delta_df = delta_df.sort_values(sort_cols, ascending=ascending, na_position="last")
            updates_df = updates_df.sort_values(sort_cols, ascending=ascending, na_position="last")

        top_delta_df = build_top_delta(delta_df, limit=args.top_delta_limit, min_score=args.top_delta_min_score)

        logging.info("Delta (NEW) rows this run: %d", len(delta_df))
        logging.info("Updates (CHANGED) rows this run: %d", len(updates_df))
        logging.info("Top delta rows this run: %d", len(top_delta_df))

        write_excel_with_table(delta_df, DELTA_XLSX, sheet_name="Delta", table_name="tbl_delta")
        logging.info("Saved delta Excel: %s | rows: %d", DELTA_XLSX, len(delta_df))

        write_excel_with_table(top_delta_df, TOP_DELTA_XLSX, sheet_name="Top_Delta", table_name="tbl_top_delta")
        logging.info("Saved top delta Excel: %s | rows: %d", TOP_DELTA_XLSX, len(top_delta_df))

        write_excel_with_table(updates_df, UPDATES_XLSX, sheet_name="Updates", table_name="tbl_updates")
        logging.info("Saved updates Excel: %s | rows: %d", UPDATES_XLSX, len(updates_df))

        updated_master = append_and_dedupe_master(master_df, mapped)
        write_excel_with_table(updated_master, MASTER_XLSX, sheet_name="Master", table_name="tbl_master")
        logging.info("Saved master Excel: %s | rows: %d", MASTER_XLSX, len(updated_master))

        if "opportunity_key" in updated_master.columns:
            _save_seen_ids(args.state_file, updated_master["opportunity_key"].tolist())

        _write_hash_version(HASH_VERSION_FILE)

        time.sleep(SLEEP_S)
        return 0

    except (requests.RequestException, ValueError, RuntimeError, OSError) as e:
        logging.exception("FAILED: %s", e)
        return 1
    finally:
        release_lock(LOCK_FILE)


if __name__ == "__main__":
    raise SystemExit(main())
