"""
dbsa_procurement_pull_and_clean (Delta + Updates + Master)

Fetch open RFP/RFQ/RFI notices from the DBSA public procurement page,
clean/normalise them, score them, and maintain:

1) MASTER Excel: latest snapshot per tender / opportunity_key
2) DELTA Excel: brand new tenders discovered this run only
3) UPDATES Excel: existing tenders whose key fields changed
4) TOP DELTA Excel: highest-scoring new tenders from the current run

The DBSA procurement page is rendered as webpage content, not a clean public API.
This scraper uses the public page only and captures the linked tender documents.

Typical daily run:
    python dbsa_procurement_pull_and_clean.py

Optional:
    python dbsa_procurement_pull_and_clean.py --include-expired
    python dbsa_procurement_pull_and_clean.py --include-cancellations
    python dbsa_procurement_pull_and_clean.py --top-delta-limit 50
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import re
import sys
from html.parser import HTMLParser
from typing import Any, Dict, List, Optional, Set
from urllib.parse import urljoin

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# pylint: disable=abstract-class-instantiated

SOURCE_NAME = "DBSA"
PORTAL_URL = "https://www.dbsa.org/procurement"
BASE_URL = "https://www.dbsa.org"

# === Folder structure ===
# Expected location of this script:
#   02_TenderScrapers\02_Sources\DBSA\00_Scripts\dbsa_procurement_pull_and_clean.py
#
# Paths are derived from the script location so the scraper remains portable
# inside the reorganised TenderScrapers architecture.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
TENDER_SCRAPERS_ROOT = os.path.abspath(os.path.join(SOURCE_ROOT, "..", ".."))

RAW_DIR = os.path.join(SOURCE_ROOT, "01_RawData")
OUTPUT_DIR = os.path.join(SOURCE_ROOT, "02_Outputs")
LOG_DIR = os.path.join(SOURCE_ROOT, "03_Logs")
STATE_DIR = os.path.join(SOURCE_ROOT, "04_State")
METADATA_DIR = os.path.join(SOURCE_ROOT, "05_Metadata")

LOG_PATH = os.path.join(LOG_DIR, "dbsa_run.log")

MASTER_XLSX = os.path.join(OUTPUT_DIR, "dbsa_master.xlsx")
DELTA_XLSX = os.path.join(OUTPUT_DIR, "dbsa_delta.xlsx")
UPDATES_XLSX = os.path.join(OUTPUT_DIR, "dbsa_updates.xlsx")
TOP_DELTA_XLSX = os.path.join(OUTPUT_DIR, "dbsa_top_delta.xlsx")

RAW_CSV = os.path.join(RAW_DIR, "dbsa_raw.csv")
RAW_XLSX = os.path.join(RAW_DIR, "dbsa_raw_full.xlsx")
RAW_HTML = os.path.join(RAW_DIR, "dbsa_raw_page.html")

STATE_FILE_DEFAULT = os.path.join(STATE_DIR, "seen_ids.json")
HASH_VERSION_FILE_DEFAULT = os.path.join(STATE_DIR, "hash_version.txt")
LOCK_FILE = os.path.join(STATE_DIR, "run.lock")

TIMEOUT_S = 60

# Hash fields that define a meaningful change in a public DBSA tender notice.
# Hash fields that define a meaningful operational change.
# Full document_urls are intentionally excluded because CMS/download URL
# representations can change without the tender itself changing. Document names
# and operational dates still capture meaningful changes.
HASH_COLS = [
    "description",
    "document_names",
    "briefing_session_details",
    "date_published",
    "closing_date",
    "closing_datetime",
    "cancellation_flag",
    "general_notice_flag",
]

HASH_VERSION = "dbsa_hash_v2_structure_2026_07"

# Top-delta shortlist defaults.
TOP_DELTA_LIMIT_DEFAULT = 20
TOP_DELTA_MIN_SCORE_DEFAULT = 0

# Scoring constants. DBSA is a priority source, so every current DBSA entry gets
# a base source score, then description keywords create the ranking difference.
WEIGHT_SOURCE_MATCH = 5
WEIGHT_DESCRIPTION_KEYWORD = 1

DESCRIPTION_KEYWORD_RULES = {
    "Water and Sanitation": ["Water and Sanitation"],
    "Housing": ["Housing", "Human Settlements"],
    "Education": ["Education", "School", "Schools"],
    "Transport": ["Transport", "Rail", "Road", "Roads", "Ports"],
    "Nuclear": ["Nuclear"],
    "Bus Rapid Transit (BRT)": ["Bus Rapid Transit", "BRT", "IRT"],
    "International Transition Plan Network (ITPN)": [
        "International Transition Plan Network",
        "ITPN",
    ],
    "Design": ["Design"],
    "Refurbishment": ["Refurbishment", "Refurbish"],
    "Engineering, Procurement, and Construction (EPC)": [
        "Engineering, Procurement, and Construction",
        "Engineering Procurement and Construction",
        "EPC",
    ],
    "Fire Engineering": ["Fire Engineering", "Fire"],
    "HVAC": ["HVAC"],
    "Mechanical": ["Mechanical"],
    "Electrical": ["Electrical"],
    "Electronic": ["Electronic"],
    "Wet Services": ["Wet Services"],
    "Project Management": ["Project Management"],
    "Programme Management": ["Programme Management", "Program Management"],
    "Project Management Office (PMO)": ["Project Management Office", "PMO"],
    "Advisory Services": ["Advisory Services", "Advisor", "Transaction Advisor"],
    "Structural Engineering": ["Structural Engineering"],
    "Civil Engineering": ["Civil Engineering", "Civil"],
    "Feasibility Study": ["Feasibility Study", "Feasibility"],
    "Professional Services": [
        "Professional Services",
        "Professional Service Provider",
        "PSP",
    ],
    "Consulting": ["Consulting", "Consultant", "Consultancy"],
    "Infrastructure": ["Infrastructure"],
    "Wastewater": ["Wastewater", "Waste Water", "Blackwater", "Greywater"],
    "Energy": ["Energy", "Power", "IPP", "IPPPP"],
    "Spatial Data": ["Spatial", "Spatialised Data", "GIS"],
}

EXPECTED_OUTPUT_COLUMNS = [
    "source",
    "opportunity_key",
    "tender_number",
    "notice_type",
    "description",
    "document_names",
    "document_urls",
    "briefing_session_details",
    "date_published",
    "closing_date",
    "closing_datetime",
    "source_url",
    "cancellation_flag",
    "general_notice_flag",
    "priority_source_match",
    "matched_description_keywords",
    "description_keyword_match_count",
    "weight_score",
    "row_hash",
    "run_timestamp",
    "change_type",
    "previous_row_hash",
]


class ProcurementPageParser(HTMLParser):
    """Tokenise DBSA page text and links in their rendered order."""

    def __init__(self):
        super().__init__()
        self.tokens: List[Dict[str, str]] = []
        self._current_link_href: Optional[str] = None
        self._current_link_text: List[str] = []
        self._skip_depth = 0

    def handle_starttag(self, tag: str, attrs: List[tuple[str, Optional[str]]]):
        tag = tag.lower()
        attrs_dict = dict(attrs)
        if tag in {"script", "style", "noscript", "svg"}:
            self._skip_depth += 1
            return
        if self._skip_depth:
            return
        if tag == "a":
            self._current_link_href = attrs_dict.get("href")
            self._current_link_text = []

    def handle_data(self, data: str):
        if self._skip_depth:
            return
        text = " ".join(str(data).split())
        if not text:
            return
        if self._current_link_href is not None:
            self._current_link_text.append(text)
        else:
            self.tokens.append({"kind": "text", "text": text, "href": ""})

    def handle_endtag(self, tag: str):
        tag = tag.lower()
        if tag in {"script", "style", "noscript", "svg"} and self._skip_depth:
            self._skip_depth -= 1
            return
        if self._skip_depth:
            return
        if tag == "a" and self._current_link_href is not None:
            link_text = " ".join(" ".join(self._current_link_text).split())
            if link_text:
                self.tokens.append(
                    {
                        "kind": "link",
                        "text": link_text,
                        "href": urljoin(BASE_URL, self._current_link_href),
                    }
                )
            self._current_link_href = None
            self._current_link_text = []


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


def _load_hash_version(path: str) -> str:
    try:
        if not os.path.exists(path):
            return ""
        with open(path, "r", encoding="utf-8") as fh:
            return fh.read().strip()
    except OSError as e:
        logging.exception("Failed to load hash version file: %s -- %s", path, e)
        return ""


def _save_hash_version(path: str, version: str) -> None:
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            fh.write(version)
        os.replace(tmp, path)
    except OSError as e:
        logging.exception("Failed to save hash version file: %s -- %s", path, e)


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


def tokenise_page(html: str) -> List[Dict[str, str]]:
    parser = ProcurementPageParser()
    parser.feed(html)
    return parser.tokens


def _is_section_start(text: str) -> bool:
    """Detect the beginning of DBSA's open procurement list.

    DBSA sometimes renders the section heading as one text node and sometimes
    splits it across several nodes. Keep this intentionally broad; the actual
    notice extraction is still controlled by _is_notice_header.
    """
    norm = _normalise_search_text(text)
    return (
        "open rfq" in norm
        or "open rfp" in norm
        or "rfp s and rfr" in norm
        or "date published closing date" in norm
    )


def _is_section_end(text: str) -> bool:
    norm = _normalise_search_text(text)
    return (
        "the dbsa does not have any procurement service provider" in norm
        or "for more information please see" in norm
        or norm == "what we offer"
    )


def _is_notice_header(text: str) -> bool:
    """Detect the start of a DBSA tender/notice block.

    The page sometimes separates the tender number from the description, so the
    colon and description are not required for RFP/RFQ/RFI/RFR headers.
    """
    cleaned = " ".join(str(text).split())
    if re.match(r"^(RFP|RFQ|RFI|RFR)\s+[A-Za-z0-9_./-]+\s*:?.*", cleaned, flags=re.IGNORECASE):
        return True
    if re.match(r"^GAS\s+IPPPP\s+Bid\s+Window\s+\d+\s*:?", cleaned, flags=re.IGNORECASE):
        return True
    if "general notice" in cleaned.casefold() and len(cleaned) > 20:
        return True
    return False


def extract_notice_blocks(tokens: List[Dict[str, str]]) -> List[List[Dict[str, str]]]:
    """Return token blocks, one block per DBSA notice.

    Primary mode starts after the open-procurement heading. Fallback mode scans
    the whole page for RFP/RFQ/RFI/RFR headers. This protects the scraper when
    DBSA changes small bits of the surrounding CMS layout.
    """
    in_section = False
    current: List[Dict[str, str]] = []
    blocks: List[List[Dict[str, str]]] = []

    for token in tokens:
        text = token.get("text", "")
        if not in_section:
            if _is_section_start(text):
                in_section = True
            continue

        if _is_section_end(text):
            if current:
                blocks.append(current)
            break

        if token.get("kind") == "text" and _is_notice_header(text):
            if current:
                blocks.append(current)
            current = [token]
        elif current:
            current.append(token)

    if current:
        blocks.append(current)

    # Fallback: if the section heading was split or renamed, scan the full page.
    if not blocks:
        current = []
        for token in tokens:
            text = token.get("text", "")
            if _is_section_end(text):
                if current:
                    blocks.append(current)
                break
            if token.get("kind") == "text" and _is_notice_header(text):
                if current:
                    blocks.append(current)
                current = [token]
            elif current:
                current.append(token)
        if current:
            blocks.append(current)

    return blocks


def _extract_tender_number_and_description(header: str) -> tuple[str, str, str]:
    """Return tender_number, notice_type, description from a DBSA notice header."""
    cleaned = " ".join(str(header).split())
    if ":" in cleaned:
        left, right = cleaned.split(":", 1)
        tender_number = left.strip()
        description = right.strip()
    else:
        tender_number = cleaned.strip()
        description = ""

    first = tender_number.split()[0].upper() if tender_number.split() else ""
    if first in {"RFP", "RFQ", "RFI", "RFR"}:
        notice_type = first
    elif "general notice" in cleaned.casefold():
        notice_type = "General Notice"
    else:
        notice_type = "Notice"

    return tender_number, notice_type, description


MONTH_PATTERN = (
    r"January|February|March|April|May|June|July|August|September|October|November|December|"
    r"Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec"
)

DATE_CORE_PATTERN = (
    r"(?:"
    r"\b\d{1,2}[\s-]+(?:" + MONTH_PATTERN + r")[\s-]+\d{4}"
    r"|"
    r"\b\d{4}[/-]\d{1,2}[/-]\d{1,2}"
    r"|"
    r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}"
    r")"
)

TIME_SUFFIX_PATTERN = r"(?:\s*(?:at|@)?\s*\d{1,2}(?:[:Hh]\d{2})?)?"

DATE_PATTERN = re.compile(
    DATE_CORE_PATTERN + TIME_SUFFIX_PATTERN,
    flags=re.IGNORECASE,
)

LABELED_CLOSING_PATTERN = re.compile(
    r"(?:closing\s+date(?:\s+and\s+time)?|closing)\s*[:\-]?\s*(" + DATE_CORE_PATTERN + TIME_SUFFIX_PATTERN + r")",
    flags=re.IGNORECASE,
)

LABELED_PUBLISHED_PATTERN = re.compile(
    r"(?:date\s+published|published)\s*[:\-]?\s*(" + DATE_CORE_PATTERN + TIME_SUFFIX_PATTERN + r")",
    flags=re.IGNORECASE,
)


def _clean_date_text(value: str) -> str:
    text = " ".join(str(value).split())
    text = re.sub(r"\bSept\b", "Sep", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*@\s*", " ", text)
    text = re.sub(r"\s+at\s+", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"(\d{1,2})[Hh](\d{2})", r"\1:\2", text)
    return text


def _parse_datetime(value: str) -> pd.Timestamp:
    text = _clean_date_text(value)
    if not text:
        return pd.NaT

    # Avoid pandas warnings by parsing ISO-style values with year-first logic.
    if re.match(r"^\d{4}[/-]\d{1,2}[/-]\d{1,2}", text):
        return pd.to_datetime(text, errors="coerce", dayfirst=False)

    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def _extract_dates(block_text: str) -> tuple[pd.Timestamp, pd.Timestamp, pd.Timestamp, List[re.Match[str]]]:
    """Return date_published, closing_date, closing_datetime and all date matches.

    DBSA changes the text layout frequently. Prefer explicitly labelled
    "Date Published" and "Closing Date and Time" values; if those are not
    present, fall back to the last two date-like values in the block.
    """
    matches = list(DATE_PATTERN.finditer(block_text))

    published_label = LABELED_PUBLISHED_PATTERN.search(block_text)
    closing_label_matches = list(LABELED_CLOSING_PATTERN.finditer(block_text))

    date_published = _parse_datetime(published_label.group(1)) if published_label else pd.NaT
    closing_datetime = _parse_datetime(closing_label_matches[-1].group(1)) if closing_label_matches else pd.NaT

    if pd.isna(date_published) or pd.isna(closing_datetime):
        if len(matches) >= 2:
            if pd.isna(date_published):
                date_published = _parse_datetime(matches[-2].group(0))
            if pd.isna(closing_datetime):
                closing_datetime = _parse_datetime(matches[-1].group(0))
        elif len(matches) == 1:
            # A single date is more often a publication/briefing date than a closing date.
            if pd.isna(date_published):
                date_published = _parse_datetime(matches[0].group(0))

    closing_date = closing_datetime.normalize() if pd.notna(closing_datetime) else pd.NaT
    return date_published, closing_date, closing_datetime, matches


def _extract_briefing_details(block_text: str, date_matches: List[re.Match[str]]) -> str:
    norm = block_text.casefold()
    idx = norm.find("briefing session")
    if idx < 0:
        return ""

    # The publication date is normally the penultimate date in the block.
    # Everything from "Briefing Session" up to that date is treated as briefing details.
    end = date_matches[-2].start() if len(date_matches) >= 2 else len(block_text)
    # Include "Compulsory" if it appears immediately before "Briefing Session".
    compulsory_idx = norm.rfind("compulsory", 0, idx)
    start = compulsory_idx if compulsory_idx >= 0 and idx - compulsory_idx < 20 else idx
    details = block_text[start:end].strip(" -:,;")
    return " ".join(details.split())


def parse_notice_blocks(blocks: List[List[Dict[str, str]]]) -> pd.DataFrame:
    """Parse DBSA notice token blocks into a raw dataframe."""
    rows: List[Dict[str, Any]] = []

    for block in blocks:
        if not block:
            continue

        header = block[0].get("text", "")
        tender_number, notice_type, description = _extract_tender_number_and_description(header)
        if not tender_number:
            continue

        # If DBSA split the number and description into separate HTML text nodes,
        # recover the description from the text immediately after the header.
        if not description:
            desc_parts: List[str] = []
            for token in block[1:]:
                text = token.get("text", "")
                text_norm = _normalise_search_text(text)
                if token.get("kind") == "link":
                    break
                if "briefing session" in text_norm or DATE_PATTERN.search(text):
                    break
                if text:
                    desc_parts.append(text)
            description = " ".join(" ".join(desc_parts).split())

        document_names: List[str] = []
        document_urls: List[str] = []
        for token in block[1:]:
            if token.get("kind") == "link" and token.get("href"):
                document_names.append(token.get("text", ""))
                document_urls.append(token.get("href", ""))

        block_text = " ".join(t.get("text", "") for t in block if t.get("text"))
        block_text = " ".join(block_text.split())
        date_published, closing_date, closing_datetime, date_matches = _extract_dates(block_text)
        briefing = _extract_briefing_details(block_text, date_matches)

        cancellation_flag = "cancellation" in _normalise_search_text(block_text)
        general_notice_flag = notice_type == "General Notice" or "general notice" in _normalise_search_text(header)

        # Prefer the displayed tender/RFP number. If DBSA ever posts a notice without a clear number,
        # fall back to a stable hash of the header text.
        key_seed = tender_number if tender_number else hashlib.sha1(header.encode("utf-8")).hexdigest()[:12]
        opportunity_key = "dbsa:" + _normalise_key(key_seed)

        rows.append(
            {
                "source": SOURCE_NAME,
                "opportunity_key": opportunity_key,
                "tender_number": tender_number,
                "notice_type": notice_type,
                "description": description,
                "document_names": "; ".join([x for x in document_names if x]),
                "document_urls": "; ".join([x for x in document_urls if x]),
                "briefing_session_details": briefing,
                "date_published": date_published,
                "closing_date": closing_date,
                "closing_datetime": closing_datetime,
                "source_url": PORTAL_URL,
                "cancellation_flag": cancellation_flag,
                "general_notice_flag": general_notice_flag,
            }
        )

    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=EXPECTED_OUTPUT_COLUMNS[:14])

    # Keep latest row per opportunity key.
    df = df.sort_values("date_published", ascending=False, na_position="last", kind="mergesort")
    df = df.drop_duplicates(subset=["opportunity_key"], keep="first")
    return df


def filter_current_tenders(df: pd.DataFrame) -> pd.DataFrame:
    """Keep only rows where the closing date has not passed."""
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


def filter_cancellations(df: pd.DataFrame) -> pd.DataFrame:
    """Exclude tender cancellation notices from operational outputs by default."""
    if "cancellation_flag" not in df.columns:
        return df
    before = len(df)
    out = df[~df["cancellation_flag"].fillna(False).astype(bool)].copy()
    logging.info("Filtered cancellation notices: %d -> %d rows", before, len(out))
    return out


def add_priority_scoring(df: pd.DataFrame) -> pd.DataFrame:
    """Add DBSA priority source + description match details and a priority weight score."""
    out = df.copy()

    if out.empty:
        out["priority_source_match"] = pd.Series(dtype="bool")
        out["matched_description_keywords"] = pd.Series(dtype="object")
        out["description_keyword_match_count"] = pd.Series(dtype="int64")
        out["weight_score"] = pd.Series(dtype="int64")
        return out

    out["priority_source_match"] = True

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
        out["priority_source_match"].astype(int) * WEIGHT_SOURCE_MATCH
        + out["description_keyword_match_count"].astype(int) * WEIGHT_DESCRIPTION_KEYWORD
    )

    return sort_for_review(out)


def sort_for_review(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    sort_cols = [c for c in ["weight_score", "closing_datetime", "closing_date", "date_published"] if c in out.columns]
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
    if df is None or df.empty:
        return pd.Series([], index=df.index if df is not None else None, dtype="object")

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


def write_excel_with_table(df: pd.DataFrame, path: str, sheet_name: str, table_name: str):
    """Write XLSX and create an Excel Table when rows exist. Atomic replace."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp.xlsx"

    out_df = df.copy() if df is not None else pd.DataFrame()
    if out_df.shape[1] == 0:
        out_df = pd.DataFrame(columns=EXPECTED_OUTPUT_COLUMNS)

    with pd.ExcelWriter(tmp, engine="xlsxwriter", datetime_format="yyyy-mm-dd", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
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
    """Append incoming rows to master and dedupe by opportunity_key. Incoming wins ties."""
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
    out = out[out["opportunity_key"].ne("")]

    for c in ["date_published", "closing_date", "closing_datetime", "run_timestamp"]:
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce", dayfirst=True)

    sort_cols = [c for c in ["date_published", "_source_priority"] if c in out.columns]
    if sort_cols:
        ascending = [False] * len(sort_cols)
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last", kind="mergesort")

    out = out.drop_duplicates(subset=["opportunity_key"], keep="first")
    out = out.drop(columns=["_source_priority"], errors="ignore")
    return sort_for_review(out)


def main() -> int:
    parser = argparse.ArgumentParser(description="Fetch and clean DBSA procurement notices")
    parser.add_argument("--include-expired", action="store_true", help="do not filter out closed tenders")
    parser.add_argument("--include-cancellations", action="store_true", help="keep tender cancellation notices in clean outputs")
    parser.add_argument("--state-file", default=STATE_FILE_DEFAULT, help="path to JSON file storing seen opportunity keys")
    parser.add_argument("--hash-version-file", default=HASH_VERSION_FILE_DEFAULT, help="path to text file storing the DBSA hash version")
    parser.add_argument("--top-delta-limit", type=int, default=TOP_DELTA_LIMIT_DEFAULT,
                        help="maximum rows to write to dbsa_top_delta.xlsx. Use 0 for all qualifying rows.")
    parser.add_argument("--top-delta-min-score", type=int, default=TOP_DELTA_MIN_SCORE_DEFAULT,
                        help="minimum weight_score required for dbsa_top_delta.xlsx.")
    args = parser.parse_args()

    setup_logging()
    for directory in [RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR]:
        os.makedirs(directory, exist_ok=True)

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
        os.makedirs(RAW_DIR, exist_ok=True)
        with open(RAW_HTML, "w", encoding="utf-8") as fh:
            fh.write(html)
        logging.info("Saved raw HTML: %s", RAW_HTML)

        tokens = tokenise_page(html)
        blocks = extract_notice_blocks(tokens)
        logging.info("Tokenised DBSA page: %d tokens | notice blocks: %d", len(tokens), len(blocks))
        raw_df = parse_notice_blocks(blocks)
        logging.info("Parsed DBSA notice rows: %d | cols: %d", len(raw_df), len(raw_df.columns))

        raw_df.to_csv(RAW_CSV, index=False)
        logging.info("Saved raw CSV: %s", RAW_CSV)

        with pd.ExcelWriter(RAW_XLSX, engine="xlsxwriter", datetime_format="yyyy-mm-dd", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
            raw_df.to_excel(writer, index=False, sheet_name="Raw_Data")
        logging.info("Saved raw Excel: %s", RAW_XLSX)

        mapped = raw_df.copy()
        if not args.include_cancellations:
            mapped = filter_cancellations(mapped)
        if not args.include_expired:
            mapped = filter_current_tenders(mapped)

        mapped = add_priority_scoring(mapped)
        mapped["row_hash"] = compute_row_hash(mapped, HASH_COLS)
        mapped["run_timestamp"] = run_ts

        master_df = load_master(MASTER_XLSX)
        if not master_df.empty and "opportunity_key" in master_df.columns:
            master_df = master_df.copy()
            master_df["opportunity_key"] = master_df["opportunity_key"].fillna("").astype(str).str.strip().str.casefold()
            master_df = master_df[master_df["opportunity_key"].ne("")]
            for c in ["row_hash", "weight_score", "priority_source_match", "matched_description_keywords", "description_keyword_match_count"]:
                if c not in master_df.columns:
                    if c == "row_hash":
                        master_df[c] = compute_row_hash(master_df, HASH_COLS)
                    elif c == "priority_source_match":
                        master_df[c] = True
                    elif c == "matched_description_keywords":
                        master_df[c] = ""
                    else:
                        master_df[c] = 0

        # Migration/state baseline:
        # If this source has an existing master but no seen_ids.json yet, or if
        # the hash version changed, do not flood Delta/Updates on the first run.
        # The master will still be refreshed and the new state files will be saved.
        seen_ids_file_exists = os.path.exists(args.state_file)
        previous_hash_version = _load_hash_version(args.hash_version_file)
        suppress_alerts_for_baseline = (
            not master_df.empty
            and (
                not seen_ids_file_exists
                or previous_hash_version != HASH_VERSION
            )
        )

        if suppress_alerts_for_baseline:
            logging.info(
                "DBSA hash/state version changed or was not recorded. "
                "Re-baselining and suppressing delta/update alerts for this run."
            )
            delta_df = mapped.iloc[0:0].copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""
            updates_df = mapped.iloc[0:0].copy()
            updates_df["change_type"] = "UPDATED"
            updates_df["previous_row_hash"] = ""
        elif not master_df.empty and "opportunity_key" in master_df.columns:
            master_ids = set(master_df["opportunity_key"].tolist())
            master_hash_map = dict(
                zip(
                    master_df["opportunity_key"].tolist(),
                    master_df["row_hash"].fillna("").astype(str).tolist(),
                )
            )

            is_new = ~mapped["opportunity_key"].isin(master_ids)
            delta_df = mapped[is_new].copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""

            exists = ~is_new
            mapped_prev_hash = mapped["opportunity_key"].map(master_hash_map).fillna("")
            is_updated = exists & (mapped["row_hash"] != mapped_prev_hash)
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
        _save_hash_version(args.hash_version_file, HASH_VERSION)

        return 0

    except (requests.RequestException, ValueError, RuntimeError, OSError) as e:
        logging.exception("FAILED: %s", e)
        return 1
    finally:
        release_lock(LOCK_FILE)


if __name__ == "__main__":
    raise SystemExit(main())
