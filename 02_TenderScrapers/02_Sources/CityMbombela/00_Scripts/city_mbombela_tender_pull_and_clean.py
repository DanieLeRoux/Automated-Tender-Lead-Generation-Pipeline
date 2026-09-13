
# city_mbombela_tender_pull_and_clean.py
# Robust City of Mbombela tender scraper
# Hotfix v25: preserve listing rows when detail pages or later pagination pages return HTTP 500.
# This scraper intentionally keeps rows with missing closing dates in the source master so that
# the central consolidation layer can resolve dates from eTenders duplicates or place unresolved
# relevant rows into Needs Manual Date Check.

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
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

SOURCE = "City of Mbombela"
SOURCE_SLUG = "city_mbombela"
SOURCE_URL = "https://www.mbombela.gov.za/downloads/tenders/"
HASH_VERSION = "city_mbombela_hash_v2_stable_20260704"

# New TenderScrapers structure:
# 02_TenderScrapers/02_Sources/CityMbombela/00_Scripts/city_mbombela_tender_pull_and_clean.py
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_SOURCE_ROOT = SCRIPT_DIR.parent if SCRIPT_DIR.name == "00_Scripts" else Path(
    r"C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\02_TenderScrapers\02_Sources\CityMbombela"
)

RAW_DIR_NAME = "01_RawData"
OUTPUT_DIR_NAME = "02_Outputs"
LOG_DIR_NAME = "03_Logs"
STATE_DIR_NAME = "04_State"
METADATA_DIR_NAME = "05_Metadata"

REQUEST_TIMEOUT = 35
REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

OUTPUT_COLUMNS = [
    "source",
    "opportunity_key",
    "tender_number",
    "reference_number",
    "title",
    "description",
    "client",
    "sector",
    "country_name",
    "province",
    "region",
    "date_published",
    "issue_date",
    "closing_date",
    "closing_datetime",
    "notice_type",
    "raw_status",
    "raw_category",
    "document_names",
    "document_urls",
    "document_count",
    "source_url",
    "action_url",
    "contact_person",
    "contact_email",
    "contact_phone",
    "briefing_session_details",
    "submission_method",
    "conditions",
    "run_timestamp",
    "row_hash",
]

DOC_EXT_RE = re.compile(r"\.(pdf|doc|docx|xls|xlsx|zip|rar)(?:$|\?)", re.I)
REFERENCE_RE = re.compile(
    r"\b(?:COM|EDP|PS|PW|ES|CORP|FIN|CSS|RFP|RFQ|SCM|BID)[\s\-_/]*\d{1,4}[\s\-_/]*\/[\s\-_/]*\d{4}\b",
    re.I,
)
EMAIL_RE = re.compile(r"[\w.\-+%]+@[\w.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:\+27|0)[\d\s()./-]{8,}")
MONTHS = (
    "January|February|March|April|May|June|July|August|September|October|November|December|"
    "Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec"
)
DATE_TEXT_RE = re.compile(
    rf"\b\d{{1,2}}\s+(?:{MONTHS})\s+\d{{4}}(?:\s*(?:at|@)?\s*\d{{1,2}}:\d{{2}}(?::\d{{2}})?)?\b",
    re.I,
)
DATE_NUMERIC_RE = re.compile(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?\b")


def setup_logging(log_dir: Path) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{SOURCE_SLUG}_run.log"
    handlers = [logging.StreamHandler(sys.stdout), logging.FileHandler(log_file, encoding="utf-8")]
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=handlers,
        force=True,
    )


def ensure_dirs(base_root: Path) -> Tuple[Path, Path, Path, Path, Path]:
    raw_dir = base_root / RAW_DIR_NAME
    output_dir = base_root / OUTPUT_DIR_NAME
    log_dir = base_root / LOG_DIR_NAME
    state_dir = base_root / STATE_DIR_NAME
    metadata_dir = base_root / METADATA_DIR_NAME
    for p in (raw_dir, output_dir, log_dir, state_dir, metadata_dir):
        p.mkdir(parents=True, exist_ok=True)
    return raw_dir, output_dir, log_dir, state_dir, metadata_dir


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


def load_seen_ids(state_dir: Path) -> Tuple[set[str], bool]:
    """Load source-level seen IDs. Returns (ids, file_exists_and_loaded)."""
    path = state_dir / "seen_ids.json"
    if not path.exists():
        return set(), False
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        ids = {str(x).strip().casefold() for x in (data or []) if str(x).strip()}
        return ids, True
    except Exception as exc:  # noqa: BLE001
        logging.warning("Could not read seen_ids.json; falling back to previous master only: %s", exc)
        return set(), False


def save_seen_ids(state_dir: Path, ids: Iterable[str]) -> None:
    """Save source-level seen IDs atomically."""
    state_dir.mkdir(parents=True, exist_ok=True)
    keys = sorted(str(x).strip().casefold() for x in ids if str(x).strip())
    tmp = state_dir / "seen_ids.json.tmp"
    final = state_dir / "seen_ids.json"
    tmp.write_text(json.dumps(keys, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(final)


def load_hash_version(state_dir: Path) -> str:
    path = state_dir / "hash_version.txt"
    if not path.exists():
        return ""
    try:
        return path.read_text(encoding="utf-8").strip()
    except Exception:  # noqa: BLE001
        return ""


def save_hash_version(state_dir: Path) -> None:
    path = state_dir / "hash_version.txt"
    tmp = state_dir / "hash_version.txt.tmp"
    tmp.write_text(HASH_VERSION, encoding="utf-8")
    tmp.replace(path)


def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text


def normalize_reference(text: str) -> str:
    text = clean_text(text)
    match = REFERENCE_RE.search(text)
    if not match:
        # Fallback for common Mbombela pattern without a preceding letter block.
        match = re.search(r"\b[A-Z]{2,6}\s*\d{1,4}\s*/\s*\d{4}\b", text, re.I)
    if not match:
        return ""
    ref = match.group(0)
    ref = re.sub(r"\s+", "", ref)
    ref = ref.replace("_", "/")
    ref = re.sub(r"-", "/", ref)
    ref = re.sub(r"/+", "/", ref)
    return ref.upper()


def make_key(reference: str, title: str, url: str) -> str:
    base = reference or title or url
    base = re.sub(r"\s+", " ", base).strip().lower()
    digest = hashlib.sha1(base.encode("utf-8", errors="ignore")).hexdigest()[:12]
    if reference:
        safe = re.sub(r"[^a-z0-9]+", "-", reference.lower()).strip("-")
        return f"{SOURCE_SLUG}:{safe}"
    return f"{SOURCE_SLUG}:{digest}"


def fetch_html(url: str, *, allow_failure: bool = False) -> Optional[str]:
    try:
        response = requests.get(url, headers=REQUEST_HEADERS, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        response.encoding = response.apparent_encoding or response.encoding
        return response.text
    except Exception as exc:
        if allow_failure:
            logging.warning("Failed to fetch Mbombela page %s: %s", url, exc)
            return None
        raise


def extract_card_text(anchor) -> str:
    # Move upward to a likely listing card, but avoid swallowing the entire page.
    parent = anchor
    best_text = clean_text(anchor.get_text(" ", strip=True))
    for _ in range(7):
        if parent is None:
            break
        text = clean_text(parent.get_text(" ", strip=True))
        if len(text) > len(best_text) and len(text) < 2500:
            best_text = text
        classes = " ".join(parent.get("class", [])).lower() if hasattr(parent, "get") else ""
        tag = getattr(parent, "name", "")
        if tag in {"article", "li"} or any(k in classes for k in ["download", "post", "entry", "item", "card"]):
            break
        parent = parent.parent
    return best_text


def extract_listing_entries(html: str, page_url: str) -> List[Dict[str, str]]:
    soup = BeautifulSoup(html, "html.parser")
    entries: List[Dict[str, str]] = []
    seen = set()

    for a in soup.find_all("a", href=True):
        href = urljoin(page_url, a["href"])
        parsed = urlparse(href)
        if "mbombela.gov.za" not in parsed.netloc:
            continue
        if "/download/" not in parsed.path:
            continue
        if "/downloads/" in parsed.path:
            continue

        title = clean_text(a.get_text(" ", strip=True))
        card_text = extract_card_text(a)
        combined = clean_text(f"{title} {card_text}")

        # The site has many irrelevant navigation links; require a likely tender reference or enough tender wording.
        if not normalize_reference(combined) and not re.search(r"\btender\b|\bappointment\b|\bpanel\b|\bservice provider\b", combined, re.I):
            continue

        if href in seen:
            continue
        seen.add(href)

        reference = normalize_reference(combined)
        if not title:
            title = combined[:180]

        entries.append({
            "listing_title": title,
            "listing_text": combined,
            "source_url": href,
            "reference_number": reference,
        })

    return entries


def extract_documents(soup: BeautifulSoup, page_url: str, detail_url: str) -> Tuple[str, str, int]:
    """Extract only actual tender document/download links from the detail page.

    Mbombela pages contain many repeated navigation links and links to other tender
    detail pages. Keeping those in document_urls creates very long Excel cells and
    unstable row hashes. This function keeps file/download links only.
    """
    docs: List[Tuple[str, str]] = []
    seen = set()
    detail_norm = detail_url.rstrip("/").lower()

    for a in soup.find_all("a", href=True):
        raw_href = clean_text(a.get("href", ""))
        if not raw_href or raw_href.startswith("#") or raw_href.lower().startswith(("javascript:", "mailto:", "tel:")):
            continue

        href = urljoin(page_url, raw_href)
        parsed = urlparse(href)
        lower_href = href.lower()
        path_lower = parsed.path.lower()
        query_lower = parsed.query.lower()
        label = clean_text(a.get_text(" ", strip=True)) or os.path.basename(parsed.path) or href

        if "mbombela.gov.za" not in parsed.netloc.lower():
            continue
        if href.rstrip("/").lower() == detail_norm:
            continue
        if path_lower.rstrip("/").endswith("/download") or "/downloads/" in path_lower:
            continue
        if "standing-orders" in lower_href or "wp-content/uploads/2024/11/standing-orders" in lower_href:
            continue

        has_file_extension = bool(DOC_EXT_RE.search(path_lower))
        has_download_token = "wpdmdl=" in query_lower or "filename=" in query_lower
        if not has_file_extension and not has_download_token:
            continue

        if href not in seen:
            docs.append((label, href))
            seen.add(href)

    if not docs and detail_url:
        docs.append(("Tender detail page", detail_url))

    names = "; ".join(name for name, _ in docs)
    urls = "; ".join(url for _, url in docs)
    return names, urls, len(docs)


def parse_date_value(text: str) -> Tuple[str, str]:
    text = clean_text(text)
    if not text:
        return "", ""

    candidates: List[str] = []

    # Prefer dates close to closing wording.
    for m in re.finditer(r"(closing|closing date|closing time|closing date and time|deadline|bid closes).{0,160}", text, re.I):
        snippet = m.group(0)
        candidates += DATE_TEXT_RE.findall(snippet)
        candidates += DATE_NUMERIC_RE.findall(snippet)

    # Regex .findall with month alternation may return month groups, so use finditer instead.
    fixed_candidates: List[str] = []
    for m in re.finditer(r"(closing|closing date|closing time|closing date and time|deadline|bid closes).{0,180}", text, re.I):
        snippet = m.group(0)
        fixed_candidates += [x.group(0) for x in DATE_TEXT_RE.finditer(snippet)]
        fixed_candidates += [x.group(0) for x in DATE_NUMERIC_RE.finditer(snippet)]

    if fixed_candidates:
        candidates = fixed_candidates
    else:
        candidates = [x.group(0) for x in DATE_TEXT_RE.finditer(text)]
        candidates += [x.group(0) for x in DATE_NUMERIC_RE.finditer(text)]

    # De-duplicate while preserving order.
    seen = set()
    ordered = []
    for c in candidates:
        c = clean_text(c)
        if c and c.lower() not in seen:
            ordered.append(c)
            seen.add(c.lower())

    now = pd.Timestamp.now().normalize()
    parsed_values = []
    for candidate in ordered:
        normalised = re.sub(r"\bat\b", " ", candidate, flags=re.I)
        normalised = normalised.replace("@", " ")
        dt = pd.to_datetime(normalised, errors="coerce", dayfirst=True)
        if pd.notna(dt):
            parsed_values.append(pd.Timestamp(dt))

    if not parsed_values:
        return "", ""

    # Prefer current/future dates, otherwise use the first parsed date.
    future = [d for d in parsed_values if d >= now]
    selected = min(future) if future else parsed_values[0]
    return selected.date().isoformat(), selected.strftime("%Y-%m-%d %H:%M:%S")


def extract_detail(entry: Dict[str, str]) -> Dict[str, str]:
    url = entry.get("source_url", "")
    html = fetch_html(url, allow_failure=True)
    if not html:
        return {}

    soup = BeautifulSoup(html, "html.parser")
    detail_text = clean_text(soup.get_text(" ", strip=True))
    closing_date, closing_datetime = parse_date_value(detail_text)
    doc_names, doc_urls, doc_count = extract_documents(soup, url, url)
    email_match = EMAIL_RE.search(detail_text)
    phone_match = PHONE_RE.search(detail_text)

    # Prefer a page heading as title if available.
    heading = ""
    for tag in ["h1", "h2", "h3"]:
        h = soup.find(tag)
        if h:
            heading = clean_text(h.get_text(" ", strip=True))
            if heading:
                break

    reference = normalize_reference(detail_text) or entry.get("reference_number", "")

    return {
        "detail_title": heading,
        "detail_text": detail_text,
        "reference_number": reference,
        "closing_date": closing_date,
        "closing_datetime": closing_datetime,
        "document_names": doc_names,
        "document_urls": doc_urls,
        "document_count": doc_count,
        "contact_email": email_match.group(0) if email_match else "",
        "contact_phone": phone_match.group(0) if phone_match else "",
    }


def fetch_and_parse(base_root: Path, raw_dir: Path, max_pages: int = 4, detail_limit: int = 0) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    run_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for page in range(1, max_pages + 1):
        url = SOURCE_URL if page == 1 else urljoin(SOURCE_URL, f"page/{page}/")
        logging.info("Fetching City of Mbombela listing page %s: %s", page, url)
        html = fetch_html(url, allow_failure=(page > 1))
        if not html:
            logging.warning("Stopping Mbombela pagination at page %s; retaining rows already parsed.", page)
            break

        if page == 1:
            (raw_dir / f"{SOURCE_SLUG}_raw_page.html").write_text(html, encoding="utf-8")

        entries = extract_listing_entries(html, url)
        logging.info("Found Mbombela listing entries on page %s: %s", page, len(entries))
        if not entries:
            if page > 1:
                break
            continue

        for entry in entries:
            if detail_limit and len(rows) >= detail_limit:
                break

            detail = extract_detail(entry)
            combined_text = clean_text(" ".join([
                entry.get("listing_title", ""),
                entry.get("listing_text", ""),
                detail.get("detail_title", ""),
                detail.get("detail_text", "")[:1500],
            ]))

            reference = detail.get("reference_number") or entry.get("reference_number") or normalize_reference(combined_text)
            title = detail.get("detail_title") or entry.get("listing_title") or reference
            description = combined_text or title
            closing_date = detail.get("closing_date", "")
            closing_datetime = detail.get("closing_datetime", "")
            document_names = detail.get("document_names", "")
            document_urls = detail.get("document_urls", "")
            document_count = detail.get("document_count", 0)
            source_url = entry.get("source_url", "")
            action_url = ""
            if document_urls:
                action_url = document_urls.split(";")[0].strip()
            else:
                action_url = source_url

            row = {
                "source": SOURCE,
                "opportunity_key": make_key(reference, title, source_url),
                "tender_number": reference,
                "reference_number": reference,
                "title": title,
                "description": description,
                "client": "City of Mbombela",
                "sector": "",
                "country_name": "South Africa",
                "province": "Mpumalanga",
                "region": "Mbombela",
                "date_published": "",
                "issue_date": "",
                "closing_date": closing_date,
                "closing_datetime": closing_datetime,
                "notice_type": "Tender",
                "raw_status": "Open / listed",
                "raw_category": "Tender",
                "document_names": document_names,
                "document_urls": document_urls,
                "document_count": document_count,
                "source_url": source_url,
                "action_url": action_url,
                "contact_person": "",
                "contact_email": detail.get("contact_email", ""),
                "contact_phone": detail.get("contact_phone", ""),
                "briefing_session_details": "",
                "submission_method": "",
                "conditions": "",
                "run_timestamp": run_ts,
            }
            rows.append(row)

        if detail_limit and len(rows) >= detail_limit:
            break

    # De-duplicate by source_url/opportunity key.
    deduped: Dict[str, Dict[str, object]] = {}
    for row in rows:
        key = str(row.get("opportunity_key") or row.get("source_url") or row.get("description"))
        if key not in deduped:
            deduped[key] = row
    return list(deduped.values())


def compute_row_hash(row: pd.Series) -> str:
    # Keep the hash stable. Mbombela download URLs often contain volatile query
    # parameters, so document_urls is intentionally excluded from change detection.
    fields = [
        "tender_number",
        "reference_number",
        "title",
        "description",
        "closing_datetime",
        "source_url",
        "document_count",
    ]
    payload = "|".join(clean_text(row.get(f, "")) for f in fields)
    return hashlib.sha256(payload.encode("utf-8", errors="ignore")).hexdigest()


def build_dataframe(rows: List[Dict[str, object]]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    if not df.empty:
        df["row_hash"] = df.apply(compute_row_hash, axis=1)
    return df[OUTPUT_COLUMNS]


def filter_master(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    out = df.copy()
    dt = pd.to_datetime(out["closing_datetime"], errors="coerce")
    now = pd.Timestamp.now()
    # Keep current rows and rows with missing dates so central consolidation can resolve eTenders duplicates
    # or route unresolved relevant rows to manual date review.
    keep = dt.isna() | (dt >= now)
    out = out[keep].copy()
    return out


def write_excel_plain(df: pd.DataFrame, path: Path, sheet_name: str) -> None:
    """Write Excel without auto-converting long URL strings into hyperlinks."""
    with pd.ExcelWriter(
        path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)


def write_outputs(
    df_raw: pd.DataFrame,
    df_master: pd.DataFrame,
    raw_dir: Path,
    output_dir: Path,
    state_dir: Path,
    *,
    preserve_master_on_failure: bool = False,
) -> None:
    raw_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    state_dir.mkdir(parents=True, exist_ok=True)

    raw_csv = raw_dir / f"{SOURCE_SLUG}_raw.csv"
    raw_xlsx = raw_dir / f"{SOURCE_SLUG}_raw_full.xlsx"
    master_xlsx = output_dir / f"{SOURCE_SLUG}_master.xlsx"
    delta_xlsx = output_dir / f"{SOURCE_SLUG}_delta.xlsx"
    updates_xlsx = output_dir / f"{SOURCE_SLUG}_updates.xlsx"
    top_delta_xlsx = output_dir / f"{SOURCE_SLUG}_top_delta.xlsx"

    previous = pd.DataFrame()
    if master_xlsx.exists():
        try:
            previous = pd.read_excel(master_xlsx, dtype=str)
        except Exception as exc:  # noqa: BLE001
            logging.warning("Could not read previous Mbombela master: %s", exc)
            previous = pd.DataFrame()

    df_raw.to_csv(raw_csv, index=False, encoding="utf-8-sig")
    write_excel_plain(df_raw, raw_xlsx, "Raw")

    if preserve_master_on_failure and not previous.empty:
        logging.warning("Fetch failed. Preserving previous master file and writing empty delta/update outputs.")
        df_master = previous.copy()
        delta = pd.DataFrame(columns=df_master.columns)
        updates = pd.DataFrame(columns=df_master.columns)
        top_delta = pd.DataFrame(columns=df_master.columns)
        write_excel_plain(delta, delta_xlsx, "Delta")
        write_excel_plain(updates, updates_xlsx, "Updates")
        write_excel_plain(top_delta, top_delta_xlsx, "Top Delta")
        logging.info("Preserved master Excel: %s | rows: %s", master_xlsx, len(df_master))
        return

    previous_keys: set[str] = set()
    previous_hash: Dict[str, str] = {}
    if not previous.empty and "opportunity_key" in previous.columns:
        previous_keys = {str(x).strip().casefold() for x in previous["opportunity_key"].dropna() if str(x).strip()}
        if "row_hash" in previous.columns:
            previous_hash = dict(
                zip(
                    previous["opportunity_key"].fillna("").astype(str).str.strip().str.casefold(),
                    previous["row_hash"].fillna("").astype(str),
                )
            )

    seen_ids, seen_ids_loaded = load_seen_ids(state_dir)
    previous_hash_version = load_hash_version(state_dir)

    baseline_mode = False
    if previous_keys and not seen_ids_loaded:
        baseline_mode = True
        seen_ids = set(previous_keys)
        logging.info(
            "No seen_ids.json found. Bootstrapping Mbombela state from existing master and suppressing alerts for this run."
        )
    elif previous_keys and previous_hash_version != HASH_VERSION:
        baseline_mode = True
        logging.info(
            "Mbombela hash version changed or was not recorded. Re-baselining master hashes and suppressing update alerts for this run."
        )

    if df_master.empty:
        delta = df_master.copy()
        updates = df_master.copy()
        top_delta = df_master.copy()
        current_keys: set[str] = set()
    else:
        current_keys = {
            str(x).strip().casefold() for x in df_master["opportunity_key"].dropna() if str(x).strip()
        }
        keys = df_master["opportunity_key"].fillna("").astype(str).str.strip().str.casefold()

        if baseline_mode:
            delta = pd.DataFrame(columns=df_master.columns)
            updates = pd.DataFrame(columns=df_master.columns)
        else:
            delta = df_master[~keys.isin(seen_ids)].copy()

            changed_mask = []
            for _, row in df_master.iterrows():
                key = str(row.get("opportunity_key", "")).strip().casefold()
                old_hash = previous_hash.get(key, "")
                new_hash = str(row.get("row_hash", ""))
                changed_mask.append(bool(key in seen_ids and old_hash and old_hash != new_hash))
            updates = df_master[pd.Series(changed_mask, index=df_master.index)].copy()

        if not delta.empty:
            delta.loc[:, "change_type"] = "NEW"
        if not updates.empty:
            updates.loc[:, "change_type"] = "UPDATED"
        top_delta = delta.head(20).copy()

    write_excel_plain(df_master, master_xlsx, "Master")
    write_excel_plain(delta, delta_xlsx, "Delta")
    write_excel_plain(updates, updates_xlsx, "Updates")
    write_excel_plain(top_delta, top_delta_xlsx, "Top Delta")

    save_seen_ids(state_dir, seen_ids.union(current_keys))
    save_hash_version(state_dir)

    logging.info("Saved raw outputs | rows: %s", len(df_raw))
    logging.info("Saved master Excel: %s | rows: %s", master_xlsx, len(df_master))
    logging.info("Saved delta Excel: %s | rows: %s", delta_xlsx, len(delta))
    logging.info("Saved updates Excel: %s | rows: %s", updates_xlsx, len(updates))
    logging.info("Saved top_delta Excel: %s | rows: %s", top_delta_xlsx, len(top_delta))


def main() -> int:
    parser = argparse.ArgumentParser(description="Scrape City of Mbombela tenders")
    parser.add_argument("--base-root", default=str(DEFAULT_SOURCE_ROOT))
    parser.add_argument("--max-pages", type=int, default=4)
    parser.add_argument("--detail-limit", type=int, default=0)
    args = parser.parse_args()

    base_root = Path(args.base_root).resolve()
    raw_dir, output_dir, log_dir, state_dir, metadata_dir = ensure_dirs(base_root)
    setup_logging(log_dir)

    lock_file = state_dir / "run.lock"
    fetch_failed = False

    logging.info("Using source root: %s", base_root)
    logging.info("Using raw directory: %s", raw_dir)
    logging.info("Using output directory: %s", output_dir)
    logging.info("Using log directory: %s", log_dir)
    logging.info("Using state directory: %s", state_dir)

    try:
        acquire_lock(lock_file)
        try:
            rows = fetch_and_parse(base_root, raw_dir, max_pages=args.max_pages, detail_limit=args.detail_limit)
        except Exception as exc:
            logging.exception("FAILED before parsed rows could be completed: %s", exc)
            rows = []
            fetch_failed = True

        df_raw = build_dataframe(rows)
        df_master = filter_master(df_raw)
        write_outputs(
            df_raw,
            df_master,
            raw_dir,
            output_dir,
            state_dir,
            preserve_master_on_failure=fetch_failed,
        )
        return 1 if fetch_failed else 0
    finally:
        release_lock(lock_file)


if __name__ == "__main__":
    raise SystemExit(main())
