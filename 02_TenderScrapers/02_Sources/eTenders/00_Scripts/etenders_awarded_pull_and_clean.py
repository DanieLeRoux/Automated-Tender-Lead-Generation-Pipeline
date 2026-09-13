r"""
etenders_awarded_pull_and_clean.py

Fetch awarded tenders from the South African eTenders portal, clean/normalise them,
and maintain operational Excel outputs for Henry Fagan review.

Recommended location:
    C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers\etenders_awarded_pull_and_clean.py

Typical run:
    python etenders_awarded_pull_and_clean.py --since-days 365

Daily run after the initial master is built:
    python etenders_awarded_pull_and_clean.py --since-days 30
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
from typing import Any, Dict, Iterable, List, Optional, Set

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ---------------------------------------------------------------------------
# Source settings
# ---------------------------------------------------------------------------

BASE_PAGE = "https://www.etenders.gov.za/Home/opportunities?id=2"
API_URL = "https://www.etenders.gov.za/Home/PaginatedTenderOpportunities"

PAGE_SIZE = 200
STATUS = 2  # eTenders Awarded tab/page
ORDER_COL = 2
ORDER_DIR = "desc"
SLEEP_S = 0.4
TIMEOUT_S = 60
MAX_PAGES = 300

# ---------------------------------------------------------------------------
# Project folders
# ---------------------------------------------------------------------------

BASE_ROOT = r"C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\TenderScrapers"

LOG_DIR = os.path.join(BASE_ROOT, "Logs", "Awarded")
OUTPUT_DIR = os.path.join(BASE_ROOT, "Outputs", "Awarded")
STATE_DIR = os.path.join(BASE_ROOT, "State", "Awarded")

LOG_PATH = os.path.join(LOG_DIR, "etenders_awarded_run.log")
LOCK_FILE = os.path.join(STATE_DIR, "run.lock")
STATE_FILE_DEFAULT = os.path.join(STATE_DIR, "seen_awarded_ids.json")

RAW_CSV = os.path.join(OUTPUT_DIR, "etenders_awarded_raw.csv")
RAW_XLSX = os.path.join(OUTPUT_DIR, "etenders_awarded_raw_full.xlsx")
MASTER_XLSX = os.path.join(OUTPUT_DIR, "etenders_awarded_master.xlsx")
DELTA_XLSX = os.path.join(OUTPUT_DIR, "etenders_awarded_delta.xlsx")
UPDATES_XLSX = os.path.join(OUTPUT_DIR, "etenders_awarded_updates.xlsx")
TOP_DELTA_XLSX = os.path.join(OUTPUT_DIR, "etenders_awarded_top_delta.xlsx")
HF_MATCHES_XLSX = os.path.join(OUTPUT_DIR, "etenders_awarded_hf_matches.xlsx")
HF_DELTA_XLSX = os.path.join(OUTPUT_DIR, "etenders_awarded_hf_delta.xlsx")

# ---------------------------------------------------------------------------
# Matching and scoring
# ---------------------------------------------------------------------------

HF_MATCH_TERMS = [
    "Henry Fagan",
    "Henry Fagan Consulting Engineers",
    "Henry Fagan Consulting Engineers and Project Managers",
    "Henry Fagan & Partners",
    "Fagan",
]

PRIORITY_CLIENTS = [
    "Umgeni Water",
    "uMngeni-uThukela Water",
    "Johannesburg Water",
    "Vaal Central Water",
    "Department of Public Works and Roads",
    "Development Bank of Southern Africa",
    "Transnet SOC Ltd",
    "Transnet National Ports Authority",
    "Airports Company South Africa",
    "Airports Company of South Africa",
    "South African National Roads Agency SOC Limited",
    "SANRAL",
    "Health",
    "South African Nuclear Energy Corporation Limited",
    "National Nuclear Regulator",
    "Polokwane Local Municipality",
]

DESCRIPTION_KEYWORD_RULES = {
    "Water and Sanitation": ["Water and Sanitation"],
    "Housing": ["Housing"],
    "Education": ["Education"],
    "Transport": ["Transport"],
    "Nuclear": ["Nuclear"],
    "BRT": ["Bus Rapid Transit", "BRT"],
    "Design": ["Design"],
    "Refurbishment": ["Refurbishment"],
    "EPC": ["Engineering, Procurement, and Construction", "Engineering Procurement and Construction", "EPC"],
    "Fire Engineering": ["Fire Engineering"],
    "HVAC": ["HVAC"],
    "Mechanical": ["Mechanical"],
    "Electrical": ["Electrical"],
    "Electronic": ["Electronic"],
    "Wet Services": ["Wet Services"],
    "Project Management": ["Project Management"],
    "Programme Management": ["Programme Management"],
    "PMO": ["Project Management Office", "PMO"],
    "Advisory Services": ["Advisory Services"],
    "Structural Engineering": ["Structural Engineering"],
    "Civil Engineering": ["Civil Engineering"],
    "Feasibility Study": ["Feasibility Study"],
}

WEIGHT_HF_MATCH = 100
WEIGHT_PRIORITY_CLIENT = 10
WEIGHT_DESCRIPTION_KEYWORD = 1
TOP_DELTA_LIMIT_DEFAULT = 50
TOP_DELTA_MIN_SCORE_DEFAULT = 0

OUTPUT_COLUMNS = [
    "opportunity_key",
    "bid_number",
    "description",
    "category",
    "client",
    "province",
    "status",
    "type",
    "date_published",
    "closing_date",
    "award_date",
    "supplier_name",
    "supplier_maaa_number",
    "award_amount",
    "award_currency",
    "conditions",
    "contact_person",
    "contact_email",
    "contact_telephone",
    "tender_validity_period",
    "hf_match",
    "matched_hf_terms",
    "priority_client_match",
    "matched_description_keywords",
    "description_keyword_match_count",
    "weight_score",
    "source_url",
    "row_hash",
    "run_timestamp",
    "change_type",
    "previous_row_hash",
]

HASH_COLS = [
    "bid_number",
    "description",
    "category",
    "client",
    "province",
    "status",
    "award_date",
    "supplier_name",
    "supplier_maaa_number",
    "award_amount",
    "award_currency",
]

COLUMN_ALIASES = {
    "bid_number": ["tender_No", "tenderNo", "tender_Number", "tenderNumber", "bid_number"],
    "description": ["description", "tenderDescription", "tender_Description"],
    "category": ["categories_name", "category", "categoryName", "category_Name"],
    "client": ["organ_of_State", "organOfState", "institution", "department", "buyer", "procuringEntity"],
    "province": ["provinces_name", "province", "provinceName", "province_Name"],
    "status": ["status", "tenderStatus", "tender_Status"],
    "type": ["type", "tenderType", "tender_Type"],
    "date_published": ["date_Published", "datePublished", "publishedDate", "publicationDate"],
    "closing_date": ["closing_Date", "closingDate", "closing_Date_Time", "closingDateTime"],
    "award_date": ["award_Date", "awardDate", "awarded_Date", "awardedDate", "date_Awarded", "dateAwarded"],
    "supplier_name": [
        "supplierName",
        "supplier_Name",
        "awardedTo",
        "awarded_To",
        "awardee",
        "successfulBidder",
        "successful_Bidder",
        "winningBidder",
        "winning_Bidder",
        "bidderName",
        "bidder_Name",
    ],
    "supplier_maaa_number": [
        "maaaNumber",
        "maaa_Number",
        "supplierMAAANumber",
        "supplier_MAAA_Number",
        "supplierMaaaNumber",
    ],
    "award_amount": [
        "awardAmount",
        "award_Amount",
        "awardedAmount",
        "awarded_Amount",
        "contractValue",
        "contract_Value",
        "totalAmount",
        "bidAmount",
        "amount",
    ],
    "award_currency": ["currency", "awardCurrency", "award_Currency"],
    "conditions": ["conditions"],
    "contact_person": ["contactPerson", "contact_Person"],
    "contact_email": ["email", "contactEmail", "contact_Email"],
    "contact_telephone": ["telephone", "contactTelephone", "contact_Telephone"],
    "tender_validity_period": ["validity", "tenderValidity", "tender_Validity"],
}

SUBSTRING_FALLBACKS = {
    "award_date": [["award", "date"], ["awarded", "date"]],
    "supplier_name": [["supplier", "name"], ["bidder", "name"], ["awardee"]],
    "supplier_maaa_number": [["maaa"]],
    "award_amount": [["award", "amount"], ["awarded", "amount"], ["contract", "value"], ["bid", "amount"]],
    "award_currency": [["currency"]],
}


def setup_logging() -> None:
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
    sess = requests.Session()
    sess.headers.update(
        {
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": BASE_PAGE,
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
        raise RuntimeError(f"Lock file exists. Another run may be active: {lock_path}")
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
    return str(value).strip().casefold()


def _normalise_search_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    cleaned = re.sub(r"[^a-z0-9]+", " ", str(value).casefold())
    return " ".join(cleaned.split())


def _contains_phrase(text_norm: str, phrase: str) -> bool:
    phrase_norm = _normalise_search_text(phrase)
    return bool(phrase_norm and f" {phrase_norm} " in f" {text_norm} ")


def _load_seen_ids(path: str) -> Set[str]:
    try:
        if not os.path.exists(path):
            return set()
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return set(_normalise_key(x) for x in (data or []))
    except (OSError, ValueError, TypeError):
        logging.exception("Failed to load state file: %s", path)
        return set()


def _save_seen_ids(path: str, ids: Iterable[str]) -> None:
    try:
        norm = sorted(_normalise_key(x) for x in ids if _normalise_key(x))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(norm, fh, ensure_ascii=False)
        os.replace(tmp, path)
    except (OSError, TypeError):
        logging.exception("Failed to save state file: %s", path)


def parse_any_date(value: Any) -> pd.Timestamp:
    return pd.to_datetime(value, errors="coerce")


def row_best_date(row: Dict[str, Any]) -> pd.Timestamp:
    for key in ["award_Date", "awardDate", "awardedDate", "date_Awarded", "date_Published", "datePublished", "closing_Date"]:
        if key in row:
            dt = parse_any_date(row.get(key))
            if pd.notna(dt):
                return dt
    return pd.NaT


def fetch_awarded_etenders(
    since_date: Optional[pd.Timestamp] = None,
    seen_ids: Optional[Set[str]] = None,
) -> List[Dict[str, Any]]:
    sess = build_session()

    # Warm up cookies/session from the Awarded page.
    r = sess.get(BASE_PAGE, timeout=TIMEOUT_S)
    r.raise_for_status()

    base_params = {
        "draw": 1,
        "length": PAGE_SIZE,
        "start": 0,
        "status": STATUS,
        "order[0][column]": ORDER_COL,
        "order[0][dir]": ORDER_DIR,
        "search[value]": "",
        "search[regex]": "false",
    }

    all_rows: List[Dict[str, Any]] = []
    start = 0
    total = None
    page_count = 0
    stop_paging = False

    while True:
        page_count += 1
        if page_count > MAX_PAGES:
            logging.warning("MAX_PAGES reached; stopping pagination.")
            break

        params = dict(base_params, start=start)
        r = sess.get(API_URL, params=params, timeout=TIMEOUT_S)
        r.raise_for_status()

        try:
            payload = r.json()
        except ValueError:
            payload = json.loads(r.text)

        rows = payload.get("data") or payload.get("aaData") or []
        if total is None:
            total = payload.get("recordsFiltered") or payload.get("recordsTotal")

        if not rows:
            break

        for row in rows:
            bid = row.get("tender_No") or row.get("tenderNo") or row.get("tenderNumber") or ""
            key = _normalise_key(bid)

            if seen_ids and key and key in seen_ids:
                stop_paging = True
                break

            if since_date is not None:
                rd = row_best_date(row)
                if pd.notna(rd) and rd < since_date:
                    stop_paging = True
                    break

            all_rows.append(row)

        start += PAGE_SIZE

        if stop_paging:
            logging.info("Stopping early due to since_date/seen_ids condition at page %d.", page_count)
            break
        if total and start >= int(total):
            break
        if len(rows) < PAGE_SIZE:
            break

        time.sleep(SLEEP_S)

    logging.info("Fetched awarded rows: %d | pages: %d | total: %s", len(all_rows), page_count, total)
    return all_rows


def _norm_col_name(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", value.casefold()).strip("_")


def _find_column(df: pd.DataFrame, aliases: List[str], fallback_tokens: Optional[List[List[str]]] = None) -> Optional[str]:
    exact = {c: c for c in df.columns}
    casefold = {c.casefold(): c for c in df.columns}
    normed = {_norm_col_name(c): c for c in df.columns}

    for alias in aliases:
        if alias in exact:
            return exact[alias]
        if alias.casefold() in casefold:
            return casefold[alias.casefold()]
        alias_norm = _norm_col_name(alias)
        if alias_norm in normed:
            return normed[alias_norm]

    for token_group in fallback_tokens or []:
        tokens = [_norm_col_name(t) for t in token_group]
        for col in df.columns:
            col_norm = _norm_col_name(col)
            if all(token in col_norm for token in tokens):
                return col

    return None


def map_and_clean(df: pd.DataFrame) -> pd.DataFrame:
    mapped = pd.DataFrame(index=df.index)

    for out_col, aliases in COLUMN_ALIASES.items():
        source_col = _find_column(df, aliases, SUBSTRING_FALLBACKS.get(out_col))
        if source_col is not None:
            mapped[out_col] = df[source_col]
        else:
            mapped[out_col] = ""

    for date_col in ["date_published", "closing_date", "award_date"]:
        mapped[date_col] = pd.to_datetime(mapped[date_col], errors="coerce")

    # Best stable identifier: bid number when available. Fallback is a hash of key text.
    mapped["bid_number"] = mapped["bid_number"].fillna("").astype(str).str.strip()

    fallback_text = (
        mapped["description"].fillna("").astype(str)
        + "|"
        + mapped["client"].fillna("").astype(str)
        + "|"
        + mapped["award_date"].astype(str)
        + "|"
        + mapped["supplier_name"].fillna("").astype(str)
    )
    fallback_hash = fallback_text.apply(lambda x: hashlib.sha256(x.encode("utf-8")).hexdigest()[:16])

    mapped["opportunity_key"] = mapped["bid_number"].apply(_normalise_key)
    mapped.loc[mapped["opportunity_key"].eq(""), "opportunity_key"] = "award:" + fallback_hash

    # Deduplicate; newest/most recently awarded row wins.
    sort_cols = ["award_date", "date_published", "closing_date"]
    mapped = mapped.sort_values(sort_cols, ascending=[False, False, False], na_position="last", kind="mergesort")
    mapped = mapped.drop_duplicates(subset=["opportunity_key"], keep="first")

    mapped["source_url"] = BASE_PAGE
    return mapped


def add_matches_and_scoring(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    priority_clients_norm = {_normalise_search_text(x) for x in PRIORITY_CLIENTS}

    def hf_terms(row: pd.Series) -> str:
        searchable = " ".join(
            str(row.get(col, ""))
            for col in ["supplier_name", "description", "client", "bid_number"]
        )
        text_norm = _normalise_search_text(searchable)
        matches = [term for term in HF_MATCH_TERMS if _contains_phrase(text_norm, term)]
        return "; ".join(matches)

    def keyword_terms(description: Any) -> str:
        text_norm = _normalise_search_text(description)
        matches = []
        for label, phrases in DESCRIPTION_KEYWORD_RULES.items():
            if any(_contains_phrase(text_norm, phrase) for phrase in phrases):
                matches.append(label)
        return "; ".join(matches)

    out["matched_hf_terms"] = out.apply(hf_terms, axis=1)
    out["hf_match"] = out["matched_hf_terms"].fillna("").astype(str).str.strip().ne("")

    out["priority_client_match"] = out["client"].apply(
        lambda x: _normalise_search_text(x) in priority_clients_norm
    )

    out["matched_description_keywords"] = out["description"].apply(keyword_terms)
    out["description_keyword_match_count"] = out["matched_description_keywords"].apply(
        lambda x: 0 if not x else len(str(x).split("; "))
    )

    out["weight_score"] = (
        out["hf_match"].astype(int) * WEIGHT_HF_MATCH
        + out["priority_client_match"].astype(int) * WEIGHT_PRIORITY_CLIENT
        + out["description_keyword_match_count"].astype(int) * WEIGHT_DESCRIPTION_KEYWORD
    )

    return sort_operational(out)


def compute_row_hash(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    use_cols = [c for c in cols if c in df.columns]
    if not use_cols:
        return pd.Series([""] * len(df), index=df.index)

    tmp = df[use_cols].copy()
    for col in use_cols:
        if pd.api.types.is_datetime64_any_dtype(tmp[col]):
            tmp[col] = tmp[col].dt.strftime("%Y-%m-%d %H:%M:%S").fillna("")
        else:
            tmp[col] = tmp[col].fillna("").astype(str).str.strip()

    joined = tmp.apply(lambda r: "||".join(r.values.tolist()), axis=1)
    return joined.apply(lambda s: hashlib.sha256(s.encode("utf-8")).hexdigest())


def sort_operational(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    sort_cols = [c for c in ["weight_score", "award_date", "date_published", "closing_date"] if c in out.columns]
    if not sort_cols:
        return out
    ascending = [False, False, False, False][: len(sort_cols)]
    return out.sort_values(sort_cols, ascending=ascending, na_position="last", kind="mergesort")


def ensure_output_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy() if df is not None else pd.DataFrame()
    for col in OUTPUT_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    extras = [c for c in out.columns if c not in OUTPUT_COLUMNS]
    return out[OUTPUT_COLUMNS + extras]


def write_excel_with_table(df: pd.DataFrame, path: str, sheet_name: str, table_name: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp.xlsx"

    out_df = ensure_output_columns(df)

    with pd.ExcelWriter(tmp, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        out_df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        nrows, ncols = out_df.shape
        columns = [{"header": col} for col in out_df.columns]

        if nrows > 0 and ncols > 0:
            worksheet.add_table(0, 0, nrows, ncols - 1, {"name": table_name, "columns": columns})
        elif ncols > 0:
            worksheet.autofilter(0, 0, 0, ncols - 1)

        worksheet.freeze_panes(1, 0)

        for idx, col in enumerate(out_df.columns):
            width = min(max(len(str(col)) + 2, 12), 45)
            worksheet.set_column(idx, idx, width)

    os.replace(tmp, path)


def load_master(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name="Master", engine="openpyxl")
    except Exception:
        logging.exception("Failed to read master file: %s", path)
        return pd.DataFrame()


def append_and_dedupe_master(master: pd.DataFrame, incoming: pd.DataFrame) -> pd.DataFrame:
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

    for col in ["award_date", "date_published", "closing_date"]:
        if col in out.columns:
            out[col] = pd.to_datetime(out[col], errors="coerce")

    out = out.sort_values(
        ["_source_priority", "award_date", "date_published", "closing_date"],
        ascending=[False, False, False, False],
        na_position="last",
        kind="mergesort",
    )
    out = out.drop_duplicates(subset=["opportunity_key"], keep="first")
    out = out.drop(columns=["_source_priority"], errors="ignore")
    return sort_operational(out)


def build_top_delta(df: pd.DataFrame, limit: int, min_score: int) -> pd.DataFrame:
    out = sort_operational(df)
    if "weight_score" not in out.columns:
        out["weight_score"] = 0
    out = out[out["weight_score"].fillna(0).astype(int) >= int(min_score)].copy()
    if limit and int(limit) > 0:
        out = out.head(int(limit)).copy()
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description="Fetch and clean awarded eTenders data")
    parser.add_argument("--since-date", help="Only fetch rows on/after this date where the API date fields allow it.")
    parser.add_argument("--since-days", type=int, help="Only fetch recent awarded rows. Suggested initial value: 365.")
    parser.add_argument("--stop-on-seen", action="store_true", help="Stop paging when a previously seen awarded tender is found.")
    parser.add_argument("--state-file", default=STATE_FILE_DEFAULT, help="JSON state file for seen awarded IDs.")
    parser.add_argument("--top-delta-limit", type=int, default=TOP_DELTA_LIMIT_DEFAULT)
    parser.add_argument("--top-delta-min-score", type=int, default=TOP_DELTA_MIN_SCORE_DEFAULT)
    args = parser.parse_args()

    setup_logging()
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(STATE_DIR, exist_ok=True)

    logging.info("Using base root: %s", BASE_ROOT)
    logging.info("Using output directory: %s", OUTPUT_DIR)
    run_ts = pd.Timestamp.now()

    try:
        acquire_lock(LOCK_FILE)

        if args.since_date:
            since_date = pd.to_datetime(args.since_date)
        elif args.since_days is not None:
            since_date = pd.Timestamp.today().normalize() - pd.Timedelta(days=args.since_days)
        else:
            since_date = None

        seen_ids = _load_seen_ids(args.state_file) if args.stop_on_seen else None

        rows = fetch_awarded_etenders(since_date=since_date, seen_ids=seen_ids)
        if not rows:
            logging.warning("No awarded rows returned from eTenders.")
            return 0

        raw_df = pd.json_normalize(rows, sep="_")
        logging.info("Raw awarded rows fetched: %d | cols: %d", len(raw_df), len(raw_df.columns))

        raw_df.to_csv(RAW_CSV, index=False)
        with pd.ExcelWriter(RAW_XLSX, engine="xlsxwriter") as writer:
            raw_df.to_excel(writer, index=False, sheet_name="Raw_Data")
        logging.info("Saved raw outputs.")

        mapped = map_and_clean(raw_df)
        mapped = add_matches_and_scoring(mapped)
        mapped["row_hash"] = compute_row_hash(mapped, HASH_COLS)
        mapped["run_timestamp"] = run_ts

        master_df = load_master(MASTER_XLSX)
        if not master_df.empty:
            master_df = ensure_output_columns(master_df)
            master_df["opportunity_key"] = master_df["opportunity_key"].fillna("").astype(str).str.strip().str.casefold()
            if "row_hash" not in master_df.columns or master_df["row_hash"].fillna("").eq("").all():
                master_df["row_hash"] = compute_row_hash(master_df, HASH_COLS)

        if not master_df.empty and "opportunity_key" in master_df.columns:
            master_ids = set(master_df["opportunity_key"].fillna("").astype(str).str.casefold())
            master_hash_map = dict(
                zip(
                    master_df["opportunity_key"].fillna("").astype(str).str.casefold(),
                    master_df["row_hash"].fillna("").astype(str),
                )
            )

            is_new = ~mapped["opportunity_key"].isin(master_ids)
            delta_df = mapped[is_new].copy()
            delta_df["change_type"] = "NEW"
            delta_df["previous_row_hash"] = ""

            previous_hash = mapped["opportunity_key"].map(master_hash_map).fillna("")
            is_updated = (~is_new) & (mapped["row_hash"] != previous_hash)
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

        top_delta_df = build_top_delta(delta_df, args.top_delta_limit, args.top_delta_min_score)

        updated_master = append_and_dedupe_master(master_df, mapped)
        hf_matches_df = updated_master[updated_master["hf_match"].fillna(False).astype(bool)].copy()
        hf_delta_df = delta_df[delta_df["hf_match"].fillna(False).astype(bool)].copy()

        write_excel_with_table(delta_df, DELTA_XLSX, "Delta", "tbl_awarded_delta")
        write_excel_with_table(updates_df, UPDATES_XLSX, "Updates", "tbl_awarded_updates")
        write_excel_with_table(top_delta_df, TOP_DELTA_XLSX, "Top_Delta", "tbl_awarded_top_delta")
        write_excel_with_table(updated_master, MASTER_XLSX, "Master", "tbl_awarded_master")
        write_excel_with_table(hf_matches_df, HF_MATCHES_XLSX, "HF_Matches", "tbl_awarded_hf_matches")
        write_excel_with_table(hf_delta_df, HF_DELTA_XLSX, "HF_Delta", "tbl_awarded_hf_delta")

        logging.info("Saved master rows: %d", len(updated_master))
        logging.info("Saved delta rows: %d", len(delta_df))
        logging.info("Saved update rows: %d", len(updates_df))
        logging.info("Saved HF match rows: %d", len(hf_matches_df))
        logging.info("Saved HF delta rows: %d", len(hf_delta_df))

        _save_seen_ids(args.state_file, updated_master["opportunity_key"].tolist())
        return 0

    except (requests.RequestException, ValueError, RuntimeError, OSError) as exc:
        logging.exception("FAILED: %s", exc)
        return 1
    finally:
        release_lock(LOCK_FILE)


if __name__ == "__main__":
    raise SystemExit(main())
