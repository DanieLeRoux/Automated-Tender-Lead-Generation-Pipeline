"""
transnet_esupplier_tender_pull_and_clean (Delta + Updates + Master)

Scrape currently advertised tenders from Transnet's eSupplier Portal and
older Azure eTender portal, clean/normalise them, and maintain:

1) MASTER Excel: latest snapshot per opportunity_key
2) DELTA Excel: brand new tenders discovered this run only
3) UPDATES Excel: existing tenders whose key operational fields changed
4) TOP DELTA Excel: highest-scoring new tenders from the current run

Notes:
- The Transnet portals are JavaScript-rendered. The script first saves a static
  HTML snapshot for diagnostics, then uses Selenium to render and paginate the
  tender tables by default.
- If Selenium is not installed and the static parse returns no rows, install it:
      pip install selenium

Typical run:
    python transnet_esupplier_tender_pull_and_clean.py

Optional:
    python transnet_esupplier_tender_pull_and_clean.py --include-expired
    python transnet_esupplier_tender_pull_and_clean.py --top-delta-limit 50
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
from pathlib import Path
import re
import sys
import time
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# -----------------------------
# Source configuration
# -----------------------------
# Transnet currently requires both portals to be checked:
# 1) the newer eSupplier Portal; and
# 2) the older Azure eTender system.
# See Transnet eTender home page wording: continue to check both systems.
ESUPPLIER_SOURCE = "Transnet eSupplier Portal"
AZURE_SOURCE = "Transnet Azure eTender Portal"

ESUPPLIER_ADVERTISED_URL = "https://esupplierportal.transnet.net/portal/advertisedTenders"
ESUPPLIER_OTHER_TENDERS_URL = "https://esupplierportal.transnet.net/portal/otherTenders"
AZURE_ADVERTISED_URL = "https://transnetetenders.azurewebsites.net/Home/AdvertisedTenders"

# Backwards-compatible names used elsewhere in the script.
SOURCE = ESUPPLIER_SOURCE
SOURCE_URL = ESUPPLIER_ADVERTISED_URL
OTHER_TENDERS_URL = ESUPPLIER_OTHER_TENDERS_URL

# === Folder configuration: new 02_TenderScrapers architecture ===
# Expected script location:
#   02_TenderScrapers\02_Sources\Transnet\00_Scripts\transnet_esupplier_tender_pull_and_clean.py
# The paths below are therefore derived from the script location instead of being hard-coded.
SCRIPT_PATH = Path(__file__).resolve()
SOURCE_ROOT = SCRIPT_PATH.parents[1]  # ...\02_Sources\Transnet
TENDER_SCRAPERS_ROOT = SCRIPT_PATH.parents[3]  # ...\02_TenderScrapers

BASE_ROOT = str(SOURCE_ROOT)
RAW_DIR = str(SOURCE_ROOT / "01_RawData")
OUTPUT_DIR = str(SOURCE_ROOT / "02_Outputs")
LOG_DIR = str(SOURCE_ROOT / "03_Logs")
STATE_DIR = str(SOURCE_ROOT / "04_State")
METADATA_DIR = str(SOURCE_ROOT / "05_Metadata")

LOG_PATH = os.path.join(LOG_DIR, "transnet_run.log")

MASTER_XLSX = os.path.join(OUTPUT_DIR, "transnet_master.xlsx")
DELTA_XLSX = os.path.join(OUTPUT_DIR, "transnet_delta.xlsx")
UPDATES_XLSX = os.path.join(OUTPUT_DIR, "transnet_updates.xlsx")
TOP_DELTA_XLSX = os.path.join(OUTPUT_DIR, "transnet_top_delta.xlsx")

RAW_CSV = os.path.join(RAW_DIR, "transnet_raw.csv")
RAW_XLSX = os.path.join(RAW_DIR, "transnet_raw_full.xlsx")
RAW_STATIC_HTML = os.path.join(RAW_DIR, "transnet_raw_page_static.html")
RAW_RENDERED_HTML = os.path.join(RAW_DIR, "transnet_raw_page_rendered.html")

STATE_FILE_DEFAULT = os.path.join(STATE_DIR, "seen_ids.json")
LOCK_FILE = os.path.join(STATE_DIR, "run.lock")

# -----------------------------
# Tunables
# -----------------------------
TIMEOUT_S = 60
SLEEP_S = 0.4
SELENIUM_WAIT_S = 45
TOP_DELTA_LIMIT_DEFAULT = 20
TOP_DELTA_MIN_SCORE_DEFAULT = 0

# Since this scraper is specifically for Transnet, all rows get the priority-client base score.
WEIGHT_PRIORITY_SOURCE = 10
WEIGHT_DESCRIPTION_KEYWORD = 1

HASH_COLS = [
    "tender_name",
    "description",
    "briefing_session",
    "closing_date",
    "tender_status",
    "document_urls",
]

DESCRIPTION_KEYWORD_RULES = {
    "Water and Sanitation": ["Water and Sanitation", "water", "sanitation"],
    "Housing": ["Housing"],
    "Education": ["Education"],
    "Transport": ["Transport", "freight", "rail", "port", "pipeline", "terminal"],
    "Nuclear": ["Nuclear"],
    "Bus Rapid Transit (BRT)": ["Bus Rapid Transit", "BRT"],
    "International Transition Plan Network (ITPN)": ["International Transition Plan Network", "ITPN"],
    "Design": ["Design"],
    "Refurbishment": ["Refurbishment", "rehabilitation", "upgrade"],
    "Engineering, Procurement, and Construction (EPC)": [
        "Engineering, Procurement, and Construction",
        "Engineering Procurement and Construction",
        "EPC",
    ],
    "Fire Engineering": ["Fire Engineering", "fire protection", "fire detection"],
    "HVAC": ["HVAC", "heating ventilation air conditioning", "air conditioning"],
    "Mechanical": ["Mechanical"],
    "Electrical": ["Electrical"],
    "Electronic": ["Electronic"],
    "Wet Services": ["Wet Services"],
    "Project Management": ["Project Management"],
    "Programme Management": ["Programme Management", "Program Management"],
    "Project Management Office (PMO)": ["Project Management Office", "PMO"],
    "Advisory Services": ["Advisory Services", "transaction advisory", "technical advisory"],
    "Structural Engineering": ["Structural Engineering"],
    "Civil Engineering": ["Civil Engineering"],
    "Feasibility Study": ["Feasibility Study", "feasibility"],
}

DEFAULT_COLUMNS = [
    "source",
    "source_subportal",
    "opportunity_key",
    "reference_number",
    "tender_name",
    "description",
    "briefing_session",
    "closing_date",
    "tender_status",
    "client",
    "source_url",
    "document_names",
    "document_urls",
    "document_count",
    "priority_source_match",
    "matched_description_keywords",
    "description_keyword_match_count",
    "weight_score",
    "row_hash",
    "run_timestamp",
    "change_type",
    "previous_row_hash",
]


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
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
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


def _save_text(path: str, text: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(text or "")


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


def _snake(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9]+", "_", value or "")
    value = re.sub(r"_+", "_", value).strip("_").lower()
    return value


def _column_alias(header: str) -> str:
    h = _normalise_search_text(header)
    aliases = {
        "reference number": "reference_number",
        "reference no": "reference_number",
        "tender number": "reference_number",
        "tender no": "reference_number",
        "contract number": "reference_number",
        "contract no": "reference_number",
        "bid number": "reference_number",
        "bid no": "reference_number",
        "rfp number": "reference_number",
        "rfp no": "reference_number",
        "rfq number": "reference_number",
        "rfq no": "reference_number",
        "rfi number": "reference_number",
        "rfi no": "reference_number",
        "enquiry number": "reference_number",
        "enquiry no": "reference_number",
        "event number": "reference_number",
        "event no": "reference_number",
        "tender name": "tender_name",
        "name of tender": "tender_name",
        "tender description": "description",
        "title": "tender_name",
        "description": "description",
        "briefing session": "briefing_session",
        "briefing": "briefing_session",
        "briefing date": "briefing_session",
        "site meeting": "briefing_session",
        "closing date": "closing_date",
        "closing date and time": "closing_date",
        "closing date time": "closing_date",
        "closing datetime": "closing_date",
        "tender status": "tender_status",
        "status": "tender_status",
    }
    return aliases.get(h, _snake(header))


def fetch_static_html(url: str) -> str:
    sess = build_session()
    response = sess.get(url, timeout=TIMEOUT_S)
    response.raise_for_status()
    return response.text


def _start_selenium_driver():
    """Start Edge first, then Chrome. Selenium Manager handles drivers in Selenium 4.6+."""
    try:
        from selenium import webdriver
        from selenium.common.exceptions import WebDriverException
        from selenium.webdriver.chrome.options import Options as ChromeOptions
        from selenium.webdriver.edge.options import Options as EdgeOptions
    except ImportError as exc:
        raise RuntimeError(
            "The Transnet tender pages are JavaScript-rendered and Selenium is required for full pagination. "
            "Install it with: pip install selenium"
        ) from exc

    def start_edge():
        opts = EdgeOptions()
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--window-size=1920,1080")
        opts.add_argument("--no-sandbox")
        return webdriver.Edge(options=opts)

    def start_chrome():
        opts = ChromeOptions()
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--window-size=1920,1080")
        opts.add_argument("--no-sandbox")
        return webdriver.Chrome(options=opts)

    driver = None
    last_error = None
    for starter in (start_edge, start_chrome):
        try:
            driver = starter()
            break
        except WebDriverException as exc:
            last_error = exc
            driver = None

    if driver is None:
        raise RuntimeError(
            "Could not start Edge or Chrome through Selenium. Ensure Microsoft Edge or Chrome is installed, "
            "then run: pip install selenium. Selenium Manager should handle the driver automatically. "
            f"Last error: {last_error}"
        )
    return driver


def _visible_tables_html(driver) -> str:
    """Return only visible table HTML to avoid parsing hidden inactive tab tables."""
    try:
        html = driver.execute_script(
            """
            const tables = Array.from(document.querySelectorAll('table'));
            const visible = tables.filter(t => !!(t.offsetWidth || t.offsetHeight || t.getClientRects().length));
            return visible.map(t => t.outerHTML).join('\n');
            """
        )
        if html:
            return f"<html><body>{html}</body></html>"
    except Exception:
        pass
    return driver.page_source


def _page_text(driver) -> str:
    try:
        return driver.find_element("tag name", "body").text or ""
    except Exception:
        return driver.page_source or ""


def _wait_for_tender_table(driver, wait_s: int) -> None:
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait

    wait = WebDriverWait(driver, wait_s)
    wait.until(
        lambda d: (
            len(d.find_elements(By.CSS_SELECTOR, "table tbody tr")) > 0
            or "Reference Number" in d.page_source
            or "Tender Name" in d.page_source
            or "Tender Description" in d.page_source
            or "Closing Date" in d.page_source
        )
    )
    # Give DataTables/Angular/Knockout a short second pass to populate rows.
    time.sleep(2)


def _click_tab_by_text(driver, tab_text: Optional[str], wait_s: int) -> None:
    if not tab_text:
        return
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait

    wait = WebDriverWait(driver, wait_s)
    xpath = (
        "//*[self::a or self::button or self::li]"
        f"[contains(normalize-space(.), {json.dumps(tab_text)})]"
    )
    try:
        elem = wait.until(lambda d: next((e for e in d.find_elements(By.XPATH, xpath) if e.is_displayed()), None))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", elem)
        time.sleep(2)
    except Exception:
        logging.warning("Could not click tab '%s'. Will parse the default visible table state.", tab_text)


def _try_set_largest_page_length(driver) -> bool:
    """Set DataTables length dropdown to All/largest value where available."""
    try:
        from selenium.webdriver.support.ui import Select
    except Exception:
        return False

    changed = False
    try:
        selects = [s for s in driver.find_elements("tag name", "select") if s.is_displayed()]
    except Exception:
        return False

    for sel_el in selects:
        try:
            select = Select(sel_el)
            options = []
            for opt in select.options:
                txt = (opt.text or "").strip()
                val = (opt.get_attribute("value") or "").strip()
                options.append((txt, val, opt))

            # Prefer explicit All/-1, otherwise largest numeric option.
            chosen = None
            for txt, val, opt in options:
                if txt.casefold() == "all" or val == "-1":
                    chosen = (txt, val)
                    break
            if chosen is None:
                numeric = []
                for txt, val, opt in options:
                    for candidate in (val, txt):
                        if str(candidate).strip().isdigit():
                            numeric.append((int(str(candidate).strip()), txt, val))
                            break
                if numeric:
                    _, txt, val = sorted(numeric, reverse=True)[0]
                    chosen = (txt, val)

            if chosen:
                txt, val = chosen
                before = driver.page_source
                try:
                    if val:
                        select.select_by_value(val)
                    else:
                        select.select_by_visible_text(txt)
                except Exception:
                    select.select_by_visible_text(txt)
                time.sleep(2)
                changed = changed or (driver.page_source != before)
        except Exception:
            continue

    return changed


def _is_disabled(elem) -> bool:
    try:
        cls = (elem.get_attribute("class") or "").casefold()
        aria = (elem.get_attribute("aria-disabled") or "").casefold()
        disabled = elem.get_attribute("disabled")
        parent_cls = ""
        try:
            parent_cls = (elem.find_element("xpath", "..").get_attribute("class") or "").casefold()
        except Exception:
            pass
        return bool(disabled) or aria == "true" or "disabled" in cls or "disabled" in parent_cls
    except Exception:
        return False


def _find_next_button(driver):
    from selenium.webdriver.common.by import By

    selectors = [
        "a.paginate_button.next",
        "li.paginate_button.next a",
        "button[aria-label*='Next']",
        "a[aria-label*='Next']",
        ".pagination a[rel='next']",
    ]
    candidates = []
    for selector in selectors:
        try:
            candidates.extend(driver.find_elements(By.CSS_SELECTOR, selector))
        except Exception:
            pass

    try:
        candidates.extend(
            driver.find_elements(
                By.XPATH,
                "//*[self::a or self::button][normalize-space(.)='Next' or contains(normalize-space(.), 'Next') or contains(@aria-label, 'Next')]",
            )
        )
    except Exception:
        pass

    for elem in candidates:
        try:
            if elem.is_displayed() and not _is_disabled(elem):
                return elem
        except Exception:
            continue
    return None


def _first_visible_row_signature(driver) -> str:
    try:
        rows = [r for r in driver.find_elements("css selector", "table tbody tr") if r.is_displayed()]
        if rows:
            return " ".join(rows[0].text.split())[:500]
    except Exception:
        pass
    return hashlib.sha256((_visible_tables_html(driver) or "").encode("utf-8")).hexdigest()


def fetch_rendered_html_with_selenium(url: str, wait_s: int = SELENIUM_WAIT_S) -> str:
    """Backwards-compatible helper: render one page and return HTML."""
    driver = _start_selenium_driver()
    try:
        driver.get(url)
        _wait_for_tender_table(driver, wait_s)
        return driver.execute_script("return document.documentElement.outerHTML;")
    finally:
        driver.quit()


def _parse_rendered_current_page(
    html: str,
    source_url: str,
    source_name: str,
    source_subportal: str,
) -> pd.DataFrame:
    df = parse_tender_tables(html, source_url, source_name=source_name, source_subportal=source_subportal)
    if df.empty:
        df = parse_from_datatables_text(html, source_url, source_name=source_name, source_subportal=source_subportal)
    return df


def fetch_rendered_table_pages_with_selenium(
    url: str,
    source_name: str,
    source_subportal: str,
    tab_text: Optional[str] = None,
    wait_s: int = SELENIUM_WAIT_S,
    max_pages: int = 100,
) -> Tuple[pd.DataFrame, str]:
    """Render a JS tender table and extract every pagination page deterministically."""
    driver = _start_selenium_driver()
    all_frames: List[pd.DataFrame] = []
    last_html = ""

    try:
        driver.get(url)
        _wait_for_tender_table(driver, wait_s)
        _click_tab_by_text(driver, tab_text, wait_s)
        _wait_for_tender_table(driver, wait_s)

        # First try to set the table to All/largest page length. If it works, pagination may not be required.
        _try_set_largest_page_length(driver)
        time.sleep(1)

        seen_page_signatures: Set[str] = set()
        page_no = 1

        while page_no <= max_pages:
            sig = _first_visible_row_signature(driver)
            if sig in seen_page_signatures:
                logging.info("Pagination loop stopped for %s because page signature repeated.", source_subportal)
                break
            seen_page_signatures.add(sig)

            html = _visible_tables_html(driver)
            last_html = html
            df = _parse_rendered_current_page(html, url, source_name, source_subportal)
            logging.info("Rendered parse rows for %s page %d: %d", source_subportal, page_no, len(df))
            if not df.empty:
                all_frames.append(df)

            next_button = _find_next_button(driver)
            if next_button is None:
                break

            before = _first_visible_row_signature(driver)
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                time.sleep(0.2)
                driver.execute_script("arguments[0].click();", next_button)
            except Exception:
                logging.exception("Failed to click Next on %s page %d", source_subportal, page_no)
                break

            # Wait until the first visible row changes, but do not fail if it is very slow.
            deadline = time.time() + wait_s
            while time.time() < deadline:
                time.sleep(0.5)
                after = _first_visible_row_signature(driver)
                if after and after != before:
                    break
            page_no += 1

        if not all_frames:
            return pd.DataFrame(columns=DEFAULT_COLUMNS), last_html

        out = pd.concat(all_frames, ignore_index=True)
        return out, last_html
    finally:
        driver.quit()


def _extract_cell_links(cell, base_url: str) -> Tuple[str, str, int]:
    names: List[str] = []
    urls: List[str] = []
    for a in cell.find_all("a", href=True):
        href = a.get("href") or ""
        url = urljoin(base_url, href)
        text = " ".join(a.get_text(" ", strip=True).split()) or url
        if url and url not in urls:
            urls.append(url)
            names.append(text)
    return "; ".join(names), "; ".join(urls), len(urls)


def parse_tender_tables(
    html: str,
    source_url: str,
    source_name: str = SOURCE,
    source_subportal: str = "",
) -> pd.DataFrame:
    soup = BeautifulSoup(html or "", "html.parser")
    rows_out: List[Dict[str, Any]] = []

    expected = {"reference_number", "tender_name", "description", "briefing_session", "closing_date", "tender_status"}

    for table in soup.find_all("table"):
        header_cells = table.find_all("th")
        headers = [" ".join(th.get_text(" ", strip=True).split()) for th in header_cells]

        # Some tables use the first row as headers.
        if not headers:
            first_tr = table.find("tr")
            if first_tr:
                first_cells = first_tr.find_all(["th", "td"])
                headers = [" ".join(c.get_text(" ", strip=True).split()) for c in first_cells]

        col_names = [_column_alias(h) for h in headers]
        if not any(c in expected for c in col_names):
            continue

        body_rows = table.select("tbody tr") or table.find_all("tr")[1:]
        for tr in body_rows:
            cells = tr.find_all("td")
            if not cells:
                continue

            row: Dict[str, Any] = {
                "source": source_name,
                "source_subportal": source_subportal,
                "client": "Transnet SOC Ltd",
                "source_url": source_url,
            }
            all_doc_names: List[str] = []
            all_doc_urls: List[str] = []

            for idx, cell in enumerate(cells):
                col = col_names[idx] if idx < len(col_names) else f"extra_col_{idx + 1}"
                text = " ".join(cell.get_text(" ", strip=True).split())
                row[col] = text

                doc_names, doc_urls, _ = _extract_cell_links(cell, source_url)
                if doc_urls:
                    for name in doc_names.split("; "):
                        if name and name not in all_doc_names:
                            all_doc_names.append(name)
                    for url in doc_urls.split("; "):
                        if url and url not in all_doc_urls:
                            all_doc_urls.append(url)

            # If the portal has a title but no explicit description column, use the title as description fallback.
            if not str(row.get("description", "")).strip() and str(row.get("tender_name", "")).strip():
                row["description"] = row.get("tender_name", "")

            if not any(str(row.get(c, "")).strip() for c in expected):
                continue

            row["document_names"] = "; ".join(all_doc_names)
            row["document_urls"] = "; ".join(all_doc_urls)
            row["document_count"] = len(all_doc_urls)
            rows_out.append(row)

    return pd.DataFrame(rows_out)


def parse_from_datatables_text(
    html: str,
    source_url: str,
    source_name: str = SOURCE,
    source_subportal: str = "",
) -> pd.DataFrame:
    """Fallback parser for rendered DataTables text where table tags are not cleanly parsed."""
    soup = BeautifulSoup(html or "", "html.parser")
    text = soup.get_text("\n", strip=True)
    lines = [" ".join(x.split()) for x in text.splitlines() if " ".join(x.split())]

    if not any(term in text for term in ["Reference Number", "Tender Number", "Contract Number", "Bid Number"]):
        return pd.DataFrame()
    if not any(term in text for term in ["Tender Name", "Description", "Tender Description"]):
        return pd.DataFrame()

    # Locate a probable header and then consume fixed-width DataTables-style rows.
    start_idx = 0
    for i, line in enumerate(lines):
        if _column_alias(line) == "reference_number":
            # Most Transnet tables expose six operational columns.
            start_idx = i + 6
            break

    data_lines = lines[start_idx:]
    rows: List[Dict[str, Any]] = []

    ref_pattern = re.compile(
        r"^(?:PPA-\d{3,}|[A-Z]{2,5}/\d{4}/\d{2}/\d{4}/\d+/(?:RFP|RFQ|RFI|RFR|EOI|TENDER)|[A-Z0-9][A-Z0-9./_-]{3,})$",
        re.I,
    )
    i = 0
    while i < len(data_lines):
        line = data_lines[i].strip()
        if not ref_pattern.search(line):
            i += 1
            continue

        ref = line
        tender_name = data_lines[i + 1] if i + 1 < len(data_lines) else ""
        desc = data_lines[i + 2] if i + 2 < len(data_lines) else ""
        briefing = data_lines[i + 3] if i + 3 < len(data_lines) else ""
        closing = data_lines[i + 4] if i + 4 < len(data_lines) else ""
        status = data_lines[i + 5] if i + 5 < len(data_lines) else ""

        rows.append(
            {
                "source": source_name,
                "source_subportal": source_subportal,
                "client": "Transnet SOC Ltd",
                "source_url": source_url,
                "reference_number": ref,
                "tender_name": tender_name,
                "description": desc or tender_name,
                "briefing_session": briefing,
                "closing_date": closing,
                "tender_status": status,
                "document_names": "",
                "document_urls": "",
                "document_count": 0,
            }
        )
        i += 6

    return pd.DataFrame(rows)


def _source_html_path(source_subportal: str, kind: str) -> str:
    safe = _snake(source_subportal) or "source"
    return os.path.join(RAW_DIR, f"transnet_raw_page_{kind}_{safe}.html")


def fetch_transnet_rows(
    use_selenium: bool = True,
    include_other_tenders: bool = True,
    include_azure: bool = True,
    max_pages: int = 100,
) -> pd.DataFrame:
    sources: List[Dict[str, Optional[str]]] = [
        {
            "source_name": ESUPPLIER_SOURCE,
            "source_subportal": "eSupplier Advertised Tenders",
            "url": ESUPPLIER_ADVERTISED_URL,
            "tab_text": None,
        }
    ]

    if include_other_tenders:
        sources.append(
            {
                "source_name": ESUPPLIER_SOURCE,
                "source_subportal": "eSupplier Other Tenders",
                "url": ESUPPLIER_OTHER_TENDERS_URL,
                "tab_text": None,
            }
        )

    if include_azure:
        sources.extend(
            [
                {
                    "source_name": AZURE_SOURCE,
                    "source_subportal": "Azure Open Tenders",
                    "url": AZURE_ADVERTISED_URL,
                    "tab_text": "Open Tenders",
                },
                {
                    "source_name": AZURE_SOURCE,
                    "source_subportal": "Azure Other Tenders",
                    "url": AZURE_ADVERTISED_URL,
                    "tab_text": "Other Tenders",
                },
            ]
        )

    all_frames: List[pd.DataFrame] = []

    for src in sources:
        url = str(src["url"])
        source_name = str(src["source_name"])
        source_subportal = str(src["source_subportal"])
        tab_text = src.get("tab_text")

        logging.info("Fetching Transnet page: %s | %s", source_subportal, url)

        df = pd.DataFrame()

        # Static parse is retained as a fallback/diagnostic, but Selenium is preferred because both portals are JS-rendered.
        try:
            static_html = fetch_static_html(url)
            static_path = _source_html_path(source_subportal, "static")
            _save_text(static_path, static_html)
            logging.info("Saved static raw HTML: %s", static_path)

            if not use_selenium:
                df = parse_tender_tables(static_html, url, source_name=source_name, source_subportal=source_subportal)
                if df.empty:
                    df = parse_from_datatables_text(static_html, url, source_name=source_name, source_subportal=source_subportal)
                logging.info("Static parse rows for %s: %d", source_subportal, len(df))
        except Exception as exc:
            logging.warning("Static fetch/parse failed for %s: %s", source_subportal, exc)

        if use_selenium:
            logging.info("Using Selenium full pagination capture for %s", source_subportal)
            rendered_df, rendered_html = fetch_rendered_table_pages_with_selenium(
                url=url,
                source_name=source_name,
                source_subportal=source_subportal,
                tab_text=str(tab_text) if tab_text else None,
                wait_s=SELENIUM_WAIT_S,
                max_pages=max_pages,
            )
            rendered_path = _source_html_path(source_subportal, "rendered")
            _save_text(rendered_path, rendered_html)
            logging.info("Saved rendered raw HTML: %s", rendered_path)
            df = rendered_df
            logging.info("Rendered total rows for %s: %d", source_subportal, len(df))

        if not df.empty:
            all_frames.append(df)

        time.sleep(SLEEP_S)

    if not all_frames:
        return pd.DataFrame(columns=DEFAULT_COLUMNS)

    out = pd.concat(all_frames, ignore_index=True)
    return out


def _parse_date_value(value: Any) -> pd.Timestamp:
    if value is None or pd.isna(value):
        return pd.NaT

    raw = str(value).strip()
    if not raw:
        return pd.NaT

    # Common clean-up for Transnet/DataTables date strings.
    cleaned = raw.replace("@", " ")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()

    # Try both SA-style and US-style interpretations because snippets sometimes expose M/D/YYYY.
    for dayfirst in (True, False):
        dt = pd.to_datetime(cleaned, errors="coerce", dayfirst=dayfirst)
        if pd.notna(dt):
            return dt

    return pd.NaT


def _row_completeness_score(df: pd.DataFrame) -> pd.Series:
    """Score how much useful information a parsed row contains.

    Used only for choosing the best duplicate row where the same reference appears
    on more than one Transnet subportal or appears in both an old master and the
    current run. It does not affect tender relevance scoring.
    """
    if df.empty:
        return pd.Series([], dtype="int64", index=df.index)

    score = pd.Series(0, index=df.index, dtype="int64")
    for col in [
        "tender_name",
        "description",
        "briefing_session",
        "closing_date",
        "tender_status",
        "document_names",
        "document_urls",
        "source_subportal",
    ]:
        if col in df.columns:
            score += df[col].fillna("").astype(str).str.strip().ne("").astype(int)

    if "document_count" in df.columns:
        score += pd.to_numeric(df["document_count"], errors="coerce").fillna(0).astype(int)

    return score


def normalise_rows(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    for col in DEFAULT_COLUMNS:
        if col not in out.columns and col not in {"row_hash", "run_timestamp", "change_type", "previous_row_hash"}:
            out[col] = "" if col != "document_count" else 0

    out["source"] = out["source"].replace("", SOURCE).fillna(SOURCE)
    out["client"] = out["client"].replace("", "Transnet SOC Ltd").fillna("Transnet SOC Ltd")
    out["source_url"] = out["source_url"].replace("", SOURCE_URL).fillna(SOURCE_URL)

    text_cols = [
        "source_subportal",
        "reference_number",
        "tender_name",
        "description",
        "briefing_session",
        "tender_status",
        "document_names",
        "document_urls",
    ]
    for col in text_cols:
        out[col] = out[col].fillna("").astype(str).str.strip()

    out["closing_date"] = out["closing_date"].apply(_parse_date_value)

    # Stable key. Reference number is the primary identifier on the portal.
    out["reference_number"] = out["reference_number"].fillna("").astype(str).str.strip()
    out = out[out["reference_number"].ne("")].copy()
    out["opportunity_key"] = "transnet:" + out["reference_number"].str.casefold()

    out["document_count"] = pd.to_numeric(out.get("document_count", 0), errors="coerce").fillna(0).astype(int)

    # If duplicate references appear across subportals, keep the richest row first.
    # If equally complete, keep the row with the nearest valid closing date.
    out["_row_completeness"] = _row_completeness_score(out)
    out = out.sort_values(
        ["_row_completeness", "closing_date"],
        ascending=[False, True],
        na_position="last",
        kind="mergesort",
    )
    out = out.drop_duplicates(subset=["opportunity_key"], keep="first")
    out = out.drop(columns=["_row_completeness"], errors="ignore")

    return out[
        [
            "source",
            "source_subportal",
            "opportunity_key",
            "reference_number",
            "tender_name",
            "description",
            "briefing_session",
            "closing_date",
            "tender_status",
            "client",
            "source_url",
            "document_names",
            "document_urls",
            "document_count",
        ]
    ].copy()


def filter_current_tenders(df: pd.DataFrame, include_expired: bool = False) -> pd.DataFrame:
    if include_expired or df.empty:
        return df.copy()

    out = df.copy()
    before = len(out)
    today = pd.Timestamp.today().normalize()

    if "closing_date" in out.columns:
        out = out[pd.notna(out["closing_date"]) & (out["closing_date"] >= today)].copy()

    if "tender_status" in out.columns:
        status_norm = out["tender_status"].fillna("").astype(str).str.casefold()
        bad_status = status_norm.str.contains("closed|cancel|awarded|withdrawn|expired", regex=True, na=False)
        out = out[~bad_status].copy()

    logging.info("Filtered current Transnet tenders: %d -> %d rows", before, len(out))
    return out


def add_priority_scoring(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["priority_source_match"] = True

    def matched_keywords(row: pd.Series) -> str:
        combined = " ".join(
            str(row.get(c, "")) for c in ["tender_name", "description", "briefing_session"]
        )
        norm = _normalise_search_text(combined)
        matches = []
        for label, phrases in DESCRIPTION_KEYWORD_RULES.items():
            if any(_contains_phrase(norm, phrase) for phrase in phrases):
                matches.append(label)
        return "; ".join(matches)

    if out.empty:
        out["matched_description_keywords"] = pd.Series(dtype="object")
        out["description_keyword_match_count"] = pd.Series(dtype="int64")
        out["weight_score"] = pd.Series(dtype="int64")
        return out

    out["matched_description_keywords"] = out.apply(matched_keywords, axis=1)
    out["description_keyword_match_count"] = out["matched_description_keywords"].apply(
        lambda x: 0 if not x else len(str(x).split("; "))
    )
    out["weight_score"] = (
        out["priority_source_match"].astype(int) * WEIGHT_PRIORITY_SOURCE
        + out["description_keyword_match_count"].astype(int) * WEIGHT_DESCRIPTION_KEYWORD
    )

    sort_cols = [c for c in ["weight_score", "closing_date"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

    return out


def compute_row_hash(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    if df.empty:
        return pd.Series([], dtype="object", index=df.index)

    use_cols = [c for c in cols if c in df.columns]
    if not use_cols:
        return pd.Series([""] * len(df), index=df.index)

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

    with pd.ExcelWriter(tmp, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm") as writer:
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
    except Exception as exc:
        logging.exception("Failed to read master file (%s): %s", path, exc)
        return pd.DataFrame()


def append_and_dedupe_master(master: pd.DataFrame, incoming: pd.DataFrame) -> pd.DataFrame:
    """Append current run rows to the master and keep the best/latest version per key.

    Important: incoming rows must win over old master rows when the same
    opportunity_key is seen again. Otherwise changes such as extended closing dates
    can be detected in the updates file but fail to update the master.
    """
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
    out = out[out["opportunity_key"].ne("")].copy()

    if "closing_date" in out.columns:
        out["closing_date"] = pd.to_datetime(out["closing_date"], errors="coerce", dayfirst=True)

    out["_row_completeness"] = _row_completeness_score(out)

    # Keep the last row after sorting: incoming beats old master; within incoming,
    # more complete rows beat sparse rows.
    out = out.sort_values(
        ["opportunity_key", "_source_priority", "_row_completeness"],
        ascending=[True, True, True],
        na_position="last",
        kind="mergesort",
    )
    out = out.drop_duplicates(subset=["opportunity_key"], keep="last")
    out = out.drop(columns=["_source_priority", "_row_completeness"], errors="ignore")

    # User-facing sort: most relevant first, then nearest closing date.
    sort_cols = [c for c in ["weight_score", "closing_date"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last", kind="mergesort")

    return out

def build_top_delta(df: pd.DataFrame, limit: int, min_score: int) -> pd.DataFrame:
    out = df.copy()
    if "weight_score" not in out.columns:
        out["weight_score"] = 0

    out = out[out["weight_score"].fillna(0).astype(int) >= int(min_score)].copy()
    sort_cols = [c for c in ["weight_score", "closing_date"] if c in out.columns]
    if sort_cols:
        ascending = [False if c == "weight_score" else True for c in sort_cols]
        out = out.sort_values(sort_cols, ascending=ascending, na_position="last")

    if limit and int(limit) > 0:
        out = out.head(int(limit)).copy()
    return out


def _load_seen_ids(path: str) -> Set[str]:
    try:
        if not os.path.exists(path):
            return set()
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return set(str(x).strip().casefold() for x in (data or []))
    except (OSError, ValueError, TypeError) as exc:
        logging.exception("Failed to load state file: %s -- %s", path, exc)
        return set()


def _save_seen_ids(path: str, ids: Iterable[str]) -> None:
    try:
        norm = sorted(str(x).strip().casefold() for x in ids if str(x).strip())
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(norm, fh, ensure_ascii=False)
        os.replace(tmp, path)
    except (OSError, TypeError) as exc:
        logging.exception("Failed to save state file: %s -- %s", path, exc)


def main() -> int:
    parser = argparse.ArgumentParser(description="Fetch and clean Transnet tenders from both Transnet tender portals")
    parser.add_argument("--include-expired", action="store_true", help="keep expired/closed tenders in clean outputs")
    parser.add_argument("--no-selenium", action="store_true", help="do not use Selenium; not recommended because Transnet pages are JS-rendered")
    parser.add_argument("--include-other-tenders", action="store_true", help="deprecated/backwards compatible; Other Tenders are included by default")
    parser.add_argument("--skip-other-tenders", action="store_true", help="skip Other Tenders pages")
    parser.add_argument("--skip-azure", action="store_true", help="skip the older Azure Transnet eTender portal")
    parser.add_argument("--max-pages", type=int, default=100, help="maximum pagination pages to click per source")
    parser.add_argument("--state-file", default=STATE_FILE_DEFAULT, help="path to JSON file storing seen opportunity keys")
    parser.add_argument("--top-delta-limit", type=int, default=TOP_DELTA_LIMIT_DEFAULT, help="top delta row limit; use 0 for all")
    parser.add_argument("--top-delta-min-score", type=int, default=TOP_DELTA_MIN_SCORE_DEFAULT, help="minimum weight_score for top delta")
    args = parser.parse_args()

    setup_logging()
    for folder in [RAW_DIR, OUTPUT_DIR, LOG_DIR, STATE_DIR, METADATA_DIR]:
        os.makedirs(folder, exist_ok=True)
    logging.info("Using tender scrapers root: %s", TENDER_SCRAPERS_ROOT)
    logging.info("Using source root: %s", BASE_ROOT)
    logging.info("Using raw directory: %s", RAW_DIR)
    logging.info("Using output directory: %s", OUTPUT_DIR)

    run_ts = pd.Timestamp.now()

    try:
        acquire_lock(LOCK_FILE)

        raw_df = fetch_transnet_rows(
            use_selenium=not args.no_selenium,
            include_other_tenders=not args.skip_other_tenders,
            include_azure=not args.skip_azure,
            max_pages=args.max_pages,
        )
        logging.info("Parsed Transnet raw rows: %d | cols: %d", len(raw_df), len(raw_df.columns))

        raw_df.to_csv(RAW_CSV, index=False)
        logging.info("Saved raw CSV: %s", RAW_CSV)

        with pd.ExcelWriter(RAW_XLSX, engine="xlsxwriter") as writer:
            raw_df.to_excel(writer, index=False, sheet_name="Raw_Data")
        logging.info("Saved raw Excel: %s", RAW_XLSX)

        mapped = normalise_rows(raw_df)
        mapped = filter_current_tenders(mapped, include_expired=args.include_expired)
        mapped = add_priority_scoring(mapped)
        mapped["row_hash"] = compute_row_hash(mapped, HASH_COLS)
        mapped["run_timestamp"] = run_ts

        master_df = load_master(MASTER_XLSX)

        if not master_df.empty and "opportunity_key" in master_df.columns:
            master_df = master_df.copy()
            master_df["opportunity_key"] = master_df["opportunity_key"].fillna("").astype(str).str.strip().str.casefold()
            master_df = master_df[master_df["opportunity_key"].ne("")]
            if "row_hash" not in master_df.columns:
                master_df["row_hash"] = compute_row_hash(master_df, HASH_COLS)

        if not master_df.empty and "opportunity_key" in master_df.columns:
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

        sort_cols = [c for c in ["weight_score", "closing_date"] if c in delta_df.columns]
        if sort_cols:
            ascending = [False if c == "weight_score" else True for c in sort_cols]
            delta_df = delta_df.sort_values(sort_cols, ascending=ascending, na_position="last")
            updates_df = updates_df.sort_values(sort_cols, ascending=ascending, na_position="last")

        top_delta_df = build_top_delta(delta_df, args.top_delta_limit, args.top_delta_min_score)

        logging.info("Delta rows this run: %d", len(delta_df))
        logging.info("Updates rows this run: %d", len(updates_df))
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

        if "opportunity_key" in updated_master.columns:
            _save_seen_ids(args.state_file, updated_master["opportunity_key"].tolist())

        return 0

    except (requests.RequestException, ValueError, RuntimeError, OSError) as exc:
        logging.exception("FAILED: %s", exc)
        return 1
    finally:
        release_lock(LOCK_FILE)


if __name__ == "__main__":
    raise SystemExit(main())
