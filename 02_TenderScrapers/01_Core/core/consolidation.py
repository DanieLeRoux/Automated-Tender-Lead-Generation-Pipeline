"""Load, standardise, score and consolidate tender outputs from all source scrapers."""
from __future__ import annotations

import hashlib
import json
import os
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List

import pandas as pd

from .date_utils import coerce_date_columns, current_flag, parse_date, parse_date_series
from .excel_utils import write_excel_with_table, write_multi_sheet_excel
from .scoring_engine import load_scoring_config, score_dataframe
from .standard_schema import STANDARD_COLUMNS, DATE_COLUMNS, HASH_COLUMNS
from .text_utils import clean_text, first_non_empty, join_unique, normalise_identifier, normalise_text


SOURCE_DEFAULTS = {
    # Source-level defaults used when a portal does not expose these fields directly.
    # This keeps consolidated scoring and filtering consistent without relying on every
    # individual scraper to populate all organisation/location fields.
    "etenders": {
        "source_url": "https://www.etenders.gov.za/Home/opportunities?id=1",
    },
    "world bank": {
        "client": "World Bank",
        "source_url": "https://projects.worldbank.org/en/projects-operations/opportunities?srce=both",
    },
    "city of cape town": {
        "client": "City of Cape Town",
        "province": "Western Cape",
        "source_url": "https://web1.capetown.gov.za/web1/tenderportal/Tender",
    },
    "dbsa": {
        "client": "Development Bank of Southern Africa",
        "source_url": "https://www.dbsa.org/procurement",
    },
    "vcwb": {
        "client": "Vaal Central Water",
        "source_url": "https://vaalcentralwater.co.za/advertised-tenders/",
    },
    "transnet": {
        "client": "Transnet",
        "source_url": "https://esupplierportal.transnet.net/portal/advertisedTenders",
    },
    "acsa": {
        "client": "Airports Company South Africa",
        "source_url": "https://www.airports.co.za/Pages/Tender-Bulletin.aspx",
    },
    "coega": {
        "client": "Coega Development Corporation",
        "province": "Eastern Cape",
        "source_url": "https://www.coega.co.za/site/tenders/current-tenders",
    },
    "dhs": {
        "client": "Department of Human Settlements",
        "source_url": "https://www.dhs.gov.za/tenders",
    },
    "gtac": {
        "client": "Government Technical Advisory Centre",
        "source_url": "https://www.gtac.gov.za/tenders/advertised-tenders/",
    },
    "necsa": {
        "client": "South African Nuclear Energy Corporation",
        "province": "North West",
        "country_name": "South Africa",
        "source_url": "https://www.necsa.co.za/tenders/",
    },
    "nnr": {
        "client": "National Nuclear Regulator",
        "province": "Gauteng",
        "country_name": "South Africa",
        "source_url": "https://nnr.co.za/tenders/",
    },
    "national nuclear regulator": {
        "client": "National Nuclear Regulator",
        "province": "Gauteng",
        "country_name": "South Africa",
        "source_url": "https://nnr.co.za/tenders/",
    },
    "rand water": {
        "client": "Rand Water",
        "province": "Gauteng",
        "country_name": "South Africa",
        "sector": "Water",
        "source_url": "https://www.randwater.co.za/availabletenders.php",
    },
    "randwater": {
        "client": "Rand Water",
        "province": "Gauteng",
        "country_name": "South Africa",
        "sector": "Water",
        "source_url": "https://www.randwater.co.za/availabletenders.php",
    },
    "tcta": {
        "client": "Trans-Caledon Tunnel Authority",
        "province": "Gauteng",
        "country_name": "South Africa",
        "sector": "Water",
        "source_url": "https://www.tcta.co.za/tenders/",
    },
    "trans caledon tunnel authority": {
        "client": "Trans-Caledon Tunnel Authority",
        "province": "Gauteng",
        "country_name": "South Africa",
        "sector": "Water",
        "source_url": "https://www.tcta.co.za/tenders/",
    },
    "city of ekurhuleni": {
        "client": "City of Ekurhuleni",
        "province": "Gauteng",
        "country_name": "South Africa",
        "source_url": "https://www.ekurhuleni.gov.za/for-my-business/tenders/open-tenders/",
    },
    "johannesburg water": {
        "client": "Johannesburg Water",
        "province": "Gauteng",
        "country_name": "South Africa",
        "sector": "Water",
        "source_url": "https://scm.johannesburgwater.co.za/supply-chain/tenders/all-open-tenders/",
    },
    "city of mbombela": {
        "client": "City of Mbombela",
        "province": "Mpumalanga",
        "country_name": "South Africa",
        "source_url": "https://www.mbombela.gov.za/downloads/tenders/",
    },
    "city of polokwane": {
        "client": "City of Polokwane",
        "province": "Limpopo",
        "country_name": "South Africa",
        "source_url": "https://www.polokwane.gov.za/tender-documents/",
    },
    "namwater": {
        "client": "NamWater",
        "country_name": "Namibia",
        "sector": "Water",
        "source_url": "https://www.namwater.com.na/index.php/bids-list",
    },
    "umngeni uthukela water": {
        "client": "uMngeni-uThukela Water",
        "province": "KwaZulu-Natal",
        "source_url": "https://www.umngeni-uthukela.co.za/tender/",
    },
    "umngeniuthukelawater": {
        "client": "uMngeni-uThukela Water",
        "province": "KwaZulu-Natal",
        "source_url": "https://www.umngeni-uthukela.co.za/tender/",
    },
}


# Daily-review exclusion rules (v8 refined). These do not delete records from the consolidated
# master register; they only prevent clearly irrelevant tenders from appearing in
# the bid-manager review tabs and consolidated top-delta shortlist.
REVIEW_EXCLUSION_KEYWORDS = {
    "Repairs": ["repairs", "repair", "repairing", "fixing"],
    "Maintenance": ["maintenance", "maintain", "servicing", "service and maintenance", "service maintenance", "preventative maintenance", "routine maintenance"],
    "Financial Advisory": ["financial advisory", "financial advisor", "financial adviser", "accounting advisory", "audit support", "accounting services", "bookkeeping"],
    "Employee Wellness": ["employee wellness", "wellness programme", "wellness program", "employee assistance programme", "employee assistance program", "eap services", "occupational wellness", "occupational health wellness"],
    "Labour": ["labour", "labor", "labour broker", "labour hire", "temporary labour", "employment agency", "recruitment agency", "staffing agency"],
    "Cleaning": ["cleaning", "cleaning services", "cleaners", "hygiene services", "deep cleaning", "sanitising", "sanitizing", "janitorial"],
    "Software / ICT / Licences": [
        "software", "software licence", "software license", "saas", "application support", "system support",
        "microsoft 365", "office 365", "azure", "azure sql", "sql server", "ict", "information technology",
        "it system", "ict system", "software solution", "software platform", "subscription", "subscription services",
        "online platform", "digital platform", "technology platform", "database subscription", "legal research platform",
        "online research platform", "research platform", "licence", "licences", "license", "licenses", "licensing"
    ],
    "Training": ["training", "skills training", "defensive driving", "driver training", "workshop facilitation", "training provider", "capacity building training"],
    "Equipment / Hardware": [
        "equipment", "machinery", "machine", "tools", "instruments", "hardware", "spares", "spare parts",
        "transformer", "distribution transformer", "circuit breaker", "substation equipment", "electrical equipment",
        "fire equipment", "meter", "meters", "pump", "pumps", "valve", "valves"
    ],
    "Appliances": ["appliance", "appliances", "white goods"],
    "Document Storage": ["document storage", "records storage", "archiving services", "offsite storage", "storage facility", "storage facilities"],
    "Blasting": ["blasting", "shot blasting", "sand blasting", "abrasive blasting"],
    "Vehicle": ["vehicle", "vehicles", "fleet", "truck", "trucks", "bakkie", "bakkies", "motor vehicle", "rental of vehicles", "cherry picker", "cherry pickers"],
    "Digital Media / Publications": [
        "digital media", "social media", "media buying", "media monitoring", "content creation", "graphic design",
        "annual report", "annual reports", "copyediting", "copy editing", "proofreading", "graphic layout", "print layout", "page layout", "printing",
        "publication", "publications", "publishing", "editorial", "interactive pdf", "corporate communications",
        "communication services", "brand identity", "corporate identity", "newsletter", "newsletters"
    ],
    "Invasive Plants / Vegetation": ["invasive plants", "invasive alien plants", "alien vegetation", "vegetation clearing", "vegetation management", "weed control", "bush clearing", "grass cutting", "tree felling", "horticulture", "invasive species"],
    "Landscaping / Gardening": ["landscaping", "gardening", "garden services", "grounds maintenance", "lawn mowing", "grass cutting"],
    "Auctioning": ["auction", "auctioning", "auctioneer", "auctioneering"],
    "Grave Preparations": ["grave preparation", "grave preparations", "grave digging", "cemetery", "burial"],
    "Animals": ["animal", "animals", "livestock", "veterinary", "pest control", "rodent control", "fauna", "game capture"],
    "Towing Services": ["towing", "towing services", "tow truck"],
    "Consumables / Supply Only": ["consumables", "canteen", "canteen consumables", "supply and delivery", "supply of goods", "supply and deliver", "supply, delivery", "supply delivery"],
    "Compliance Certificates": ["certificate of compliance", "certificates of compliance", "electrical coc", "fire coc", "coc certificate", "coc certificates", "compliance certificate"],
    "Security / Guarding": [
        "security services", "physical security", "guarding", "guarding services", "security guards",
        "access control services", "surveillance services", "armed response", "security personnel"
    ],
    "Travel / Shuttle / Staff Transport": [
        "travel management", "travel management services", "travel agency", "travel agencies",
        "shuttle services", "shuttle staff", "staff transport", "passenger transport",
        "transport services to shuttle", "transportation of staff", "car hire", "flight bookings", "accommodation bookings"
    ],
    "SAP / ERP / ICT Consulting": [
        "sap", "s/4 hana", "s4 hana", "erp", "jde", "oracle erp", "system upgrade",
        "system conversion", "functional support services", "application management services",
        "metadata automation", "digital display", "digital signage", "signage solution", "ict consulting",
        "it consulting", "business systems", "information systems", "managed services"
    ],
    "Furniture / Office / Facilities": [
        "furniture", "office furniture", "interior design", "interior layout", "office equipment",
        "office fit-out", "office fit out", "facilities management", "space planning", "workspace planning"
    ],
    "HR / Language / Medical / Admin Support": [
        "translation", "translating", "interpreting", "interpretation", "language practitioners",
        "medical surveillance", "medical testing", "occupational medical", "locum", "pharmacist", "pharmacy",
        "leadership assessment", "assessment tool", "psychometric", "board evaluation", "company secretarial",
        "secretarial services", "minutes taking", "administrative support", "hr services", "human resources services"
    ],
    "Legal / Audit / Financial Admin": [
        "legal services", "legal service", "legal panel", "legal practitioners", "attorneys", "conveyancer",
        "internal audit", "external audit", "forensic services", "forensic audit", "vat recovery", "tax recovery",
        "annual financial statements", "financial statements", "technical financial support", "expected credit loss",
        "ecl model", "audit firms", "audit firm", "accounting firm", "accounting support", "payroll"
    ],
    "Waste / Refuse / Sludge / Recycling": [
        "sludge disposal", "off-site disposal of sludge", "off site disposal of sludge", "refuse collection",
        "waste collection", "waste disposal", "wet refuse", "dry recyclables", "recyclables", "recycling",
        "landfill", "landfill site", "weighbridge", "sewer sludge", "solid waste", "waste management services"
    ],
}
REVIEW_EXCLUSION_REGEX = {
    # Standalone AI / Artificial Intelligence only. Avoid matching words that merely contain "ai".
    "AI / Artificial Intelligence": [r"\bai\b", r"\ba\.i\.\b", r"\bartificial\s+intelligence\b"],
    "Compliance Certificate COC": [r"\bcoc\b"],
    # Exclude CIDB grades/classes 1 to 6, including patterns such as CIDB 6CE,
    # Grade 4GB, CIDB grading of 5, etc. Grades 7 to 9 remain eligible.
    "CIDB below 7": [
        r"\bcidb\b.{0,50}\b[1-6]\s*(?:ce|gb|me|ep|eb|so|sq|sl|sb|sh|sm|sn|sk|sf|sg|sc|sd|se|sw|si|sa)?\b",
        r"\b(?:grade|grading|class)\s*(?:of)?\s*[1-6]\s*(?:ce|gb|me|ep|eb|so|sq|sl|sb|sh|sm|sn|sk|sf|sg|sc|sd|se|sw|si|sa)?\b.{0,50}\bcidb\b",
        r"\b[1-6]\s*(?:ce|gb|me|ep|eb|so|sq|sl|sb|sh|sm|sn|sk|sf|sg|sc|sd|se|sw|si|sa)\b",
    ],
}


# Core professional / built-environment scope indicators for bid-manager review.
# A priority client alone must not make a tender review-eligible. The daily review
# should show tenders that have at least one genuine consulting, professional,
# engineering, infrastructure planning, project/programme management or advisory
# indicator, unless future rules deliberately override this.
CORE_SCOPE_KEYWORDS = {
    "Professional Services / PSP": [
        "professional service provider", "professional services", "psp", "psp services",
        "service provider for professional", "professional engineering services",
        "built environment professional", "built environment services"
    ],
    "Consulting / Specialist Services": [
        "consultant", "consultants", "consultancy", "consulting", "consulting services",
        "specialist consultant", "specialist consultants", "specialist services", "expert panel", "panel of experts"
    ],
    "Engineering Services": [
        "engineering services", "civil engineering", "structural engineering", "mechanical engineering",
        "electrical engineering", "fire engineering", "wet services", "mep services",
        "engineering consultant", "engineering consultants", "engineers", "professional engineer"
    ],
    "Project / Programme Management": [
        "project management", "project manager", "programme management", "program management",
        "programme manager", "project management office", "pmo", "contract administration",
        "contract management", "principal agent", "employer's agent", "employers agent", "project implementation support"
    ],
    "Planning / Studies / Advisory": [
        "feasibility study", "feasibility studies", "prefeasibility", "pre-feasibility",
        "business case", "master plan", "masterplan", "infrastructure master plan",
        "implementation plan", "precinct plan", "development plan", "spatial plan", "spatial planning",
        "sector plan", "technical advisory", "technical advisor", "transaction advisory", "transaction advisor",
        "advisory services", "technical assistance", "due diligence", "options analysis", "needs assessment"
    ],
    "Design / Monitoring / Infrastructure Delivery": [
        "detailed design", "concept design", "preliminary design", "design services", "design review",
        "design and upgrade", "design and upgrades", "facility design", "facility upgrade", "facility upgrades",
        "upgrade of existing", "upgrades of existing", "production facility", "hvac system", "cgmp",
        "construction monitoring", "site supervision", "works supervision", "quality assurance services",
        "infrastructure delivery", "infrastructure programme", "infrastructure planning", "asset management plan"
    ],
    "Management Systems Previously Confirmed Relevant": [
        "performance management system", "performance management framework"
    ],
}


def _match_core_scope(row: pd.Series | Dict[str, Any]) -> tuple[bool, str]:
    """Return whether a tender has a genuine professional/built-environment scope indicator."""
    text = _review_filter_text(row)
    text_norm = normalise_text(text)
    matches: list[str] = []
    for label, phrases in CORE_SCOPE_KEYWORDS.items():
        if any(_contains_normalised_phrase(text_norm, phrase) for phrase in phrases):
            matches.append(label)
    matches = list(dict.fromkeys(matches))
    return bool(matches), "; ".join(matches)


def _review_filter_text(row: pd.Series | Dict[str, Any]) -> str:
    """Text used for exclusion matching.

    The filter intentionally considers more than description only, because some
    sites place the relevant exclusion cue in category, notice type or conditions.
    """
    fields = [
        "title",
        "description",
        "client",
        "sector",
        "category",
        "notice_type",
        "conditions",
        "raw_category",
        "raw_status",
        "briefing_session_details",
    ]
    values = []
    for field in fields:
        try:
            values.append(clean_text(row.get(field, "")))
        except AttributeError:
            values.append(clean_text(row[field] if field in row else ""))
    return " ".join(v for v in values if v)


def _contains_normalised_phrase(text_norm: str, phrase: str) -> bool:
    phrase_norm = normalise_text(phrase)
    if not phrase_norm:
        return False
    return f" {phrase_norm} " in f" {text_norm} "


def _match_review_exclusions(row: pd.Series | Dict[str, Any]) -> tuple[bool, str, str]:
    """Return (excluded, reason, matched terms) for bid-manager review filtering."""
    text = _review_filter_text(row)
    text_norm = normalise_text(text)
    matched_reasons: list[str] = []
    matched_terms: list[str] = []

    for reason, phrases in REVIEW_EXCLUSION_KEYWORDS.items():
        hits = [phrase for phrase in phrases if _contains_normalised_phrase(text_norm, phrase)]
        if hits:
            matched_reasons.append(reason)
            matched_terms.extend(hits)

    text_raw = clean_text(text).casefold()
    for reason, patterns in REVIEW_EXCLUSION_REGEX.items():
        hits = []
        for pattern in patterns:
            if re.search(pattern, text_raw, flags=re.IGNORECASE):
                hits.append(pattern)
        if hits:
            matched_reasons.append(reason)
            matched_terms.extend(hits)

    matched_reasons = list(dict.fromkeys(matched_reasons))
    matched_terms = list(dict.fromkeys(matched_terms))
    excluded = bool(matched_reasons)
    return excluded, "; ".join(matched_reasons), "; ".join(matched_terms)


def _apply_review_filter(df: pd.DataFrame) -> pd.DataFrame:
    """Mark records that should be excluded from bid-manager daily review.

    The records remain in consolidated_master/delta/updates for audit and future
    tuning. Daily review priority tabs and consolidated_top_delta use only
    review_eligible=True.
    """
    out = ensure_standard_columns(df).copy()
    if out.empty:
        out["review_eligible"] = []
        out["excluded_by_filter"] = []
        out["exclusion_reason"] = []
        out["excluded_keyword_matches"] = []
        return out

    excluded_flags = []
    reasons = []
    terms = []
    for _, row in out.iterrows():
        excluded, reason, matched = _match_review_exclusions(row)
        has_scope, scope_matches = _match_core_scope(row)

        # Scope gate: priority clients boost relevant work, but should not cause
        # operational/admin/procurement-only tenders to appear in the bid-manager
        # daily review. If no core scope indicator is present, exclude from the
        # daily review while keeping the record in the master register.
        if not has_scope:
            excluded = True
            reason = "; ".join(x for x in [reason, "No core professional/built-environment scope"] if x)
            matched = "; ".join(x for x in [matched, "No core scope keyword matched"] if x)
        elif scope_matches and matched:
            # Keep a note of scope matches in the matched term audit where useful.
            matched = "; ".join(x for x in [matched, f"Core scope: {scope_matches}"] if x)

        excluded_flags.append(bool(excluded))
        reasons.append(reason)
        terms.append(matched)

    out["excluded_by_filter"] = excluded_flags
    out["review_eligible"] = [not flag for flag in excluded_flags]
    out["exclusion_reason"] = reasons
    out["excluded_keyword_matches"] = terms
    return ensure_standard_columns(out)


def _source_key(source_name: str) -> str:
    """Normalise source labels so source-level defaults still work with punctuation changes."""
    return normalise_identifier(source_name).replace("-", "").replace("_", " ").strip()


def _apply_source_defaults(std: Dict[str, Any], source_name: str) -> Dict[str, Any]:
    """Fill stable source-specific defaults where source outputs omit or misuse fields."""
    source_norm = normalise_text(source_name)
    source_ident = _source_key(source_name)
    defaults = SOURCE_DEFAULTS.get(source_norm) or SOURCE_DEFAULTS.get(source_ident) or {}

    # World Bank rows expose country/region in country_name. The client should not
    # become the country/region; it should remain World Bank for consolidated scoring.
    if source_norm == "world bank" or source_ident == "world bank":
        std["client"] = "World Bank"

    for field, value in defaults.items():
        if field == "client" and (source_norm == "world bank" or source_ident == "world bank"):
            std[field] = value
        elif not clean_text(std.get(field, "")):
            std[field] = value

    # Keep source_url_list aligned when a fallback source_url was inserted.
    if clean_text(std.get("source_url", "")) and not clean_text(std.get("source_url_list", "")):
        std["source_url_list"] = std["source_url"]

    return std


def _recalculate_current_flags(df: pd.DataFrame) -> pd.DataFrame:
    """Recalculate current status centrally, ignoring inconsistent source flags."""
    out = ensure_standard_columns(df).copy()
    if out.empty:
        return out

    flags = []
    for _, row in out.iterrows():
        archived = _booleanish(row.get("is_archived", False))
        cancelled = _booleanish(row.get("is_cancelled", False))
        status = normalise_text(row.get("status", ""))
        if any(term in status for term in ["cancelled", "canceled", "withdrawn"]):
            cancelled = True
        if "archived" in status:
            archived = True
        closing_value = row.get("closing_datetime", None)
        if pd.isna(closing_value) if not isinstance(closing_value, str) else not closing_value.strip():
            closing_value = row.get("closing_date", None)
        flags.append(current_flag(closing_value, archived=archived, cancelled=cancelled))

    out["is_current"] = flags
    return out


def _parse_datetime_series(values: pd.Series) -> pd.Series:
    """Parse tender date values using central robust date logic.

    This handles normal strings, pandas timestamps and Excel serial dates such as
    46191.4166666667, which represent valid future dates in source workbooks.
    """
    if values is None:
        return pd.Series(dtype="datetime64[ns]")
    return parse_date_series(values)


def _coalesced_closing_series(df: pd.DataFrame) -> pd.Series:
    """Return closing_datetime where valid, otherwise closing_date, using robust parsing."""
    if df is None or df.empty:
        return pd.Series(dtype="datetime64[ns]")
    close_dt = _parse_datetime_series(df["closing_datetime"]) if "closing_datetime" in df.columns else pd.Series(pd.NaT, index=df.index)
    close_date = _parse_datetime_series(df["closing_date"]) if "closing_date" in df.columns else pd.Series(pd.NaT, index=df.index)
    return close_dt.where(close_dt.notna(), close_date)


def _has_closing_date(df: pd.DataFrame) -> pd.Series:
    if df is None or df.empty:
        return pd.Series(dtype=bool)
    return _coalesced_closing_series(df).notna()

def _align_closing_date_to_datetime(df: pd.DataFrame) -> pd.DataFrame:
    """Keep closing_date and closing_datetime consistent.

    Some source workbooks contain both fields, and one of the two may have been
    parsed before consolidation with month/day assumptions. For review purposes,
    closing_datetime is the authoritative field because it is used by current
    status and closing-soon logic. Where it exists, force closing_date to the
    date component of closing_datetime. This prevents confusing workbook rows
    such as closing_date=09-Jun-2026 and closing_datetime=06-Sep-2026.
    """
    out = ensure_standard_columns(df).copy()
    if out.empty:
        return out

    closing_dt = _parse_datetime_series(out["closing_datetime"])
    closing_date = _parse_datetime_series(out["closing_date"])

    has_dt = closing_dt.notna()
    # First normalise the datetime field itself, because some source outputs arrive
    # as text or Excel serial values. Keeping it parsed avoids later month/day drift.
    out.loc[has_dt, "closing_datetime"] = closing_dt.loc[has_dt]

    # Where a datetime exists, make the date column the date part of that datetime.
    out.loc[has_dt, "closing_date"] = closing_dt.loc[has_dt].dt.normalize()

    # Where only closing_date exists, use it as the datetime fallback. This keeps
    # filtering and duplicate sorting consistent for sources that only expose dates.
    missing_dt = closing_dt.isna() & closing_date.notna()
    out.loc[missing_dt, "closing_datetime"] = closing_date.loc[missing_dt]
    out.loc[missing_dt, "closing_date"] = closing_date.loc[missing_dt].dt.normalize()

    return out


def _truthy_series(series: pd.Series) -> pd.Series:
    return series.apply(_booleanish)


def load_json(path: str | Path) -> dict[str, Any]:
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)



def _normalise_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except Exception:
        pass
    return str(value).strip().casefold() in {"true", "yes", "y", "1", "active", "enabled"}


def _resolve_relative_path(root: Path, path_value: Any) -> Path:
    text = clean_text(path_value)
    if not text:
        return Path()
    path = Path(text.replace("\\", "/"))
    if path.is_absolute():
        return path
    return root / path


def load_sources_registry(registry_path: str | Path) -> pd.DataFrame:
    registry_path = Path(registry_path)
    if not registry_path.exists():
        raise FileNotFoundError(f"Source registry not found: {registry_path}")
    df = pd.read_excel(registry_path, sheet_name="Sources_Registry", engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    required = {"source_name", "source_key", "active", "folder_name"}
    missing = sorted(required - set(df.columns))
    if missing:
        raise ValueError(f"Source registry is missing required column(s): {', '.join(missing)}")
    df = df[df["source_name"].notna()].copy()
    df["active_bool"] = df["active"].apply(_normalise_bool)
    return df


def sources_config_from_registry(registry_path: str | Path, project_root: str | Path) -> dict[str, Any]:
    """Build the consolidation source config directly from 00_Config/sources_registry.xlsx.

    This replaces the legacy sources_config.json dependency and uses the actual
    restructured source output paths from the registry.
    """
    project_root = Path(project_root)
    registry = load_sources_registry(registry_path)
    sources: list[dict[str, Any]] = []

    output_columns = {
        "master": "master_output_path",
        "delta": "delta_output_path",
        "updates": "updates_output_path",
        "top_delta": "top_delta_output_path",
    }

    for _, row in registry.iterrows():
        file_paths: dict[str, str] = {}
        files: dict[str, str] = {}
        for key, column in output_columns.items():
            value = clean_text(row.get(column, ""))
            if value:
                file_paths[key] = value
                files[key] = Path(value.replace("\\", "/")).name

        folder_name = clean_text(row.get("folder_name", ""))
        sources.append(
            {
                "name": clean_text(row.get("source_name", "")),
                "source_key": clean_text(row.get("source_key", "")),
                "enabled": bool(row.get("active_bool", False)),
                "folder": f"02_Sources/{folder_name}" if folder_name else "",
                "outputs_folder": "02_Outputs",
                "files": files,
                "file_paths": file_paths,
                "notes": clean_text(row.get("notes", "")),
            }
        )

    return {
        "data_analytics_root_default": str(project_root),
        "consolidated_output_folder": "04_FinalOutputs",
        "sources": sources,
    }


def read_excel_file(path: str | Path) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path, engine="openpyxl")
    except Exception:
        # Some source files have specific sheet names; fallback to first sheet.
        return pd.read_excel(path, sheet_name=0, engine="openpyxl")


def _get(row: pd.Series, *names: str) -> str:
    for name in names:
        if name in row.index:
            text = clean_text(row.get(name, ""))
            if text:
                return text
    return ""


def _booleanish(value: Any) -> bool:
    text = normalise_text(value)
    return text in {"true", "yes", "y", "1", "archived", "cancelled", "canceled", "withdrawn"}


def _document_count(row: pd.Series) -> int:
    existing = _get(row, "document_count")
    if existing:
        try:
            return int(float(existing))
        except Exception:
            pass
    urls = _get(row, "document_urls", "document_url", "primary_document_url", "additional_document_urls")
    if not urls:
        return 0
    parts = [p for p in str(urls).replace("|", ";").split(";") if p.strip()]
    return len(parts) if parts else 1


def _split_multi_values(value: Any) -> list[str]:
    """Split semicolon/pipe separated URL or label fields into unique values."""
    text = clean_text(value)
    if not text:
        return []
    values: list[str] = []
    for part in str(text).replace("|", ";").split(";"):
        part = clean_text(part)
        if part and part not in values:
            values.append(part)
    return values


def _first_multi_value(value: Any) -> str:
    values = _split_multi_values(value)
    return values[0] if values else ""


def _resolve_source_url(row: pd.Series, source_name: str) -> str:
    """Resolve the best source/detail URL from source output fields.

    For uMngeni-uThukela, prefer the row-level tender detail link over the
    generic portal URL. For other sources, keep explicit source_url first and
    then fall back to detail/view links if present.
    """
    source_norm = normalise_text(source_name)
    source_ident = _source_key(source_name)
    if "umngeni" in source_norm or "umngeni" in source_ident:
        return first_non_empty(
            _get(row, "view_url"),
            _get(row, "title_url"),
            _get(row, "source_url"),
            _get(row, "url"),
            _get(row, "scraped_url"),
        )
    return first_non_empty(
        _get(row, "source_url"),
        _get(row, "url"),
        _get(row, "scraped_url"),
        _get(row, "view_url"),
        _get(row, "title_url"),
    )


def _resolve_document_urls(row: pd.Series) -> str:
    return _get(row, "document_urls", "document_url", "primary_document_url", "additional_document_urls")


def _apply_link_fields(std: Dict[str, Any], source_name: str) -> Dict[str, Any]:
    """Populate action_url and document_access_status consistently.

    eTenders is intentionally treated as a discovery feed: where no direct
    document URL is available, the action link is the eTenders portal and the
    status tells the reviewer that manual lookup is required.
    """
    source_norm = normalise_text(source_name)
    docs = clean_text(std.get("document_urls", ""))
    src_url = clean_text(std.get("source_url", ""))
    src_url_list = clean_text(std.get("source_url_list", ""))

    if docs:
        std["action_url"] = _first_multi_value(docs)
        std["document_access_status"] = "Direct document available"
    elif source_norm == "etenders":
        std["action_url"] = _first_multi_value(src_url or src_url_list)
        std["document_access_status"] = "Manual eTenders lookup required"
    elif src_url or src_url_list:
        std["action_url"] = _first_multi_value(src_url or src_url_list)
        std["document_access_status"] = "Source/detail page only"
    else:
        std["action_url"] = ""
        std["document_access_status"] = "No link captured"

    return std

def _apply_merged_link_fields(base: Dict[str, Any], group: pd.DataFrame) -> Dict[str, Any]:
    """Resolve the best reviewer action link after duplicate merging.

    If a tender appears in both eTenders and a source portal, prefer the source
    portal's direct document/detail link. eTenders remains a discovery feed and
    should only be the action link when no better source exists.
    """
    docs = join_unique(group["document_urls"].tolist()) if "document_urls" in group.columns else clean_text(base.get("document_urls", ""))
    base["document_urls"] = docs

    if docs:
        base["action_url"] = _first_multi_value(docs)
        base["document_access_status"] = "Direct document available"
        return base

    non_etender_urls: list[str] = []
    etender_urls: list[str] = []
    other_urls: list[str] = []

    for _, row in group.iterrows():
        src = normalise_text(row.get("source", ""))
        candidates = []
        candidates.extend(_split_multi_values(row.get("source_url", "")))
        candidates.extend(_split_multi_values(row.get("source_url_list", "")))
        for url in candidates:
            if not url:
                continue
            if src == "etenders":
                if url not in etender_urls:
                    etender_urls.append(url)
            elif src:
                if url not in non_etender_urls:
                    non_etender_urls.append(url)
            else:
                if url not in other_urls:
                    other_urls.append(url)

    if non_etender_urls:
        base["action_url"] = non_etender_urls[0]
        base["document_access_status"] = "Source/detail page only"
    elif other_urls:
        base["action_url"] = other_urls[0]
        base["document_access_status"] = "Source/detail page only"
    elif etender_urls:
        base["action_url"] = etender_urls[0]
        base["document_access_status"] = "Manual eTenders lookup required"
    else:
        base["action_url"] = ""
        base["document_access_status"] = "No link captured"

    return base


def _build_opportunity_key(source_name: str, row: pd.Series) -> str:
    existing = _get(row, "opportunity_key")
    if existing:
        return existing
    ident = first_non_empty(
        _get(row, "tender_number"),
        _get(row, "bid_number"),
        _get(row, "reference_number"),
        _get(row, "notice_id"),
        _get(row, "id"),
        _get(row, "project_id"),
    )
    if not ident:
        payload = "||".join([
            _get(row, "description", "bid_description", "title"),
            _get(row, "closing_date", "closing_datetime", "deadline_date"),
            source_name,
        ])
        ident = hashlib.sha1(payload.encode("utf-8")).hexdigest()[:16]
    return f"{normalise_identifier(source_name) or 'source'}:{normalise_identifier(ident) or ident}"


def _compute_hash(row: pd.Series) -> str:
    payload = []
    for col in HASH_COLUMNS:
        value = row.get(col, "")
        if col in DATE_COLUMNS:
            dt = parse_date(value)
            payload.append("" if pd.isna(dt) else dt.strftime("%Y-%m-%d %H:%M:%S"))
        else:
            payload.append(clean_text(value))
    return hashlib.sha256("||".join(payload).encode("utf-8")).hexdigest()


def standardise_frame(df: pd.DataFrame, source_name: str, output_type: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=STANDARD_COLUMNS)

    rows: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        source_original_weight = _get(row, "weight_score")
        title = _get(row, "title", "Name", "name")
        description_original = _get(row, "description_original", "bid_description_original", "bid_description", "description", "Tender Name", "tender_name")
        description_en = _get(row, "description_en", "bid_description_en")
        # Use English translation for consolidated review/scoring where available,
        # while retaining the original source text in description_original.
        description = first_non_empty(description_en, _get(row, "description", "bid_description", "Tender Name", "tender_name"), description_original)
        if not title:
            title = description[:180] if description else ""

        is_archived = _booleanish(_get(row, "is_archived", "archived_flag"))
        is_cancelled = _booleanish(_get(row, "is_cancelled", "cancellation_flag"))
        status_text = _get(row, "status", "raw_status", "Tender Status")
        if any(term in normalise_text(status_text) for term in ["cancelled", "canceled", "withdrawn"]):
            is_cancelled = True
        if "archived" in normalise_text(status_text):
            is_archived = True

        closing_datetime = _get(row, "closing_datetime", "deadline_date", "Closing Date", "closing_date")
        closing_date = _get(row, "closing_date", "deadline_date", "Closing Date") or closing_datetime

        std = {
            "source": source_name,
            "source_output_type": output_type,
            "source_url": _resolve_source_url(row, source_name),
            "source_list": source_name,
            "source_url_list": _resolve_source_url(row, source_name),
            "action_url": "",
            "document_access_status": "",
            "duplicate_group_id": "",
            "duplicate_count": 1,
            "opportunity_key": _build_opportunity_key(source_name, row),
            "tender_number": _get(row, "tender_number", "bid_number", "tender_No"),
            "reference_number": _get(row, "reference_number", "Reference", "ref", "project_id", "id"),
            "notice_type": _get(row, "notice_type", "type", "procurement_method"),
            "title": title,
            "description": description,
            "description_original": description_original,
            "description_en": description_en,
            "detected_language": _get(row, "detected_language"),
            "translation_status": _get(row, "translation_status"),
            "translation_provider": _get(row, "translation_provider"),
            "client": _get(row, "client", "organ_of_State"),
            "sector": _get(row, "sector"),
            "category": _get(row, "category", "procurement_category", "TCategory"),
            "country_name": _get(row, "country_name"),
            "province": _get(row, "province"),
            "region": _get(row, "region"),
            "date_published": _get(row, "date_published", "publication_date", "posted_date", "Date Published"),
            "issue_date": _get(row, "issue_date", "issued_date", "date_uploaded", "Issue Date"),
            "briefing_date": _get(row, "briefing_date", "Briefing Date"),
            "closing_date": closing_date,
            "closing_datetime": closing_datetime,
            "closing_datetime_source": "",
            "date_resolution_notes": "",
            "status": status_text,
            "is_current": current_flag(closing_datetime or closing_date, archived=is_archived, cancelled=is_cancelled),
            "is_archived": is_archived,
            "is_cancelled": is_cancelled,
            "document_names": _get(row, "document_names", "primary_document_name", "additional_document_names"),
            "document_urls": _resolve_document_urls(row),
            "document_count": _document_count(row),
            "contact_person": _get(row, "contact_person", "contactPerson"),
            "contact_email": _get(row, "contact_email", "email"),
            "contact_telephone": _get(row, "contact_telephone", "telephone"),
            "submission_method": _get(row, "submission_method"),
            "briefing_session_details": _get(row, "briefing_session_details", "briefingVenue", "briefing_venue"),
            "conditions": _get(row, "conditions"),
            "raw_category": _get(row, "raw_category", "category", "procurement_category"),
            "raw_status": _get(row, "raw_status", "status"),
            "run_timestamp": _get(row, "run_timestamp"),
            "row_hash": "",
            "change_type": _get(row, "change_type"),
            "previous_row_hash": _get(row, "previous_row_hash"),
            "priority_client_match": False,
            "priority_client_matches": "",
            "sector_keyword_matches": "",
            "technical_keyword_matches": "",
            "service_keyword_matches": "",
            "procurement_keyword_matches": "",
            "strategic_keyword_matches": "",
            "matched_keywords_all": "",
            "weight_score": 0,
            "score_band": "",
            "scoring_notes": "",
            "source_original_weight_score": source_original_weight,
            "review_eligible": True,
            "excluded_by_filter": False,
            "exclusion_reason": "",
            "excluded_keyword_matches": "",
        }
        std = _apply_source_defaults(std, source_name)
        std = _apply_link_fields(std, source_name)
        rows.append(std)

    out = pd.DataFrame(rows)
    out = coerce_date_columns(out, DATE_COLUMNS)
    out = _align_closing_date_to_datetime(out)
    out = _recalculate_current_flags(out)
    out["row_hash"] = out.apply(_compute_hash, axis=1) if not out.empty else []
    return ensure_standard_columns(out)


def ensure_standard_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy() if df is not None else pd.DataFrame()
    for col in STANDARD_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    return out[STANDARD_COLUMNS]


def load_source_outputs(config: dict[str, Any], data_root: str | Path) -> dict[str, pd.DataFrame]:
    data_root = Path(data_root)
    combined = {"master": [], "delta": [], "updates": [], "top_delta": []}

    for source in config.get("sources", []):
        if not source.get("enabled", True):
            continue
        source_name = source["name"]
        output_keys = set((source.get("files", {}) or {}).keys()) | set((source.get("file_paths", {}) or {}).keys())
        for output_type in sorted(output_keys):
            if output_type not in combined:
                continue
            path = _source_output_path(source, data_root, output_type)
            if not path.exists():
                continue
            raw = read_excel_file(path)
            std = standardise_frame(raw, source_name, output_type)
            if not std.empty:
                combined.setdefault(output_type, []).append(std)

    return {
        key: ensure_standard_columns(pd.concat(frames, ignore_index=True)) if frames else pd.DataFrame(columns=STANDARD_COLUMNS)
        for key, frames in combined.items()
    }


def _exact_duplicate_key(row: pd.Series) -> str:
    ref = first_non_empty(row.get("tender_number", ""), row.get("reference_number", ""))
    ref_norm = normalise_identifier(ref)
    if ref_norm:
        return f"ref:{ref_norm}"

    # Direct document URLs are stable enough to merge duplicated portal rows,
    # especially for ACSA where the same tender can appear twice with one
    # ambiguous date parsed differently.
    doc_url = _first_multi_value(row.get("document_urls", ""))
    doc_norm = normalise_identifier(doc_url)
    if doc_norm:
        return f"doc:{doc_norm}"

    # Fallback avoids unsafe merging when no identifier exists.
    return f"opp:{normalise_identifier(row.get('opportunity_key', ''))}"



def _valid_datetime(value: Any) -> pd.Timestamp | pd.NaT:
    """Parse a date/datetime value and return NaT when it is not usable."""
    dt = parse_date(value)
    if pd.isna(dt):
        return pd.NaT
    return dt


def _choose_best_closing_datetime(group: pd.DataFrame) -> tuple[Any, str, str]:
    """Choose the best closing date from any duplicate source row.

    Duplicate groups may merge eTenders with a source portal such as TCTA. In
    that case, the primary row may have a blank eTenders date while the source
    portal row has the correct closing datetime. The merged row must use the
    best available date from the group, not blindly keep the primary row.

    Rule:
    - parse closing_datetime first, falling back to closing_date;
    - prefer the earliest future/current datetime where available;
    - if all dates are in the past, use the latest parsed date so audit output
      still retains the best available source date;
    - if all dates are missing, return blank.
    """
    candidates: list[tuple[pd.Timestamp, str]] = []
    for _, row in group.iterrows():
        src = clean_text(row.get("source", ""))
        dt = _valid_datetime(row.get("closing_datetime", ""))
        if pd.isna(dt):
            dt = _valid_datetime(row.get("closing_date", ""))
        if pd.notna(dt):
            candidates.append((dt, src))

    if not candidates:
        return "", "", "No valid closing date found in duplicate group"

    today = pd.Timestamp.now().normalize()
    future = [(dt, src) for dt, src in candidates if dt.normalize() >= today]
    if future:
        chosen_dt, chosen_src = sorted(future, key=lambda x: x[0])[0]
        note = f"Closing datetime selected from {chosen_src or 'source row'}; earliest current/future date in duplicate group"
    else:
        chosen_dt, chosen_src = sorted(candidates, key=lambda x: x[0])[-1]
        note = f"Closing datetime selected from {chosen_src or 'source row'}; all duplicate dates appear expired"

    return chosen_dt, chosen_src, note


def _first_non_empty_from_group(group: pd.DataFrame, field: str) -> str:
    """First non-empty field value in a duplicate group, preferring non-eTenders rows."""
    if field not in group.columns:
        return ""
    ordered = group.copy()
    if "source" in ordered.columns:
        ordered["_source_is_etenders"] = ordered["source"].apply(lambda x: normalise_text(x) == "etenders")
        ordered = ordered.sort_values("_source_is_etenders", ascending=True)
    for _, row in ordered.iterrows():
        value = clean_text(row.get(field, ""))
        if value:
            return value
    return ""


def _fill_merged_scalar_fields(base: pd.Series, group: pd.DataFrame) -> pd.Series:
    """Fill missing fields on a merged duplicate row from other source rows.

    This avoids losing useful source-portal information when eTenders is chosen
    as the primary row during duplicate sorting.
    """
    for field in [
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
        "status",
        "contact_telephone",
        "submission_method",
        "briefing_session_details",
        "conditions",
        "raw_category",
        "raw_status",
    ]:
        if not clean_text(base.get(field, "")):
            value = _first_non_empty_from_group(group, field)
            if value:
                base[field] = value
    return base


def _apply_merged_date_fields(base: pd.Series, group: pd.DataFrame) -> pd.Series:
    """Apply field-level date coalescing after duplicate grouping."""
    chosen_dt, chosen_src, note = _choose_best_closing_datetime(group)
    if clean_text(chosen_dt):
        dt = parse_date(chosen_dt)
        if pd.notna(dt):
            base["closing_datetime"] = dt
            base["closing_date"] = dt.normalize()
            base["closing_datetime_source"] = chosen_src
            base["date_resolution_notes"] = note
    else:
        base["closing_datetime"] = ""
        base["closing_date"] = ""
        base["closing_datetime_source"] = ""
        base["date_resolution_notes"] = note
    return base

def dedupe_records(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return ensure_standard_columns(df)

    work = ensure_standard_columns(df).copy()
    work["_dup_key"] = work.apply(_exact_duplicate_key, axis=1)
    work["_sort_closing"] = _coalesced_closing_series(work)
    work["_sort_score"] = pd.to_numeric(work["weight_score"], errors="coerce").fillna(0)

    rows = []
    for dup_key, group in work.groupby("_dup_key", dropna=False):
        group = group.sort_values(["_sort_score", "_sort_closing"], ascending=[False, True], na_position="last")
        base = group.iloc[0].copy()
        base["duplicate_group_id"] = hashlib.sha1(str(dup_key).encode("utf-8")).hexdigest()[:12]
        base["duplicate_count"] = len(group)
        base["source_list"] = join_unique(group["source"].tolist())
        base["source_url_list"] = join_unique(group["source_url"].tolist())
        base["document_names"] = join_unique(group["document_names"].tolist())
        base["document_urls"] = join_unique(group["document_urls"].tolist())
        base["contact_email"] = join_unique(group["contact_email"].tolist())
        base["contact_person"] = join_unique(group["contact_person"].tolist())
        base = _fill_merged_scalar_fields(base, group)
        base = _apply_merged_date_fields(base, group)
        base["document_count"] = len(_split_multi_values(base.get("document_urls", "")))
        base_dict = _apply_merged_link_fields(base.to_dict(), group)
        base = pd.Series(base_dict)
        rows.append(base.drop(labels=["_dup_key", "_sort_closing", "_sort_score", "_source_is_etenders"], errors="ignore"))

    out = pd.DataFrame(rows)
    out = ensure_standard_columns(out)
    out = _align_closing_date_to_datetime(out)
    out = _recalculate_current_flags(out)
    return sort_for_review(out)


def sort_for_review(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return ensure_standard_columns(df)
    out = _align_closing_date_to_datetime(ensure_standard_columns(df)).copy()
    out["_sort_score"] = pd.to_numeric(out["weight_score"], errors="coerce").fillna(0)
    out["_sort_closing"] = _coalesced_closing_series(out)
    out = out.sort_values(["_sort_score", "_sort_closing", "source"], ascending=[False, True, True], na_position="last")
    return out.drop(columns=["_sort_score", "_sort_closing"], errors="ignore")



# Columns shown in the bid-manager daily review sheets. The full technical/audit
# schema is still preserved in consolidated_master.xlsx, consolidated_delta.xlsx
# and consolidated_updates.xlsx. These display columns keep the daily review as
# an action list rather than a raw data dump.
DAILY_REVIEW_COLUMNS = [
    "priority_rank",
    "score_band",
    "weight_score",
    "source",
    "source_list",
    "tender_reference",
    "client",
    "description",
    "location",
    "closing_date_time",
    "days_to_close",
    "document_access_status",
    "action_url",
    "matched_keywords_all",
    "scoring_notes",
    "change_type",
    "notice_type",
    "date_published",
    "document_count",
    "contact_email",
]

SA_SOURCE_NAMES = {
    "etenders",
    "city of cape town",
    "dbsa",
    "vcwb",
    "vaal central water",
    "transnet",
    "acsa",
    "coega",
    "dhs",
    "gtac",
    "government technical advisory centre",
    "necsa",
    "south african nuclear energy corporation",
    "nnr",
    "national nuclear regulator",
    "umngeni uthukela water",
    "umngeniuthukelawater",
    "rand water",
    "randwater",
    "tcta",
    "trans caledon tunnel authority",
    "city of ekurhuleni",
    "johannesburg water",
    "city of mbombela",
    "city of polokwane",
}


def _format_display_datetime(value: Any) -> str:
    """Readable date/date-time for the daily review workbook."""
    dt = parse_date(value)
    if pd.isna(dt):
        return ""
    if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
        return dt.strftime("%Y-%m-%d")
    return dt.strftime("%Y-%m-%d %H:%M")


def _display_tender_reference(row: pd.Series) -> str:
    return first_non_empty(
        row.get("tender_number", ""),
        row.get("reference_number", ""),
        row.get("opportunity_key", ""),
    )


def _display_location(row: pd.Series) -> str:
    """Create one practical location column for the daily review.

    For South African sources, province is usually more useful than country.
    If province exists, show "Province, South Africa". If province is missing
    but the source is known to be South African, show "South Africa".
    For World Bank and similar sources, use country_name first, then region.
    """
    province = clean_text(row.get("province", ""))
    country = clean_text(row.get("country_name", ""))
    region = clean_text(row.get("region", ""))
    source_norm = normalise_text(row.get("source", ""))

    if province:
        return f"{province}, South Africa"
    if country:
        return country
    if region:
        return region
    if source_norm in SA_SOURCE_NAMES or _source_key(source_norm) in SA_SOURCE_NAMES:
        return "South Africa"
    return ""


def _prepare_daily_review_display(df: pd.DataFrame) -> pd.DataFrame:
    """Reduce a standardised tender dataframe to bid-manager review columns.

    This function is intentionally used only for consolidated_daily_review.xlsx.
    It does not remove fields from the consolidated master/delta/update files.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=DAILY_REVIEW_COLUMNS)

    out = _align_closing_date_to_datetime(ensure_standard_columns(df)).copy()
    out = sort_for_review(out)

    out["priority_rank"] = range(1, len(out) + 1)
    out["tender_reference"] = out.apply(_display_tender_reference, axis=1)

    # If title and description differ in future, title remains available in the
    # master, but the daily review uses one description field only.
    out["description"] = out.apply(
        lambda r: clean_text(r.get("description", "")) or clean_text(r.get("title", "")),
        axis=1,
    )
    out["location"] = out.apply(_display_location, axis=1)

    closing_dt = _coalesced_closing_series(out)
    closing_values = closing_dt
    out["closing_date_time"] = closing_values.apply(_format_display_datetime)
    out["days_to_close"] = (closing_dt.dt.normalize() - pd.Timestamp.now().normalize()).dt.days.astype("object")
    out.loc[closing_dt.isna(), "days_to_close"] = ""

    for col in DAILY_REVIEW_COLUMNS:
        if col not in out.columns:
            out[col] = ""

    return out[DAILY_REVIEW_COLUMNS].copy()



def _configured_sources(sources_config: dict[str, Any] | None) -> list[dict[str, Any]]:
    if not sources_config:
        return []
    return list(sources_config.get("sources", []) or [])


def _source_aliases(name: str) -> set[str]:
    """Create a loose set of source labels used for matching source/source_list values."""
    raw = clean_text(name)
    norm = normalise_text(raw)
    ident = _source_key(raw)
    aliases = {raw, norm, ident}
    # Common aliases already present in source outputs.
    if "vaal central" in norm or norm == "vcwb":
        aliases.update({"vcwb", "vaal central water", "vaal central water board"})
    if "umngeni" in norm:
        aliases.update({"umngeni uthukela water", "umngeniuthukelawater", "umngeni-uthukela water", "umngeni water"})
    if "city of cape town" in norm:
        aliases.update({"city of cape town", "coct"})
    if norm == "nnr" or "national nuclear regulator" in norm:
        aliases.update({"nnr", "national nuclear regulator"})
    if norm == "necsa" or "south african nuclear energy" in norm:
        aliases.update({"necsa", "south african nuclear energy corporation"})
    if norm == "tcta" or "trans caledon" in norm:
        aliases.update({"tcta", "trans-caledon tunnel authority", "trans caledon tunnel authority"})
    if norm == "rand water" or norm == "randwater":
        aliases.update({"rand water", "randwater"})
    if "ekurhuleni" in norm:
        aliases.update({"city of ekurhuleni", "ekurhuleni"})
    if "johannesburg water" in norm or "joburg water" in norm:
        aliases.update({"johannesburg water", "joburg water"})
    if "mbombela" in norm:
        aliases.update({"city of mbombela", "mbombela"})
    if "polokwane" in norm:
        aliases.update({"city of polokwane", "polokwane"})
    if "namwater" in norm or "nam water" in norm:
        aliases.update({"namwater", "nam water"})
    return {normalise_text(a) for a in aliases if clean_text(a)}


def _source_names_for_row(row: pd.Series) -> list[str]:
    values = _split_multi_values(row.get("source_list", ""))
    if not values:
        values = _split_multi_values(row.get("source", ""))
    if not values and clean_text(row.get("source", "")):
        values = [clean_text(row.get("source", ""))]
    return values


def _expanded_source_counts(df: pd.DataFrame, source_names: list[str], mask: pd.Series | None = None) -> dict[str, int]:
    """Count rows per source using expanded source_list values.

    A duplicated/consolidated row may have several sources in source_list. This
    function counts coverage for each listed source. This is intentionally different
    from the primary-source count used in older versions of the daily review.
    """
    counts = {name: 0 for name in source_names}
    if df is None or df.empty:
        return counts
    work = df.copy()
    if mask is not None and len(mask) == len(work):
        work = work[mask].copy()
    alias_map = {name: _source_aliases(name) for name in source_names}
    for _, row in work.iterrows():
        row_sources = {normalise_text(s) for s in _source_names_for_row(row)}
        if not row_sources and clean_text(row.get("source", "")):
            row_sources = {normalise_text(row.get("source", ""))}
        matched_any = False
        for name, aliases in alias_map.items():
            if row_sources & aliases:
                counts[name] += 1
                matched_any = True
        # If a source was not configured under the exact label, it is deliberately
        # omitted from the all-source summary but will still be visible in the
        # source-coverage detail table below.
    return counts


def _safe_excel_row_count(path: str | Path) -> int | str:
    if not os.path.exists(path):
        return "Missing"
    try:
        return len(read_excel_file(path))
    except Exception:
        return "Unreadable"


def _safe_file_modified(path: str | Path) -> str:
    if not os.path.exists(path):
        return ""
    try:
        return pd.to_datetime(os.path.getmtime(path), unit="s").strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return ""


def _source_output_path(src: dict[str, Any], data_root: Path, file_key: str) -> Path:
    file_paths = src.get("file_paths", {}) or {}
    direct = file_paths.get(file_key, "")
    if clean_text(direct):
        return _resolve_relative_path(data_root, direct)
    return data_root / src.get("folder", "") / src.get("outputs_folder", "Outputs") / src.get("files", {}).get(file_key, "")


def _build_by_source_summary(
    sources_config: dict[str, Any] | None,
    data_root: Path | None,
    master: pd.DataFrame,
    delta: pd.DataFrame,
    updates: pd.DataFrame,
    active_master: pd.DataFrame,
    missing_closing: pd.DataFrame,
    closing_soon: pd.DataFrame,
) -> pd.DataFrame:
    """All configured sources summary for the daily review workbook.

    This replaces the older By Source tab, which only counted sources represented
    in the current review list. The new table answers a more operational question:
    which scrapers are configured, did their output files exist, and how many rows
    ended up in the review after filtering?
    """
    sources = _configured_sources(sources_config)
    if not sources:
        return pd.DataFrame()
    data_root = Path(data_root or ".")
    names = [clean_text(src.get("name", "")) for src in sources]

    current_mask = _truthy_series(master["is_current"]) if master is not None and not master.empty else pd.Series(dtype=bool)
    eligible_mask = _truthy_series(master["review_eligible"]) if master is not None and not master.empty else pd.Series(dtype=bool)
    excluded_mask = _truthy_series(master["excluded_by_filter"]) if master is not None and not master.empty else pd.Series(dtype=bool)

    all_master_counts = _expanded_source_counts(master, names)
    current_master_counts = _expanded_source_counts(master, names, current_mask) if len(current_mask) else {n: 0 for n in names}
    eligible_current_counts = _expanded_source_counts(master, names, current_mask & eligible_mask) if len(current_mask) else {n: 0 for n in names}
    excluded_counts = _expanded_source_counts(master, names, excluded_mask) if len(excluded_mask) else {n: 0 for n in names}
    daily_review_counts = _expanded_source_counts(active_master, names)
    manual_date_counts = _expanded_source_counts(missing_closing, names)
    closing_soon_counts = _expanded_source_counts(closing_soon, names)
    new_today_counts = _expanded_source_counts(delta, names, _truthy_series(delta["is_current"]) & _truthy_series(delta["review_eligible"])) if delta is not None and not delta.empty else {n: 0 for n in names}
    updated_today_counts = _expanded_source_counts(updates, names, _truthy_series(updates["is_current"]) & _truthy_series(updates["review_eligible"])) if updates is not None and not updates.empty else {n: 0 for n in names}

    rows = []
    for src in sources:
        name = clean_text(src.get("name", ""))
        enabled = bool(src.get("enabled", True))
        master_path = _source_output_path(src, data_root, "master")
        delta_path = _source_output_path(src, data_root, "delta")
        updates_path = _source_output_path(src, data_root, "updates")
        top_delta_path = _source_output_path(src, data_root, "top_delta")
        file_status_parts = []
        for label, path in [("master", master_path), ("delta", delta_path), ("updates", updates_path), ("top_delta", top_delta_path)]:
            if not os.path.exists(path):
                file_status_parts.append(f"Missing {label}")
        output_status = "Disabled" if not enabled else ("OK" if not file_status_parts else "; ".join(file_status_parts))
        note = clean_text(src.get("note", "")) or clean_text(src.get("notes", ""))
        if enabled and daily_review_counts.get(name, 0) == 0:
            if all_master_counts.get(name, 0) == 0:
                note = first_non_empty(note, "No source rows in current master/output.")
            elif eligible_current_counts.get(name, 0) == 0:
                note = first_non_empty(note, "Rows scraped but none currently pass daily-review eligibility/current-date filters.")
            else:
                note = first_non_empty(note, "Rows available but none fell into the current review tabs.")
        rows.append({
            "source": name,
            "enabled": enabled,
            "output_status": output_status,
            "source_master_rows": _safe_excel_row_count(master_path),
            "source_delta_rows": _safe_excel_row_count(delta_path),
            "source_updates_rows": _safe_excel_row_count(updates_path),
            "source_top_delta_rows": _safe_excel_row_count(top_delta_path),
            "consolidated_master_rows": all_master_counts.get(name, 0),
            "current_rows": current_master_counts.get(name, 0),
            "review_eligible_current_rows": eligible_current_counts.get(name, 0),
            "excluded_rows": excluded_counts.get(name, 0),
            "daily_review_rows": daily_review_counts.get(name, 0),
            "closing_soon_rows": closing_soon_counts.get(name, 0),
            "new_today_rows": new_today_counts.get(name, 0),
            "updated_today_rows": updated_today_counts.get(name, 0),
            "manual_date_check_rows": manual_date_counts.get(name, 0),
            "latest_master_modified": _safe_file_modified(master_path),
            "folder": src.get("folder", ""),
            "notes": note,
        })
    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values(["enabled", "daily_review_rows", "review_eligible_current_rows", "source"], ascending=[False, False, False, True])
    return out


def _build_source_coverage_detail(active_master: pd.DataFrame) -> pd.DataFrame:
    """Sources represented in the bid-manager review list using expanded source_list.

    This is intentionally a review-coverage view, not a run-status view. For all
    configured sources and zero-row sources, use the By Source / Scraper Run Summary tabs.
    """
    if active_master is None or active_master.empty:
        return pd.DataFrame(columns=[
            "source", "review_rows", "primary_source_rows", "duplicate_coverage_rows",
            "avg_score", "max_score", "very_high_rows", "high_rows", "medium_rows",
            "direct_document_rows", "source_detail_only_rows", "manual_lookup_rows"
        ])
    rows = []
    for _, row in active_master.iterrows():
        sources = _source_names_for_row(row)
        if not sources:
            sources = [clean_text(row.get("source", "")) or "Unknown"]
        primary = clean_text(row.get("source", "")) or "Unknown"
        for src in sources:
            rows.append({
                "source": src,
                "primary_source_match": normalise_text(src) == normalise_text(primary),
                "opportunity_key": row.get("opportunity_key", ""),
                "weight_score": pd.to_numeric(row.get("weight_score", 0), errors="coerce"),
                "score_band": row.get("score_band", ""),
                "document_access_status": row.get("document_access_status", ""),
            })
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame()
    grouped = df.groupby("source", dropna=False).agg(
        review_rows=("opportunity_key", "count"),
        primary_source_rows=("primary_source_match", "sum"),
        avg_score=("weight_score", "mean"),
        max_score=("weight_score", "max"),
    ).reset_index()
    band_counts = pd.crosstab(df["source"], df["score_band"]).reset_index()
    status_counts = pd.crosstab(df["source"], df["document_access_status"]).reset_index()
    out = grouped.merge(band_counts, on="source", how="left").merge(status_counts, on="source", how="left")
    # Friendly stable column names even if a band/status is absent.
    rename_map = {
        "Very High": "very_high_rows",
        "High": "high_rows",
        "Medium": "medium_rows",
        "Direct document available": "direct_document_rows",
        "Source/detail page only": "source_detail_only_rows",
        "Manual eTenders lookup required": "manual_lookup_rows",
    }
    out = out.rename(columns=rename_map)
    for col in ["very_high_rows", "high_rows", "medium_rows", "direct_document_rows", "source_detail_only_rows", "manual_lookup_rows"]:
        if col not in out.columns:
            out[col] = 0
    out["duplicate_coverage_rows"] = out["review_rows"] - out["primary_source_rows"]
    cols = [
        "source", "review_rows", "primary_source_rows", "duplicate_coverage_rows",
        "avg_score", "max_score", "very_high_rows", "high_rows", "medium_rows",
        "direct_document_rows", "source_detail_only_rows", "manual_lookup_rows",
    ]
    return out[cols].sort_values(["max_score", "review_rows"], ascending=[False, False])



def _build_daily_review_qa(sheets: dict[str, pd.DataFrame], master: pd.DataFrame) -> pd.DataFrame:
    """Small operational QA table for the generated daily review workbook."""
    rows: list[dict[str, Any]] = []
    review_sheet_names = [
        "Very High Priority",
        "High Priority",
        "Medium Priority",
        "Closing Soon",
        "New Today",
        "Updated Today",
        "Needs Manual Date Check",
    ]
    total_review_rows = sum(len(sheets.get(name, pd.DataFrame())) for name in review_sheet_names)
    rows.append({"check": "daily_review_rows_total", "status": "INFO", "value": total_review_rows, "notes": "Total rows across operational review tabs."})

    manual_df = sheets.get("Needs Manual Date Check", pd.DataFrame())
    rows.append({"check": "manual_date_check_rows", "status": "INFO", "value": len(manual_df), "notes": "Should contain only review-eligible rows with no reliable closing date."})

    if master is not None and not master.empty:
        eligible = _truthy_series(master["review_eligible"])
        excluded = _truthy_series(master["excluded_by_filter"])
        has_close = _has_closing_date(master)
        current = _truthy_series(master["is_current"])
        expected_priority_pool = int((eligible & ~excluded & has_close & current).sum())
        expected_manual_pool = int((eligible & ~excluded & ~has_close).sum())
        excluded_current_count = int((excluded & current).sum())
    else:
        expected_priority_pool = 0
        expected_manual_pool = 0
        excluded_current_count = 0

    rows.append({"check": "eligible_current_with_dates", "status": "INFO", "value": expected_priority_pool, "notes": "Potential rows available for priority/closing-soon tabs before score-band splitting."})
    rows.append({"check": "eligible_missing_dates", "status": "PASS" if len(manual_df) == expected_manual_pool else "REVIEW", "value": expected_manual_pool, "notes": "Expected manual-date rows from consolidated master."})
    rows.append({"check": "excluded_current_rows_retained_in_master", "status": "INFO", "value": excluded_current_count, "notes": "Excluded rows remain in consolidated_master but must not appear in daily-review operational tabs."})
    rows.append({"check": "scraper_run_summary_present", "status": "PASS" if "Scraper Run Summary" in sheets else "FAIL", "value": "Yes" if "Scraper Run Summary" in sheets else "No", "notes": "Operational source summary sheet should be present."})
    rows.append({"check": "by_source_present", "status": "PASS" if "By Source" in sheets else "FAIL", "value": "Yes" if "By Source" in sheets else "No", "notes": "All-source reporting tab should be present."})
    rows.append({"check": "source_coverage_present", "status": "PASS" if "Source Coverage" in sheets else "FAIL", "value": "Yes" if "Source Coverage" in sheets else "No", "notes": "Review-list source representation tab should be present."})
    return pd.DataFrame(rows)


def build_daily_review(
    master: pd.DataFrame,
    delta: pd.DataFrame,
    updates: pd.DataFrame,
    sources_config: dict[str, Any] | None = None,
    data_root: str | Path | None = None,
) -> dict[str, pd.DataFrame]:
    """Build the daily review workbook.

    Priority review tabs use centrally recalculated current-status logic, not
    source-level flags. Rows with no valid closing date are separated into a
    manual-check tab instead of appearing in priority review tabs.
    """
    master = _apply_review_filter(_recalculate_current_flags(_align_closing_date_to_datetime(sort_for_review(master))))
    delta = _apply_review_filter(_recalculate_current_flags(_align_closing_date_to_datetime(sort_for_review(delta))))
    updates = _apply_review_filter(_recalculate_current_flags(_align_closing_date_to_datetime(sort_for_review(updates))))

    if master.empty:
        active_master = master.copy()
        missing_closing = master.copy()
    else:
        has_close = _has_closing_date(master)
        archived = _truthy_series(master["is_archived"])
        cancelled = _truthy_series(master["is_cancelled"])
        current = _truthy_series(master["is_current"])
        eligible = _truthy_series(master["review_eligible"])
        excluded = _truthy_series(master["excluded_by_filter"])
        active_master = master[current & has_close & ~archived & ~cancelled & eligible & ~excluded].copy()
        missing_closing = master[~has_close & ~archived & ~cancelled & eligible & ~excluded].copy()

    if active_master.empty:
        closing_dt = pd.Series(dtype="datetime64[ns]")
        days = pd.Series(dtype="float")
    else:
        closing_dt = _coalesced_closing_series(active_master)
        days = (closing_dt.dt.normalize() - pd.Timestamp.now().normalize()).dt.days

    current_delta = delta[_truthy_series(delta["is_current"]) & _has_closing_date(delta) & _truthy_series(delta["review_eligible"]) & ~_truthy_series(delta["excluded_by_filter"])] if not delta.empty else delta.copy()
    current_updates = updates[_truthy_series(updates["is_current"]) & _has_closing_date(updates) & _truthy_series(updates["review_eligible"]) & ~_truthy_series(updates["excluded_by_filter"])] if not updates.empty else updates.copy()

    very_high = active_master[active_master["score_band"].eq("Very High")].copy()
    high = active_master[active_master["score_band"].eq("High")].copy()
    medium = active_master[active_master["score_band"].eq("Medium")].copy()
    closing_soon = active_master[(days >= 0) & (days <= 14)].copy() if not active_master.empty else active_master.copy()

    by_source = _build_by_source_summary(sources_config, Path(data_root) if data_root else None, master, delta, updates, active_master, missing_closing, closing_soon)
    source_coverage = _build_source_coverage_detail(active_master)

    sheets = {
        "Very High Priority": _prepare_daily_review_display(very_high),
        "High Priority": _prepare_daily_review_display(high),
        "Medium Priority": _prepare_daily_review_display(medium),
        "Closing Soon": _prepare_daily_review_display(closing_soon),
        "New Today": _prepare_daily_review_display(current_delta.copy()),
        "Updated Today": _prepare_daily_review_display(current_updates.copy()),
        "Needs Manual Date Check": _prepare_daily_review_display(sort_for_review(missing_closing)),
        "By Source": by_source,
        "Source Coverage": source_coverage,
    }
    # Dedicated duplicate of the operational source status table, making the
    # purpose explicit for users who expect Source Coverage to include all scrapers.
    if not by_source.empty:
        sheets["Scraper Run Summary"] = by_source.copy()

    sheets["QA Checks"] = _build_daily_review_qa(sheets, master)
    return sheets



def _row_identifier_keys(row: pd.Series) -> set[str]:
    """Return stable identifier keys for matching source records back to a merged row."""
    keys: set[str] = set()
    for field in ["tender_number", "reference_number", "opportunity_key"]:
        value = clean_text(row.get(field, ""))
        if not value:
            continue
        values = [value]
        # Opportunity keys often look like tcta:051-2025...; include the suffix as well.
        if ":" in value:
            values.append(value.split(":", 1)[1])
        for v in values:
            norm = normalise_identifier(v)
            if norm:
                keys.add(norm)
    return keys


def _row_source_names(row: pd.Series) -> set[str]:
    """Return normalised source names from source and source_list."""
    sources: set[str] = set()
    for field in ["source", "source_list"]:
        for part in _split_multi_values(row.get(field, "")):
            norm = normalise_text(part)
            if norm:
                sources.add(norm)
    return sources


def _best_date_candidate_from_rows(candidates: list[tuple[pd.Timestamp, str]]) -> tuple[pd.Timestamp | None, str, str]:
    """Choose earliest current/future date, otherwise latest parsed historical date."""
    if not candidates:
        return None, "", "No source record with valid closing date found"
    today = pd.Timestamp.now().normalize()
    future = [(dt, src) for dt, src in candidates if dt.normalize() >= today]
    if future:
        dt, src = sorted(future, key=lambda x: x[0])[0]
        return dt, src, "Filled from matched source record; earliest current/future date selected"
    dt, src = sorted(candidates, key=lambda x: x[0])[-1]
    return dt, src, "Filled from matched source record; all matching dates appear expired"


def _build_date_candidate_index(source_records: pd.DataFrame) -> dict[str, list[tuple[pd.Timestamp, str]]]:
    """Build ref/opportunity-key -> [(closing datetime, source)] lookup from all source rows.

    This is a defensive post-dedupe repair step. If a duplicate row is merged with
    eTenders as the primary record, but a source portal row such as TCTA has the
    actual closing date, the consolidated record must use the source-portal date.
    """
    index: dict[str, list[tuple[pd.Timestamp, str]]] = {}
    if source_records is None or source_records.empty:
        return index
    src_df = ensure_standard_columns(source_records).copy()
    for _, src_row in src_df.iterrows():
        dt = _valid_datetime(src_row.get("closing_datetime", ""))
        if pd.isna(dt):
            dt = _valid_datetime(src_row.get("closing_date", ""))
        if pd.isna(dt):
            continue
        src_name = clean_text(src_row.get("source", ""))
        for key in _row_identifier_keys(src_row):
            index.setdefault(key, []).append((dt, src_name))
    return index


def _repair_missing_closing_dates_from_sources(df: pd.DataFrame, source_records: pd.DataFrame) -> pd.DataFrame:
    """Fill missing consolidated closing dates from matched source records.

    This specifically protects cross-source duplicates such as eTenders + TCTA.
    If the selected consolidated base row has a blank date, but any matched source
    row has a valid date, the date is restored before daily-review classification.
    """
    out = ensure_standard_columns(df).copy()
    if out.empty:
        return out

    date_index = _build_date_candidate_index(source_records)
    if not date_index:
        return out

    for idx, row in out.iterrows():
        current_dt = _valid_datetime(row.get("closing_datetime", ""))
        if pd.isna(current_dt):
            current_dt = _valid_datetime(row.get("closing_date", ""))
        if pd.notna(current_dt):
            continue

        candidates: list[tuple[pd.Timestamp, str]] = []
        for key in _row_identifier_keys(row):
            candidates.extend(date_index.get(key, []))

        # Deduplicate candidate pairs while preserving values.
        unique_candidates: list[tuple[pd.Timestamp, str]] = []
        seen: set[tuple[str, str]] = set()
        row_sources = _row_source_names(row)
        for dt, src in candidates:
            # Prefer candidates from sources actually represented in source/source_list,
            # but do not discard all candidates if source_list is incomplete.
            norm_src = normalise_text(src)
            candidate_key = (str(pd.Timestamp(dt)), norm_src)
            if candidate_key in seen:
                continue
            seen.add(candidate_key)
            unique_candidates.append((dt, src))

        if row_sources:
            source_matched = [(dt, src) for dt, src in unique_candidates if normalise_text(src) in row_sources]
            if source_matched:
                unique_candidates = source_matched

        chosen_dt, chosen_src, note = _best_date_candidate_from_rows(unique_candidates)
        if chosen_dt is not None and pd.notna(chosen_dt):
            out.at[idx, "closing_datetime"] = chosen_dt
            out.at[idx, "closing_date"] = pd.Timestamp(chosen_dt).normalize()
            out.at[idx, "closing_datetime_source"] = chosen_src
            prev_note = clean_text(out.at[idx, "date_resolution_notes"]) if "date_resolution_notes" in out.columns else ""
            out.at[idx, "date_resolution_notes"] = "; ".join(x for x in [prev_note, note] if x)

    out = _align_closing_date_to_datetime(out)
    out = _recalculate_current_flags(out)
    return out


def _combined_source_records(scored: dict[str, pd.DataFrame]) -> pd.DataFrame:
    frames = []
    for key in ["master", "delta", "updates", "top_delta"]:
        df = scored.get(key)
        if df is not None and not df.empty:
            frames.append(ensure_standard_columns(df).copy())
    if not frames:
        return pd.DataFrame(columns=STANDARD_COLUMNS)
    return pd.concat(frames, ignore_index=True, sort=False)

def consolidate(
    sources_config_path: str | Path | None = None,
    scoring_config_path: str | Path | None = None,
    data_root: str | Path | None = None,
    output_root: str | Path | None = None,
    sources_registry_path: str | Path | None = None,
    project_root: str | Path | None = None,
) -> dict[str, str]:
    if scoring_config_path is None:
        raise ValueError("scoring_config_path is required")

    scoring_config = load_scoring_config(scoring_config_path)

    if sources_registry_path:
        root = Path(project_root or data_root or Path(sources_registry_path).resolve().parents[1])
        sources_config = sources_config_from_registry(sources_registry_path, root)
        data_root = Path(data_root or root)
    elif sources_config_path:
        sources_config = load_json(sources_config_path)
        data_root = Path(data_root or sources_config.get("data_analytics_root_default"))
    else:
        raise ValueError("Either sources_registry_path or sources_config_path is required")

    output_root = Path(output_root or (Path(data_root) / sources_config.get("consolidated_output_folder", "04_FinalOutputs")))
    output_root.mkdir(parents=True, exist_ok=True)

    loaded = load_source_outputs(sources_config, data_root)

    scored = {}
    for key, df in loaded.items():
        scored[key] = _apply_review_filter(sort_for_review(score_dataframe(df, scoring_config)))

    all_source_records = _combined_source_records(scored)

    master = dedupe_records(scored.get("master", pd.DataFrame(columns=STANDARD_COLUMNS)))
    master = _repair_missing_closing_dates_from_sources(master, all_source_records)
    master = _apply_review_filter(_recalculate_current_flags(master))

    delta = dedupe_records(scored.get("delta", pd.DataFrame(columns=STANDARD_COLUMNS)))
    delta = _repair_missing_closing_dates_from_sources(delta, all_source_records)
    delta = _apply_review_filter(_recalculate_current_flags(delta))

    updates = dedupe_records(scored.get("updates", pd.DataFrame(columns=STANDARD_COLUMNS)))
    updates = _repair_missing_closing_dates_from_sources(updates, all_source_records)
    updates = _apply_review_filter(_recalculate_current_flags(updates))

    top_delta_source = _align_closing_date_to_datetime(scored.get("top_delta", pd.DataFrame(columns=STANDARD_COLUMNS)))
    top_delta_source = _repair_missing_closing_dates_from_sources(top_delta_source, all_source_records)
    top_delta_source = _apply_review_filter(_recalculate_current_flags(top_delta_source))

    # Consolidated top delta is recalculated from current consolidated delta using the central score.
    if not delta.empty:
        delta_current = delta[_truthy_series(delta["is_current"]) & _has_closing_date(delta) & _truthy_series(delta["review_eligible"])].copy()
    else:
        delta_current = delta.copy()
    top_delta = sort_for_review(delta_current).head(50).copy()

    paths = {
        "master": str(output_root / "consolidated_master.xlsx"),
        "delta": str(output_root / "consolidated_delta.xlsx"),
        "updates": str(output_root / "consolidated_updates.xlsx"),
        "top_delta": str(output_root / "consolidated_top_delta.xlsx"),
        "daily_review": str(output_root / "consolidated_daily_review.xlsx"),
        "source_top_delta_standardised": str(output_root / "source_top_delta_standardised.xlsx"),
    }

    write_excel_with_table(master, paths["master"], "Master", "tbl_consolidated_master")
    write_excel_with_table(delta, paths["delta"], "Delta", "tbl_consolidated_delta")
    write_excel_with_table(updates, paths["updates"], "Updates", "tbl_consolidated_updates")
    write_excel_with_table(top_delta, paths["top_delta"], "Top_Delta", "tbl_consolidated_top_delta")
    write_excel_with_table(sort_for_review(top_delta_source), paths["source_top_delta_standardised"], "Source_Top_Delta", "tbl_source_top_delta")
    write_multi_sheet_excel(build_daily_review(master, delta, updates, sources_config=sources_config, data_root=data_root), paths["daily_review"])

    return paths
