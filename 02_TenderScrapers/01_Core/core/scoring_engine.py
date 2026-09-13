"""Central scoring engine for tender records."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pandas as pd

from .date_utils import days_until
from .text_utils import clean_text, match_keyword_groups, normalise_text, join_unique


def load_scoring_config(path: str | Path) -> dict[str, Any]:
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def _combined_search_text(row: pd.Series) -> str:
    fields = [
        "title",
        "description",
        "description_en",
        "description_original",
        "client",
        "sector",
        "category",
        "notice_type",
        "conditions",
        "briefing_session_details",
    ]
    return " ".join(clean_text(row.get(field, "")) for field in fields if clean_text(row.get(field, "")))


def _priority_client_matches(row: pd.Series, config: dict[str, Any]) -> list[str]:
    search = normalise_text(" ".join([
        clean_text(row.get("client", "")),
        clean_text(row.get("source", "")),
        clean_text(row.get("description", "")),
    ]))
    matches = []
    for item in config.get("priority_clients", []):
        item_norm = normalise_text(item)
        if item_norm and f" {item_norm} " in f" {search} ":
            matches.append(item)
    return matches


def _notice_type_match(row: pd.Series, config: dict[str, Any]) -> bool:
    notice = normalise_text(row.get("notice_type", ""))
    title_desc = normalise_text(f"{row.get('title', '')} {row.get('description', '')}")
    for item in config.get("priority_notice_types", []):
        item_norm = normalise_text(item)
        if item_norm and (f" {item_norm} " in f" {notice} " or f" {item_norm} " in f" {title_desc} "):
            return True
    return False


def score_band(score: int, config: dict[str, Any]) -> str:
    bands = config.get("score_bands", [])
    for band in sorted(bands, key=lambda b: b.get("min", 0), reverse=True):
        if score >= int(band.get("min", 0)):
            return band.get("band", "")
    return ""


def score_dataframe(df: pd.DataFrame, config: dict[str, Any]) -> pd.DataFrame:
    out = df.copy()
    weights = config.get("weights", {})

    if out.empty:
        for col in [
            "priority_client_match",
            "priority_client_matches",
            "sector_keyword_matches",
            "technical_keyword_matches",
            "service_keyword_matches",
            "procurement_keyword_matches",
            "strategic_keyword_matches",
            "matched_keywords_all",
            "weight_score",
            "score_band",
            "scoring_notes",
        ]:
            if col not in out.columns:
                out[col] = [] if col == "weight_score" else ""
        return out

    scores = []
    bands = []
    notes = []
    priority_flags = []
    priority_matches_col = []
    sector_matches_col = []
    technical_matches_col = []
    service_matches_col = []
    procurement_matches_col = []
    strategic_matches_col = []
    all_matches_col = []

    for _, row in out.iterrows():
        score = 0
        row_notes: list[str] = []
        search_text = _combined_search_text(row)

        priority_matches = _priority_client_matches(row, config)
        priority_match = bool(priority_matches)
        if priority_match:
            add = int(weights.get("priority_client_match", 0))
            score += add
            row_notes.append(f"Priority client/source match +{add}: {', '.join(priority_matches)}")

        sector_matches = match_keyword_groups(search_text, config.get("strategic_sector_keywords", {}))
        if sector_matches:
            add = len(sector_matches) * int(weights.get("strategic_sector_keyword", 0))
            score += add
            row_notes.append(f"Strategic sector keywords +{add}: {', '.join(sector_matches)}")

        technical_matches = match_keyword_groups(search_text, config.get("technical_keywords", {}))
        if technical_matches:
            add = len(technical_matches) * int(weights.get("technical_keyword", 0))
            score += add
            row_notes.append(f"Technical keywords +{add}: {', '.join(technical_matches)}")

        service_matches = match_keyword_groups(search_text, config.get("service_keywords", {}))
        if service_matches:
            add = len(service_matches) * int(weights.get("service_keyword", 0))
            score += add
            row_notes.append(f"Service keywords +{add}: {', '.join(service_matches)}")

        procurement_matches = match_keyword_groups(search_text, config.get("procurement_delivery_keywords", {}))
        if procurement_matches:
            add = len(procurement_matches) * int(weights.get("procurement_delivery_keyword", 0))
            score += add
            row_notes.append(f"Procurement/delivery keywords +{add}: {', '.join(procurement_matches)}")

        if _notice_type_match(row, config):
            add = int(weights.get("notice_type_match", 0))
            score += add
            row_notes.append(f"Priority notice type +{add}")

        d_until = days_until(row.get("closing_datetime", None))
        if d_until is None:
            d_until = days_until(row.get("closing_date", None))
        if d_until is not None and 0 <= d_until <= int(config.get("closing_soon_days", 14)):
            add = int(weights.get("closing_soon", 0))
            score += add
            row_notes.append(f"Closing within {config.get('closing_soon_days', 14)} days +{add}")

        doc_count = row.get("document_count", 0)
        try:
            has_docs = int(float(doc_count or 0)) > 0
        except Exception:
            has_docs = bool(clean_text(row.get("document_urls", "")))
        if has_docs:
            add = int(weights.get("has_documents", 0))
            score += add
            row_notes.append(f"Has documents +{add}")

        if clean_text(row.get("contact_email", "")):
            add = int(weights.get("has_contact_email", 0))
            score += add
            row_notes.append(f"Has contact email +{add}")

        all_matches = join_unique(
            sector_matches + technical_matches + service_matches + procurement_matches,
            sep="; ",
        )

        scores.append(int(score))
        bands.append(score_band(int(score), config))
        notes.append(" | ".join(row_notes))
        priority_flags.append(priority_match)
        priority_matches_col.append("; ".join(priority_matches))
        sector_matches_col.append("; ".join(sector_matches))
        technical_matches_col.append("; ".join(technical_matches))
        service_matches_col.append("; ".join(service_matches))
        procurement_matches_col.append("; ".join(procurement_matches))
        strategic_matches_col.append("; ".join(sector_matches))
        all_matches_col.append(all_matches)

    out["priority_client_match"] = priority_flags
    out["priority_client_matches"] = priority_matches_col
    out["sector_keyword_matches"] = sector_matches_col
    out["technical_keyword_matches"] = technical_matches_col
    out["service_keyword_matches"] = service_matches_col
    out["procurement_keyword_matches"] = procurement_matches_col
    out["strategic_keyword_matches"] = strategic_matches_col
    out["matched_keywords_all"] = all_matches_col
    out["weight_score"] = scores
    out["score_band"] = bands
    out["scoring_notes"] = notes
    return out
