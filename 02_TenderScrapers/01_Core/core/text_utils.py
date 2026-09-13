"""Text cleaning and matching helpers."""
from __future__ import annotations

import re
from typing import Any, Iterable, List

import pandas as pd


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalise_text(value: Any) -> str:
    text = clean_text(value).casefold()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def normalise_identifier(value: Any) -> str:
    text = clean_text(value).casefold()
    text = re.sub(r"[^a-z0-9]+", "", text)
    return text


def contains_phrase(text_norm: str, phrase: str) -> bool:
    phrase_norm = normalise_text(phrase)
    if not phrase_norm:
        return False
    return f" {phrase_norm} " in f" {text_norm} "


def match_keyword_groups(text: Any, keyword_groups: dict[str, list[str]]) -> List[str]:
    text_norm = normalise_text(text)
    matches: list[str] = []
    for label, phrases in keyword_groups.items():
        if any(contains_phrase(text_norm, phrase) for phrase in phrases):
            matches.append(label)
    return matches


def join_unique(values: Iterable[Any], sep: str = "; ") -> str:
    seen = set()
    out = []
    for value in values:
        if value is None:
            continue
        text = clean_text(value)
        if not text:
            continue
        # Split existing semi-colon delimited lists and dedupe atomically.
        parts = [p.strip() for p in text.split(";") if p.strip()]
        for part in parts:
            key = normalise_text(part)
            if key and key not in seen:
                seen.add(key)
                out.append(part)
    return sep.join(out)


def first_non_empty(*values: Any) -> str:
    for value in values:
        text = clean_text(value)
        if text:
            return text
    return ""
