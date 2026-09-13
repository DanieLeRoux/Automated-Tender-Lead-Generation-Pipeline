"""Date parsing and tender status helpers.

v21: Handles Excel serial dates written/read from source workbooks.
Several source scraper outputs, especially eTenders, may store closing dates as
Excel serial values (e.g. 46191.4166666667 = 2026-06-18 10:00). Pandas' normal
pd.to_datetime parser treats numeric values as nanoseconds from 1970, which causes
valid future tender dates to become invalid/expired during consolidation.
"""
from __future__ import annotations

from datetime import date, datetime
from typing import Any
import re

import pandas as pd

EXCEL_DATE_MIN = 20000   # safely above Excel serial dates before 1954
EXCEL_DATE_MAX = 90000   # safely beyond current operational tender horizons


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    if isinstance(value, str) and not value.strip():
        return True
    return False


def _parse_excel_serial(value: Any):
    """Parse Excel serial date values, including fractional days."""
    try:
        number = float(value)
    except Exception:
        return pd.NaT
    if EXCEL_DATE_MIN <= number <= EXCEL_DATE_MAX:
        return pd.to_datetime(number, unit="D", origin="1899-12-30", errors="coerce").round("s")
    return pd.NaT


def parse_date(value: Any):
    """Robustly parse a date/datetime value used by tender outputs.

    Handles:
    - pandas/native datetime values;
    - normal date strings, including SA day-first formats;
    - ISO date/time strings;
    - Excel serial date numbers and numeric strings.
    """
    if _is_blank(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value if pd.notna(value) else pd.NaT

    if isinstance(value, (datetime, date)):
        return pd.to_datetime(value, errors="coerce")

    # Excel often stores dates as floats/ints. Numeric strings must also be caught
    # before pd.to_datetime, otherwise they become NaT or 1970-based nanoseconds.
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        serial_dt = _parse_excel_serial(value)
        if pd.notna(serial_dt):
            return serial_dt

    if isinstance(value, str):
        text = value.strip()
        if not text:
            return pd.NaT
        # Remove common Excel escaped carriage-return artefacts and redundant spaces.
        text = text.replace("_x000D_", " ").replace("\r", " ").replace("\n", " ")
        text = " ".join(text.split())
        # Numeric string may be an Excel serial date.
        if text.replace(".", "", 1).isdigit():
            serial_dt = _parse_excel_serial(text)
            if pd.notna(serial_dt):
                return serial_dt
        value = text

    # ISO-style dates are year-first, so do not parse them day-first. This avoids
    # pandas warnings and prevents accidental month/day ambiguity.
    if isinstance(value, str) and re.match(r"^\d{4}-\d{1,2}-\d{1,2}(?:[T\s].*)?$", value):
        dt = pd.to_datetime(value, errors="coerce", dayfirst=False)
    else:
        dt = pd.to_datetime(value, errors="coerce", dayfirst=True)

    # Defensive correction: if a numeric input slipped through and pandas treated it
    # as nanoseconds after Unix epoch, prefer Excel serial interpretation.
    if pd.notna(dt) and isinstance(value, (int, float)) and dt.year < 1990:
        serial_dt = _parse_excel_serial(value)
        if pd.notna(serial_dt):
            return serial_dt

    return dt


def parse_date_series(series: pd.Series) -> pd.Series:
    """Parse a Series using parse_date while preserving the original index."""
    if series is None:
        return pd.Series(dtype="datetime64[ns]")
    return series.apply(parse_date)


def coerce_date_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = df.copy()
    for col in columns:
        if col in out.columns:
            out[col] = parse_date_series(out[col])
    return out


def current_flag(closing_value: Any, archived: bool = False, cancelled: bool = False) -> bool:
    if archived or cancelled:
        return False
    dt = parse_date(closing_value)
    if pd.isna(dt):
        return False
    now = pd.Timestamp.now()
    if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
        return dt.normalize() >= now.normalize()
    return dt >= now


def days_until(value: Any):
    dt = parse_date(value)
    if pd.isna(dt):
        return None
    return int((dt.normalize() - pd.Timestamp.now().normalize()).days)
