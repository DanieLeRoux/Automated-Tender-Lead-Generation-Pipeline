"""Excel output helpers."""
from __future__ import annotations

import os
from typing import Optional

import pandas as pd


def safe_table_name(name: str) -> str:
    keep = "".join(ch if ch.isalnum() or ch == "_" else "_" for ch in name)
    if not keep or keep[0].isdigit():
        keep = "tbl_" + keep
    return keep[:250]


def write_excel_with_table(df: pd.DataFrame, path: str, sheet_name: str, table_name: Optional[str] = None) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    out_df = df.copy() if df is not None else pd.DataFrame()
    tmp = path + ".tmp.xlsx"

    with pd.ExcelWriter(
        tmp,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
        datetime_format="yyyy-mm-dd hh:mm",
    ) as writer:
        out_df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        worksheet = writer.sheets[sheet_name[:31]]
        nrows, ncols = out_df.shape
        if ncols:
            columns = [{"header": str(col)} for col in out_df.columns]
            if nrows:
                worksheet.add_table(
                    0,
                    0,
                    nrows,
                    ncols - 1,
                    {"name": safe_table_name(table_name or f"tbl_{sheet_name}"), "columns": columns},
                )
            else:
                worksheet.autofilter(0, 0, 0, ncols - 1)
            worksheet.freeze_panes(1, 0)
            for idx, col in enumerate(out_df.columns):
                width = min(max(len(str(col)) + 2, 12), 50)
                worksheet.set_column(idx, idx, width)

    os.replace(tmp, path)


def write_multi_sheet_excel(sheets: dict[str, pd.DataFrame], path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp.xlsx"
    used_names = set()
    with pd.ExcelWriter(
        tmp,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
        datetime_format="yyyy-mm-dd hh:mm",
    ) as writer:
        for raw_sheet_name, df in sheets.items():
            sheet_name = raw_sheet_name[:31]
            base = sheet_name
            i = 1
            while sheet_name in used_names:
                suffix = f"_{i}"
                sheet_name = (base[: 31 - len(suffix)] + suffix)[:31]
                i += 1
            used_names.add(sheet_name)

            out_df = df.copy() if df is not None else pd.DataFrame()
            out_df.to_excel(writer, index=False, sheet_name=sheet_name)
            worksheet = writer.sheets[sheet_name]
            nrows, ncols = out_df.shape
            if ncols:
                columns = [{"header": str(col)} for col in out_df.columns]
                if nrows:
                    worksheet.add_table(
                        0,
                        0,
                        nrows,
                        ncols - 1,
                        {"name": safe_table_name(f"tbl_{sheet_name}"), "columns": columns},
                    )
                else:
                    worksheet.autofilter(0, 0, 0, ncols - 1)
                worksheet.freeze_panes(1, 0)
                for idx, col in enumerate(out_df.columns):
                    width = min(max(len(str(col)) + 2, 12), 50)
                    worksheet.set_column(idx, idx, width)
    os.replace(tmp, path)
