# gui/program/excel_logic.py

from __future__ import annotations
import math
from typing import Dict, Any

import pandas as pd

TABLE_NAMES = ("ACCA Long Term", "ACCA", "DCwAC")


# ───────────────────────── LOADING ───────────────────────── #

def load_workbook(path: str) -> Dict[str, Dict[str, pd.DataFrame]]:
    """
    Load an Excel workbook and parse key tables from each sheet.

    Returns:
        {
          "SheetName": {
              "ACCA Long Term": df,
              "ACCA": df,
              "DCwAC": df,
          },
          ...
        }
    """
    # header=None so we can find the header rows ourselves
    raw_sheets: Dict[str, pd.DataFrame] = pd.read_excel(
        path, sheet_name=None, header=None, engine="openpyxl"
    )

    parsed: Dict[str, Dict[str, pd.DataFrame]] = {}

    print("=== Parsing workbook ===")
    for sheet_name, df in raw_sheets.items():
        tables = extract_tables_from_sheet(df)
        parsed[sheet_name] = tables
        print(f"Sheet '{sheet_name}' → found tables: {list(tables.keys())}")

    return parsed


def extract_tables_from_sheet(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """
    Find each block that belongs to "ACCA Long Term", "ACCA", "DCwAC".

    Pattern (based on your screenshot):

    Row X-1 : contains table name ("ACCA Long Term", "ACCA", "DCwAC")
    Row X   : first cell is "Contingency Events"
    Row X+1..Y-1 : data rows until a fully blank row
    """
    tables: Dict[str, pd.DataFrame] = {}

    nrows, ncols = df.shape

    def row_is_blank(idx: int) -> bool:
        # Entire row is NaN -> blank
        return bool(df.iloc[idx, :].isna().all())

    for i in range(nrows):
        cell0 = df.iloc[i, 0]

        if not isinstance(cell0, str):
            continue
        if cell0.strip().lower() != "contingency events":
            continue

        # Look one row above for the table name
        table_name = None
        if i > 0:
            header_row = df.iloc[i - 1, :]
            for c in range(ncols):
                val = header_row[c]
                if isinstance(val, str):
                    text = val.strip()
                    for candidate in TABLE_NAMES:
                        if text.lower() == candidate.lower():
                            table_name = candidate
                            break
                if table_name:
                    break

        if not table_name:
            # Couldn't match a known table name
            continue

        # Determine data region
        start_data = i + 1
        end_data = start_data
        while end_data < nrows and not row_is_blank(end_data):
            end_data += 1

        # Build DataFrame with row i as header
        header = df.iloc[i, :].tolist()
        block = df.iloc[start_data:end_data, :].copy()
        block.columns = header
        block = block.dropna(how="all")

        if table_name in tables:
            tables[table_name] = pd.concat(
                [tables[table_name], block], ignore_index=True
            )
        else:
            tables[table_name] = block

    return tables


# ───────────────────────── COMPARISON ───────────────────────── #

def _find_col_by_prefix(columns: Any, prefix: str) -> str | None:
    prefix = prefix.lower()
    for col in columns:
        if isinstance(col, str) and col.strip().lower().startswith(prefix):
            return col
    return None


def _to_float_series(series: pd.Series) -> pd.Series:
    def _conv(x):
        if x is None or (isinstance(x, float) and math.isnan(x)):
            return math.nan
        s = str(x).strip().replace("%", "")
        if not s:
            return math.nan
        try:
            return float(s)
        except ValueError:
            return math.nan

    return series.apply(_conv)


def compare_tables(
    table1: pd.DataFrame, table2: pd.DataFrame, name1: str, name2: str
) -> pd.DataFrame:
    """
    Compare two tables of the same type from two different sheets.
    """

    # Auto-detect important columns
    c1 = _find_col_by_prefix(table1.columns, "contingency")
    r1 = _find_col_by_prefix(table1.columns, "resulting")
    v1 = _find_col_by_prefix(table1.columns, "contingency value")
    p1 = _find_col_by_prefix(table1.columns, "percent")

    c2 = _find_col_by_prefix(table2.columns, "contingency")
    r2 = _find_col_by_prefix(table2.columns, "resulting")
    v2 = _find_col_by_prefix(table2.columns, "contingency value")
    p2 = _find_col_by_prefix(table2.columns, "percent")

    required = [("contingency", c1, c2), ("resulting issue", r1, r2),
                ("value", v1, v2), ("percent", p1, p2)]
    for label, left, right in required:
        if left is None or right is None:
            raise ValueError(f"Could not find '{label}' columns on both tables.")

    left = pd.DataFrame(
        {
            "contingency": table1[c1].astype(str),
            "issue": table1[r1].astype(str),
            "value_1": _to_float_series(table1[v1]),
            "percent_1": _to_float_series(table1[p1]),
        }
    )

    right = pd.DataFrame(
        {
            "contingency": table2[c2].astype(str),
            "issue": table2[r2].astype(str),
            "value_2": _to_float_series(table2[v2]),
            "percent_2": _to_float_series(table2[p2]),
        }
    )

    merged = pd.merge(
        left,
        right,
        on=["contingency", "issue"],
        how="outer",
        sort=False,
    )

    merged["sheet_left"] = name1
    merged["sheet_right"] = name2
    merged["delta_percent"] = merged["percent_2"] - merged["percent_1"]

    def _status(row):
        in_left = not math.isnan(row.get("percent_1", math.nan))
        in_right = not math.isnan(row.get("percent_2", math.nan))
        if in_left and in_right:
            return "both"
        if in_left:
            return "only in left"
        if in_right:
            return "only in right"
        return "unknown"

    merged["status"] = merged.apply(_status, axis=1)

    merged = merged.sort_values(
        by=["status", "delta_percent"], ascending=[True, False], ignore_index=True
    )

    return merged


def compare_sheet_pair(
    workbook_data: Dict[str, Dict[str, pd.DataFrame]],
    sheet_left: str,
    sheet_right: str,
) -> Dict[str, pd.DataFrame]:
    """
    Compare two sheets across all known table types.
    Only returns tables that exist on BOTH sheets.
    """
    tables_left = workbook_data.get(sheet_left, {})
    tables_right = workbook_data.get(sheet_right, {})

    results: Dict[str, pd.DataFrame] = {}

    for tname in TABLE_NAMES:
        t1 = tables_left.get(tname)
        t2 = tables_right.get(tname)
        if t1 is None or t2 is None:
            continue
        results[tname] = compare_tables(t1, t2, sheet_left, sheet_right)

    return results

