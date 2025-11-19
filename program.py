# program.py
"""
Core logic for comparing contingency tables between Excel sheets.

This module does NOT contain any GUI code.  It:
- Loads an .xlsx workbook
- Extracts the 3 tables ("ACCA Long Term", "ACCA", "DCwAC") from each sheet
- Compares tables between any two sheets
"""

from __future__ import annotations
import math
from typing import Dict, Any
import pandas as pd


TABLE_NAMES = ("ACCA Long Term", "ACCA", "DCwAC")


def load_workbook(path: str) -> Dict[str, Dict[str, pd.DataFrame]]:
    """
    Load an Excel workbook and parse the important tables from every sheet.

    Returns:
        {
          "Sheet1": {
              "ACCA Long Term": df,
              "ACCA": df,
              "DCwAC": df,
          },
          "Sheet2": { ... },
          ...
        }
    """
    # Read whole workbook as raw text (no header, we’ll detect it ourselves)
    raw_sheets: Dict[str, pd.DataFrame] = pd.read_excel(
        path, sheet_name=None, header=None, dtype=str, engine="openpyxl"
    )

    parsed: Dict[str, Dict[str, pd.DataFrame]] = {}
    for sheet_name, df in raw_sheets.items():
        parsed[sheet_name] = extract_tables_from_sheet(df)

    return parsed


def extract_tables_from_sheet(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """
    Given a raw sheet (no headers), find and extract each contingency table.

    Assumptions (based on your screenshot):
    - Each block has a header row whose first cell is "Contingency Events".
    - The row *above* that header row contains the table name somewhere
      in the row: "ACCA Long Term", "ACCA", or "DCwAC".
    - Rows between blocks are blank.
    """
    tables: Dict[str, pd.DataFrame] = {}

    nrows, ncols = df.shape

    def row_is_blank(idx: int) -> bool:
        row = df.iloc[idx, :]
        return all(
            (isinstance(v, float) and math.isnan(v))
            or (isinstance(v, str) and v.strip() == "")
            or (v is None)
            for v in row
        )

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
                if not isinstance(val, str):
                    continue
                text = val.strip()
                for candidate in TABLE_NAMES:
                    if text.lower() == candidate.lower():
                        table_name = candidate
                        break
                if table_name:
                    break

        if not table_name:
            # We found a "Contingency Events" header but no recognizable table
            continue

        # Find the bottom of this block (first fully blank row after the header)
        start_data = i + 1
        end_data = start_data
        while end_data < nrows and not row_is_blank(end_data):
            end_data += 1

        # Build a DataFrame with header row = row i
        header = df.iloc[i, :].tolist()
        block = df.iloc[start_data:end_data, :].copy()
        block.columns = header
        block = block.dropna(how="all")

        if table_name in tables:
            # If the same table appears more than once, append
            tables[table_name] = pd.concat([tables[table_name], block], ignore_index=True)
        else:
            tables[table_name] = block

    return tables


def _find_col_by_prefix(columns: Any, prefix: str) -> str | None:
    """Find the first column whose name starts with the given prefix (case-insensitive)."""
    prefix = prefix.lower()
    for col in columns:
        if not isinstance(col, str):
            continue
        if col.strip().lower().startswith(prefix):
            return col
    return None


def _to_float_series(series: pd.Series | None) -> pd.Series:
    """Convert a Series to float, stripping % and ignoring non-numeric junk."""
    if series is None:
        # Create empty float series if column missing
        return pd.Series([math.nan] * 0, dtype=float)

    def _conv(x):
        if x is None:
            return math.nan
        s = str(x).strip()
        if s == "" or s.lower() == "nan":
            return math.nan
        s = s.replace("%", "")
        try:
            return float(s)
        except ValueError:
            return math.nan

    return series.apply(_conv)


def compare_tables(
    table1: pd.DataFrame, table2: pd.DataFrame, name1: str, name2: str
) -> pd.DataFrame:
    """
    Compare two contingency tables (same type, different sheets).

    Returns a DataFrame with:
    - contingency
    - issue
    - value_1, percent_1
    - value_2, percent_2
    - delta_percent (2 - 1)
    - status: 'both', 'only in left', 'only in right'
    """

    # Try to auto-locate important columns
    c_col = _find_col_by_prefix(table1.columns, "contingency")
    r_col = _find_col_by_prefix(table1.columns, "resulting")
    v1_col = _find_col_by_prefix(table1.columns, "contingency value")
    p1_col = _find_col_by_prefix(table1.columns, "percent")

    c_col2 = _find_col_by_prefix(table2.columns, "contingency")
    r_col2 = _find_col_by_prefix(table2.columns, "resulting")
    v2_col = _find_col_by_prefix(table2.columns, "contingency value")
    p2_col = _find_col_by_prefix(table2.columns, "percent")

    # Build normalized versions
    left = pd.DataFrame(
        {
            "contingency": table1[c_col].astype(str),
            "issue": table1[r_col].astype(str),
            "value_1": _to_float_series(table1.get(v1_col)),
            "percent_1": _to_float_series(table1.get(p1_col)),
        }
    )

    right = pd.DataFrame(
        {
            "contingency": table2[c_col2].astype(str),
            "issue": table2[r_col2].astype(str),
            "value_2": _to_float_series(table2.get(v2_col)),
            "percent_2": _to_float_series(table2.get(p2_col)),
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

    # Optional sort: biggest % increase to the top
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
    Compare two sheets for all known table types.

    Returns:
        {
          "ACCA Long Term": df_comparison,
          "ACCA": df_comparison,
          "DCwAC": df_comparison
        }
    Only tables present on BOTH sheets are returned.
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
