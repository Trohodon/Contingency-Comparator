from __future__ import annotations
import math
from typing import Dict, Any, List

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
    raw_sheets: Dict[str, pd.DataFrame] = pd.read_excel(
        path, sheet_name=None, header=None, engine="openpyxl"
    )

    parsed: Dict[str, Dict[str, pd.DataFrame]] = {}

    print("=== Parsing workbook ===")
    for sheet_name, df in raw_sheets.items():
        print(f"\n--- DEBUG SHEET: {sheet_name} ---")
        # show first 40 rows so we can sanity-check
        try:
            print(df.head(40).to_string())
        except Exception:
            print(df.head(40))

        tables = extract_tables_from_sheet(df)
        parsed[sheet_name] = tables
        print(f"Sheet '{sheet_name}' → found tables: {list(tables.keys())}")

    return parsed


def extract_tables_from_sheet(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """
    Find ACCA Long Term / ACCA / DCwAC blocks using the 'Percent\\nLoading' row.

    Pattern (based on your workbook):

    - There is a header row where one of the cells contains "Percent" and
      another contains "Case".
    - The row immediately *below* that header is the first data row.
    - Data rows continue until a fully blank row OR another header row.
    - The *first* such block on the sheet is "ACCA Long Term",
      the second is "ACCA", and the third is "DCwAC".
    """

    tables: Dict[str, pd.DataFrame] = {}
    nrows, ncols = df.shape

    # 1) Find all header rows (those with a cell containing "percent")
    header_rows: List[int] = []
    for i in range(nrows):
        row = df.iloc[i, :]
        if any(isinstance(x, str) and "percent" in x.lower() for x in row):
            header_rows.append(i)

    # 2) For each header row, slice its data block and normalize columns
    for block_index, header_idx in enumerate(header_rows):
        if block_index >= len(TABLE_NAMES):
            # ignore extra blocks
            break

        table_name = TABLE_NAMES[block_index]

        # find key columns on the HEADER row
        percent_col = None
        case_col = None
        for col in range(ncols):
            val = df.iloc[header_idx, col]
            if isinstance(val, str):
                txt = val.lower()
                if "percent" in txt and percent_col is None:
                    percent_col = col
                elif "case" in txt and case_col is None:
                    case_col = col

        if percent_col is None:
            # without this we can't really parse the block
            continue

        # data starts on the next row
        first_data_row = header_idx + 1
        if first_data_row >= nrows:
            continue

        # data ends at the first fully blank row OR next header row
        def row_is_blank(idx: int) -> bool:
            return bool(df.iloc[idx, :].isna().all())

        last_data_row = first_data_row
        while last_data_row < nrows and not row_is_blank(last_data_row):
            # also stop if we hit another 'Percent' row
            if last_data_row in header_rows and last_data_row != header_idx:
                break
            last_data_row += 1

        # slice data block
        data_block = df.iloc[first_data_row:last_data_row, :]

        if data_block.empty:
            continue

        # Determine which columns are Contingency / Issue / Value by looking
        # at the first non-blank row in the block (usually first_data_row)
        sample_row = data_block.iloc[0, :]

        # columns to the LEFT of the Percent column that have non-NaN sample values
        left_cols = [c for c in range(percent_col)
                     if not pd.isna(sample_row.iloc[c])]

        contingency_col = left_cols[0] if len(left_cols) >= 1 else None
        issue_col = left_cols[1] if len(left_cols) >= 2 else None
        value_col = left_cols[2] if len(left_cols) >= 3 else None

        # If we can't identify at least contingency & issue, skip
        if contingency_col is None or issue_col is None:
            continue

        # Build a canonical DataFrame with fixed column names
        def safe_series(col_idx):
            if col_idx is None or col_idx >= ncols:
                return pd.Series([None] * len(data_block), index=data_block.index)
            return data_block.iloc[:, col_idx]

        block_norm = pd.DataFrame(
            {
                "Contingency": safe_series(contingency_col).astype(str),
                "Resulting Issue": safe_series(issue_col).astype(str),
                "Contingency Value": safe_series(value_col),
                "Percent Loading": safe_series(percent_col),
                "Case": safe_series(case_col),
            }
        )

        # Drop fully blank rows (in case of noise)
        block_norm = block_norm.dropna(how="all")
        tables[table_name] = block_norm.reset_index(drop=True)

    return tables


# ───────────────────────── COMPARISON ───────────────────────── #

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

    We now rely on our OWN canonical column names set in extract_tables_from_sheet():
      "Contingency", "Resulting Issue", "Contingency Value",
      "Percent Loading"
    """

    left = pd.DataFrame(
        {
            "contingency": table1["Contingency"].astype(str),
            "issue": table1["Resulting Issue"].astype(str),
            "value_1": _to_float_series(table1["Contingency Value"]),
            "percent_1": _to_float_series(table1["Percent Loading"]),
        }
    )

    right = pd.DataFrame(
        {
            "contingency": table2["Contingency"].astype(str),
            "issue": table2["Resulting Issue"].astype(str),
            "value_2": _to_float_series(table2["Contingency Value"]),
            "percent_2": _to_float_series(table2["Percent Loading"]),
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
    Compare two sheets across ACCA Long Term / ACCA / DCwAC.
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