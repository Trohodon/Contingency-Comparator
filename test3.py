import os
import pandas as pd

def find_column(df, aliases):
    """
    Find a column in df whose name (lowercased) either matches or contains
    one of the alias strings. Returns the column name or None.
    """
    cols_lower = {col.lower(): col for col in df.columns}

    # Exact matches first
    for alias in aliases:
        if alias in cols_lower:
            return cols_lower[alias]

    # Then partial matches
    for col in df.columns:
        col_lower = col.lower()
        for alias in aliases:
            if alias in col_lower:
                return col
    return None


def main():
    print("=== PowerWorld Test 2 Results Filter ===")
    print("This will take your full Test 2 CSV and output a filtered CSV")
    print("containing only Line / Transformer limit violations.\n")

    # ---- 1. Ask user for input/output paths ----
    input_path = input("Enter path to Test 2 CSV (press Enter for 'test2_output.csv'): ").strip()
    if not input_path:
        input_path = "test2_output.csv"

    if not os.path.isfile(input_path):
        print(f"\n[ERROR] File not found: {input_path}")
        return

    default_output = "filtered_contingencies.csv"
    output_path = input(f"Enter path for filtered CSV (press Enter for '{default_output}'): ").strip()
    if not output_path:
        output_path = default_output

    print(f"\n[INFO] Reading input file: {input_path}")
    df = pd.read_csv(input_path)

    print(f"[INFO] Loaded {len(df)} rows with {len(df.columns)} columns.")
    print("[INFO] Columns found:")
    for col in df.columns:
        print(f"   - {col}")
    print()

    # ---- 2. Locate key columns (best-effort, flexible) ----
    contingency_col = find_column(df, [
        "contingency", "ctgname", "ctg_name", "contingency name"
    ])

    object_name_col = find_column(df, [
        "device", "element", "branch", "object", "element name", "object name"
    ])

    from_bus_col = find_column(df, [
        "from bus", "frombus", "bus from", "busfrom", "bus num from", "bus from num"
    ])

    to_bus_col = find_column(df, [
        "to bus", "tobus", "bus to", "busto", "bus num to", "bus to num"
    ])

    limit_col = find_column(df, [
        "limit", "rating", "mw limit", "amp limit", "mva limit"
    ])

    value_col = find_column(df, [
        "value", "flow", "mw", "mvar", "amps", "mva"
    ])

    percent_col = find_column(df, [
        "percent", "% of limit", "pct", "pctoflimit", "percent of limit"
    ])

    object_type_col = find_column(df, [
        "objecttype", "object type", "category", "device type", "element type"
    ])

    print("[INFO] Mapped columns:")
    print(f"   Contingency:  {contingency_col}")
    print(f"   Element name: {object_name_col}")
    print(f"   From bus:     {from_bus_col}")
    print(f"   To bus:       {to_bus_col}")
    print(f"   Limit:        {limit_col}")
    print(f"   Value:        {value_col}")
    print(f"   Percent:      {percent_col}")
    print(f"   Object type:  {object_type_col}")
    print()

    # ---- 3. Filter to lines / transformers (if we can) ----
    if object_type_col is not None:
        print("[INFO] Filtering rows to LINE / TRANSFORMER types using object type column...")
        type_series = df[object_type_col].astype(str).str.lower()

        mask_type = type_series.str.contains("branch|line|xfmr|transformer", regex=True, na=False)
        df = df[mask_type].copy()
        print(f"[INFO] After type filter: {len(df)} rows remain.")
    else:
        print("[WARN] No object type column found. Skipping line/transformer filter.")
        print("       (You may want to adjust the script once you know the exact column name.)")

    # ---- 4. Filter to actual violations using percent > 100, if we have that ----
    if percent_col is not None:
        print("[INFO] Filtering to actual limit violations where Percent > 100...")
        # Convert to numeric safely
        pct = pd.to_numeric(df[percent_col], errors="coerce")
        df = df[pct > 100].copy()
        print(f"[INFO] After violation filter: {len(df)} rows remain.")
    else:
        print("[WARN] No percent-of-limit column found. Keeping all rows (no violation filter).")

    # ---- 5. Build final output dataframe with just the columns you care about ----
    selected_cols = []
    col_order = [
        ("Contingency", contingency_col),
        ("Element", object_name_col),
        ("FromBus", from_bus_col),
        ("ToBus", to_bus_col),
        ("Limit", limit_col),
        ("Value", value_col),
        ("PercentOfLimit", percent_col)
    ]

    for friendly_name, real_col in col_order:
        if real_col is not None:
            df[friendly_name] = df[real_col]
            selected_cols.append(friendly_name)
        else:
            print(f"[WARN] Could not find column for: {friendly_name}")

    if not selected_cols:
        print("[ERROR] No output columns could be mapped. Nothing to write.")
        return

    filtered_df = df[selected_cols].copy()

    # ---- 6. Sort by highest percent violation (if we have it) ----
    if "PercentOfLimit" in filtered_df.columns:
        filtered_df["PercentOfLimit"] = pd.to_numeric(filtered_df["PercentOfLimit"], errors="coerce")
        filtered_df = filtered_df.sort_values(by="PercentOfLimit", ascending=False)

    # ---- 7. Save result ----
    filtered_df.to_csv(output_path, index=False)
    print(f"\n[DONE] Wrote {len(filtered_df)} filtered violation rows to:")
    print(f"       {output_path}\n")


if __name__ == "__main__":
    main()