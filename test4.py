import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np

# -----------------------------
# HARDCODED LEGEND (from your screenshot)
# High Side V, Low Side V, Size MVA -> Copper/Core losses
# -----------------------------
LEGEND_ROWS = [
    # 230 kV high side
    {"High_kV": 230, "Low_kV": 115, "MVA": 336,   "Copper_W": 150000, "Core_W": 100000},
    {"High_kV": 230, "Low_kV": 115, "MVA": 224,   "Copper_W": 120000, "Core_W":  80000},
    {"High_kV": 230, "Low_kV":  46, "MVA":  28,   "Copper_W":  80000, "Core_W":  40000},
    {"High_kV": 230, "Low_kV":  33, "MVA":  93.3, "Copper_W":  47000, "Core_W":  11500},
    {"High_kV": 230, "Low_kV":  23, "MVA":  37.3, "Copper_W":  66500, "Core_W":  23000},
    {"High_kV": 230, "Low_kV":  13, "MVA":  56,   "Copper_W":  57000, "Core_W":  29000},

    # 115 kV high side
    {"High_kV": 115, "Low_kV":  46, "MVA":  28,   "Copper_W":  50000, "Core_W":  12000},
    {"High_kV": 115, "Low_kV":  33, "MVA":  50,   "Copper_W":  50000, "Core_W":  19000},
    {"High_kV": 115, "Low_kV":  23, "MVA":  37.3, "Copper_W":  55000, "Core_W":  18000},
    {"High_kV": 115, "Low_kV":  23, "MVA":  28,   "Copper_W":  45000, "Core_W":  15000},
    {"High_kV": 115, "Low_kV":  23, "MVA":  22.4, "Copper_W":  40000, "Core_W":  14000},
    {"High_kV": 115, "Low_kV":  23, "MVA":  10.5, "Copper_W":  30000, "Core_W":  11000},
    {"High_kV": 115, "Low_kV":  13, "MVA":  22.4, "Copper_W":  42000, "Core_W":  11000},
    {"High_kV": 115, "Low_kV":   8, "MVA":  22.4, "Copper_W":  40000, "Core_W":  11000},
    {"High_kV": 115, "Low_kV":   4, "MVA":  22.4, "Copper_W":  40000, "Core_W":  11000},

    # 46 kV high side
    {"High_kV":  46, "Low_kV":  23, "MVA":  10.5, "Copper_W":  35000, "Core_W":  10000},
    {"High_kV":  46, "Low_kV":  13, "MVA":  10.5, "Copper_W":  33000, "Core_W":   7000},
    {"High_kV":  46, "Low_kV":   8, "MVA":  10.5, "Copper_W":  25000, "Core_W":   6000},
    {"High_kV":  46, "Low_kV":  0.5,"MVA":   3,   "Copper_W":  12000, "Core_W":   5800},

    # 33 kV high side
    {"High_kV":  33, "Low_kV":  13, "MVA":  20,   "Copper_W":   7000, "Core_W":   3000},
    {"High_kV":  33, "Low_kV":   8, "MVA":  10.5, "Copper_W":   5000, "Core_W":   2000},
]

LEGEND = pd.DataFrame(LEGEND_ROWS).copy()
LEGEND["High_kV"] = LEGEND["High_kV"].astype(float).round(6)
LEGEND["Low_kV"]  = LEGEND["Low_kV"].astype(float).round(6)
LEGEND["MVA"]     = LEGEND["MVA"].astype(float).round(6)

# -----------------------------
# Column name aliases (your sheet has "Lim MVA A" sometimes)
# -----------------------------
DATA_COL_ALIASES = {
    "From Nom kV": ["From Nom kV", "FromNomkV", "From kV", "From KV", "From Nom KV"],
    "To Nom kV":   ["To Nom kV", "ToNomkV", "To kV", "To KV", "To Nom KV"],
    "Lim MVA":     ["Lim MVA", "Lim MVA A", "Limit MVA", "MVA", "Size MVA", "Rating MVA"],
}

def _normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _pick_col(df: pd.DataFrame, wanted: str, aliases: dict) -> str:
    cols = set(df.columns)
    for cand in aliases.get(wanted, [wanted]):
        if cand in cols:
            return cand
    lower_map = {c.lower(): c for c in df.columns}
    for cand in aliases.get(wanted, [wanted]):
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    raise KeyError(f"Could not find required column '{wanted}'. Tried: {aliases.get(wanted, [wanted])}")

def _to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")

def read_sheet_with_auto_header(xlsx_path: str, sheet_name: str, scan_rows: int = 200) -> pd.DataFrame:
    """
    Your workbook has rows above the actual table header.
    This scans the first `scan_rows` rows to find the row that contains
    From Nom kV / To Nom kV / Lim MVA (or Lim MVA A), then reads with that header.
    """
    preview = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None, nrows=scan_rows, dtype=str)
    preview = preview.fillna("").astype(str).applymap(lambda x: x.strip())

    req_from = set(a.lower() for a in DATA_COL_ALIASES["From Nom kV"])
    req_to   = set(a.lower() for a in DATA_COL_ALIASES["To Nom kV"])
    req_mva  = set(a.lower() for a in DATA_COL_ALIASES["Lim MVA"])

    header_row = None
    for r in range(len(preview)):
        row_vals = set(v.lower() for v in preview.iloc[r].tolist() if v)
        if (row_vals & req_from) and (row_vals & req_to) and (row_vals & req_mva):
            header_row = r
            break

    if header_row is None:
        raise ValueError(
            "Could not auto-detect header row.\n"
            "I scanned the first rows but did not see From Nom kV / To Nom kV / Lim MVA.\n"
            "If your headers are different, tell me the exact text."
        )

    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=header_row)
    df = _normalize_cols(df)

    # Drop unnamed columns that come from formatting
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed", case=False)]
    return df

def annotate_data(data_df: pd.DataFrame, allow_nearest_mva: bool, mva_tol: float) -> pd.DataFrame:
    df = _normalize_cols(data_df)

    c_from = _pick_col(df, "From Nom kV", DATA_COL_ALIASES)
    c_to   = _pick_col(df, "To Nom kV",   DATA_COL_ALIASES)
    c_mva  = _pick_col(df, "Lim MVA",     DATA_COL_ALIASES)

    out = df.copy()
    out["_From_kV"] = _to_num(out[c_from])
    out["_To_kV"]   = _to_num(out[c_to])
    out["_MVA"]     = _to_num(out[c_mva])

    out["High_kV"] = out[["_From_kV", "_To_kV"]].max(axis=1).round(6)
    out["Low_kV"]  = out[["_From_kV", "_To_kV"]].min(axis=1).round(6)
    out["MVA"]     = out["_MVA"].round(6)

    merged = out.merge(
        LEGEND,
        how="left",
        on=["High_kV", "Low_kV", "MVA"],
        suffixes=("", "_legend"),
    )

    merged["Match_Status"] = np.where(
        merged["Copper_W"].notna() & merged["Core_W"].notna(),
        "OK",
        "NO EXACT MATCH"
    )

    if allow_nearest_mva:
        need = merged["Match_Status"] != "OK"
        if need.any():
            legend_groups = {k: g.sort_values("MVA").reset_index(drop=True)
                             for k, g in LEGEND.groupby(["High_kV", "Low_kV"])}

            used_mva = pd.Series([np.nan] * len(merged), index=merged.index)

            for i in merged.index[need]:
                hk = merged.at[i, "High_kV"]
                lk = merged.at[i, "Low_kV"]
                mva = merged.at[i, "MVA"]

                if pd.isna(hk) or pd.isna(lk) or pd.isna(mva):
                    merged.at[i, "Match_Status"] = "MISSING kV/MVA"
                    continue

                g = legend_groups.get((hk, lk))
                if g is None or g.empty:
                    merged.at[i, "Match_Status"] = "NO LEGEND FOR kV PAIR"
                    continue

                diffs = (g["MVA"] - mva).abs()
                j = int(diffs.idxmin())
                best = g.loc[j]
                diff_val = float(diffs.loc[j])

                if diff_val <= mva_tol:
                    merged.at[i, "Copper_W"] = best["Copper_W"]
                    merged.at[i, "Core_W"] = best["Core_W"]
                    used_mva.at[i] = best["MVA"]
                    merged.at[i, "Match_Status"] = f"NEAREST MVA ({best['MVA']})"
                else:
                    merged.at[i, "Match_Status"] = f"NO MATCH (nearest {best['MVA']}, diff {diff_val:g})"

            merged["Legend_MVA_Used"] = used_mva

    merged["Total_W"] = merged[["Copper_W", "Core_W"]].sum(axis=1, min_count=1)
    merged["Pair"] = merged["High_kV"].astype("Int64").astype(str) + "-" + merged["Low_kV"].astype("Int64").astype(str)

    merged.drop(columns=[c for c in ["_From_kV", "_To_kV", "_MVA"] if c in merged.columns], inplace=True)
    return merged

def summarize_pairs(annotated: pd.DataFrame) -> pd.DataFrame:
    df = annotated.copy()
    summary = (
        df.groupby(["High_kV", "Low_kV", "Pair"], dropna=False)
          .agg(
              Count=("Pair", "size"),
              Copper_W_Sum=("Copper_W", "sum"),
              Core_W_Sum=("Core_W", "sum"),
              Total_W_Sum=("Total_W", "sum"),
              OK_Matches=("Match_Status", lambda s: int((s == "OK").sum())),
              NonExact=("Match_Status", lambda s: int((s != "OK").sum())),
          )
          .reset_index()
          .sort_values(["High_kV", "Low_kV"])
    )
    return summary

def run_process(input_path: str, sheet_name: str, output_path: str,
                allow_nearest_mva: bool, mva_tol: float) -> None:
    data_df = read_sheet_with_auto_header(input_path, sheet_name)
    annotated = annotate_data(data_df, allow_nearest_mva, mva_tol)
    summary = summarize_pairs(annotated)

    with pd.ExcelWriter(output_path, engine="openpyxl") as w:
        annotated.to_excel(w, sheet_name="Annotated", index=False)
        summary.to_excel(w, sheet_name="Summary", index=False)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("XFMR Loss Summarizer (Hardcoded Legend)")
        self.geometry("860x440")

        self.in_path = tk.StringVar()
        self.out_path = tk.StringVar()
        self.sheet_name = tk.StringVar()

        self.allow_nearest = tk.BooleanVar(value=True)
        self.mva_tol = tk.StringVar(value="0.0")

        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(frm, text="Input .xlsx").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.in_path, width=78).grid(row=0, column=1, sticky="we", **pad)
        ttk.Button(frm, text="Browse", command=self.pick_input).grid(row=0, column=2, **pad)

        ttk.Label(frm, text="Sheet to process").grid(row=1, column=0, sticky="w", **pad)
        self.sheet_combo = ttk.Combobox(frm, textvariable=self.sheet_name, width=52, state="readonly", values=[])
        self.sheet_combo.grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(frm, text="Output .xlsx").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.out_path, width=78).grid(row=2, column=1, sticky="we", **pad)
        ttk.Button(frm, text="Browse", command=self.pick_output).grid(row=2, column=2, **pad)

        ttk.Checkbutton(frm, text="Allow nearest-MVA fallback (within tolerance)",
                        variable=self.allow_nearest).grid(row=3, column=1, sticky="w", **pad)

        tol_row = ttk.Frame(frm)
        tol_row.grid(row=4, column=1, sticky="w", **pad)
        ttk.Label(tol_row, text="MVA tolerance:").pack(side="left")
        ttk.Entry(tol_row, textvariable=self.mva_tol, width=10).pack(side="left", padx=8)
        ttk.Label(tol_row, text="(0 = exact only; try 1, 5, etc)").pack(side="left")

        ttk.Separator(frm).grid(row=5, column=0, columnspan=3, sticky="we", pady=10)

        ttk.Button(frm, text="Run", command=self.run).grid(row=6, column=1, sticky="e", padx=10, pady=10)

        frm.columnconfigure(1, weight=1)

        note = (
            "This auto-detects the header row (because your sheet has rows above the table).\n"
            "Required columns must exist somewhere in the header row:\n"
            "  • From Nom kV\n"
            "  • To Nom kV\n"
            "  • Lim MVA (or Lim MVA A)\n\n"
            "Output:\n"
            "  • Annotated: adds High/Low kV + Copper/Core/Total + Match_Status\n"
            "  • Summary: sums grouped by High-Low pair (230-115, 230-46, ...)\n"
        )
        ttk.Label(frm, text=note, justify="left").grid(row=7, column=0, columnspan=3, sticky="w", padx=10, pady=8)

    def pick_input(self):
        path = filedialog.askopenfilename(
            title="Select input Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not path:
            return

        self.in_path.set(path)

        # Populate sheet dropdown
        try:
            xl = pd.ExcelFile(path)
            sheets = xl.sheet_names
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read workbook sheets:\n{e}")
            return

        self.sheet_combo["values"] = sheets
        if sheets:
            # Guess a likely sheet
            guess = None
            for s in sheets:
                if "xfmr" in s.lower() or "xfmrs" in s.lower():
                    guess = s
                    break
            self.sheet_name.set(guess if guess else sheets[0])

        if not self.out_path.get():
            self.out_path.set(path.replace(".xlsx", "_xfmr_loss_summary.xlsx"))

    def pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Save output Excel file as",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.out_path.set(path)

    def run(self):
        in_path = self.in_path.get().strip()
        out_path = self.out_path.get().strip()
        sheet = self.sheet_name.get().strip()

        if not in_path:
            messagebox.showerror("Missing input", "Pick an input .xlsx file.")
            return
        if not sheet:
            messagebox.showerror("Missing sheet", "Pick a sheet from the dropdown.")
            return
        if not out_path:
            messagebox.showerror("Missing output", "Pick an output .xlsx file.")
            return

        try:
            tol = float(self.mva_tol.get().strip())
            if tol < 0:
                raise ValueError("MVA tolerance must be >= 0")
        except Exception:
            messagebox.showerror("Invalid tolerance", "MVA tolerance must be a number (>= 0).")
            return

        try:
            run_process(in_path, sheet, out_path, bool(self.allow_nearest.get()), tol)
        except Exception as e:
            messagebox.showerror("Failed", str(e))
            return

        messagebox.showinfo("Done", f"Saved:\n{out_path}")

if __name__ == "__main__":
    App().mainloop()
