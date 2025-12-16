import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np

LEGEND_ROWS = [
    {"High_kV": 230, "Low_kV": 115, "MVA": 336,  "Core_W": 100000},
    {"High_kV": 230, "Low_kV": 115, "MVA": 224,  "Core_W":  80000},
    {"High_kV": 230, "Low_kV":  46, "MVA":  28,  "Core_W":  40000},
    {"High_kV": 230, "Low_kV":  33, "MVA":  93.3,"Core_W":  11500},
    {"High_kV": 230, "Low_kV":  23, "MVA":  37.3,"Core_W":  23000},
    {"High_kV": 230, "Low_kV":  13, "MVA":  56,  "Core_W":  29000},

    {"High_kV": 115, "Low_kV":  46, "MVA":  28,  "Core_W":  12000},
    {"High_kV": 115, "Low_kV":  33, "MVA":  50,  "Core_W":  19000},
    {"High_kV": 115, "Low_kV":  23, "MVA":  37.3,"Core_W":  18000},
    {"High_kV": 115, "Low_kV":  23, "MVA":  28,  "Core_W":  15000},
    {"High_kV": 115, "Low_kV":  23, "MVA":  22.4,"Core_W":  14000},
    {"High_kV": 115, "Low_kV":  23, "MVA":  10.5,"Core_W":  11000},
    {"High_kV": 115, "Low_kV":  13, "MVA":  22.4,"Core_W":  11000},
    {"High_kV": 115, "Low_kV":   8, "MVA":  22.4,"Core_W":  11000},
    {"High_kV": 115, "Low_kV":   4, "MVA":  22.4,"Core_W":  11000},

    {"High_kV":  46, "Low_kV":  23, "MVA":  10.5,"Core_W":  10000},
    {"High_kV":  46, "Low_kV":  13, "MVA":  10.5,"Core_W":   7000},
    {"High_kV":  46, "Low_kV":   8, "MVA":  10.5,"Core_W":   6000},
    {"High_kV":  46, "Low_kV":  0.5,"MVA":   3,  "Core_W":   5800},

    {"High_kV":  33, "Low_kV":  13, "MVA":  20,  "Core_W":   3000},
    {"High_kV":  33, "Low_kV":   8, "MVA":  10.5,"Core_W":   2000},
]

LEGEND = pd.DataFrame(LEGEND_ROWS).copy()
LEGEND["High_kV"] = pd.to_numeric(LEGEND["High_kV"], errors="coerce").astype(float).round(6)
LEGEND["Low_kV"]  = pd.to_numeric(LEGEND["Low_kV"], errors="coerce").astype(float).round(6)
LEGEND["MVA"]     = pd.to_numeric(LEGEND["MVA"], errors="coerce").astype(float).round(6)
LEGEND["Core_W"]  = pd.to_numeric(LEGEND["Core_W"], errors="coerce")

DATA_COL_ALIASES = {
    "From": ["From", "FROM", "From kV", "From KV"],
    "To": ["To", "TO", "To kV", "To KV"],
    "Lim MVA": ["Lim MVA", "Lim MVA A", "Limit MVA", "MVA", "Size MVA", "Rating MVA"],
}

def _normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _pick_col(df: pd.DataFrame, wanted: str) -> str:
    cols = list(df.columns)
    for cand in DATA_COL_ALIASES.get(wanted, [wanted]):
        if cand in cols:
            return cand
    lower_map = {c.lower(): c for c in cols}
    for cand in DATA_COL_ALIASES.get(wanted, [wanted]):
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    raise KeyError(f"Could not find required column '{wanted}'")

def _to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")

def fmt_kv(x) -> str:
    if pd.isna(x):
        return ""
    x = float(x)
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    return str(x).rstrip("0").rstrip(".")

def read_sheet_with_auto_header(xlsx_path: str, sheet_name: str, scan_rows: int = 200) -> pd.DataFrame:
    preview = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None, nrows=scan_rows, dtype=str)
    preview = preview.fillna("").astype(str).apply(lambda col: col.str.strip())

    req_from = set(a.lower() for a in DATA_COL_ALIASES["From"])
    req_to   = set(a.lower() for a in DATA_COL_ALIASES["To"])
    req_mva  = set(a.lower() for a in DATA_COL_ALIASES["Lim MVA"])

    header_row = None
    for r in range(len(preview)):
        row_vals = set(v.lower() for v in preview.iloc[r].tolist() if v)
        if (row_vals & req_from) and (row_vals & req_to) and (row_vals & req_mva):
            header_row = r
            break

    if header_row is None:
        raise ValueError("Could not auto-detect header row (missing From/To/Lim MVA).")

    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=header_row)
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    return df

def match_core_losses(df: pd.DataFrame, allow_nearest_mva: bool, mva_tol: float) -> pd.DataFrame:
    df = _normalize_cols(df)
    c_from = _pick_col(df, "From")
    c_to   = _pick_col(df, "To")
    c_mva  = _pick_col(df, "Lim MVA")

    out = df.copy()
    out["From_kV"] = _to_num(out[c_from]).astype(float).round(6)
    out["To_kV"]   = _to_num(out[c_to]).astype(float).round(6)
    out["MVA"]     = _to_num(out[c_mva]).astype(float).round(6)

    out["High_kV"] = out[["From_kV", "To_kV"]].max(axis=1).round(6)
    out["Low_kV"]  = out[["From_kV", "To_kV"]].min(axis=1).round(6)

    merged = out.merge(LEGEND, how="left", on=["High_kV", "Low_kV", "MVA"])
    merged["Legend_MVA_Used"] = np.where(merged["Core_W"].notna(), merged["MVA"], np.nan)
    merged["Match_Status"] = np.where(merged["Core_W"].notna(), "OK", "NO EXACT MATCH")

    if allow_nearest_mva:
        need = merged["Core_W"].isna() & merged["High_kV"].notna() & merged["Low_kV"].notna() & merged["MVA"].notna()
        if need.any():
            legend_groups = {k: g.sort_values("MVA").reset_index(drop=True)
                             for k, g in LEGEND.groupby(["High_kV", "Low_kV"])}

            for i in merged.index[need]:
                hk = merged.at[i, "High_kV"]
                lk = merged.at[i, "Low_kV"]
                mva = merged.at[i, "MVA"]

                g = legend_groups.get((hk, lk))
                if g is None or g.empty:
                    merged.at[i, "Match_Status"] = "NO LEGEND FOR kV PAIR"
                    continue

                diffs = (g["MVA"] - mva).abs()
                j = int(diffs.idxmin())
                best = g.loc[j]
                diff_val = float(diffs.loc[j])

                if diff_val <= mva_tol:
                    merged.at[i, "Core_W"] = best["Core_W"]
                    merged.at[i, "Legend_MVA_Used"] = best["MVA"]
                    merged.at[i, "Match_Status"] = "NEAREST"
                else:
                    merged.at[i, "Match_Status"] = "NO MATCH"

    merged["HighToLow"] = merged["High_kV"].apply(fmt_kv) + "-" + merged["Low_kV"].apply(fmt_kv)
    return merged

def build_breakdown_and_summary(merged: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    matched = merged[merged["Core_W"].notna()].copy()
    unmatched = merged[merged["Core_W"].isna()].copy()

    breakdown = (
        matched.groupby(["High_kV", "Low_kV", "HighToLow", "Core_W"], dropna=False)
               .size()
               .reset_index(name="Count")
    )
    breakdown["Bucket_Total_Core_W"] = breakdown["Count"] * breakdown["Core_W"]
    breakdown = breakdown.sort_values(["High_kV", "Low_kV", "Core_W"], ascending=[False, False, False])

    summary = (
        breakdown.groupby(["High_kV", "Low_kV", "HighToLow"], dropna=False)
                 .agg(
                     Total_Count=("Count", "sum"),
                     Core_W_Sum=("Bucket_Total_Core_W", "sum"),
                 )
                 .reset_index()
                 .sort_values(["High_kV", "Low_kV"], ascending=[False, False])
    )

    if not unmatched.empty:
        unmatched_view = unmatched[["From_kV", "To_kV", "MVA", "High_kV", "Low_kV", "HighToLow", "Match_Status"]].copy()
        unmatched_view = unmatched_view.sort_values(["High_kV", "Low_kV"], ascending=[False, False])
    else:
        unmatched_view = pd.DataFrame(columns=["From_kV", "To_kV", "MVA", "High_kV", "Low_kV", "HighToLow", "Match_Status"])

    return breakdown, summary, unmatched_view

def run_process(input_path: str, sheet_name: str, output_path: str,
                allow_nearest_mva: bool, mva_tol: float) -> None:
    data_df = read_sheet_with_auto_header(input_path, sheet_name)
    merged = match_core_losses(data_df, allow_nearest_mva, mva_tol)
    breakdown, summary, unmatched = build_breakdown_and_summary(merged)

    with pd.ExcelWriter(output_path, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name="Summary", index=False)
        breakdown.to_excel(w, sheet_name="Breakdown", index=False)
        unmatched.to_excel(w, sheet_name="Unmatched", index=False)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Core Loss Summarizer")
        self.geometry("820x330")

        self.in_path = tk.StringVar()
        self.out_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.allow_nearest = tk.BooleanVar(value=True)
        self.mva_tol = tk.StringVar(value="0.0")

        pad = {"padx": 10, "pady": 6}
        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(frm, text="Input .xlsx").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.in_path, width=75).grid(row=0, column=1, sticky="we", **pad)
        ttk.Button(frm, text="Browse", command=self.pick_input).grid(row=0, column=2, **pad)

        ttk.Label(frm, text="Sheet to process").grid(row=1, column=0, sticky="w", **pad)
        self.sheet_combo = ttk.Combobox(frm, textvariable=self.sheet_name, width=45, state="readonly", values=[])
        self.sheet_combo.grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(frm, text="Output .xlsx").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.out_path, width=75).grid(row=2, column=1, sticky="we", **pad)
        ttk.Button(frm, text="Browse", command=self.pick_output).grid(row=2, column=2, **pad)

        ttk.Checkbutton(frm, text="Allow nearest-MVA lookup (within tolerance)", variable=self.allow_nearest)\
            .grid(row=3, column=1, sticky="w", **pad)

        tol_row = ttk.Frame(frm)
        tol_row.grid(row=4, column=1, sticky="w", **pad)
        ttk.Label(tol_row, text="MVA tolerance:").pack(side="left")
        ttk.Entry(tol_row, textvariable=self.mva_tol, width=10).pack(side="left", padx=8)

        ttk.Separator(frm).grid(row=5, column=0, columnspan=3, sticky="we", pady=10)
        ttk.Button(frm, text="Run", command=self.run).grid(row=6, column=1, sticky="e", padx=10, pady=10)

        frm.columnconfigure(1, weight=1)

    def pick_input(self):
        path = filedialog.askopenfilename(
            title="Select input Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not path:
            return

        self.in_path.set(path)
        try:
            xl = pd.ExcelFile(path)
            sheets = xl.sheet_names
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read workbook sheets:\n{e}")
            return

        self.sheet_combo["values"] = sheets
        if sheets:
            self.sheet_name.set(sheets[0])

        if not self.out_path.get():
            self.out_path.set(path.replace(".xlsx", "_coreloss_summary.xlsx"))

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
                raise ValueError()
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