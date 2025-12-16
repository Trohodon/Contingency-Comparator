import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np

EXPECTED_LEGEND_COLS = [
    "High Side V",
    "Low Side V",
    "Size MVA",
    "Typical Copper Losses (Watts)",
    "Typical Core Losses (Watts)",
]

# Common variants seen in real sheets (we normalize these)
LEGEND_COL_ALIASES = {
    "High Side V": ["High Side V", "High Side kV", "High kV", "HighSideV", "HighSide"],
    "Low Side V": ["Low Side V", "Low Side kV", "Low kV", "LowSideV", "LowSide"],
    "Size MVA": ["Size MVA", "MVA", "Lim MVA", "Rating MVA", "Size"],
    "Typical Copper Losses (Watts)": ["Typical Copper Losses (Watts)", "Copper Losses", "Copper Loss", "Load Typical Copper Losses (Watts)"],
    "Typical Core Losses (Watts)": ["Typical Core Losses (Watts)", "Core Losses", "Core Loss", "No load Typical Core Losses (Watts)"],
}

DATA_COL_ALIASES = {
    "From Nom kV": ["From Nom kV", "FromNomkV", "From kV", "From KV", "From Nom KV"],
    "To Nom kV": ["To Nom kV", "ToNomkV", "To kV", "To KV", "To Nom KV"],
    "Lim MVA": ["Lim MVA", "Limit MVA", "MVA", "Size MVA", "Rating MVA"],
}

def _normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _pick_col(df: pd.DataFrame, wanted: str, aliases: dict) -> str:
    """Find a column name in df that matches wanted or any alias."""
    cols = set(df.columns)
    for cand in aliases.get(wanted, [wanted]):
        if cand in cols:
            return cand
    # fallback: case-insensitive match
    lower_map = {c.lower(): c for c in df.columns}
    for cand in aliases.get(wanted, [wanted]):
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    raise KeyError(f"Could not find required column '{wanted}' (aliases tried: {aliases.get(wanted, [wanted])})")

def _to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")

def build_legend_lookup(legend_df: pd.DataFrame) -> pd.DataFrame:
    legend_df = _normalize_cols(legend_df)

    col_high = _pick_col(legend_df, "High Side V", LEGEND_COL_ALIASES)
    col_low  = _pick_col(legend_df, "Low Side V", LEGEND_COL_ALIASES)
    col_mva  = _pick_col(legend_df, "Size MVA", LEGEND_COL_ALIASES)
    col_cu   = _pick_col(legend_df, "Typical Copper Losses (Watts)", LEGEND_COL_ALIASES)
    col_core = _pick_col(legend_df, "Typical Core Losses (Watts)", LEGEND_COL_ALIASES)

    out = legend_df[[col_high, col_low, col_mva, col_cu, col_core]].copy()
    out.columns = ["High_kV", "Low_kV", "MVA", "Copper_W", "Core_W"]

    out["High_kV"] = _to_num(out["High_kV"])
    out["Low_kV"]  = _to_num(out["Low_kV"])
    out["MVA"]     = _to_num(out["MVA"])
    out["Copper_W"]= _to_num(out["Copper_W"])
    out["Core_W"]  = _to_num(out["Core_W"])

    out = out.dropna(subset=["High_kV", "Low_kV", "MVA", "Copper_W", "Core_W"])
    # Normalize as ints where appropriate (legend often uses whole kV values)
    out["High_kV"] = out["High_kV"].round(6)
    out["Low_kV"]  = out["Low_kV"].round(6)
    out["MVA"]     = out["MVA"].round(6)

    return out

def annotate_data(data_df: pd.DataFrame, legend: pd.DataFrame, allow_nearest_mva: bool, mva_tol: float) -> pd.DataFrame:
    data_df = _normalize_cols(data_df)

    col_from = _pick_col(data_df, "From Nom kV", DATA_COL_ALIASES)
    col_to   = _pick_col(data_df, "To Nom kV", DATA_COL_ALIASES)
    col_mva  = _pick_col(data_df, "Lim MVA", DATA_COL_ALIASES)

    out = data_df.copy()
    out["_From_kV"] = _to_num(out[col_from])
    out["_To_kV"]   = _to_num(out[col_to])
    out["_Lim_MVA"] = _to_num(out[col_mva])

    out["High_kV"] = out[["_From_kV", "_To_kV"]].max(axis=1)
    out["Low_kV"]  = out[["_From_kV", "_To_kV"]].min(axis=1)
    out["High_kV"] = out["High_kV"].round(6)
    out["Low_kV"]  = out["Low_kV"].round(6)
    out["MVA"]     = out["_Lim_MVA"].round(6)

    # Fast exact merge
    merged = out.merge(
        legend,
        how="left",
        left_on=["High_kV", "Low_kV", "MVA"],
        right_on=["High_kV", "Low_kV", "MVA"],
        suffixes=("", "_legend"),
    )

    merged["Match_Status"] = np.where(
        merged["Copper_W"].notna() & merged["Core_W"].notna(),
        "OK",
        "NO EXACT MATCH"
    )

    if allow_nearest_mva:
        # For rows with no exact match, find nearest MVA within same High/Low group
        need = merged["Match_Status"] != "OK"
        if need.any():
            legend_groups = {}
            for (hk, lk), g in legend.groupby(["High_kV", "Low_kV"]):
                legend_groups[(hk, lk)] = g.sort_values("MVA").reset_index(drop=True)

            copper = merged["Copper_W"].copy()
            core = merged["Core_W"].copy()
            status = merged["Match_Status"].copy()
            used_mva = pd.Series([np.nan]*len(merged), index=merged.index)

            idxs = merged.index[need].tolist()
            for i in idxs:
                hk = merged.at[i, "High_kV"]
                lk = merged.at[i, "Low_kV"]
                mva = merged.at[i, "MVA"]
                if pd.isna(hk) or pd.isna(lk) or pd.isna(mva):
                    status.at[i] = "MISSING kV/MVA"
                    continue
                g = legend_groups.get((hk, lk))
                if g is None or g.empty:
                    status.at[i] = "NO LEGEND FOR kV PAIR"
                    continue
                # nearest mva
                diffs = (g["MVA"] - mva).abs()
                j = int(diffs.idxmin())
                best = g.loc[j]
                if float(diffs.loc[j]) <= mva_tol:
                    copper.at[i] = best["Copper_W"]
                    core.at[i] = best["Core_W"]
                    used_mva.at[i] = best["MVA"]
                    status.at[i] = f"NEAREST MVA ({best['MVA']})"
                else:
                    status.at[i] = f"NO MATCH (nearest {best['MVA']}, diff {float(diffs.loc[j]):.3g})"

            merged["Copper_W"] = copper
            merged["Core_W"] = core
            merged["Legend_MVA_Used"] = used_mva
            merged["Match_Status"] = status

    merged["Total_W"] = merged[["Copper_W", "Core_W"]].sum(axis=1, min_count=1)

    # Clean helper cols
    merged = merged.drop(columns=[c for c in ["_From_kV", "_To_kV", "_Lim_MVA"] if c in merged.columns])
    return merged

def summarize_pairs(annotated: pd.DataFrame) -> pd.DataFrame:
    # Only sum numeric rows where we got values
    df = annotated.copy()
    df["Pair"] = df["High_kV"].astype("Int64").astype(str) + "-" + df["Low_kV"].astype("Int64").astype(str)

    summary = (
        df.groupby(["High_kV", "Low_kV", "Pair"], dropna=False)
        .agg(
            Count=("Pair", "size"),
            Matched=("Match_Status", lambda s: int((s == "OK").sum()) if hasattr(s, "__iter__") else 0),
            Copper_W_Sum=("Copper_W", "sum"),
            Core_W_Sum=("Core_W", "sum"),
            Total_W_Sum=("Total_W", "sum"),
        )
        .reset_index()
        .sort_values(["High_kV", "Low_kV"])
    )
    return summary

def run_process(input_path: str, output_path: str, legend_sheet: str, data_sheet: str,
                allow_nearest_mva: bool, mva_tol: float) -> None:
    xl = pd.ExcelFile(input_path)

    if legend_sheet not in xl.sheet_names:
        raise ValueError(f"Legend sheet '{legend_sheet}' not found. Available: {xl.sheet_names}")
    if data_sheet not in xl.sheet_names:
        raise ValueError(f"Data sheet '{data_sheet}' not found. Available: {xl.sheet_names}")

    legend_df = xl.parse(legend_sheet)
    data_df = xl.parse(data_sheet)

    legend = build_legend_lookup(legend_df)
    annotated = annotate_data(data_df, legend, allow_nearest_mva, mva_tol)
    summary = summarize_pairs(annotated)

    with pd.ExcelWriter(output_path, engine="openpyxl") as w:
        annotated.to_excel(w, sheet_name="Annotated", index=False)
        summary.to_excel(w, sheet_name="Summary", index=False)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("XFMR Loss Summarizer (Legend → Loss Lookup)")
        self.geometry("720x360")

        self.in_path = tk.StringVar()
        self.out_path = tk.StringVar()
        self.legend_sheet = tk.StringVar(value="Legend")
        self.data_sheet = tk.StringVar(value="2019 PF Case XFMRS (2)")
        self.allow_nearest = tk.BooleanVar(value=True)
        self.mva_tol = tk.StringVar(value="0.0")  # exact by default; set e.g. 5 for nearest within 5 MVA

        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(frm, text="Input .xlsx").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.in_path, width=70).grid(row=0, column=1, sticky="we", **pad)
        ttk.Button(frm, text="Browse", command=self.pick_input).grid(row=0, column=2, **pad)

        ttk.Label(frm, text="Output .xlsx").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.out_path, width=70).grid(row=1, column=1, sticky="we", **pad)
        ttk.Button(frm, text="Browse", command=self.pick_output).grid(row=1, column=2, **pad)

        ttk.Label(frm, text="Legend sheet name").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.legend_sheet, width=30).grid(row=2, column=1, sticky="w", **pad)

        ttk.Label(frm, text="Data sheet name").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.data_sheet, width=30).grid(row=3, column=1, sticky="w", **pad)

        ttk.Checkbutton(frm, text="Allow nearest-MVA fallback (within tolerance)",
                        variable=self.allow_nearest).grid(row=4, column=1, sticky="w", **pad)

        tol_row = ttk.Frame(frm)
        tol_row.grid(row=5, column=1, sticky="w", **pad)
        ttk.Label(tol_row, text="MVA tolerance:").pack(side="left")
        ttk.Entry(tol_row, textvariable=self.mva_tol, width=10).pack(side="left", padx=8)
        ttk.Label(tol_row, text="(0 = exact only; try 1, 5, etc if your Lim MVA isn’t exact)").pack(side="left")

        ttk.Separator(frm).grid(row=6, column=0, columnspan=3, sticky="we", pady=10)

        ttk.Button(frm, text="Run", command=self.run).grid(row=7, column=1, sticky="e", padx=10, pady=8)

        frm.columnconfigure(1, weight=1)

        note = (
            "Output file will contain:\n"
            " • Annotated: original rows + High_kV/Low_kV + Copper_W/Core_W/Total_W + Match_Status\n"
            " • Summary: sums grouped by High-Low pair (e.g., 230-115, 230-46, ...)\n\n"
            "Required Legend columns (aliases handled): High Side V, Low Side V, Size MVA, Typical Copper Losses, Typical Core Losses.\n"
            "Required Data columns (aliases handled): From Nom kV, To Nom kV, Lim MVA."
        )
        ttk.Label(frm, text=note, justify="left").grid(row=8, column=0, columnspan=3, sticky="w", padx=10, pady=8)

    def pick_input(self):
        path = filedialog.askopenfilename(
            title="Select input Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.in_path.set(path)
            if not self.out_path.get():
                self.out_path.set(path.replace(".xlsx", "_xfmr_losses_out.xlsx"))

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
        legend_sheet = self.legend_sheet.get().strip()
        data_sheet = self.data_sheet.get().strip()
        allow_nearest = bool(self.allow_nearest.get())

        try:
            tol = float(self.mva_tol.get().strip())
            if tol < 0:
                raise ValueError("MVA tolerance must be >= 0")
        except Exception:
            messagebox.showerror("Invalid tolerance", "MVA tolerance must be a number (>= 0).")
            return

        if not in_path:
            messagebox.showerror("Missing input", "Pick an input .xlsx file.")
            return
        if not out_path:
            messagebox.showerror("Missing output", "Pick an output .xlsx file.")
            return

        try:
            run_process(in_path, out_path, legend_sheet, data_sheet, allow_nearest, tol)
        except Exception as e:
            messagebox.showerror("Failed", str(e))
            return

        messagebox.showinfo("Done", f"Saved:\n{out_path}")

if __name__ == "__main__":
    App().mainloop()
