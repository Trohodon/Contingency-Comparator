import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import win32com.client
import pandas as pd


class PwbExportApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("PowerWorld Contingency Results Export")
        self.geometry("800x500")

        self.pwb_path = tk.StringVar(value="No .pwb file selected")
        self.csv_path = None

        self._build_gui()

    # ───────────── GUI LAYOUT ───────────── #

    def _build_gui(self):
        top = ttk.Frame(self)
        top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        ttk.Label(top, text="Selected .pwb case:").grid(row=0, column=0, sticky="w")
        ttk.Label(top, textvariable=self.pwb_path, width=80).grid(
            row=1, column=0, columnspan=2, sticky="w"
        )

        browse_btn = ttk.Button(top, text="Browse...", command=self.browse_pwb)
        browse_btn.grid(row=1, column=2, padx=(5, 0), sticky="e")

        run_btn = ttk.Button(
            top,
            text="Export existing contingency results",
            command=self.run_export,
        )
        run_btn.grid(row=2, column=0, columnspan=3, pady=(10, 0), sticky="w")

        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=10, pady=5)

        log_frame = ttk.Frame(self)
        log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(log_frame, text="Log:").pack(anchor="w")

        self.log_text = tk.Text(log_frame, wrap="word", height=15)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log_text.yview,
        )
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scroll.set)

    def log(self, msg: str):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    # ───────────── CALLBACKS ───────────── #

    def browse_pwb(self):
        path = filedialog.askopenfilename(
            title="Select PowerWorld case (.pwb)",
            filetypes=[("PowerWorld case", "*.pwb"), ("All files", "*.*")],
        )
        if path:
            self.pwb_path.set(path)
            self.csv_path = None
            self.log(f"Selected case: {path}")

    def run_export(self):
        pwb = self.pwb_path.get()
        if not pwb.lower().endswith(".pwb") or not os.path.exists(pwb):
            messagebox.showwarning(
                "No case selected",
                "Please select a valid .pwb file.",
            )
            return

        base, _ = os.path.splitext(pwb)
        csv_out = base + "_Violations.csv"
        self.csv_path = csv_out

        try:
            self._export_existing_results(pwb, csv_out)
        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))

    # ────────── POWERWORLD EXPORT (existing results only) ────────── #

    def _export_existing_results(self, pwb_path: str, csv_out: str):
        self.log("Connecting to PowerWorld via SimAuto...")
        simauto = win32com.client.Dispatch("pwrworld.SimulatorAuto")
        self.log("Connected.")

        # 1) Open case – we assume contingencies are already solved in this file
        self.log(f"Opening case: {pwb_path}")
        (err,) = simauto.OpenCase(pwb_path)
        if err:
            raise RuntimeError(f"OpenCase error: {err}")
        self.log("Case opened successfully (existing results will be used).")

        # 2) Go to Contingency mode (does NOT re-run analysis)
        self.log("Entering Contingency mode...")
        (err,) = simauto.RunScriptCommand("EnterMode(Contingency);")
        if err:
            raise RuntimeError(f"EnterMode(Contingency) error: {err}")

        # 3) Export the stored violation matrices to CSV (branches = lines + transformers)
        self.log(f"Saving stored violation matrices to CSV:\n  {csv_out}")
        clean_csv = csv_out.replace("\\", "/")
        cmd = (
            f'CTGSaveViolationMatrices("{clean_csv}", CSVCOLHEADER, '
            "YES, [BRANCH], YES, NO);"
        )
        (err,) = simauto.RunScriptCommand(cmd)
        if err:
            raise RuntimeError(f"CTGSaveViolationMatrices error: {err}")
        self.log("CSV export complete (using existing CA results).")

        # Close SimAuto
        try:
            simauto.CloseCase()
        except Exception:
            pass
        del simauto

        # 4) Load the raw CSV and build a filtered summary
        if os.path.exists(csv_out):
            self.log("\nPreview of first few rows (raw CTG matrix):")
            try:
                df = pd.read_csv(csv_out)
                self.log(df.head(5).to_string(index=False))

                summary = self._build_branch_summary(df)

                if summary is not None and not summary.empty:
                    summary_path = csv_out.replace(".csv", "_summary.csv")
                    summary.to_csv(summary_path, index=False)

                    self.log("\nSaved filtered line/transformer summary to:")
                    self.log(f"  {summary_path}")
                    self.log("\nPreview of summary (first 20 rows):")
                    self.log(summary.head(20).to_string(index=False))
                else:
                    self.log(
                        "\nNo matching branch/transformer violations found "
                        "for summary (or could not auto-detect columns)."
                    )

            except Exception as e:
                self.log(f"(Could not read or summarize CSV: {e})")
        else:
            self.log("WARNING: CSV file does not exist after export.")

        messagebox.showinfo(
            "Done",
            f"Violations exported to:\n{csv_out}\n\n"
            f"Filtered summary (lines/transformers) is in:\n"
            f"{csv_out.replace('.csv', '_summary.csv')}",
        )

    # ────────── SUMMARY BUILDER ────────── #

    def _build_branch_summary(self, df: pd.DataFrame):
        """
        Build a compact summary for line/transformer contingencies:
        - Contingency name
        - Affected line (label + optional From/To)
        - Limit
        - Violation value
        - Percent loading
        """

        if df is None or df.empty:
            return None

        df2 = df.copy()

        # Helper to find first column whose name contains any of given substrings
        def find_col(substrings, exclude=None):
            if exclude is None:
                exclude_local = []
            else:
                exclude_local = [e.lower() for e in exclude]

            for c in df2.columns:
                cl = c.lower()
                if any(s in cl for s in substrings) and not any(
                    e in cl for e in exclude_local
                ):
                    return c
            return None

        # Contingency column
        ctg_col = find_col(["ctg", "contingency", "cont"])
        # Element / line label column
        elem_col = None
        for c in df2.columns:
            cl = c.lower()
            if (
                any(s in cl for s in ["mon", "element", "object", "branch", "line"])
                and any(s in cl for s in ["label", "name", "id"])
            ):
                elem_col = c
                break
        if elem_col is None and len(df2.columns) > 0:
            # Fallback to first column if we cannot detect anything better
            elem_col = df2.columns[0]

        # From / To bus columns (if available)
        from_bus_col = None
        to_bus_col = None
        for c in df2.columns:
            cl = c.lower()
            if "from" in cl and "bus" in cl:
                from_bus_col = c
            if "to" in cl and "bus" in cl:
                to_bus_col = c

        # Limit, value, percent columns
        percent_col = find_col(["percent", "%"])
        limit_col = find_col(["limit"], exclude=["percent"])
        value_col = find_col(
            ["value", "flow", "mw", "mva", "amp"],
            exclude=["limit", "percent"],
        )

        # Optional type column to restrict to branches (in case other types exist)
        type_col = None
        for c in df2.columns:
            cl = c.lower()
            if "objecttype" in cl or cl == "type" or "type " in cl:
                type_col = c
                break

        # Start with all rows
        mask = pd.Series(True, index=df2.index)

        # Restrict to branches (lines/transformers) if type column exists
        if type_col is not None:
            tvals = df2[type_col].astype(str).str.lower()
            mask &= tvals.str.contains("branch") | tvals.str.contains("xfmr")

        # Restrict to actual violations: Percent >= 100 or Value > Limit
        if percent_col is not None:
            try:
                perc = pd.to_numeric(df2[percent_col], errors="coerce")
                mask &= perc >= 100.0
            except Exception:
                pass
        elif value_col is not None and limit_col is not None:
            try:
                val = pd.to_numeric(df2[value_col], errors="coerce")
                lim = pd.to_numeric(df2[limit_col], errors="coerce")
                mask &= (val.notna() & lim.notna() & (val > lim))
            except Exception:
                pass

        filtered = df2[mask].copy()
        if filtered.empty:
            return None

        cols = []
        if ctg_col:
            cols.append(ctg_col)
        if elem_col:
            cols.append(elem_col)
        if from_bus_col:
            cols.append(from_bus_col)
        if to_bus_col:
            cols.append(to_bus_col)
        if limit_col:
            cols.append(limit_col)
        if value_col:
            cols.append(value_col)
        if percent_col:
            cols.append(percent_col)

        # Drop duplicates while keeping order
        seen = set()
        ordered_cols = []
        for c in cols:
            if c and c not in seen and c in filtered.columns:
                seen.add(c)
                ordered_cols.append(c)

        if not ordered_cols:
            return None

        return filtered[ordered_cols]


if __name__ == "__main__":
    app = PwbExportApp()
    app.mainloop()