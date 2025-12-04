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

        # 1) Open case – contingencies already solved
        self.log(f"Opening case: {pwb_path}")
        (err,) = simauto.OpenCase(pwb_path)
        if err:
            raise RuntimeError(f"OpenCase error: {err}")
        self.log("Case opened successfully (existing results will be used).")

        # 2) Enter Contingency mode
        self.log("Entering Contingency mode...")
        (err,) = simauto.RunScriptCommand("EnterMode(Contingency);")
        if err:
            raise RuntimeError(f"EnterMode(Contingency) error: {err}")

        # 3) Export violation matrices for branches (lines + xfmrs)
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

        # 4) Load raw CSV and build summary
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
                        "\nNo matching Branch MVA rows found, "
                        "or required columns were missing."
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
        Build a compact summary for line/transformer contingencies using
        the specific columns you requested:

        - Contingency        (from CTGLabel)
        - LineID             (from LimViolID)
        - Limit_MVA          (from LimViolLimit)
        - Flow_MVA           (computed = Limit_MVA * Percent / 100)
        - PercentOfLimit     (from CTGVioIMaxLine)
        """

        if df is None or df.empty:
            return None

        df2 = df.copy()

        # Only keep Branch MVA rows if LimViolCat exists
        if "LimViolCat" in df2.columns:
            mask = df2["LimViolCat"].astype(str).str.contains(
                "Branch MVA", case=False, na=False
            )
            df2 = df2[mask]

        if df2.empty:
            return None

        required = ["CTGLabel", "LimViolID", "LimViolLimit", "CTGVioIMaxLine"]
        missing = [c for c in required if c not in df2.columns]
        if missing:
            # If something is missing, bail out so you know about it
            self.log(f"Required columns missing from CSV: {missing}")
            return None

        # Drop rows where LimViolID is blank (no actual violated element)
        df2 = df2[df2["LimViolID"].astype(str).str.strip() != ""]
        if df2.empty:
            return None

        # Convert numeric columns
        df2["LimViolLimit"] = pd.to_numeric(df2["LimViolLimit"], errors="coerce")
        df2["CTGVioIMaxLine"] = pd.to_numeric(
            df2["CTGVioIMaxLine"], errors="coerce"
        )

        # Compute actual flow MVA from limit * percent / 100
        df2["Flow_MVA"] = df2["LimViolLimit"] * df2["CTGVioIMaxLine"] / 100.0

        # Build the final summary table with nice column names
        summary = pd.DataFrame(
            {
                "Contingency": df2["CTGLabel"],
                "LineID": df2["LimViolID"],
                "Limit_MVA": df2["LimViolLimit"],
                "Flow_MVA": df2["Flow_MVA"],
                "PercentOfLimit": df2["CTGVioIMaxLine"],
            }
        )

        # Sort by highest loading first
        summary = summary.sort_values(
            by=["PercentOfLimit", "Contingency"], ascending=[False, True]
        )

        return summary


if __name__ == "__main__":
    app = PwbExportApp()
    app.mainloop()