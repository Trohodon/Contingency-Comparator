import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import win32com.client
import pandas as pd


class PwbExportApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("PowerWorld Contingency Violations Export (Trimmed)")
        self.geometry("900x550")

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

        browse_btn = ttk.Button(top, text="Browse…", command=self.browse_pwb)
        browse_btn.grid(row=1, column=2, padx=(5, 0), sticky="e")

        run_btn = ttk.Button(
            top,
            text="Export existing contingency results (trimmed)",
            command=self.run_export,
        )
        run_btn.grid(row=2, column=0, columnspan=3, pady=(10, 0), sticky="w")

        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=10, pady=5)

        # Log / output area
        log_frame = ttk.Frame(self)
        log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(log_frame, text="Log:").pack(anchor="w")

        self.log_text = tk.Text(log_frame, wrap="word", height=18)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.log_text.yview
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
            messagebox.showwarning("No case selected", "Please select a valid .pwb file.")
            return

        base, _ = os.path.splitext(pwb)
        csv_out = base + "_CTGViolation_trimmed.csv"
        self.csv_path = csv_out

        try:
            self._export_ctgviolation_trimmed(pwb, csv_out)
        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))

    # ───────────── POWERWORLD EXPORT (CTGViolation, trimmed) ───────────── #

    def _export_ctgviolation_trimmed(self, pwb_path: str, csv_out: str):
        self.log("Connecting to PowerWorld via SimAuto...")
        simauto = win32com.client.Dispatch("pwrworld.SimulatorAuto")
        self.log("Connected.")

        # 1) Open the case
        self.log(f"Opening case: {pwb_path}")
        (err,) = simauto.OpenCase(pwb_path)
        if err:
            raise RuntimeError(f"OpenCase error: {err}")
        self.log("Case opened successfully; using stored contingency violations.")

        # Enter contingency mode (usually not strictly required, but safe)
        self.log("Entering Contingency mode...")
        (err,) = simauto.RunScriptCommand("EnterMode(Contingency);")
        if err:
            raise RuntimeError(f"EnterMode(Contingency) error: {err}")

        # 2) Export CTGViolation to a temporary CSV
        tmp_csv = csv_out + ".tmp"
        clean_tmp = tmp_csv.replace("\\", "/")

        self.log(f"Saving CTGViolation to temporary CSV:\n  {tmp_csv}")

        # SaveData("file", CSV, CTGViolation, [ALL], [], "")
        cmd = 'SaveData("{}", CSV, CTGViolation, [ALL], [], "")'.format(clean_tmp)
        (err,) = simauto.RunScriptCommand(cmd)
        if err:
            raise RuntimeError(f"SaveData(CTGViolation) error: {err}")

        self.log("Stored CTGViolation CSV export complete.")

        # 3) Close case / SimAuto
        try:
            simauto.CloseCase()
        except Exception:
            pass
        del simauto

        # 4) Load temporary CSV
        if not os.path.exists(tmp_csv):
            raise RuntimeError("Temporary CSV not found after export.")

        self.log("Loading temporary CSV into pandas...")
        df = pd.read_csv(tmp_csv)

        self.log(f"Columns found: {list(df.columns)}")

        # 5) Auto-detect key columns

        cols = list(df.columns)

        # Contingency name column
        col_ctg = next(
            (c for c in cols if "CTG" in c or c == "Name"), None
        )
        # Object / element name (branch, xfmr, bus, etc.)
        col_obj = next(
            (
                c
                for c in cols
                if "ObjectName" in c
                or "Element" in c
                or "BranchName" in c
                or "ViolObject" in c
            ),
            None,
        )
        # Category (Branch MVA, Bus Voltage, etc.)
        col_cat = next((c for c in cols if "Category" in c or "ViolCat" in c), None)
        # Violation numeric value
        col_val = next((c for c in cols if c.lower() == "value"), None)
        # Limit
        col_lim = next((c for c in cols if c.lower() == "limit"), None)
        # Percent of limit
        col_pct = next(
            (
                c
                for c in cols
                if "Percent" in c
                or "PctOfLimit" in c
                or "PercentOfLimit" in c
            ),
            None,
        )

        keep_cols = [c for c in [col_ctg, col_obj, col_cat, col_val, col_lim, col_pct] if c]

        if not keep_cols:
            raise RuntimeError(
                "Could not detect any relevant CTGViolation columns to keep.\n"
                "Check the 'Columns found' list in the log."
            )

        self.log(f"Keeping only columns: {keep_cols}")

        trimmed = df[keep_cols]

        # 6) Write final trimmed CSV
        trimmed.to_csv(csv_out, index=False)
        self.log(f"\nTrimmed CSV written to:\n  {csv_out}")

        # Remove temporary file
        try:
            os.remove(tmp_csv)
            self.log(f"Temporary file removed: {tmp_csv}")
        except OSError:
            self.log(f"(Could not remove temporary file: {tmp_csv})")

        # 7) Preview
        self.log("\nPreview of trimmed CSV:")
        self.log(f"Columns: {list(trimmed.columns)}")
        self.log(trimmed.head(10).to_string(index=False))

        messagebox.showinfo("Done", f"Stored CTGViolation exported to:\n{csv_out}")


if __name__ == "__main__":
    app = PwbExportApp()
    app.mainloop() 