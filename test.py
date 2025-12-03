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

        log_frame = ttk.Frame(self)
        log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(log_frame, text="Log:").pack(anchor="w")

        self.log_text = tk.Text(log_frame, wrap="word", height=18)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scroll.set)

    def log(self, msg: str):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    # -------- GUI Actions -------- #

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
        csv_out = base + "_ViolationCTG_trimmed.csv"
        self.csv_path = csv_out

        try:
            self._export_violationctg_trimmed(pwb, csv_out)
        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))

    # -------- ACTUAL EXPORT -------- #

    def _export_violationctg_trimmed(self, pwb_path: str, csv_out: str):
        self.log("Connecting to PowerWorld via SimAuto...")
        simauto = win32com.client.Dispatch("pwrworld.SimulatorAuto")
        self.log("Connected.")

        # 1) Open case
        self.log(f"Opening case: {pwb_path}")
        (err,) = simauto.OpenCase(pwb_path)
        if err:
            raise RuntimeError(f"OpenCase error: {err}")
        self.log("Case opened successfully.")

        # 2) Export ViolationCTG table
        tmp_csv = csv_out + ".tmp"
        clean_tmp = tmp_csv.replace("\\", "/")

        self.log(f"Saving ViolationCTG to temporary file:\n  {tmp_csv}")

        cmd = f'SaveData("{clean_tmp}", CSV, ViolationCTG, [ALL], [], "")'
        (err,) = simauto.RunScriptCommand(cmd)
        if err:
            raise RuntimeError(f"SaveData(ViolationCTG) error: {err}")

        self.log("ViolationCTG export complete.")

        simauto.CloseCase()
        del simauto

        # 3) Load temporary CSV
        df = pd.read_csv(tmp_csv)
        self.log("CSV loaded into pandas.")

        # 4) Identify desired columns
        cols = df.columns.tolist()
        self.log(f"Columns found: {cols}")

        col_ctg = next((c for c in cols if "CTG" in c or "ViolCTG" in c), None)
        col_obj = next((c for c in cols if "Object" in c or "Element" in c), None)
        col_cat = next((c for c in cols if "Category" in c), None)
        col_val = next((c for c in cols if c.lower() == "value"), None)
        col_lim = next((c for c in cols if c.lower() == "limit"), None)
        col_pct = next((c for c in cols if "Percent" in c), None)

        keep = [c for c in [col_ctg, col_obj, col_cat, col_val, col_lim, col_pct] if c]

        if not keep:
            raise RuntimeError("No usable columns detected.")

        self.log(f"Keeping columns: {keep}")

        trimmed = df[keep]
        trimmed.to_csv(csv_out, index=False)
        self.log(f"Trimmed CSV saved to:\n  {csv_out}")

        os.remove(tmp_csv)
        self.log("Temporary file removed.")

        self.log("\nPreview:")
        self.log(trimmed.head(10).to_string(index=False))

        messagebox.showinfo("Done", f"Stored ViolationCTG exported to:\n{csv_out}")


# -------- Run GUI -------- #

if __name__ == "__main__":
    app = PwbExportApp()
    app.mainloop()