import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import win32com.client
import pandas as pd


class PwbExportApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("PowerWorld Contingency Violations Export (ViolationCTG)")
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
            text="Export existing contingency violations (ViolationCTG)",
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
        csv_out = base + "_ViolationCTG.csv"
        self.csv_path = csv_out

        try:
            self._export_violation_ctg(pwb, csv_out)
        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))

    # ───────────── POWERWORLD EXPORT (ViolationCTG) ───────────── #

    def _export_violation_ctg(self, pwb_path: str, csv_out: str):
        self.log("Connecting to PowerWorld via SimAuto...")
        simauto = win32com.client.Dispatch("pwrworld.SimulatorAuto")
        self.log("Connected.")

        # 1) Open the case (must already have contingency results)
        self.log(f"Opening case: {pwb_path}")
        (err,) = simauto.OpenCase(pwb_path)
        if err:
            raise RuntimeError(f"OpenCase error: {err}")
        self.log("Case opened successfully; using existing contingency results.")

        # 2) Enter Contingency mode (so ViolationCTG objects reflect CTG results)
        self.log("Entering Contingency mode...")
        (err,) = simauto.RunScriptCommand("EnterMode(Contingency);")
        if err:
            raise RuntimeError(f"EnterMode(Contingency) error: {err}")

        # 3) Use SaveData on ViolationCTG.
self.log(f"Saving ViolationCTG data to CSV:\n  {csv_out}")

# Fix Windows slashes BEFORE the f-string:
clean_csv = csv_out.replace("\\", "/")

cmd = (
    f'SaveData("{clean_csv}", CSV, ViolationCTG, '
    '[ALL], [], "");'
)
(err,) = simauto.RunScriptCommand(cmd)
if err:
    raise RuntimeError(f"SaveData(ViolationCTG) error: {err}")
self.log("CSV export complete for ViolationCTG.")

        # 4) Clean up SimAuto
        try:
            simauto.CloseCase()
        except Exception:
            pass
        del simauto

        # 5) Quick preview of the CSV so you can see the Value column
        if os.path.exists(csv_out):
            self.log("\nPreview of first few rows:")
            try:
                df = pd.read_csv(csv_out)
                self.log(f"Columns: {list(df.columns)}")
                preview = df.head(10).to_string(index=False)
                self.log(preview)
            except Exception as e:
                self.log(f"(Could not read CSV preview: {e})")
        else:
            self.log("WARNING: CSV file does not exist after export.")

        messagebox.showinfo("Done", f"ViolationCTG exported to:\n{csv_out}")


if __name__ == "__main__":
    app = PwbExportApp()
    app.mainloop()