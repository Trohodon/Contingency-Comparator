import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import win32com.client
import pandas as pd


class PwbExportApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("PowerWorld Contingency Export Tool")
        self.geometry("800x500")

        self.pwb_path = tk.StringVar(value="No .pwb file selected")
        self.csv_path = None

        self._build_gui()

    # ───────────────── GUI LAYOUT ───────────────── #

    def _build_gui(self):
        top = ttk.Frame(self)
        top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        # Row: Select file
        ttk.Label(top, text="Selected .pwb case:").grid(row=0, column=0, sticky="w")
        ttk.Label(top, textvariable=self.pwb_path, width=80).grid(
            row=1, column=0, columnspan=2, sticky="w"
        )

        browse_btn = ttk.Button(top, text="Browse...", command=self.browse_pwb)
        browse_btn.grid(row=1, column=2, padx=(5, 0), sticky="e")

        # Row: Export button
        run_btn = ttk.Button(
            top, text="Run Contingency Export", command=self.run_export
        )
        run_btn.grid(row=2, column=0, columnspan=3, pady=(10, 0), sticky="w")

        # Separator
        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=10, pady=5)

        # Log / output area
        log_frame = ttk.Frame(self)
        log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(log_frame, text="Log:").pack(anchor="w")

        self.log_text = tk.Text(log_frame, wrap="word", height=15)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scroll.set)

    def log(self, msg: str):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    # ───────────────── CALLBACKS ───────────────── #

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

        # Decide CSV output path (same folder + _Violations.csv)
        base, _ = os.path.splitext(pwb)
        csv_out = base + "_Violations.csv"
        self.csv_path = csv_out

        try:
            self._do_export(pwb, csv_out)
        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))

    # ───────────────── POWERWORLD LOGIC ───────────────── #

    def _do_export(self, pwb_path: str, csv_out: str):
        self.log("Connecting to PowerWorld via SimAuto...")
        simauto = win32com.client.Dispatch("pwrworld.SimulatorAuto")
        self.log("Connected.")

        # 1) Open case
        self.log(f"Opening case: {pwb_path}")
        err, = simauto.OpenCase(pwb_path)
        if err:
            raise RuntimeError(f"OpenCase error: {err}")
        self.log("Case opened successfully.")

        # 2) Solve power flow
        self.log("Entering PowerFlow mode...")
        err, = simauto.RunScriptCommand("EnterMode(PowerFlow);")
        if err:
            raise RuntimeError(f"EnterMode(PowerFlow) error: {err}")

        self.log("Solving power flow...")
        err, = simauto.RunScriptCommand("SolvePowerFlow;")
        if err:
            raise RuntimeError(f"SolvePowerFlow error: {err}")
        self.log("Power flow solved.")

        # 3) Run contingency analysis
        self.log("Entering Contingency mode...")
        err, = simauto.RunScriptCommand("EnterMode(Contingency);")
        if err:
            raise RuntimeError(f"EnterMode(Contingency) error: {err}")

        self.log("Running CTGSolveAll (YES, YES)...")
        err, = simauto.RunScriptCommand("CTGSolveAll(YES, YES);")
        if err:
            raise RuntimeError(f"CTGSolveAll error: {err}")
        self.log("Contingency analysis complete.")

        # 4) Export violations matrix to CSV
        self.log(f"Saving violation matrices to CSV:\n  {csv_out}")
        cmd = (
            f'CTGSaveViolationMatrices("{csv_out}", CSVCOLHEADER, '
            'YES, [BRANCH], YES, NO);'
        )
        err, = simauto.RunScriptCommand(cmd)
        if err:
            raise RuntimeError(f"CTGSaveViolationMatrices error: {err}")
        self.log("CSV export complete.")

        # Close SimAuto connection
        del simauto

        # 5) Try previewing first few rows
        if os.path.exists(csv_out):
            self.log("\nPreview of first few rows:")
            try:
                df = pd.read_csv(csv_out)
                preview = df.head(10).to_string(index=False)
                self.log(preview)
            except Exception as e:
                self.log(f"(Could not read CSV preview: {e})")
        else:
            self.log("WARNING: CSV file does not exist after export.")

        messagebox.showinfo("Done", f"Violations exported to:\n{csv_out}")


if __name__ == "__main__":
    app = PwbExportApp()
    app.mainloop()