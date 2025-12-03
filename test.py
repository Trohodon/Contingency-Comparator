import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client


def export_ctg_results(pwb_path: str, csv_path: str):
    """
    Use SimAuto to open a PWB that already has contingency results
    and export the Combined Tables 'CTG_Results' info to CSV.

    The CSV will have one row per contingency result with columns like:
    Name (contingency), Category, Value, Limit, Percent, LimitScale.
    """
    simauto = None
    try:
        simauto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

        # Make Simulator invisible so it doesn't pop up all over the place
        try:
            simauto.UIVisible = False
        except Exception:
            # Older versions may not have UIVisible; ignore if so
            pass

        # Open the case
        err, = simauto.OpenCase(pwb_path)
        if err:
            raise RuntimeError(f"OpenCase error: {err}")

        # Make sure we’re in RUN mode (usually already true if results exist)
        err, = simauto.RunScriptCommand('EnterMode(RUN);')
        if err:
            raise RuntimeError(f"EnterMode error: {err}")

        # Build a SaveData command for the CTG_Results table.
        #
        # SaveData("filename", filetype, objecttype,
        #          [fieldlist], [subdatalist], filter,
        #          [SortFieldList], Transpose, Append);
        #
        # Field list below:
        #   Name        -> Contingency name/label
        #   Category    -> Branch MVA, Bus Voltage, etc.
        #   Value       -> The MVA value you see in the grid
        #   Limit       -> Limit used for that violation
        #   Percent     -> Value / Limit * 100
        #   LimitScale  -> Scale of limit (emergency etc.)
        #
        # If any of these names are slightly different in your version,
        # PowerWorld will report the invalid field name in the error text.
        cmd = (
            'SaveData("{fname}", CSV, CTG_Results, '
            '[Name, Category, Value, Limit, Percent, LimitScale], '
            '[], "", [], NO, NO);'
        ).format(fname=csv_path.replace('\\', '/'))

        err, = simauto.RunScriptCommand(cmd)
        if err:
            raise RuntimeError(f"SaveData error: {err}")

    finally:
        if simauto is not None:
            try:
                simauto.CloseCase()
            except Exception:
                pass
            del simauto


class ExportGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PowerWorld Contingency Exporter")

        self.pwb_path = tk.StringVar()
        self.csv_path = tk.StringVar()

        # Row 0: PWB selection
        tk.Label(self, text="PowerWorld case (.pwb):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        tk.Entry(self, textvariable=self.pwb_path, width=70).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self, text="Browse…", command=self.browse_pwb).grid(row=0, column=2, padx=5, pady=5)

        # Row 1: CSV output
        tk.Label(self, text="Output CSV:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        tk.Entry(self, textvariable=self.csv_path, width=70).grid(row=1, column=1, padx=5, pady=5)
        tk.Button(self, text="Browse…", command=self.browse_csv).grid(row=1, column=2, padx=5, pady=5)

        # Row 2: Run button
        tk.Button(self, text="Export CTG Results", command=self.run_export, width=25).grid(
            row=2, column=0, columnspan=3, pady=10
        )

    def browse_pwb(self):
        filename = filedialog.askopenfilename(
            title="Select PowerWorld Case",
            filetypes=[("PowerWorld case", "*.pwb"), ("All files", "*.*")]
        )
        if filename:
            self.pwb_path.set(filename)
            # Suggest a CSV name in the same folder
            base = os.path.splitext(os.path.basename(filename))[0]
            suggested = os.path.join(os.path.dirname(filename), f"{base}_CTGResults.csv")
            if not self.csv_path.get():
                self.csv_path.set(suggested)

    def browse_csv(self):
        filename = filedialog.asksaveasfilename(
            title="Select Output CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.csv_path.set(filename)

    def run_export(self):
        pwb = self.pwb_path.get().strip()
        csv = self.csv_path.get().strip()

        if not pwb or not os.path.isfile(pwb):
            messagebox.showerror("Error", "Please select a valid .pwb file.")
            return
        if not csv:
            messagebox.showerror("Error", "Please choose an output CSV file.")
            return

        try:
            export_ctg_results(pwb, csv)
            messagebox.showinfo("Done", f"Export complete.\n\nSaved:\n{csv}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))


if __name__ == "__main__":
    # Simple Tkinter startup
    app = ExportGUI()
    app.mainloop()