import os
import csv
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import win32com.client


# -------------------------- Core PowerWorld Logic -------------------------- #

def open_case(simauto, case_path, log):
    """Open a PowerWorld case and log any errors."""
    log(f"Opening case: {case_path}")
    result = simauto.OpenCase(case_path)
    if result:
        # PowerWorld returns an error string if something went wrong
        log(f"[ERROR] OpenCase returned: {result}")
        raise RuntimeError(f"PowerWorld OpenCase error: {result}")
    log("Case opened successfully.")


def get_branch_results(simauto, log):
    """
    Pull contingency branch results (lines/transformers) from PowerWorld.

    We use the CTG_Results_Branch table which already ONLY contains
    branch elements (lines/transformers).
    """
    table_name = "CTG_Results_Branch"

    # Fields we ask PowerWorld for. These names are standard for this table.
    fields = [
        "CTGName",   # Contingency name
        "BusNum",    # From bus number
        "BusNum:1",  # To bus number
        "LineID",    # Circuit ID / line ID
        "Limit",     # Limit value
        "MW",        # Flow in MW (violation value)
        "PctOfLimit" # Percent of limit
    ]

    log(f"Requesting contingency results from table '{table_name}'...")
    error, returned_fields, values = simauto.GetParametersMultipleElement(
        table_name,
        fields,
        ""  # empty filter = all rows
    )

    if error:
        log(f"[ERROR] GetParametersMultipleElement returned: {error}")
        raise RuntimeError(f"PowerWorld GetParametersMultipleElement error: {error}")

    log(f"Returned {len(values)} branch result rows.")
    return values  # list of lists in the same order as 'fields'


def write_test1_overview(values, output_path, log):
    """
    Test 1: simple overview.
    Columns: Contingency, Element (From->To LineID), PercentOfLimit
    """
    log(f"Writing Test 1 overview CSV: {output_path}")
    with open(output_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Contingency", "Element", "PercentOfLimit"])

        for row in values:
            # row = [CTGName, BusNum, BusNum:1, LineID, Limit, MW, PctOfLimit]
            ctg_name = row[0]
            bus_from = row[1]
            bus_to   = row[2]
            line_id  = row[3]
            pct      = row[6]

            element = f"{bus_from}->{bus_to} {line_id}"
            writer.writerow([ctg_name, element, pct])

    log(f"Test 1 CSV written with {len(values)} rows.")


def write_test2_filtered(values, output_path, log):
    """
    Test 2: filtered detailed branch/transformer info.
    Columns:
      Contingency, FromBus, ToBus, LineID, Limit, ValueMW, PercentOfLimit
    """
    log(f"Writing Test 2 filtered CSV: {output_path}")
    with open(output_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([
            "Contingency",
            "FromBus",
            "ToBus",
            "LineID",
            "Limit",
            "ValueMW",
            "PercentOfLimit"
        ])

        for row in values:
            # row = [CTGName, BusNum, BusNum:1, LineID, Limit, MW, PctOfLimit]
            ctg_name = row[0]
            bus_from = row[1]
            bus_to   = row[2]
            line_id  = row[3]
            limit    = row[4]
            mw_val   = row[5]
            pct      = row[6]

            writer.writerow([
                ctg_name,
                bus_from,
                bus_to,
                line_id,
                limit,
                mw_val,
                pct
            ])

    log(f"Test 2 CSV written with {len(values)} rows.")


def run_test(case_path, test_type, log):
    """
    Connects to PowerWorld, opens the case, pulls branch results,
    and writes the CSV according to test_type: 'test1' or 'test2'.
    """
    if not os.path.isfile(case_path):
        raise FileNotFoundError(f"Case not found: {case_path}")

    # Build an output filename next to the .pwb
    base_dir = os.path.dirname(case_path)
    base_name = os.path.splitext(os.path.basename(case_path))[0]
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    if test_type == "test1":
        out_name = f"{base_name}_CTG_Test1_Overview_{timestamp}.csv"
    else:
        out_name = f"{base_name}_CTG_Test2_Filtered_{timestamp}.csv"

    output_path = os.path.join(base_dir, out_name)

    # Connect to SimAuto
    log("Connecting to PowerWorld SimAuto...")
    simauto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

    try:
        open_case(simauto, case_path, log)
        values = get_branch_results(simauto, log)

        if test_type == "test1":
            write_test1_overview(values, output_path, log)
        else:
            write_test2_filtered(values, output_path, log)

        log(f"[DONE] Output file created:\n{output_path}")
        return output_path

    finally:
        try:
            log("Closing case in PowerWorld...")
            simauto.CloseCase()
            log("Case closed.")
        except Exception as e:
            log(f"[WARN] Error while closing case: {e}")
        simauto = None


# -------------------------- GUI Definition -------------------------- #

class PowerWorldGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("PowerWorld Contingency Results Viewer")

        # ------------- Case selection ------------- #
        frame_top = tk.Frame(master)
        frame_top.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_top, text="Case (.pwb):").grid(row=0, column=0, sticky="w")
        self.case_entry = tk.Entry(frame_top, width=60)
        self.case_entry.grid(row=0, column=1, padx=5)

        browse_btn = tk.Button(frame_top, text="Browse...", command=self.browse_case)
        browse_btn.grid(row=0, column=2, padx=5)

        # ------------- Buttons for Test 1 / Test 2 ------------- #
        frame_buttons = tk.Frame(master)
        frame_buttons.pack(fill="x", padx=10, pady=5)

        self.btn_test1 = tk.Button(
            frame_buttons,
            text="Run Test 1 (Simple Overview)",
            command=self.run_test1_clicked
        )
        self.btn_test1.pack(side="left", padx=5)

        self.btn_test2 = tk.Button(
            frame_buttons,
            text="Run Test 2 (Filtered Lines/Transformers)",
            command=self.run_test2_clicked
        )
        self.btn_test2.pack(side="left", padx=5)

        # ------------- Log window ------------- #
        frame_log = tk.Frame(master)
        frame_log.pack(fill="both", expand=True, padx=10, pady=5)

        tk.Label(frame_log, text="Log:").pack(anchor="w")

        self.log_text = scrolledtext.ScrolledText(frame_log, height=15, state="disabled")
        self.log_text.pack(fill="both", expand=True)

    # ------------------------- GUI Helper Methods ------------------------- #

    def log(self, message: str):
        """Append a message to the log text box."""
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        print(message)  # also print to console for debugging if run from terminal

    def browse_case(self):
        """Open file dialog to select a .pwb case file."""
        filename = filedialog.askopenfilename(
            title="Select PowerWorld Case",
            filetypes=[("PowerWorld Case", "*.pwb"), ("All files", "*.*")]
        )
        if filename:
            self.case_entry.delete(0, "end")
            self.case_entry.insert(0, filename)

    def disable_buttons(self):
        self.btn_test1.config(state="disabled")
        self.btn_test2.config(state="disabled")

    def enable_buttons(self):
        self.btn_test1.config(state="normal")
        self.btn_test2.config(state="normal")

    def run_test_common(self, test_type: str):
        """Shared logic for running Test 1 or Test 2."""
        case_path = self.case_entry.get().strip()
        if not case_path:
            messagebox.showerror("Error", "Please select a .pwb case file first.")
            return

        self.log("\n" + "=" * 70)
        self.log(f"Starting {test_type.upper()} on case: {case_path}")
        self.disable_buttons()
        self.master.update_idletasks()

        try:
            output_path = run_test(case_path, test_type, self.log)
            messagebox.showinfo(
                "Done",
                f"{test_type.upper()} completed successfully.\n\nOutput file:\n{output_path}"
            )
        except Exception as e:
            self.log(f"[EXCEPTION] {e}")
            messagebox.showerror("Error", str(e))
        finally:
            self.enable_buttons()

    def run_test1_clicked(self):
        self.run_test_common("test1")

    def run_test2_clicked(self):
        # This is the one you care about: line/transformer filtered output
        self.run_test_common("test2")


# -------------------------- Program Entry Point -------------------------- #

if __name__ == "__main__":
    root = tk.Tk()
    app = PowerWorldGUI(root)
    root.mainloop()