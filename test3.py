import csv
import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext


def log(msg, log_widget=None):
    """Append a line to the log box and print to console."""
    print(msg)
    if log_widget is not None:
        log_widget.insert(tk.END, msg + "\n")
        log_widget.see(tk.END)
        log_widget.update_idletasks()


def find_column(headers, candidates):
    """
    Find the first header that contains any of the candidate substrings
    (case-insensitive). Returns the header name or None.
    """
    lower_headers = [h.lower() for h in headers]
    for cand in candidates:
        cand_lower = cand.lower()
        for h, lh in zip(headers, lower_headers):
            if cand_lower in lh:
                return h
    return None


def filter_ctg_csv(input_path, log_widget=None, pct_threshold=100.0):
    """
    Read a big contingency RESULTS CSV (TEST2 output),
    filter to line/transformer violations, and write a new CSV.

    pct_threshold: keep rows where PercentOfLimit >= pct_threshold.
    """
    log(f"Loading CSV: {input_path}", log_widget)

    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    with open(input_path, "r", newline="", encoding="utf-8-sig") as f_in:
        reader = csv.DictReader(f_in)
        headers = reader.fieldnames
        if headers is None:
            raise ValueError("CSV appears to have no header row.")

        log(f"Detected {len(headers)} columns.", log_widget)
        log("Headers:\n  " + "\n  ".join(headers), log_widget)

        # Try to auto-detect useful columns
        col_ctg = find_column(headers, ["ctg name", "contingency", "ctg"])
        col_type = find_column(headers, ["object type", "element type", "device type", "type"])
        col_from = find_column(headers, ["from bus name", "frombus", "from bus", "from name"])
        col_to = find_column(headers, ["to bus name", "tobus", "to bus", "to name"])
        col_limit = find_column(headers, ["limit", "rating"])
        col_value = find_column(headers, ["value", "flow", "actual"])
        col_pct = find_column(headers, ["pctoflimit", "% of limit", "% limit", "percent"])

        # Log what we found
        log("\nDetected columns:", log_widget)
        log(f"  Contingency : {col_ctg}", log_widget)
        log(f"  Type        : {col_type}", log_widget)
        log(f"  From bus    : {col_from}", log_widget)
        log(f"  To bus      : {col_to}", log_widget)
        log(f"  Limit       : {col_limit}", log_widget)
        log(f"  Value       : {col_value}", log_widget)
        log(f"  % of limit  : {col_pct}", log_widget)

        # Basic sanity check
        required_for_output = [col_ctg, col_from, col_to, col_limit, col_value, col_pct]
        if any(c is None for c in required_for_output):
            missing = [name for name, c in zip(
                ["Contingency", "FromBus", "ToBus", "Limit", "Value", "PctOfLimit"],
                required_for_output
            ) if c is None]
            raise ValueError(
                "Could not auto-detect some necessary columns.\n"
                f"Missing logical columns: {', '.join(missing)}.\n"
                "Check the header names in your TEST2 CSV or send me a screenshot of the first few lines."
            )

        rows_out = []

        # Loop over rows and apply filters
        for row in reader:
            # Filter by type (line/xfmr)
            if col_type is not None:
                obj_type = (row.get(col_type, "") or "").lower()
                # Only keep lines/transformers
                if not any(key in obj_type for key in ["branch", "line", "xfmr", "transformer"]):
                    continue

            # Filter by Percent of Limit
            pct_raw = (row.get(col_pct, "") or "").strip()
            try:
                pct_val = float(pct_raw)
            except ValueError:
                # If it cannot be parsed, skip it (likely blank / non-violation)
                continue

            if pct_val < pct_threshold:
                continue

            # Build output row
            from_bus = (row.get(col_from, "") or "").strip()
            to_bus = (row.get(col_to, "") or "").strip()

            out_row = {
                "Contingency": (row.get(col_ctg, "") or "").strip(),
                "ElementType": (row.get(col_type, "") or "").strip() if col_type else "",
                "FromBus": from_bus,
                "ToBus": to_bus,
                "From->To": f"{from_bus} -> {to_bus}",
                "Limit": (row.get(col_limit, "") or "").strip(),
                "Value": (row.get(col_value, "") or "").strip(),
                "PercentOfLimit": pct_raw,
            }
            rows_out.append(out_row)

    if not rows_out:
        log("\nNo rows matched the filter (line/xfmr & percent >= threshold).", log_widget)
    else:
        log(f"\nFiltered down to {len(rows_out)} rows.", log_widget)

    # Build output path
    base, ext = os.path.splitext(input_path)
    output_path = base + "_filtered_lines.csv"

    # Write output CSV
    fieldnames = [
        "Contingency",
        "ElementType",
        "FromBus",
        "ToBus",
        "From->To",
        "Limit",
        "Value",
        "PercentOfLimit",
    ]

    with open(output_path, "w", newline="", encoding="utf-8") as f_out:
        writer = csv.DictWriter(f_out, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows_out)

    log(f"\nFiltered CSV written to:\n  {output_path}", log_widget)
    return output_path


# =========================
# Tkinter GUI
# =========================

class CTGFilterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Contingency Results Filter (Lines/Transformers)")
        self.geometry("800x450")

        self.input_path_var = tk.StringVar()
        self.pct_threshold_var = tk.StringVar(value="100")  # default 100%

        self.create_widgets()

    def create_widgets(self):
        # Input file selector
        frm_top = tk.Frame(self)
        frm_top.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(frm_top, text="TEST2 CSV file:").grid(row=0, column=0, sticky="w")

        ent = tk.Entry(frm_top, textvariable=self.input_path_var)
        ent.grid(row=0, column=1, sticky="we", padx=5)
        frm_top.columnconfigure(1, weight=1)

        btn_browse = tk.Button(frm_top, text="Browse...", command=self.browse_file)
        btn_browse.grid(row=0, column=2, padx=5)

        # Threshold
        tk.Label(frm_top, text="Min % of Limit (>=):").grid(row=1, column=0, sticky="w", pady=(8, 0))
        tk.Entry(frm_top, textvariable=self.pct_threshold_var, width=8).grid(
            row=1, column=1, sticky="w", pady=(8, 0)
        )

        # Run button
        btn_run = tk.Button(frm_top, text="Run Filter", command=self.run_filter, width=15)
        btn_run.grid(row=1, column=2, padx=5, pady=(8, 0))

        # Log area
        tk.Label(self, text="Log:").pack(anchor="w", padx=10)
        self.txt_log = scrolledtext.ScrolledText(self, height=18)
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select TEST2 contingency results CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            self.input_path_var.set(path)

    def run_filter(self):
        input_path = self.input_path_var.get().strip()
        if not input_path:
            messagebox.showerror("Error", "Please select a TEST2 CSV file first.")
            return

        try:
            pct_str = self.pct_threshold_var.get().strip()
            pct_threshold = float(pct_str)
        except ValueError:
            messagebox.showerror("Error", f"Invalid percent threshold: {self.pct_threshold_var.get()}")
            return

        self.txt_log.delete("1.0", tk.END)
        try:
            out_path = filter_ctg_csv(input_path, log_widget=self.txt_log, pct_threshold=pct_threshold)
            messagebox.showinfo("Done", f"Filtered results written to:\n{out_path}")
        except Exception as e:
            err_text = "".join(traceback.format_exception_only(type(e), e)).strip()
            log("\n[ERROR] " + err_text, self.txt_log)
            log(traceback.format_exc(), self.txt_log)
            messagebox.showerror("Error", f"Filtering failed:\n{err_text}")


if __name__ == "__main__":
    app = CTGFilterApp()
    app.mainloop()