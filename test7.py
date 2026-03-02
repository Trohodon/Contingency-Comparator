import os
import re
import subprocess
import sys
from collections import OrderedDict
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


SLOTS = OrderedDict(
    [
        ("P1_2026", ("P1", "2026")),
        ("P1_2030", ("P1", "2030")),
        ("P2_2026", ("P2", "2026")),
        ("P2_2030", ("P2", "2030")),
    ]
)

CONTINGENCY_RE = re.compile(r"^\s*CONTINGENCY\s+['\"](?P<name>[^'\"]+)['\"]\s*$", re.IGNORECASE)
END_RE = re.compile(r"^\s*END\s*$", re.IGNORECASE)
WHITESPACE_RE = re.compile(r"\s+")


def normalize_action_line(line):
    return WHITESPACE_RE.sub(" ", line.strip()).upper()


def join_actions(actions):
    return "\n".join(actions)


def format_timestamp(timestamp):
    return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")


def find_con_files(folder):
    folder_path = Path(folder)
    slot_matches = {slot: [] for slot in SLOTS}

    for entry in folder_path.iterdir():
        if not entry.is_file() or entry.suffix.lower() != ".con":
            continue

        upper_name = entry.name.upper()
        for slot, (group, year) in SLOTS.items():
            if group in upper_name and year in upper_name:
                slot_matches[slot].append(entry)

    selected = {}
    warnings = []
    missing = []

    for slot, matches in slot_matches.items():
        if not matches:
            missing.append(slot)
            continue

        matches.sort(key=lambda path: path.stat().st_mtime, reverse=True)
        selected[slot] = str(matches[0])

        if len(matches) > 1:
            chosen = matches[0]
            warnings.append(
                f"{slot}: multiple matches found; using newest file "
                f"{chosen.name} ({format_timestamp(chosen.stat().st_mtime)})"
            )

    return {
        "selected": selected,
        "warnings": warnings,
        "missing": missing,
        "all_matches": {slot: [str(path) for path in paths] for slot, paths in slot_matches.items()},
    }


def parse_con_file(path):
    contingencies = OrderedDict()
    current_name = None
    current_actions = []

    with open(path, "r", encoding="utf-8-sig", errors="replace") as handle:
        for raw_line in handle:
            line = raw_line.rstrip("\r\n")
            match = CONTINGENCY_RE.match(line)
            if match:
                if current_name is not None:
                    contingencies[current_name] = current_actions
                current_name = match.group("name")
                current_actions = []
                continue

            if current_name is None:
                continue

            if END_RE.match(line):
                contingencies[current_name] = current_actions
                current_name = None
                current_actions = []
                continue

            current_actions.append(line)

    if current_name is not None:
        contingencies[current_name] = current_actions

    return contingencies


def compare_sets(dict_2026, dict_2030):
    names_2026 = set(dict_2026)
    names_2030 = set(dict_2030)

    unchanged = []
    modified = []
    removed = []
    added = []

    for name in sorted(names_2026 & names_2030):
        actions_2026 = dict_2026[name]
        actions_2030 = dict_2030[name]

        normalized_2026 = [normalize_action_line(line) for line in actions_2026]
        normalized_2030 = [normalize_action_line(line) for line in actions_2030]

        if normalized_2026 == normalized_2030:
            unchanged.append({"Contingency": name, "Actions": join_actions(actions_2026)})
        else:
            modified.append(
                {
                    "Contingency": name,
                    "Actions_2026": join_actions(actions_2026),
                    "Actions_2030": join_actions(actions_2030),
                }
            )

    for name in sorted(names_2026 - names_2030):
        removed.append({"Contingency": name, "Actions": join_actions(dict_2026[name])})

    for name in sorted(names_2030 - names_2026):
        added.append({"Contingency": name, "Actions": join_actions(dict_2030[name])})

    return {
        "summary": {
            "Unchanged": len(unchanged),
            "Modified": len(modified),
            "Added": len(added),
            "Removed": len(removed),
        },
        "unchanged": unchanged,
        "modified": modified,
        "removed": removed,
        "added": added,
    }


def autofit_columns(ws, widths=None):
    widths = widths or {}

    for column_cells in ws.columns:
        column_index = column_cells[0].column
        letter = get_column_letter(column_index)
        if letter in widths:
            ws.column_dimensions[letter].width = widths[letter]
            continue

        max_length = 0
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        ws.column_dimensions[letter].width = min(max(max_length + 2, 12), 40)


def style_header_row(ws, row_number):
    for cell in ws[row_number]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="left", vertical="top")


def write_table(ws, start_row, headers, rows, wrap_columns=None, fixed_widths=None):
    wrap_columns = set(wrap_columns or [])
    fixed_widths = fixed_widths or {}

    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=start_row, column=col_idx, value=header)
    style_header_row(ws, start_row)

    for row_offset, row_data in enumerate(rows, start=1):
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=start_row + row_offset, column=col_idx, value=row_data.get(header, ""))
            cell.alignment = Alignment(
                wrap_text=header in wrap_columns,
                vertical="top",
                horizontal="left",
            )

    header_to_letter = {header: get_column_letter(idx) for idx, header in enumerate(headers, start=1)}
    autofit_columns(
        ws,
        widths={header_to_letter[header]: width for header, width in fixed_widths.items() if header in header_to_letter},
    )
    ws.freeze_panes = "A2"


def write_summary_sheet(ws, prefix, results, file_paths):
    ws.title = f"{prefix}_Summary"

    path_rows = [
        {"Label": f"{prefix} 2026", "Path": file_paths[f"{prefix}_2026"]},
        {"Label": f"{prefix} 2030", "Path": file_paths[f"{prefix}_2030"]},
        {"Label": "Generated", "Path": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
    ]

    write_table(
        ws,
        start_row=1,
        headers=["Label", "Path"],
        rows=path_rows,
        wrap_columns={"Path"},
        fixed_widths={"Path": 90},
    )

    summary_start = len(path_rows) + 3
    metric_rows = [{"Metric": metric, "Count": count} for metric, count in results["summary"].items()]
    write_table(ws, start_row=summary_start, headers=["Metric", "Count"], rows=metric_rows)
    ws.freeze_panes = "A2"


def write_detail_sheet(ws, title, headers, rows):
    ws.title = title
    wrap_columns = {header for header in headers if "Actions" in header}
    fixed_widths = {header: 70 for header in headers if "Actions" in header}
    write_table(
        ws,
        start_row=1,
        headers=headers,
        rows=rows,
        wrap_columns=wrap_columns,
        fixed_widths=fixed_widths,
    )


def write_inputs_sheet(ws, file_paths):
    ws.title = "Inputs"
    rows = []
    for slot in SLOTS:
        path = file_paths.get(slot, "")
        rows.append(
            {
                "Slot": slot,
                "Filename": Path(path).name if path else "",
                "Full_Path": path,
            }
        )

    write_table(
        ws,
        start_row=1,
        headers=["Slot", "Filename", "Full_Path"],
        rows=rows,
        wrap_columns={"Full_Path"},
        fixed_widths={"Full_Path": 90},
    )


def write_excel_report(out_path, results_p1, results_p2, file_paths):
    workbook = Workbook()
    workbook.remove(workbook.active)

    for prefix, results in (("P1", results_p1), ("P2", results_p2)):
        write_summary_sheet(workbook.create_sheet(), prefix, results, file_paths)
        write_detail_sheet(workbook.create_sheet(), f"{prefix}_Added", ["Contingency", "Actions"], results["added"])
        write_detail_sheet(workbook.create_sheet(), f"{prefix}_Removed", ["Contingency", "Actions"], results["removed"])
        write_detail_sheet(
            workbook.create_sheet(),
            f"{prefix}_Modified",
            ["Contingency", "Actions_2026", "Actions_2030"],
            results["modified"],
        )
        write_detail_sheet(
            workbook.create_sheet(),
            f"{prefix}_Unchanged",
            ["Contingency", "Actions"],
            results["unchanged"],
        )

    write_inputs_sheet(workbook.create_sheet(), file_paths)
    workbook.save(out_path)


class ContingencyCompareApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Contingency Compare")
        self.root.geometry("860x520")
        self.root.minsize(760, 460)

        self.selected_folder = ""
        self.file_paths = {}
        self.path_vars = {slot: tk.StringVar(value="Not found") for slot in SLOTS}

        self._build_ui()

    def _build_ui(self):
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(6, weight=1)

        title = ttk.Label(container, text="Contingency Compare", font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        select_button = tk.Button(
            container,
            text="Select Folder",
            font=("Segoe UI", 14, "bold"),
            padx=16,
            pady=12,
            command=self.select_folder,
        )
        select_button.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 16))

        for row_idx, slot in enumerate(SLOTS, start=2):
            ttk.Label(container, text=slot.replace("_", " "), width=12).grid(row=row_idx, column=0, sticky="w", pady=4)
            entry = ttk.Entry(container, textvariable=self.path_vars[slot], state="readonly")
            entry.grid(row=row_idx, column=1, sticky="ew", pady=4)

        self.status_label = ttk.Label(container, text="Select a folder to begin.", foreground="#333333")
        self.status_label.grid(row=6, column=0, columnspan=2, sticky="w", pady=(12, 4))

        self.status_text = tk.Text(container, height=10, wrap="word", state="disabled")
        self.status_text.grid(row=7, column=0, columnspan=2, sticky="nsew", pady=(0, 12))

        self.run_button = ttk.Button(container, text="Run Compare", command=self.run_compare, state="disabled")
        self.run_button.grid(row=8, column=0, columnspan=2, sticky="ew")

    def set_status(self, message, append=False, error=False):
        self.status_label.config(
            text=message.splitlines()[0] if message else "",
            foreground="#b00020" if error else "#333333",
        )
        self.status_text.config(state="normal")
        if not append:
            self.status_text.delete("1.0", tk.END)
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state="disabled")
        self.root.update_idletasks()

    def select_folder(self):
        folder = filedialog.askdirectory()
        if not folder:
            return

        self.selected_folder = folder
        detection = find_con_files(folder)
        self.file_paths = detection["selected"]

        for slot in SLOTS:
            path = self.file_paths.get(slot, "")
            self.path_vars[slot].set(path if path else "Not found")

        messages = [f"Selected folder: {folder}"]
        if detection["warnings"]:
            messages.extend(f"Warning: {warning}" for warning in detection["warnings"])

        if detection["missing"]:
            missing_text = ", ".join(detection["missing"])
            messages.append(f"Missing: {missing_text}")
            self.run_button.config(state="disabled")
            self.set_status("\n".join(messages), error=True)
        else:
            messages.append("All required files found.")
            self.run_button.config(state="normal")
            self.set_status("\n".join(messages))

    def run_compare(self):
        if not self.selected_folder or len(self.file_paths) != len(SLOTS):
            self.set_status("Missing required files. Select a valid folder first.", error=True)
            self.run_button.config(state="disabled")
            return

        try:
            self.set_status("Parsing...")
            p1_2026 = parse_con_file(self.file_paths["P1_2026"])
            p1_2030 = parse_con_file(self.file_paths["P1_2030"])
            p2_2026 = parse_con_file(self.file_paths["P2_2026"])
            p2_2030 = parse_con_file(self.file_paths["P2_2030"])

            self.set_status("Comparing...", append=True)
            results_p1 = compare_sets(p1_2026, p1_2030)
            results_p2 = compare_sets(p2_2026, p2_2030)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            output_path = os.path.join(self.selected_folder, f"Contingency_Compare_{timestamp}.xlsx")

            self.set_status("Writing report...", append=True)
            write_excel_report(output_path, results_p1, results_p2, self.file_paths)

            self.set_status(f"Done: {output_path}", append=True)
            self._show_completion(output_path)
        except Exception as exc:
            self.set_status(f"Error: {exc}", append=True, error=True)
            messagebox.showerror("Contingency Compare", str(exc))

    def _show_completion(self, output_path):
        messagebox.showinfo("Contingency Compare", f"Report written to:\n{output_path}")
        if messagebox.askyesno("Open Folder", "Open the output folder in File Explorer?"):
            self.open_output_location(output_path)

    @staticmethod
    def open_output_location(output_path):
        folder = os.path.dirname(output_path)
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception:
            pass


def main():
    root = tk.Tk()
    ttk.Style().theme_use("clam")
    app = ContingencyCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
