# gui.py
"""
Tkinter GUI for visualizing comparisons between Excel sheets.

Layout:
- Top: "Open Workbook" button + label with filename.
- Middle: two side-by-side comparison panels.
  Each panel lets you choose:
    - Left sheet
    - Right sheet
    - Compare button
  and shows results for: "ACCA Long Term", "ACCA", "DCwAC"
"""

from __future__ import annotations
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import program  # our comparison logic module


class ComparisonApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Contingency Sheet Comparison")
        self.geometry("1400x800")

        self.workbook_data = None  # filled by program.load_workbook()
        self.sheet_names = []

        self._build_top_bar()
        self._build_comparison_frames()

    # ---------- UI construction ----------

    def _build_top_bar(self):
        top = ttk.Frame(self)
        top.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        btn = ttk.Button(top, text="Open Excel Workbook", command=self.open_workbook)
        btn.pack(side=tk.LEFT)

        self.file_label_var = tk.StringVar(value="No file loaded")
        lbl = ttk.Label(top, textvariable=self.file_label_var)
        lbl.pack(side=tk.LEFT, padx=10)

    def _build_comparison_frames(self):
        container = ttk.Frame(self)
        container.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.comp_frame_a = SheetComparisonFrame(container, title="Comparison A")
        self.comp_frame_a.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        self.comp_frame_b = SheetComparisonFrame(container, title="Comparison B")
        self.comp_frame_b.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # Give frames a handle back to the app so they can see workbook data
        self.comp_frame_a.app = self
        self.comp_frame_b.app = self

    # ---------- callbacks ----------

    def open_workbook(self):
        path = filedialog.askopenfilename(
            title="Select Excel Workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
        )
        if not path:
            return

        try:
            data = program.load_workbook(path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not load workbook:\n{e}")
            return

        self.workbook_data = data
        self.sheet_names = list(data.keys())

        fname = os.path.basename(path)
        self.file_label_var.set(f"Loaded: {fname}")

        # Update drop-downs in both panels
        self.comp_frame_a.set_sheet_options(self.sheet_names)
        self.comp_frame_b.set_sheet_options(self.sheet_names)


class SheetComparisonFrame(ttk.LabelFrame):
    """
    One comparison "window" (left or right side).

    Lets the user choose two sheets and shows comparison
    for ACCA Long Term, ACCA, and DCwAC in a Notebook.
    """

    def __init__(self, parent, title="Comparison"):
        super().__init__(parent, text=title)
        self.app: ComparisonApp | None = None

        # Top controls (sheet selectors + button)
        ctrl = ttk.Frame(self)
        ctrl.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        ttk.Label(ctrl, text="Left sheet:").pack(side=tk.LEFT)
        self.left_sheet_var = tk.StringVar()
        self.left_combo = ttk.Combobox(ctrl, textvariable=self.left_sheet_var, width=18, state="readonly")
        self.left_combo.pack(side=tk.LEFT, padx=3)

        ttk.Label(ctrl, text="Right sheet:").pack(side=tk.LEFT)
        self.right_sheet_var = tk.StringVar()
        self.right_combo = ttk.Combobox(ctrl, textvariable=self.right_sheet_var, width=18, state="readonly")
        self.right_combo.pack(side=tk.LEFT, padx=3)

        btn = ttk.Button(ctrl, text="Compare", command=self.do_compare)
        btn.pack(side=tk.LEFT, padx=5)

        # Notebook with one tab per table type
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.tables = {}
        for tname in program.TABLE_NAMES:
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=tname)
            self.tables[tname] = self._create_table_widget(frame)

    def set_sheet_options(self, sheet_names):
        """Called by app after workbook is loaded."""
        values = list(sheet_names)
        self.left_combo["values"] = values
        self.right_combo["values"] = values

        if values:
            self.left_sheet_var.set(values[0])
            if len(values) > 1:
                self.right_sheet_var.set(values[1])
            else:
                self.right_sheet_var.set(values[0])

    # ----- Treeview creation & population -----

    def _create_table_widget(self, parent):
        """
        Create a Treeview + vertical scrollbar in the given parent frame.

        Columns:
        contingency, issue, percent_1, percent_2, delta_percent, status
        """
        cols = ("contingency", "issue", "percent_1", "percent_2", "delta_percent", "status")
        tree = ttk.Treeview(parent, columns=cols, show="headings")

        tree.heading("contingency", text="Contingency")
        tree.heading("issue", text="Resulting Issue")
        tree.heading("percent_1", text="Left %")
        tree.heading("percent_2", text="Right %")
        tree.heading("delta_percent", text="Δ% (Right - Left)")
        tree.heading("status", text="Status")

        tree.column("contingency", width=260, anchor=tk.W)
        tree.column("issue", width=420, anchor=tk.W)
        tree.column("percent_1", width=90, anchor=tk.E)
        tree.column("percent_2", width=90, anchor=tk.E)
        tree.column("delta_percent", width=120, anchor=tk.E)
        tree.column("status", width=110, anchor=tk.W)

        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        return tree

    def _clear_tree(self, tree: ttk.Treeview):
        for item in tree.get_children():
            tree.delete(item)

    def _populate_tree(self, tree: ttk.Treeview, df):
        self._clear_tree(tree)
        if df is None or df.empty:
            return

        for _, row in df.iterrows():
            p1 = row.get("percent_1")
            p2 = row.get("percent_2")
            dlt = row.get("delta_percent")
            def _fmt(x):
                try:
                    return f"{float(x):.2f}"
                except Exception:
                    return ""
            vals = (
                str(row.get("contingency", "")),
                str(row.get("issue", "")),
                _fmt(p1),
                _fmt(p2),
                _fmt(dlt),
                str(row.get("status", "")),
            )
            tree.insert("", tk.END, values=vals)

    # ----- comparison button -----

    def do_compare(self):
        if self.app is None or self.app.workbook_data is None:
            messagebox.showwarning("No workbook", "Please load an Excel workbook first.")
            return

        left = self.left_sheet_var.get()
        right = self.right_sheet_var.get()
        if not left or not right:
            messagebox.showwarning("Select sheets", "Please choose both left and right sheets.")
            return
        if left == right:
            res = messagebox.askyesno(
                "Same sheet selected",
                "Left and right sheets are the same. Continue anyway?"
            )
            if not res:
                return

        try:
            comparisons = program.compare_sheet_pair(self.app.workbook_data, left, right)
        except Exception as e:
            messagebox.showerror("Error", f"Comparison failed:\n{e}")
            return

        # Update each tab
        for tname, tree in self.tables.items():
            df = comparisons.get(tname)
            self._populate_tree(tree, df)
