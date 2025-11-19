# gui/app.py

from __future__ import annotations
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont

from .splitscreen import SplitscreenManager
from .program import excel_logic


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Contingency Sheet Comparison")
        self.geometry("1500x850")

        self.workbook_data = None
        self.sheet_names: list[str] = []

        self._build_top_bar()

        # container for split-screen frames
        self.center_container = ttk.Frame(self)
        self.center_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.split_manager = SplitscreenManager(self.center_container)
        self.comparison_panels: list[ComparisonPanel] = []

        # default = 2 panels
        self.num_panels_var.set(2)
        self._rebuild_panels(2)

    # ───────────────── top bar ───────────────── #

    def _build_top_bar(self):
        bar = ttk.Frame(self)
        bar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        btn_open = ttk.Button(bar, text="Open Excel Workbook", command=self.open_workbook)
        btn_open.pack(side=tk.LEFT)

        self.file_label_var = tk.StringVar(value="No file loaded")
        ttk.Label(bar, textvariable=self.file_label_var).pack(side=tk.LEFT, padx=10)

        ttk.Label(bar, text="Number of comparisons:").pack(side=tk.LEFT, padx=(20, 5))
        self.num_panels_var = tk.IntVar(value=2)
        spin = ttk.Spinbox(
            bar,
            from_=1,
            to=4,
            width=3,
            textvariable=self.num_panels_var,
            command=self._on_num_panels_changed,
        )
        spin.pack(side=tk.LEFT)

    # ───────────────── workbook ───────────────── #

    def open_workbook(self):
        path = filedialog.askopenfilename(
            title="Select Excel Workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
        )
        if not path:
            return

        try:
            data = excel_logic.load_workbook(path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not load workbook:\n{e}")
            return

        self.workbook_data = data
        self.sheet_names = list(data.keys())
        self.file_label_var.set(f"Loaded: {os.path.basename(path)}")

        for panel in self.comparison_panels:
            panel.set_sheet_options(self.sheet_names)

    # ───────────────── panels / split-screen ───────────────── #

    def _on_num_panels_changed(self):
        n = self.num_panels_var.get()
        n = min(max(n, 1), 4)
        self.num_panels_var.set(n)
        self._rebuild_panels(n)

    def _rebuild_panels(self, num_panels: int):
        frames = self.split_manager.create_layout(num_panels)

        for p in self.comparison_panels:
            p.destroy()
        self.comparison_panels.clear()

        for idx, frame in enumerate(frames, start=1):
            panel = ComparisonPanel(frame, title=f"Comparison {idx}", app=self)
            panel.pack(fill=tk.BOTH, expand=True)
            if self.sheet_names:
                panel.set_sheet_options(self.sheet_names)
            self.comparison_panels.append(panel)


# ───────────────── Single comparison panel ───────────────── #

class ComparisonPanel(ttk.LabelFrame):
    def __init__(self, parent, title: str, app: MainApp):
        super().__init__(parent, text=title)
        self.app = app

        # base font size for this panel
        self.base_font_size = 9
        self.row_font = tkfont.Font(family="Segoe UI", size=self.base_font_size)

        # top controls (sheet selectors + Compare)
        ctrl = ttk.Frame(self)
        ctrl.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(5, 0))

        ttk.Label(ctrl, text="Left sheet:").pack(side=tk.LEFT)
        self.left_sheet_var = tk.StringVar()
        self.left_combo = ttk.Combobox(
            ctrl, textvariable=self.left_sheet_var, state="readonly", width=22
        )
        self.left_combo.pack(side=tk.LEFT, padx=3)

        ttk.Label(ctrl, text="Right sheet:").pack(side=tk.LEFT)
        self.right_sheet_var = tk.StringVar()
        self.right_combo = ttk.Combobox(
            ctrl, textvariable=self.right_sheet_var, state="readonly", width=22
        )
        self.right_combo.pack(side=tk.LEFT, padx=3)

        btn = ttk.Button(ctrl, text="Compare", command=self.do_compare)
        btn.pack(side=tk.LEFT, padx=5)

        # zoom slider
        zoom_frame = ttk.Frame(self)
        zoom_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(2, 5))

        ttk.Label(zoom_frame, text="Zoom:").pack(side=tk.LEFT)
        self.zoom_var = tk.DoubleVar(value=1.0)
        zoom_slider = ttk.Scale(
            zoom_frame,
            from_=0.6,
            to=1.6,
            orient="horizontal",
            variable=self.zoom_var,
            command=self._on_zoom_changed,
        )
        zoom_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

        # notebook for tables
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.table_views = {}
        for tname in excel_logic.TABLE_NAMES:
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=tname)
            self.table_views[tname] = self._create_table_widget(frame)

    # ───── zoom handling ───── #

    def _on_zoom_changed(self, *_):
        factor = float(self.zoom_var.get())
        size = max(6, int(self.base_font_size * factor))
        self.row_font.configure(size=size)

    # ───── combobox options ───── #

    def set_sheet_options(self, sheet_names: list[str]):
        vals = list(sheet_names)
        self.left_combo["values"] = vals
        self.right_combo["values"] = vals

        if vals:
            self.left_sheet_var.set(vals[0])
            if len(vals) > 1:
                self.right_sheet_var.set(vals[1])
            else:
                self.right_sheet_var.set(vals[0])

    # ───── Treeview helpers ───── #

    def _create_table_widget(self, parent):
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

        # apply font via tag
        tree.tag_configure("zoom", font=self.row_font)

        return tree

    def _clear_tree(self, tree: ttk.Treeview):
        for item in tree.get_children():
            tree.delete(item)

    def _populate_tree(self, tree: ttk.Treeview, df):
        self._clear_tree(tree)
        if df is None or df.empty:
            return

        for _, row in df.iterrows():
            def fmt(x):
                try:
                    return f"{float(x):.2f}"
                except Exception:
                    return ""
            values = (
                str(row.get("contingency", "")),
                str(row.get("issue", "")),
                fmt(row.get("percent_1")),
                fmt(row.get("percent_2")),
                fmt(row.get("delta_percent")),
                str(row.get("status", "")),
            )
            tree.insert("", tk.END, values=values, tags=("zoom",))

    # ───── compare button ───── #

    def do_compare(self):
        if self.app.workbook_data is None:
            messagebox.showwarning("No workbook", "Load an Excel workbook first.")
            return

        left = self.left_sheet_var.get()
        right = self.right_sheet_var.get()
        if not left or not right:
            messagebox.showwarning("Select sheets", "Choose both left and right sheets.")
            return

        try:
            results = excel_logic.compare_sheet_pair(
                self.app.workbook_data, left, right
            )
        except Exception as e:
            messagebox.showerror("Error", f"Comparison failed:\n{e}")
            return

        if not results:
            messagebox.showinfo(
                "No tables",
                "No matching ACCA Long Term/ACCA/DCwAC tables were found "
                "on both selected sheets.\n\nCheck the console for 'found tables' "
                "messages to see what was detected.",
            )

        for tname, tree in self.table_views.items():
            df = results.get(tname)
            self._populate_tree(tree, df)


# ───────────────── public entry ───────────────── #

def run_app():
    app = MainApp()
    app.mainloop()