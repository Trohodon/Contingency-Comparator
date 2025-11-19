# gui/splitscreen.py

from __future__ import annotations
import tkinter as tk
from tkinter import ttk


class SplitscreenManager:
    """
    Responsible for arranging N comparison panels (1–4) in a grid.

    Layouts:
      1 → 1 big panel
      2 → 2 vertical panels (left/right)
      3 → 2 top, 1 full-width bottom
      4 → 2x2 grid
    """

    def __init__(self, parent: tk.Widget):
        self.parent = parent
        self._reset_grid_config()

    def _reset_grid_config(self):
        for r in range(2):
            self.parent.rowconfigure(r, weight=1)
        for c in range(2):
            self.parent.columnconfigure(c, weight=1)

    def create_layout(self, num_panels: int) -> list[ttk.Frame]:
        # clear old children
        for child in self.parent.winfo_children():
            child.destroy()

        self._reset_grid_config()
        frames: list[ttk.Frame] = []

        def make_frame(row, col, rowspan=1, colspan=1):
            f = ttk.Frame(self.parent, borderwidth=2, relief="groove")
            f.grid(row=row, column=col,
                   rowspan=rowspan, columnspan=colspan,
                   sticky="nsew", padx=3, pady=3)
            frames.append(f)

        if num_panels <= 1:
            make_frame(0, 0, rowspan=2, colspan=2)
        elif num_panels == 2:
            make_frame(0, 0, rowspan=2)
            make_frame(0, 1, rowspan=2)
        elif num_panels == 3:
            make_frame(0, 0)
            make_frame(0, 1)
            make_frame(1, 0, colspan=2)
        else:  # 4
            make_frame(0, 0)
            make_frame(0, 1)
            make_frame(1, 0)
            make_frame(1, 1)

        return frames

