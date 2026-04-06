from __future__ import annotations

import os
import tkinter as tk
from tkinter import ttk, messagebox

from core.conductor_loader import load_conductor_database, ConductorDatabase
from models.conductor import Conductor


class LineRatingApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        self.title("Line Rating Calculator")
        self.geometry("1100x700")
        self.minsize(950, 600)

        self.database: ConductorDatabase | None = None
        self.selected_conductor: Conductor | None = None

        self.family_var = tk.StringVar()
        self.conductor_var = tk.StringVar()

        self._build_ui()
        self._load_database()

    def _build_ui(self) -> None:
        top_frame = ttk.Frame(self, padding=10)
        top_frame.pack(fill="x")

        ttk.Label(top_frame, text="Conductor Family:").grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.family_combo = ttk.Combobox(
            top_frame,
            textvariable=self.family_var,
            state="readonly",
            width=20
        )
        self.family_combo.grid(row=0, column=1, sticky="w", padx=(0, 20))
        self.family_combo.bind("<<ComboboxSelected>>", self._on_family_changed)

        ttk.Label(top_frame, text="Conductor:").grid(row=0, column=2, sticky="w", padx=(0, 8))

        self.conductor_combo = ttk.Combobox(
            top_frame,
            textvariable=self.conductor_var,
            state="readonly",
            width=30
        )
        self.conductor_combo.grid(row=0, column=3, sticky="w", padx=(0, 20))
        self.conductor_combo.bind("<<ComboboxSelected>>", self._on_conductor_changed)

        self.reload_button = ttk.Button(top_frame, text="Reload Data", command=self._load_database)
        self.reload_button.grid(row=0, column=4, sticky="w")

        middle_frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        middle_frame.pack(fill="both", expand=True)

        left_frame = ttk.LabelFrame(middle_frame, text="Conductor Data", padding=10)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        self.data_tree = ttk.Treeview(
            left_frame,
            columns=("property", "value"),
            show="headings",
            height=25
        )
        self.data_tree.heading("property", text="Property")
        self.data_tree.heading("value", text="Value")
        self.data_tree.column("property", width=260, anchor="w")
        self.data_tree.column("value", width=220, anchor="w")
        self.data_tree.pack(fill="both", expand=True)

        right_frame = ttk.LabelFrame(middle_frame, text="Rating Inputs (Next Step)", padding=10)
        right_frame.pack(side="right", fill="y", padx=(5, 0))

        self.input_vars = {
            "ambient_temp_c": tk.StringVar(value="25"),
            "wind_speed_mps": tk.StringVar(value="0.61"),
            "wind_angle_deg": tk.StringVar(value="90"),
            "elevation_m": tk.StringVar(value="0"),
            "solar_w_m2": tk.StringVar(value="1000"),
            "target_temp_c": tk.StringVar(value="75"),
            "emissivity": tk.StringVar(value="0.5"),
            "absorptivity": tk.StringVar(value="0.5"),
        }

        input_rows = [
            ("Ambient Temp (C)", "ambient_temp_c"),
            ("Wind Speed (m/s)", "wind_speed_mps"),
            ("Wind Angle (deg)", "wind_angle_deg"),
            ("Elevation (m)", "elevation_m"),
            ("Solar (W/m^2)", "solar_w_m2"),
            ("Target Temp (C)", "target_temp_c"),
            ("Emissivity", "emissivity"),
            ("Absorptivity", "absorptivity"),
        ]

        for row_idx, (label, key) in enumerate(input_rows):
            ttk.Label(right_frame, text=label).grid(row=row_idx, column=0, sticky="w", pady=4, padx=(0, 8))
            ttk.Entry(right_frame, textvariable=self.input_vars[key], width=18).grid(
                row=row_idx, column=1, sticky="w", pady=4
            )

        ttk.Separator(right_frame, orient="horizontal").grid(
            row=len(input_rows), column=0, columnspan=2, sticky="ew", pady=12
        )

        self.calculate_button = ttk.Button(
            right_frame,
            text="Calculate Rating",
            command=self._calculate_placeholder
        )
        self.calculate_button.grid(row=len(input_rows) + 1, column=0, columnspan=2, sticky="ew")

        self.status_var = tk.StringVar(value="Ready.")
        status_bar = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w", padding=6)
        status_bar.pack(fill="x", side="bottom")

    def _load_database(self) -> None:
        try:
            filepath = os.path.join("Resources", "ConData.xlsx")
            self.database = load_conductor_database(filepath)

            families = self.database.get_families()
            self.family_combo["values"] = families

            if not families:
                self.family_var.set("")
                self.conductor_combo["values"] = []
                self._clear_data_tree()
                self.status_var.set("No conductor sheets found.")
                return

            first_family = families[0]
            self.family_var.set(first_family)
            self._populate_conductors(first_family)

            total_count = sum(len(self.database.get_conductors(f)) for f in families)
            self.status_var.set(f"Loaded {total_count} conductors from {len(families)} sheet(s).")

        except Exception as exc:
            messagebox.showerror("Load Error", str(exc))
            self.status_var.set("Failed to load conductor database.")

    def _populate_conductors(self, family: str) -> None:
        if self.database is None:
            return

        conductors = self.database.get_conductors(family)
        names = [c.code_word for c in conductors if c.code_word]

        self.conductor_combo["values"] = names

        if names:
            self.conductor_var.set(names[0])
            self._display_selected_conductor(family, names[0])
        else:
            self.conductor_var.set("")
            self.selected_conductor = None
            self._clear_data_tree()
            self.status_var.set(f"No conductors found in sheet '{family}'.")

    def _on_family_changed(self, _event=None) -> None:
        family = self.family_var.get().strip()
        self._populate_conductors(family)

    def _on_conductor_changed(self, _event=None) -> None:
        family = self.family_var.get().strip()
        code_word = self.conductor_var.get().strip()
        self._display_selected_conductor(family, code_word)

    def _display_selected_conductor(self, family: str, code_word: str) -> None:
        if self.database is None:
            return

        conductor = self.database.find_conductor(family, code_word)
        self.selected_conductor = conductor
        self._clear_data_tree()

        if conductor is None:
            self.status_var.set(f"Could not find conductor '{code_word}' in '{family}'.")
            return

        if conductor.max_temp_c is not None:
            self.input_vars["target_temp_c"].set(str(conductor.max_temp_c))
        if conductor.emissivity is not None:
            self.input_vars["emissivity"].set(str(conductor.emissivity))
        if conductor.absorptivity is not None:
            self.input_vars["absorptivity"].set(str(conductor.absorptivity))

        data_rows = [
            ("Family", conductor.family),
            ("Code Word", conductor.code_word),
            ("Name", conductor.name),
            ("Size (kcmil)", conductor.size_kcmil),
            ("Stranding", conductor.stranding),
            ("Al Area (in^2)", conductor.al_area_in2),
            ("Total Area (in^2)", conductor.total_area_in2),
            ("Al Layers", conductor.al_layers),
            ("Al Strand Dia (in)", conductor.al_strand_dia_in),
            ("Steel Strand Dia (in)", conductor.steel_strand_dia_in),
            ("Steel Core Dia (in)", conductor.steel_core_dia_in),
            ("OD (in)", conductor.od_in),
            ("Al Weight (lb/kft)", conductor.al_weight_lb_per_kft),
            ("Steel Weight (lb/kft)", conductor.steel_weight_lb_per_kft),
            ("Total Weight (lb/kft)", conductor.total_weight_lb_per_kft),
            ("Al Percent", conductor.al_percent),
            ("Steel Percent", conductor.steel_percent),
            ("RBS (klb)", conductor.rbs_klb),
            ("DC Res @20C (ohm/mile)", conductor.dc_res_20c_ohm_per_mile),
            ("AC Res @25C (ohm/mile)", conductor.ac_res_25c_ohm_per_mile),
            ("AC Res @50C (ohm/mile)", conductor.ac_res_50c_ohm_per_mile),
            ("AC Res @75C (ohm/mile)", conductor.ac_res_75c_ohm_per_mile),
            ("GMR (ft)", conductor.gmr_ft),
            ("Xa @60Hz (ohm/mile)", conductor.xa_60hz_ohm_per_mile),
            ("Capacitive Reactance", conductor.capacitive_reactance),
            ("Ampacity @75C (amp)", conductor.ampacity_75c_amp),
            ("Emissivity", conductor.emissivity),
            ("Absorptivity", conductor.absorptivity),
            ("Max Temp (C)", conductor.max_temp_c),
        ]

        for prop, value in data_rows:
            self.data_tree.insert("", "end", values=(prop, "" if value is None else value))

        self.status_var.set(f"Selected {family} / {code_word}")

    def _clear_data_tree(self) -> None:
        for item in self.data_tree.get_children():
            self.data_tree.delete(item)

    def _calculate_placeholder(self) -> None:
        if self.selected_conductor is None:
            messagebox.showwarning("No Conductor", "Please select a conductor first.")
            return

        messagebox.showinfo(
            "Coming Next",
            f"GUI and conductor loading are working.\n\nSelected conductor: {self.selected_conductor.code_word}"
        )