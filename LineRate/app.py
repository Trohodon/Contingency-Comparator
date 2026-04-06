from __future__ import annotations

import os
import tkinter as tk
from tkinter import ttk, messagebox

from core.conductor_loader import load_conductor_database, ConductorDatabase
from core.ieee738 import calculate_steady_state_rating
from core.solar_ieee738 import parse_date_input, parse_time_input
from models.conductor import Conductor


class LineRatingApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        self.title("Line Rating Calculator")
        self.geometry("1260x820")
        self.minsize(1080, 700)

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
            height=18
        )
        self.data_tree.heading("property", text="Property")
        self.data_tree.heading("value", text="Value")
        self.data_tree.column("property", width=290, anchor="w")
        self.data_tree.column("value", width=250, anchor="w")
        self.data_tree.pack(fill="both", expand=True)

        right_frame = ttk.LabelFrame(middle_frame, text="Rating Inputs", padding=10)
        right_frame.pack(side="right", fill="y", padx=(5, 0))

        self.input_vars = {
            "ambient_temp_c": tk.StringVar(value="40"),
            "wind_speed_mps": tk.StringVar(value="2.0013"),
            "wind_angle_deg": tk.StringVar(value="45"),
            "elevation_m": tk.StringVar(value="500"),
            "target_temp_c": tk.StringVar(value="125"),
            "emissivity": tk.StringVar(value="0.5"),
            "absorptivity": tk.StringVar(value="0.5"),
            "latitude_deg": tk.StringVar(value="32.234"),
            "line_azimuth_deg": tk.StringVar(value="0"),
            "date_text": tk.StringVar(value="7/1/20"),
            "time_text": tk.StringVar(value="7:00 AM"),
            "atmosphere_type": tk.StringVar(value="clear"),
        }

        input_rows = [
            ("Ambient Temp (C)", "ambient_temp_c"),
            ("Wind Speed (m/s)", "wind_speed_mps"),
            ("Wind Angle to Axis (deg)", "wind_angle_deg"),
            ("Elevation (m)", "elevation_m"),
            ("Target Temp (C)", "target_temp_c"),
            ("Emissivity", "emissivity"),
            ("Absorptivity", "absorptivity"),
            ("Latitude (deg)", "latitude_deg"),
            ("Line Azimuth (deg)", "line_azimuth_deg"),
            ("Date", "date_text"),
            ("Time", "time_text"),
        ]

        for row_idx, (label, key) in enumerate(input_rows):
            ttk.Label(right_frame, text=label).grid(row=row_idx, column=0, sticky="w", pady=4, padx=(0, 8))
            ttk.Entry(right_frame, textvariable=self.input_vars[key], width=18).grid(
                row=row_idx, column=1, sticky="w", pady=4
            )

        atmosphere_row = len(input_rows)
        ttk.Label(right_frame, text="Atmosphere").grid(row=atmosphere_row, column=0, sticky="w", pady=4, padx=(0, 8))
        atmosphere_combo = ttk.Combobox(
            right_frame,
            textvariable=self.input_vars["atmosphere_type"],
            state="readonly",
            values=["clear", "industrial"],
            width=15
        )
        atmosphere_combo.grid(row=atmosphere_row, column=1, sticky="w", pady=4)

        ttk.Separator(right_frame, orient="horizontal").grid(
            row=atmosphere_row + 1, column=0, columnspan=2, sticky="ew", pady=12
        )

        self.calculate_button = ttk.Button(
            right_frame,
            text="Calculate Rating",
            command=self._calculate_rating
        )
        self.calculate_button.grid(row=atmosphere_row + 2, column=0, columnspan=2, sticky="ew")

        result_frame = ttk.LabelFrame(self, text="Calculation Result", padding=10)
        result_frame.pack(fill="both", expand=False, padx=10, pady=(0, 10))

        self.result_text = tk.Text(result_frame, height=14, wrap="word")
        self.result_text.pack(fill="both", expand=True)
        self.result_text.insert("1.0", "Results will appear here after calculation.\n")
        self.result_text.configure(state="disabled")

        self.status_var = tk.StringVar(value="Ready.")
        status_bar = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w", padding=6)
        status_bar.pack(fill="x", side="bottom")

    def _find_data_source(self) -> str:
        resources_dir = "Resources"
        preferred_files = [
            "ConSizes.xlsx",
            "ConData.xlsx",
        ]

        for filename in preferred_files:
            path = os.path.join(resources_dir, filename)
            if os.path.exists(path):
                return path

        raise FileNotFoundError(
            "No conductor workbook found in Resources. Expected ConSizes.xlsx or ConData.xlsx."
        )

    def _load_database(self) -> None:
        try:
            filepath = self._find_data_source()
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
            source_name = os.path.basename(filepath)
            self.status_var.set(
                f"Loaded {total_count} conductors from {len(families)} family/families using {source_name}."
            )

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
            self.status_var.set(f"No conductors found in family '{family}'.")

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
        else:
            if family.strip().upper() == "ACSR":
                self.input_vars["target_temp_c"].set("125")

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
            ("Southwire Ampacity / STDOL", conductor.ampacity_75c_amp),
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

    def _set_result_text(self, text: str) -> None:
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", text)
        self.result_text.configure(state="disabled")

    def _get_float_input(self, key: str, label: str) -> float:
        raw = self.input_vars[key].get().strip()
        try:
            return float(raw)
        except ValueError:
            raise ValueError(f"Invalid numeric value for {label}: '{raw}'")

    def _calculate_rating(self) -> None:
        if self.selected_conductor is None:
            messagebox.showwarning("No Conductor", "Please select a conductor first.")
            return

        try:
            ambient_temp_c = self._get_float_input("ambient_temp_c", "Ambient Temp")
            wind_speed_mps = self._get_float_input("wind_speed_mps", "Wind Speed")
            wind_angle_deg = self._get_float_input("wind_angle_deg", "Wind Angle")
            elevation_m = self._get_float_input("elevation_m", "Elevation")
            target_temp_c = self._get_float_input("target_temp_c", "Target Temp")
            emissivity = self._get_float_input("emissivity", "Emissivity")
            absorptivity = self._get_float_input("absorptivity", "Absorptivity")
            latitude_deg = self._get_float_input("latitude_deg", "Latitude")
            line_azimuth_deg = self._get_float_input("line_azimuth_deg", "Line Azimuth")

            input_date = parse_date_input(self.input_vars["date_text"].get())
            input_time = parse_time_input(self.input_vars["time_text"].get())
            atmosphere_type = self.input_vars["atmosphere_type"].get().strip().lower()

            result = calculate_steady_state_rating(
                conductor=self.selected_conductor,
                ambient_temp_c=ambient_temp_c,
                wind_speed_mps=wind_speed_mps,
                wind_angle_deg=wind_angle_deg,
                elevation_m=elevation_m,
                target_temp_c=target_temp_c,
                emissivity=emissivity,
                absorptivity=absorptivity,
                latitude_deg=latitude_deg,
                line_azimuth_deg=line_azimuth_deg,
                input_date=input_date,
                input_time=input_time,
                atmosphere_type=atmosphere_type,
            )

            source_ref = os.path.basename(self.database.source_path) if self.database and self.database.source_path else "unknown"
            southwire_75c = self.selected_conductor.ampacity_75c_amp
            solar = result["solar"]

            result_lines = [
                f"Conductor: {self.selected_conductor.code_word}",
                f"Family: {self.selected_conductor.family}",
                f"Source Workbook: {source_ref}",
                "",
                "Inputs",
                f"  Ambient Temp (C): {ambient_temp_c:.3f}",
                f"  Wind Speed (m/s): {wind_speed_mps:.6f}",
                f"  Wind Speed (ft/s): {result['wind_speed_fps']:.6f}",
                f"  Wind Angle to Axis φ (deg): {wind_angle_deg:.3f}",
                f"  Wind Angle to Perpendicular β (deg): {result['beta_deg']:.3f}",
                f"  Elevation (m): {elevation_m:.3f}",
                f"  Elevation (ft): {result['elevation_ft']:.3f}",
                f"  Target Temp (C): {target_temp_c:.3f}",
                f"  Emissivity: {emissivity:.3f}",
                f"  Absorptivity: {absorptivity:.3f}",
                f"  Latitude (deg): {latitude_deg:.3f}",
                f"  Line Azimuth (deg): {line_azimuth_deg:.3f}",
                f"  Date: {input_date.isoformat()}",
                f"  Time: {input_time.strftime('%H:%M:%S')}",
                f"  Atmosphere: {atmosphere_type}",
                "",
                "Calculated Rating",
                f"  Ampacity (A): {result['amps']:.6f}",
                "",
                "Heat Terms (US internal match mode)",
                f"  qc total (W/ft): {result['qc_w_per_ft']:.6f}",
                f"  qcn natural (W/ft): {result['qcn_w_per_ft']:.6f}",
                f"  qc1 forced low-Re (W/ft): {result['qc1_w_per_ft']:.6f}",
                f"  qc2 forced high-Re (W/ft): {result['qc2_w_per_ft']:.6f}",
                f"  qr radiated (W/ft): {result['qr_w_per_ft']:.6f}",
                f"  qs solar (W/ft): {result['qs_w_per_ft']:.6f}",
                "",
                "Resistance / Air Properties",
                f"  Diameter (in): {result['diameter_in']:.6f}",
                f"  Diameter (ft): {result['diameter_ft']:.9f}",
                f"  Resistance (ohm/mile): {result['resistance_ohm_per_mile']:.6f}",
                f"  Resistance (ohm/ft): {result['resistance_ohm_per_ft']:.10f}",
                f"  Mean film temp (C): {result['tfilm_c']:.6f}",
                f"  Air density (lb/ft^3): {result['rho_f_lb_per_ft3']:.8f}",
                f"  Air viscosity (lb/ft-s): {result['mu_f_lb_per_ft_s']:.10f}",
                f"  Air conductivity (W/ft-C): {result['k_f_w_per_ft_c']:.8f}",
                f"  Reynolds number: {result['n_re']:.6f}",
                f"  Wind direction factor: {result['k_angle']:.6f}",
                "",
                "IEEE 738 Solar Details",
                f"  Day of year N: {solar['n_day']}",
                f"  Decimal hour: {solar['hour_decimal']:.6f}",
                f"  Hour angle ω (deg): {solar['omega_deg']:.6f}",
                f"  Solar declination δ (deg): {solar['delta_deg']:.6f}",
                f"  Solar altitude Hc (deg): {solar['hc_deg']:.6f}",
                f"  Solar azimuth variable χ: {solar['chi']:.6f}",
                f"  Solar azimuth constant C (deg): {solar['c_constant']:.6f}",
                f"  Solar azimuth Zc (deg): {solar['zc_deg']:.6f}",
                f"  Incidence angle θ (deg): {solar['theta_deg']:.6f}",
                f"  Sea-level solar intensity (W/ft^2): {solar['qs_sea_level_w_per_ft2']:.6f}",
                f"  Corrected solar intensity (W/ft^2): {solar['qse_w_per_ft2']:.6f}",
                f"  Elevation correction Ksolar: {solar['ksolar']:.6f}",
            ]

            if southwire_75c is not None:
                result_lines.extend([
                    "",
                    "Reference",
                    f"  Ampacity / STDOL from workbook: {southwire_75c:.3f}",
                ])

            self._set_result_text("\n".join(result_lines))
            self.status_var.set(f"Calculated steady-state rating for {self.selected_conductor.code_word}")

        except Exception as exc:
            messagebox.showerror("Calculation Error", str(exc))
            self.status_var.set("Calculation failed.")