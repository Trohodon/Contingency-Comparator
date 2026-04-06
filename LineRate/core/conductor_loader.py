from __future__ import annotations

import os
from typing import Dict, List, Optional

import pandas as pd

from models.conductor import Conductor


class ConductorDatabase:
    def __init__(self) -> None:
        self.by_family: Dict[str, List[Conductor]] = {}

    def add_family(self, family: str, conductors: List[Conductor]) -> None:
        self.by_family[family] = conductors

    def get_families(self) -> List[str]:
        return sorted(self.by_family.keys())

    def get_conductors(self, family: str) -> List[Conductor]:
        return self.by_family.get(family, [])

    def find_conductor(self, family: str, code_word: str) -> Optional[Conductor]:
        for conductor in self.get_conductors(family):
            if conductor.code_word.strip().upper() == code_word.strip().upper():
                return conductor
        return None


def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip().upper() for col in df.columns]
    return df


def _to_float(value) -> Optional[float]:
    if pd.isna(value) or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _to_int(value) -> Optional[int]:
    if pd.isna(value) or value == "":
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def load_conductor_database(filepath: str) -> ConductorDatabase:
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Conductor data file not found: {filepath}")

    workbook = pd.read_excel(filepath, sheet_name=None, engine="openpyxl")
    database = ConductorDatabase()

    for sheet_name, df in workbook.items():
        df = _clean_columns(df)
        df = df.dropna(how="all")

        conductors: List[Conductor] = []

        for _, row in df.iterrows():
            code_word = str(row.get("CODE_WORD", "")).strip()
            if not code_word:
                continue

            conductor = Conductor(
                family=sheet_name,
                code_word=code_word,
                size_kcmil=_to_float(row.get("SIZE_KCMIL")),
                stranding=str(row.get("STRANDING", "")).strip() or None,

                al_area_in2=_to_float(row.get("AL_AREA_IN2")),
                total_area_in2=_to_float(row.get("TOTAL_AREA_IN2")),
                al_layers=_to_int(row.get("AL_LAYERS")),

                al_strand_dia_in=_to_float(row.get("AL_STRAND_DIA_IN")),
                steel_strand_dia_in=_to_float(row.get("STEEL_STRAND_DIA_IN")),
                steel_core_dia_in=_to_float(row.get("STEEL_CORE_DIA_IN")),
                od_in=_to_float(row.get("OD_IN")),

                al_weight_lb_per_kft=_to_float(row.get("AL_WEIGHT_LB_PER_KFT")),
                steel_weight_lb_per_kft=_to_float(row.get("STEEL_WEIGHT_LB_PER_KFT")),
                total_weight_lb_per_kft=_to_float(row.get("TOTAL_WEIGHT_LB_PER_KFT")),

                al_percent=_to_float(row.get("AL_PERCENT")),
                steel_percent=_to_float(row.get("STEEL_PERCENT")),
                rbs_klb=_to_float(row.get("RBS_KLB")),

                dc_res_20c_ohm_per_mile=_to_float(row.get("DC_RES_20C_OHM_PER_MILE")),
                ac_res_25c_ohm_per_mile=_to_float(row.get("AC_RES_25C_OHM_PER_MILE")),
                ac_res_50c_ohm_per_mile=_to_float(row.get("AC_RES_50C_OHM_PER_MILE")),
                ac_res_75c_ohm_per_mile=_to_float(row.get("AC_RES_75C_OHM_PER_MILE")),

                gmr_ft=_to_float(row.get("GMR_FT")),
                xa_60hz_ohm_per_mile=_to_float(row.get("XA_60HZ_OHM_PER_MILE")),
                capacitive_reactance=_to_float(row.get("CAPACITIVE_REACTANCE")),
                ampacity_75c_amp=_to_float(row.get("AMPACITY_75C_AMP")),

                name=str(row.get("NAME", "")).strip() or code_word,
                emissivity=_to_float(row.get("EMISSIVITY")),
                absorptivity=_to_float(row.get("ABSORPTIVITY")),
                max_temp_c=_to_float(row.get("MAX_TEMP_C")),
            )

            conductors.append(conductor)

        database.add_family(sheet_name, conductors)

    return database