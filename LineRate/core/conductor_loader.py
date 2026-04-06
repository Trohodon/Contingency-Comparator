from __future__ import annotations

import os
from typing import Dict, List, Optional

import pandas as pd

from models.conductor import Conductor


class ConductorDatabase:
    def __init__(self) -> None:
        self.by_family: Dict[str, List[Conductor]] = {}
        self.source_path: Optional[str] = None

    def add_family(self, family: str, conductors: List[Conductor]) -> None:
        self.by_family[family] = conductors

    def get_families(self) -> List[str]:
        return sorted(self.by_family.keys())

    def get_conductors(self, family: str) -> List[Conductor]:
        return self.by_family.get(family, [])

    def find_conductor(self, family: str, code_word: str) -> Optional[Conductor]:
        for conductor in self.get_conductors(family):
            if str(conductor.code_word).strip().upper() == str(code_word).strip().upper():
                return conductor
        return None


def _normalize_column_name(col_name: str) -> str:
    name = str(col_name).strip().upper()
    name = name.replace("\n", "_")
    name = name.replace(" ", "_")
    name = name.replace("-", "_")
    name = name.replace("/", "_")
    name = name.replace("(", "")
    name = name.replace(")", "")
    name = name.replace(".", "")
    return name


def _clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_normalize_column_name(col) for col in df.columns]
    df = df.dropna(how="all")
    return df


def _is_blank(value) -> bool:
    if value is None:
        return True
    if pd.isna(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _to_float(value) -> Optional[float]:
    if _is_blank(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _to_int(value) -> Optional[int]:
    if _is_blank(value):
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def _to_str(value) -> Optional[str]:
    if _is_blank(value):
        return None
    return str(value).strip()


def _get_first_present(row, possible_names: list[str]):
    for name in possible_names:
        if name in row.index:
            return row.get(name)
    return None


def _looks_like_consizes_workbook(df: pd.DataFrame) -> bool:
    cols = set(df.columns)
    required = {"CODE_NAME", "TYPE", "R25", "R75", "OD_IN"}
    return required.issubset(cols)


def _build_conductor_from_row(sheet_name: str, row, family_override: Optional[str] = None) -> Optional[Conductor]:
    code_word_raw = _get_first_present(row, ["CODE_WORD", "CODE_NAME", "CODEWORD", "CODE"])
    code_word = _to_str(code_word_raw)

    if code_word is None:
        return None

    family = family_override if family_override else sheet_name

    conductor = Conductor(
        family=family,
        code_word=code_word,
        size_kcmil=_to_float(_get_first_present(row, ["SIZE_KCMIL", "SIZE"])),
        stranding=_to_str(_get_first_present(row, ["STRANDING", "STRAND"])),

        al_area_in2=_to_float(_get_first_present(row, ["AL_AREA_IN2"])),
        total_area_in2=_to_float(_get_first_present(row, ["TOTAL_AREA_IN2", "AREA_SQIN"])),
        al_layers=_to_int(_get_first_present(row, ["AL_LAYERS"])),

        al_strand_dia_in=_to_float(_get_first_present(row, ["AL_STRAND_DIA_IN", "DIAM_OUTERIN", "DIAM_OUTER_IN"])),
        steel_strand_dia_in=_to_float(_get_first_present(row, ["STEEL_STRAND_DIA_IN", "DIAM_INNERIN", "DIAM_INNER_IN"])),
        steel_core_dia_in=_to_float(_get_first_present(row, ["STEEL_CORE_DIA_IN"])),
        od_in=_to_float(_get_first_present(row, ["OD_IN", "COMPLETE_DIAMETER_IN"])),

        al_weight_lb_per_kft=_to_float(_get_first_present(row, ["AL_WEIGHT_LB_PER_KFT", "LBS_KFT_OUTER"])),
        steel_weight_lb_per_kft=_to_float(_get_first_present(row, ["STEEL_WEIGHT_LB_PER_KFT", "LBS_KFT_INNER"])),
        total_weight_lb_per_kft=_to_float(_get_first_present(row, ["TOTAL_WEIGHT_LB_PER_KFT"])),

        al_percent=_to_float(_get_first_present(row, ["AL_PERCENT"])),
        steel_percent=_to_float(_get_first_present(row, ["STEEL_PERCENT"])),
        rbs_klb=_to_float(_get_first_present(row, ["RBS_KLB", "RBS"])),

        dc_res_20c_ohm_per_mile=_to_float(_get_first_present(row, ["DC_RES_20C_OHM_PER_MILE"])),
        ac_res_25c_ohm_per_mile=_to_float(_get_first_present(row, ["AC_RES_25C_OHM_PER_MILE", "R25"])),
        ac_res_50c_ohm_per_mile=_to_float(_get_first_present(row, ["AC_RES_50C_OHM_PER_MILE"])),
        ac_res_75c_ohm_per_mile=_to_float(_get_first_present(row, ["AC_RES_75C_OHM_PER_MILE", "R75"])),

        gmr_ft=_to_float(_get_first_present(row, ["GMR_FT"])),
        xa_60hz_ohm_per_mile=_to_float(_get_first_present(row, ["XA_60HZ_OHM_PER_MILE"])),
        capacitive_reactance=_to_float(_get_first_present(row, ["CAPACITIVE_REACTANCE"])),
        ampacity_75c_amp=_to_float(_get_first_present(row, ["AMPACITY_75C_AMP", "STDOL"])),

        name=_to_str(_get_first_present(row, ["NAME"])) or code_word,
        emissivity=_to_float(_get_first_present(row, ["EMISSIVITY"])),
        absorptivity=_to_float(_get_first_present(row, ["ABSORPTIVITY"])),
        max_temp_c=_to_float(_get_first_present(row, ["MAX_TEMP_C"])),
    )

    return conductor


def load_conductor_database(filepath: str) -> ConductorDatabase:
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Conductor data file not found: {filepath}")

    workbook = pd.read_excel(filepath, sheet_name=None, engine="openpyxl")
    database = ConductorDatabase()
    database.source_path = filepath

    for sheet_name, raw_df in workbook.items():
        df = _clean_dataframe(raw_df)
        conductors: List[Conductor] = []

        if _looks_like_consizes_workbook(df):
            for _, row in df.iterrows():
                family_value = _to_str(_get_first_present(row, ["TYPE"])) or sheet_name
                conductor = _build_conductor_from_row(sheet_name, row, family_override=family_value)
                if conductor is not None:
                    conductors.append(conductor)

            grouped: Dict[str, List[Conductor]] = {}
            for conductor in conductors:
                grouped.setdefault(conductor.family, []).append(conductor)

            for family_name, family_conductors in grouped.items():
                existing = database.get_conductors(family_name)
                database.add_family(family_name, existing + family_conductors)

        else:
            for _, row in df.iterrows():
                conductor = _build_conductor_from_row(sheet_name, row)
                if conductor is not None:
                    conductors.append(conductor)

            database.add_family(str(sheet_name), conductors)

    return database