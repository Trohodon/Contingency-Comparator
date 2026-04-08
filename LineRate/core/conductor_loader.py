from __future__ import annotations

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
    replacements = {
        "\n": "_",
        " ": "_",
        "-": "_",
        "/": "_",
        "(": "",
        ")": "",
        ".": "",
    }
    for old, new in replacements.items():
        name = name.replace(old, new)
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


def _looks_like_conductordata_workbook(df: pd.DataFrame) -> bool:
    cols = set(df.columns)
    required = {"TYPE", "CODE", "NAME", "RADIUSFT", "ROHMS_M", "GMRFT"}
    return required.issubset(cols)


def _extract_size_from_name(raw_name: Optional[str]) -> Optional[float]:
    if not raw_name:
        return None
    try:
        return float(raw_name)
    except ValueError:
        return None


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


def _build_conductor_from_conductordata_row(sheet_name: str, row) -> Optional[Conductor]:
    code_word = _to_str(_get_first_present(row, ["CODE"]))
    if code_word is None:
        return None

    family = _to_str(_get_first_present(row, ["TYPE"])) or sheet_name
    raw_name = _to_str(_get_first_present(row, ["NAME"]))
    pretty_name = _to_str(_get_first_present(row, ["NAME_1", "NAME1", "FULL_NAME", "FULLNAME"])) or raw_name or code_word

    radius_ft = _to_float(_get_first_present(row, ["RADIUSFT"]))
    od_in = radius_ft * 24.0 if radius_ft is not None else None

    resistance = _to_float(_get_first_present(row, ["ROHMS_M"]))
    gmr_ft = _to_float(_get_first_present(row, ["GMRFT"]))

    rate_a = _to_float(_get_first_present(row, ["RATEAA"]))
    rate_b = _to_float(_get_first_present(row, ["RATEBA"]))
    rate_c = _to_float(_get_first_present(row, ["RATECA"]))

    conductor = Conductor(
        family=family,
        code_word=code_word,
        size_kcmil=_extract_size_from_name(raw_name),
        stranding=None,

        al_area_in2=None,
        total_area_in2=None,
        al_layers=None,

        al_strand_dia_in=None,
        steel_strand_dia_in=None,
        steel_core_dia_in=None,
        od_in=od_in,

        al_weight_lb_per_kft=None,
        steel_weight_lb_per_kft=None,
        total_weight_lb_per_kft=None,

        al_percent=None,
        steel_percent=None,
        rbs_klb=None,

        dc_res_20c_ohm_per_mile=resistance,
        ac_res_25c_ohm_per_mile=resistance,
        ac_res_50c_ohm_per_mile=resistance,
        ac_res_75c_ohm_per_mile=resistance,

        gmr_ft=gmr_ft,
        xa_60hz_ohm_per_mile=_to_float(_get_first_present(row, ["XLOHMS_R", "XL_OHMS_R", "XLOHMSR"])),
        capacitive_reactance=_to_float(_get_first_present(row, ["XCOHMS_T", "XC_OHMS_T", "XCOHMST"])),
        ampacity_75c_amp=rate_b if rate_b is not None else (rate_a if rate_a is not None else rate_c),

        name=pretty_name,
        emissivity=None,
        absorptivity=None,
        max_temp_c=None,
    )

    return conductor


def load_conductor_database(filepath: str) -> ConductorDatabase:
    workbook = pd.read_excel(filepath, sheet_name=None, engine="openpyxl")
    database = ConductorDatabase()
    database.source_path = filepath

    for sheet_name, raw_df in workbook.items():
        df = _clean_dataframe(raw_df)

        if _looks_like_conductordata_workbook(df):
            grouped: Dict[str, List[Conductor]] = {}
            for _, row in df.iterrows():
                conductor = _build_conductor_from_conductordata_row(sheet_name, row)
                if conductor is None:
                    continue
                grouped.setdefault(conductor.family, []).append(conductor)

            for family_name, family_conductors in grouped.items():
                existing = database.get_conductors(family_name)
                database.add_family(family_name, existing + family_conductors)
            continue

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
