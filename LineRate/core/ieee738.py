from __future__ import annotations

import math
from typing import Optional

from core.solar_ieee738 import solar_heat_gain
from models.conductor import Conductor


INCH_TO_FT = 1.0 / 12.0
FT_PER_M = 3.280839895013123
M_PER_FT = 1.0 / FT_PER_M
OHM_PER_MILE_TO_OHM_PER_FT = 1.0 / 5280.0
MPS_TO_FPS = FT_PER_M


def inch_to_foot(value_in: float) -> float:
    return value_in * INCH_TO_FT


def ohm_per_mile_to_ohm_per_ft(value: float) -> float:
    return value * OHM_PER_MILE_TO_OHM_PER_FT


def interpolate_resistance_ohm_per_mile(conductor: Conductor, temp_c: float) -> float:
    points = []

    if conductor.ac_res_25c_ohm_per_mile is not None:
        points.append((25.0, conductor.ac_res_25c_ohm_per_mile))
    if conductor.ac_res_50c_ohm_per_mile is not None:
        points.append((50.0, conductor.ac_res_50c_ohm_per_mile))
    if conductor.ac_res_75c_ohm_per_mile is not None:
        points.append((75.0, conductor.ac_res_75c_ohm_per_mile))

    points.sort(key=lambda x: x[0])

    if len(points) >= 2:
        if temp_c <= points[0][0]:
            return _linear_interp(points[0], points[1], temp_c)

        for i in range(len(points) - 1):
            t1, r1 = points[i]
            t2, r2 = points[i + 1]
            if t1 <= temp_c <= t2:
                return _linear_interp((t1, r1), (t2, r2), temp_c)

        return _linear_interp(points[-2], points[-1], temp_c)

    if len(points) == 1:
        return points[0][1]

    if conductor.dc_res_20c_ohm_per_mile is not None:
        return conductor.dc_res_20c_ohm_per_mile

    raise ValueError(
        f"No usable resistance data found for conductor '{conductor.code_word}'. "
        f"Need AC resistance columns or DC resistance fallback."
    )


def _linear_interp(p1: tuple[float, float], p2: tuple[float, float], x: float) -> float:
    x1, y1 = p1
    x2, y2 = p2
    if x2 == x1:
        return y1
    return y1 + (y2 - y1) * (x - x1) / (x2 - x1)


def mean_film_temperature(ts_c: float, ta_c: float) -> float:
    return (ts_c + ta_c) / 2.0


def air_dynamic_viscosity_lb_per_ft_s(tfilm_c: float) -> float:
    # workbook-style US form
    return (9.806e-7 * (tfilm_c + 273.0) ** 1.5) / (tfilm_c + 383.4)


def air_density_lb_per_ft3(tfilm_c: float, elevation_ft: float) -> float:
    numerator = 0.080695 - 2.901e-6 * elevation_ft + 3.7e-11 * (elevation_ft ** 2)
    denominator = 1.0 + 0.00367 * tfilm_c
    return numerator / denominator


def air_thermal_conductivity_w_per_ft_c(tfilm_c: float) -> float:
    return 7.388e-3 + 2.279e-5 * tfilm_c - 1.343e-9 * (tfilm_c ** 2)


def wind_direction_factor_from_beta(beta_deg: float) -> float:
    beta = math.radians(beta_deg)
    return 1.194 - math.sin(beta) - 0.194 * math.cos(2.0 * beta) + 0.368 * math.sin(2.0 * beta)


def reynolds_number(diameter_ft: float, rho_f_lb_per_ft3: float, wind_fps: float, mu_f_lb_per_ft_s: float) -> float:
    return diameter_ft * rho_f_lb_per_ft3 * wind_fps / mu_f_lb_per_ft_s


def natural_convection_loss_w_per_ft(ts_c: float, ta_c: float, diameter_ft: float, rho_f_lb_per_ft3: float) -> float:
    delta_t = max(ts_c - ta_c, 0.0)
    if delta_t <= 0.0:
        return 0.0
    return 1.825 * (rho_f_lb_per_ft3 ** 0.5) * (diameter_ft ** 0.75) * (delta_t ** 1.25)


def forced_convection_losses_w_per_ft(
    ts_c: float,
    ta_c: float,
    diameter_ft: float,
    wind_fps: float,
    beta_deg: float,
    rho_f_lb_per_ft3: float,
    mu_f_lb_per_ft_s: float,
    k_f_w_per_ft_c: float,
) -> tuple[float, float]:
    delta_t = max(ts_c - ta_c, 0.0)
    if delta_t <= 0.0 or wind_fps <= 0.0:
        return 0.0, 0.0

    k_angle = wind_direction_factor_from_beta(beta_deg)
    n_re = reynolds_number(diameter_ft, rho_f_lb_per_ft3, wind_fps, mu_f_lb_per_ft_s)

    qc1 = k_angle * (1.01 + 1.35 * (n_re ** 0.52)) * k_f_w_per_ft_c * delta_t
    qc2 = k_angle * 0.754 * (n_re ** 0.60) * k_f_w_per_ft_c * delta_t
    return qc1, qc2


def convection_loss(
    ts_c: float,
    ta_c: float,
    diameter_ft: float,
    wind_speed_mps: float,
    wind_angle_deg: float,
    elevation_m: float,
) -> dict:
    tfilm = mean_film_temperature(ts_c, ta_c)
    wind_fps = wind_speed_mps * MPS_TO_FPS
    elevation_ft = elevation_m * FT_PER_M

    mu_f = air_dynamic_viscosity_lb_per_ft_s(tfilm)
    rho_f = air_density_lb_per_ft3(tfilm, elevation_ft)
    k_f = air_thermal_conductivity_w_per_ft_c(tfilm)

    # user input is angle to conductor axis; workbook uses beta angle to perpendicular
    beta_deg = 90.0 - wind_angle_deg
    qcn = natural_convection_loss_w_per_ft(ts_c, ta_c, diameter_ft, rho_f)
    qc1, qc2 = forced_convection_losses_w_per_ft(ts_c, ta_c, diameter_ft, wind_fps, beta_deg, rho_f, mu_f, k_f)
    qc = max(qcn, qc1, qc2)

    n_re = reynolds_number(diameter_ft, rho_f, wind_fps, mu_f) if wind_fps > 0 else 0.0

    return {
        "qc_w_per_ft": qc,
        "qcn_w_per_ft": qcn,
        "qc1_w_per_ft": qc1,
        "qc2_w_per_ft": qc2,
        "qc_w_per_m": qc * FT_PER_M,
        "qcn_w_per_m": qcn * FT_PER_M,
        "qc1_w_per_m": qc1 * FT_PER_M,
        "qc2_w_per_m": qc2 * FT_PER_M,
        "tfilm_c": tfilm,
        "rho_f_lb_per_ft3": rho_f,
        "mu_f_lb_per_ft_s": mu_f,
        "k_f_w_per_ft_c": k_f,
        "n_re": n_re,
        "k_angle": wind_direction_factor_from_beta(beta_deg),
        "beta_deg": beta_deg,
        "wind_fps": wind_fps,
        "elevation_ft": elevation_ft,
    }


def radiated_heat_loss(ts_c: float, ta_c: float, diameter_ft: float, emissivity: float) -> dict:
    qr_w_per_ft = 1.656 * diameter_ft * emissivity * ((((ts_c + 273.0) / 100.0) ** 4) - (((ta_c + 273.0) / 100.0) ** 4))
    return {
        "qr_w_per_ft": qr_w_per_ft,
        "qr_w_per_m": qr_w_per_ft * FT_PER_M,
    }


def calculate_steady_state_rating(
    conductor: Conductor,
    ambient_temp_c: float,
    wind_speed_mps: float,
    wind_angle_deg: float,
    elevation_m: float,
    target_temp_c: float,
    emissivity: Optional[float] = None,
    absorptivity: Optional[float] = None,
    latitude_deg: Optional[float] = None,
    line_azimuth_deg: Optional[float] = None,
    input_date=None,
    input_time=None,
    atmosphere_type: str = "clear",
) -> dict:
    if conductor.od_in is None:
        raise ValueError(f"Conductor '{conductor.code_word}' is missing OD_IN.")

    if latitude_deg is None or line_azimuth_deg is None or input_date is None or input_time is None:
        raise ValueError("Latitude, line azimuth, date, and time are required for the full IEEE 738 solar model.")

    eps = emissivity if emissivity is not None else (conductor.emissivity if conductor.emissivity is not None else 0.5)
    alpha = absorptivity if absorptivity is not None else (conductor.absorptivity if conductor.absorptivity is not None else 0.5)

    diameter_ft = inch_to_foot(conductor.od_in)
    diameter_m = diameter_ft * M_PER_FT

    resistance_ohm_per_mile = interpolate_resistance_ohm_per_mile(conductor, target_temp_c)
    resistance_ohm_per_ft = ohm_per_mile_to_ohm_per_ft(resistance_ohm_per_mile)
    resistance_ohm_per_m = resistance_ohm_per_ft * FT_PER_M

    convection = convection_loss(
        ts_c=target_temp_c,
        ta_c=ambient_temp_c,
        diameter_ft=diameter_ft,
        wind_speed_mps=wind_speed_mps,
        wind_angle_deg=wind_angle_deg,
        elevation_m=elevation_m,
    )

    radiation = radiated_heat_loss(
        ts_c=target_temp_c,
        ta_c=ambient_temp_c,
        diameter_ft=diameter_ft,
        emissivity=eps,
    )

    solar = solar_heat_gain(
        absorptivity=alpha,
        diameter_ft=diameter_ft,
        latitude_deg=latitude_deg,
        line_azimuth_deg=line_azimuth_deg,
        input_date=input_date,
        input_time=input_time,
        elevation_m=elevation_m,
        atmosphere_type=atmosphere_type,
    )

    net_ft = convection["qc_w_per_ft"] + radiation["qr_w_per_ft"] - solar["qs_w_per_ft"]
    amps = math.sqrt(net_ft / resistance_ohm_per_ft) if net_ft > 0.0 and resistance_ohm_per_ft > 0.0 else 0.0

    return {
        "code_word": conductor.code_word,
        "target_temp_c": target_temp_c,
        "ambient_temp_c": ambient_temp_c,
        "wind_speed_mps": wind_speed_mps,
        "wind_speed_fps": convection["wind_fps"],
        "wind_angle_deg": wind_angle_deg,
        "beta_deg": convection["beta_deg"],
        "elevation_m": elevation_m,
        "elevation_ft": convection["elevation_ft"],
        "diameter_m": diameter_m,
        "diameter_ft": diameter_ft,
        "diameter_in": conductor.od_in,
        "resistance_ohm_per_mile": resistance_ohm_per_mile,
        "resistance_ohm_per_ft": resistance_ohm_per_ft,
        "resistance_ohm_per_m": resistance_ohm_per_m,
        "qc_w_per_ft": convection["qc_w_per_ft"],
        "qcn_w_per_ft": convection["qcn_w_per_ft"],
        "qc1_w_per_ft": convection["qc1_w_per_ft"],
        "qc2_w_per_ft": convection["qc2_w_per_ft"],
        "qc_w_per_m": convection["qc_w_per_m"],
        "qcn_w_per_m": convection["qcn_w_per_m"],
        "qc1_w_per_m": convection["qc1_w_per_m"],
        "qc2_w_per_m": convection["qc2_w_per_m"],
        "qr_w_per_ft": radiation["qr_w_per_ft"],
        "qr_w_per_m": radiation["qr_w_per_m"],
        "qs_w_per_ft": solar["qs_w_per_ft"],
        "qs_w_per_m": solar["qs_w_per_m"],
        "amps": amps,
        "tfilm_c": convection["tfilm_c"],
        "rho_f_lb_per_ft3": convection["rho_f_lb_per_ft3"],
        "mu_f_lb_per_ft_s": convection["mu_f_lb_per_ft_s"],
        "k_f_w_per_ft_c": convection["k_f_w_per_ft_c"],
        "n_re": convection["n_re"],
        "k_angle": convection["k_angle"],
        "emissivity": eps,
        "absorptivity": alpha,
        "solar": solar,
    }