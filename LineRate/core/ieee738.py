from __future__ import annotations

import math
from typing import Optional

from models.conductor import Conductor


INCH_TO_M = 0.0254
MILE_TO_M = 1609.344


def inch_to_meter(value_in: float) -> float:
    return value_in * INCH_TO_M


def ohm_per_mile_to_ohm_per_meter(value: float) -> float:
    return value / MILE_TO_M


def interpolate_resistance_ohm_per_mile(conductor: Conductor, temp_c: float) -> float:
    """
    Interpolate conductor AC resistance using available points in the conductor table.
    Preference order:
    1) AC 25C / 50C / 75C with piecewise linear interpolation (and edge extrapolation)
    2) AC 25C / 75C
    3) DC 20C fallback (constant)
    """
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


def air_dynamic_viscosity(tfilm_c: float) -> float:
    # IEEE 738 Eq. (13), SI
    return (1.458e-6 * (tfilm_c + 273.0) ** 1.5) / (tfilm_c + 383.4)


def air_density(tfilm_c: float, elevation_m: float) -> float:
    # IEEE 738 Eq. (15), SI
    numerator = 1.293 - 1.525e-4 * elevation_m + 6.379e-9 * (elevation_m ** 2)
    denominator = 1.0 + 0.00367 * tfilm_c
    return numerator / denominator


def air_thermal_conductivity(tfilm_c: float) -> float:
    # IEEE 738 Eq. (17), SI
    return 2.424e-2 + 7.477e-5 * tfilm_c - 4.407e-9 * (tfilm_c ** 2)


def wind_direction_factor(phi_deg: float) -> float:
    # IEEE 738 Eq. (19), using phi = angle between wind and conductor axis
    phi = math.radians(phi_deg)
    return 1.194 - math.cos(phi) + 0.194 * math.cos(2.0 * phi) + 0.368 * math.sin(2.0 * phi)


def reynolds_number(diameter_m: float, rho_f: float, wind_mps: float, mu_f: float) -> float:
    # IEEE 738 Eq. (11)
    return diameter_m * rho_f * wind_mps / mu_f


def natural_convection_loss(ts_c: float, ta_c: float, diameter_m: float, rho_f: float) -> float:
    # IEEE 738 Eq. (7), SI
    delta_t = max(ts_c - ta_c, 0.0)
    if delta_t <= 0.0:
        return 0.0
    return 3.645 * (rho_f ** 0.5) * (diameter_m ** 0.75) * (delta_t ** 1.25)


def forced_convection_losses(
    ts_c: float,
    ta_c: float,
    diameter_m: float,
    wind_mps: float,
    phi_deg: float,
    rho_f: float,
    mu_f: float,
    k_f: float,
) -> tuple[float, float]:
    # IEEE 738 Eq. (9) and Eq. (10), SI
    delta_t = max(ts_c - ta_c, 0.0)
    if delta_t <= 0.0 or wind_mps <= 0.0:
        return 0.0, 0.0

    k_angle = wind_direction_factor(phi_deg)
    n_re = reynolds_number(diameter_m, rho_f, wind_mps, mu_f)

    qc1 = k_angle * (1.01 + 1.35 * (n_re ** 0.52)) * k_f * delta_t
    qc2 = k_angle * 0.754 * (n_re ** 0.60) * k_f * delta_t
    return qc1, qc2


def convection_loss(ts_c: float, ta_c: float, diameter_m: float, wind_mps: float, phi_deg: float, elevation_m: float) -> dict:
    tfilm = mean_film_temperature(ts_c, ta_c)
    mu_f = air_dynamic_viscosity(tfilm)
    rho_f = air_density(tfilm, elevation_m)
    k_f = air_thermal_conductivity(tfilm)

    qcn = natural_convection_loss(ts_c, ta_c, diameter_m, rho_f)
    qc1, qc2 = forced_convection_losses(ts_c, ta_c, diameter_m, wind_mps, phi_deg, rho_f, mu_f, k_f)
    qc = max(qcn, qc1, qc2)

    return {
        "qc": qc,
        "qcn": qcn,
        "qc1": qc1,
        "qc2": qc2,
        "tfilm_c": tfilm,
        "rho_f": rho_f,
        "mu_f": mu_f,
        "k_f": k_f,
        "n_re": reynolds_number(diameter_m, rho_f, wind_mps, mu_f) if wind_mps > 0 else 0.0,
        "k_angle": wind_direction_factor(phi_deg),
    }


def radiated_heat_loss(ts_c: float, ta_c: float, diameter_m: float, emissivity: float) -> float:
    # IEEE 738 Eq. (21), SI
    return 17.8 * diameter_m * emissivity * ((((ts_c + 273.0) / 100.0) ** 4) - (((ta_c + 273.0) / 100.0) ** 4))


def solar_heat_gain_simple(solar_w_per_m2: float, absorptivity: float, diameter_m: float) -> float:
    """
    Simplified solar gain using user-input irradiance:
        qs = alpha * G * A'
    where projected area per unit length A' = D0.
    This uses the user's entered solar irradiance directly rather than the full
    IEEE 738 sun-position model.
    """
    return max(solar_w_per_m2, 0.0) * max(absorptivity, 0.0) * diameter_m


def calculate_steady_state_rating(
    conductor: Conductor,
    ambient_temp_c: float,
    wind_speed_mps: float,
    wind_angle_deg: float,
    elevation_m: float,
    solar_w_per_m2: float,
    target_temp_c: float,
    emissivity: Optional[float] = None,
    absorptivity: Optional[float] = None,
) -> dict:
    if conductor.od_in is None:
        raise ValueError(f"Conductor '{conductor.code_word}' is missing OD_IN.")

    eps = emissivity if emissivity is not None else (conductor.emissivity if conductor.emissivity is not None else 0.5)
    alpha = absorptivity if absorptivity is not None else (conductor.absorptivity if conductor.absorptivity is not None else 0.5)

    diameter_m = inch_to_meter(conductor.od_in)
    resistance_ohm_per_mile = interpolate_resistance_ohm_per_mile(conductor, target_temp_c)
    resistance_ohm_per_m = ohm_per_mile_to_ohm_per_meter(resistance_ohm_per_mile)

    convection = convection_loss(
        ts_c=target_temp_c,
        ta_c=ambient_temp_c,
        diameter_m=diameter_m,
        wind_mps=wind_speed_mps,
        phi_deg=wind_angle_deg,
        elevation_m=elevation_m,
    )

    qr = radiated_heat_loss(
        ts_c=target_temp_c,
        ta_c=ambient_temp_c,
        diameter_m=diameter_m,
        emissivity=eps,
    )

    qs = solar_heat_gain_simple(
        solar_w_per_m2=solar_w_per_m2,
        absorptivity=alpha,
        diameter_m=diameter_m,
    )

    net = convection["qc"] + qr - qs
    amps = math.sqrt(net / resistance_ohm_per_m) if net > 0.0 and resistance_ohm_per_m > 0.0 else 0.0

    return {
        "code_word": conductor.code_word,
        "target_temp_c": target_temp_c,
        "ambient_temp_c": ambient_temp_c,
        "wind_speed_mps": wind_speed_mps,
        "wind_angle_deg": wind_angle_deg,
        "elevation_m": elevation_m,
        "solar_w_per_m2": solar_w_per_m2,
        "diameter_m": diameter_m,
        "diameter_in": conductor.od_in,
        "resistance_ohm_per_mile": resistance_ohm_per_mile,
        "resistance_ohm_per_m": resistance_ohm_per_m,
        "qc_w_per_m": convection["qc"],
        "qcn_w_per_m": convection["qcn"],
        "qc1_w_per_m": convection["qc1"],
        "qc2_w_per_m": convection["qc2"],
        "qr_w_per_m": qr,
        "qs_w_per_m": qs,
        "amps": amps,
        "tfilm_c": convection["tfilm_c"],
        "rho_f": convection["rho_f"],
        "mu_f": convection["mu_f"],
        "k_f": convection["k_f"],
        "n_re": convection["n_re"],
        "k_angle": convection["k_angle"],
        "emissivity": eps,
        "absorptivity": alpha,
    }