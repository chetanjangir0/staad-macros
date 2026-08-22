"""Pure-Python port of the validated 'IS_Mid Frame_2015wind.xlsx' workbook
(`src/staad_ext/wind_profiles/IS_Mid Frame_2015wind.xlsx`, sheets
'WL-Single Gable-MBS', 'Sheet1', 'For STAAD') that computes IS 875 Part 3
wind loads for a single-span gable frame.

Every function here is a direct, faithful translation of one named cell (or
small formula chain) in that workbook -- not a reinterpretation. Each
docstring cites the Excel cell(s) it replicates. Table constants are
transcribed verbatim from a full formula+value dump of the workbook
(see scripts/generate_wind_golden_fixtures.py for how to regenerate the
golden-master fixtures this module is verified against).

This module deliberately preserves several workbook quirks rather than
"fixing" them, because the workbook is the authoritative, validated source:

  * K1 (risk coefficient), K3 (topography factor) and K4 (cyclonic region
    factor) are hard-coded to 1 in the workbook (cells F39, F147, F148 are
    plain VALUEs, not formulas referencing the Table-1 risk data or the
    `design_life` input). So `design_life` and `basic_wind_speed`'s Table-1
    A/B lookup have NO EFFECT on the final wind pressure in this workbook
    version. We replicate that exactly rather than "helpfully" wiring up
    the textbook IS 875 risk-coefficient formula.
  * F8 ("h2-Eaves height from FFL", used in the Pd 0.8x-below-10m rule and
    in the wall-area Ka lookup) is treated as equal to F10 ("h-Eaves height
    from FGL", i.e. `height`). This mirrors F10's own formula (`=F7+F8`)
    with the plinth height F7 pinned at 0 -- there is no separate "plinth
    height" input in this tool's UI, so the only self-consistent reading of
    F8 is that it equals `height` too. (An earlier revision of this port
    hard-coded F8 to the workbook's shipped demo value of 28.5 m, matching
    a 9-cell Excel COM driver that never wrote F7/F8 at all; that was a
    genuine bug, not a faithful quirk, because it silently disconnected
    `height` from the Ka wall-area lookup and the K152 "<10m" branch for
    every real building shorter than 10 m -- i.e. almost all of them. Fixed
    here, and the golden fixtures were regenerated with a corrected driver
    that writes F7=0 and F8=height.)
  * Sheet1's "parallel to ridge" roof force (B97/B98) and its 90 degree
    suction counterpart (B106/B107) both reference the SAME near-roof-plane
    cell (B50, i.e. WL sheet F264) for what are labelled "Near Roof Plane"
    and "Back Roof Plane" -- the far roof plane (G264) is computed but
    never actually used in the final STAAD output. WL5/WL6 (gable-1/gable-2
    parallel pressure) and WL7/WL8 (parallel suction) are therefore
    literal duplicates of each other. We replicate this exactly.
  * The near-roof-plane coefficient for the 0 degree (perpendicular to
    ridge) case is passed through Excel's ROUNDUP(...,3) before use
    (Sheet1!B46), while every other Cpe value is used at full precision.
"""
from __future__ import annotations

import math
from dataclasses import dataclass

# ---------------------------------------------------------------------------
# Constants transcribed from 'WL-Single Gable-MBS'
# ---------------------------------------------------------------------------

#: Directional factor Kd -- WL-Single Gable-MBS!J146 (fixed value, not an
#: input).
_KD_DIRECTIONAL_FACTOR = 0.9

#: Combination factor Kc -- WL-Single Gable-MBS!J148 (fixed value).
_KC_COMBINATION_FACTOR = 0.9

#: K3 - Topography factor -- WL-Single Gable-MBS!F147 (plain VALUE, not a
#: formula; always 1 in this workbook).
_K3_TOPOGRAPHY_FACTOR = 1.0

#: K4 - Cyclonic region factor -- WL-Single Gable-MBS!F148 (plain VALUE,
#: always 1 in this workbook).
_K4_CYCLONIC_FACTOR = 1.0

#: K1 - Risk coefficient -- WL-Single Gable-MBS!F39 (plain VALUE = 1, NOT a
#: formula driven by Table 1 / PN / design life; see module docstring).
_K1_RISK_COEFFICIENT = 1.0

#: "h1-Height of plinth" -- WL-Single Gable-MBS!F7. Pinned at 0: this
#: tool's UI has no separate plinth-height input, so F10 (=F7+F8, eaves
#: height from FGL, which we set directly from the user's `height`) is
#: only self-consistent with F7=0 and F8=height. See module docstring.
_PLINTH_HEIGHT_M = 0.0

#: Table 1, IS 875 (Part 3) 1987 -- WL-Single Gable-MBS!E29:G34. Not used
#: in the final calculation chain (K1 is hard-coded, see above) but kept
#: here, transcribed verbatim, since it is a genuine cell in the workbook
#: and a future revision of the sheet may wire it back up.
TABLE1_RISK_COEFFICIENT_AB = {
    33: (83.2, 9.2),
    39: (84.0, 14.0),
    44: (88.0, 18.0),
    47: (88.0, 20.5),
    50: (88.8, 22.8),
    55: (90.8, 27.3),
}
#: PN - probability factor -- WL-Single Gable-MBS!F35.
TABLE1_PN_PROBABILITY_FACTOR = 0.63

#: Table 2, IS 875 (Part 3) 1987 - Terrain, height & structure size factor
#: k2, category A (largest structures class). WL-Single Gable-MBS!
#: E45:F65 (category 1), E69:F89 (category 2), E93:F113 (category 3),
#: E117:F137 (category 4). Values are the workbook's own *cached* results,
#: i.e. already fully evaluated (some cells are raw table values, some are
#: formulas interpolating between the surrounding raw values) -- so this is
#: already a clean height -> k2 breakpoint table at 5 m spacing.
TABLE2_TERRAIN_HEIGHT_SIZE_FACTOR = {
    1: [
        (0, 1.05), (5, 1.05), (10, 1.05), (15, 1.09), (20, 1.12),
        (25, 1.135), (30, 1.15), (35, 1.1624999999999999),
        (40, 1.1749999999999998), (45, 1.1875), (50, 1.2), (55, 1.206),
        (60, 1.212), (65, 1.218), (70, 1.224), (75, 1.23), (80, 1.236),
        (85, 1.242), (90, 1.248), (95, 1.254), (100, 1.26),
    ],
    2: [
        (0, 1.0), (5, 1.0), (10, 1.0), (15, 1.05), (20, 1.07),
        (25, 1.0950000000000002), (30, 1.12), (35, 1.1325), (40, 1.145),
        (45, 1.1575), (50, 1.17), (55, 1.177), (60, 1.184),
        (65, 1.1909999999999998), (70, 1.198), (75, 1.205), (80, 1.212),
        (85, 1.2189999999999999), (90, 1.226), (95, 1.233), (100, 1.24),
    ],
    3: [
        (0, 0.91), (5, 0.91), (10, 0.91), (15, 0.97), (20, 1.01),
        (25, 1.0350000000000001), (30, 1.06), (35, 1.0750000000000002),
        (40, 1.09), (45, 1.105), (50, 1.12), (55, 1.1280000000000001),
        (60, 1.1360000000000001), (65, 1.1440000000000001),
        (70, 1.1520000000000001), (75, 1.1600000000000001), (80, 1.168),
        (85, 1.176), (90, 1.184), (95, 1.192), (100, 1.2),
    ],
    4: [
        (0, 0.8), (5, 0.8), (10, 0.8), (15, 0.8), (20, 0.8), (25, 0.885),
        (30, 0.97), (35, 1.0025), (40, 1.0350000000000001),
        (45, 1.0675000000000001), (50, 1.1), (55, 1.11), (60, 1.12),
        (65, 1.1300000000000001), (70, 1.1400000000000001), (75, 1.15),
        (80, 1.16), (85, 1.17), (90, 1.18), (95, 1.19), (100, 1.2),
    ],
}

#: Area-averaging reduction factor Ka breakpoints -- WL-Single Gable-MBS!
#: P43:Q150 (roof area) / S43:T150 (wall area). Both tables share the same
#: breakpoints: <=10 sqm -> 1.0, 25 sqm -> 0.9, >=100 sqm -> 0.8, linear
#: interpolation in between.
_KA_AREA_BREAKPOINTS = [(10.0, 1.0), (25.0, 0.9), (100.0, 0.8)]

#: Table No. 4, IS 875 (Part 3) 1987 - external pressure coefficients Cpe
#: for walls of a rectangular clad building. WL-Single Gable-MBS!D164:H190.
#: Keyed by (h/w band, l/w band) -> {"0": (A,B,C,D,Local), "90": (...)}.
#: h/w bands: <=0.5, (0.5,1.5], (1.5,6], >6. l/w bands vary by h/w band
#: (see wall_pressure_coefficients() for the exact nested-IF thresholds,
#: transcribed from WL-Single Gable-MBS!D193:H194).
TABLE4_WALL_CPE = {
    ("h<=0.5", "l<=1.5"): {"0": (0.7, -0.2, -0.5, -0.5, -0.8), "90": (-0.5, -0.5, 0.7, -0.2, -0.8)},
    ("h<=0.5", "l>1.5"): {"0": (0.7, -0.25, -0.6, -0.6, -1.0), "90": (-0.5, -0.5, 0.7, -0.1, -1.0)},
    ("0.5<h<=1.5", "l<=1.5"): {"0": (0.7, -0.25, -0.6, -0.6, -1.1), "90": (-0.6, -0.6, 0.7, -0.25, -1.1)},
    ("0.5<h<=1.5", "l>1.5"): {"0": (0.7, -0.3, -0.7, -0.7, -1.1), "90": (-0.5, -0.5, 0.7, -0.1, -1.1)},
    ("1.5<h<=6", "l<=1.5"): {"0": (0.8, -0.25, -0.8, -0.8, -1.2), "90": (-0.8, -0.8, 0.8, -0.25, -1.2)},
    ("1.5<h<=6", "l>1.5"): {"0": (0.7, -0.4, -0.7, -0.7, -1.2), "90": (-0.5, -0.5, 0.8, -0.1, -1.2)},
    ("h>6", "l<=1"): {"0": (0.95, -1.25, -0.7, -0.7, -1.25), "90": (-0.7, -0.7, 0.95, -1.25, -1.25)},
    ("h>6", "1<l<=1.5"): {"0": (0.95, -1.85, -0.9, -0.9, -1.25), "90": (-0.8, -0.8, 0.9, -0.85, -1.25)},
    ("h>6", "l>1.5"): {"0": (0.85, -0.75, -0.75, -0.75, -1.25), "90": (-0.75, -0.75, 0.85, -0.75, -1.25)},
}

#: Table No. 5, IS 875 (Part 3) 1987 - external pressure coefficients Cpe
#: for a pitched roof of a single-span building. WL-Single Gable-MBS!
#: A218:E227 (h/w<=0.5), A232:E241 (0.5<h/w<=1.5), A246:E255 (1.5<h/w<6).
#: Each row: roof_angle_deg -> (EF, GH, EG, FH), i.e. (windward slope @
#: wind 0 deg, leeward slope @ wind 0 deg, near half @ wind 90 deg, far
#: half @ wind 90 deg). Values are the workbook's cached results (already
#: interpolated where the source cell was a formula).
TABLE5_ROOF_CPE = {
    "h<=0.5": [
        (0, (-0.8, -0.4, -0.8, -0.4)),
        (5, (-0.9, -0.4, -0.8, -0.4)),
        (10, (-1.2, -0.4, -0.8, -0.6)),
        (15, (-0.8, -0.4, -0.75, -0.6)),
        (20, (-0.4, -0.4, -0.7, -0.6)),
        (25, (-0.2, -0.4, -0.7, -0.6)),
        (30, (0.0, -0.4, -0.7, -0.6)),
        (35, (0.1, -0.43333333333333335, -0.7, -0.6)),
        (40, (0.2, -0.4666666666666667, -0.7, -0.6)),
        (45, (0.3, -0.5, -0.7, -0.6)),
    ],
    "0.5<h<=1.5": [
        (0, (-0.8, -0.6, -1.0, -0.6)),
        (5, (-0.9, -0.6, -0.9, -0.6)),
        (10, (-1.1, -0.6, -0.8, -0.6)),
        (15, (-0.9, -0.55, -0.8, -0.6)),
        (20, (-0.7, -0.5, -0.8, -0.6)),
        (25, (-0.45, -0.5, -0.8, -0.7000000000000001)),
        (30, (-0.2, -0.5, -0.8, -0.8)),
        (35, (-0.06666666666666668, -0.5, -0.8, -0.8)),
        (40, (0.06666666666666665, -0.5, -0.8, -0.8)),
        (45, (0.2, -0.5, -0.8, -0.8)),
    ],
    "1.5<h<6": [
        (0, (-0.7, -0.6, -0.9, -0.7)),
        (5, (-0.7, -0.6, -0.8, -0.8)),
        (10, (-0.7, -0.6, -0.8, -0.8)),
        (15, (-0.75, -0.6, -0.8, -0.8)),
        (20, (-0.8, -0.6, -0.8, -0.8)),
        (25, (-0.9, -0.55, -0.8, -0.75)),
        (30, (-1.0, -0.5, -0.8, -0.7)),
        (35, (-0.6, -0.5, -0.8, -0.7)),
        (40, (-0.2, -0.5, -0.8, -0.7)),
        (45, (0.0, -0.5, -0.8, -0.7)),
    ],
}

# IS 875 Part 3 dropdown options exposed by the UI (mirrors wind_load.py).
BASIC_WIND_SPEEDS = (33.0, 39.0, 44.0, 47.0, 50.0, 55.0)
DESIGN_LIVES = (5, 25, 50, 100)
TERRAIN_CATEGORIES = (1, 2, 3, 4)
OPENING_OPTIONS = ("<5%", "5-20%", ">20%")

_CPI_BY_OPENING = {"<5%": 0.2, "5-20%": 0.5, ">20%": 0.7}


class WindLoadError(RuntimeError):
    """Raised when the wind load parameters cannot be resolved to a result."""


@dataclass
class Is875WindParameters:
    """User-facing inputs for the IS 875 Part 3 'WL-Single Gable-MBS' workbook."""

    building_length: float  # m (F5)
    width: float  # m (F6)
    height: float  # m, from FGL (F10, overwrites the F7+F8 formula)
    roof_slope_x: float  # x in 1:x (F12)
    basic_wind_speed: float  # m/s, one of BASIC_WIND_SPEEDS (F13)
    design_life: int  # years, one of DESIGN_LIVES (F14) -- see module
    # docstring: has NO EFFECT on the result in this workbook version.
    terrain_category: int  # 1-4 (F15)
    bay_spacing: float  # m (I7)
    opening: str  # one of OPENING_OPTIONS

    def cpi(self) -> float:
        try:
            return _CPI_BY_OPENING[self.opening]
        except KeyError as exc:
            raise ValueError(f"Unknown opening percentage option: {self.opening!r}") from exc


# ---------------------------------------------------------------------------
# Small numeric helpers
# ---------------------------------------------------------------------------


def _vlookup_approx(table: list[tuple[float, float]], key: float) -> float:
    """Excel VLOOKUP(key, table, 2, TRUE) semantics: return the value for
    the largest breakpoint <= key; if key is below the first breakpoint,
    return the first row's value (matches this workbook's tables, which
    all start at a breakpoint <= any realistic input).
    """
    best = table[0][1]
    for breakpoint, value in table:
        if breakpoint <= key:
            best = value
        else:
            break
    return best


def _interp(x: float, x1: float, y1: float, x2: float, y2: float) -> float:
    """Linear interpolation, matching the workbook's repeated
    `=Y1+(Y2-Y1)*(X-X1)/(X2-X1)` formula pattern. Returns y1 if x2==x1
    (matching the sheet's own div/0 guard, e.g. F146's outer IF).
    """
    if x2 == x1:
        return y1
    return y1 + (y2 - y1) * (x - x1) / (x2 - x1)


def _excel_roundup(value: float, digits: int) -> float:
    """Excel ROUNDUP(value, digits): rounds away from zero (used by
    Sheet1!B46 on the near-roof-plane 0 deg Cpe).
    """
    factor = 10 ** digits
    if value >= 0:
        return math.ceil(value * factor) / factor
    return math.floor(value * factor) / factor


# ---------------------------------------------------------------------------
# Table 1 / K1 - risk coefficient (WL-Single Gable-MBS!F36,F37,F39)
# ---------------------------------------------------------------------------


def table1_risk_coefficient_ab(basic_wind_speed: float) -> tuple[float, float]:
    """A, B lookup from Table 1 -- WL-Single Gable-MBS!F36
    (`=VLOOKUP(F13,E29:G34,2)`) and F37 (`=VLOOKUP(F13,E29:G34,3)`).

    NOTE: these values are computed by the workbook but not actually used
    anywhere in the final calculation chain -- see risk_coefficient_k1().
    Kept here for completeness/debugging parity with the golden fixtures.
    """
    key = int(round(basic_wind_speed))
    if key not in TABLE1_RISK_COEFFICIENT_AB:
        raise WindLoadError(f"Unsupported basic wind speed for Table 1 lookup: {basic_wind_speed}")
    return TABLE1_RISK_COEFFICIENT_AB[key]


def risk_coefficient_k1(design_life: int) -> float:
    """K1 - Risk coefficient -- WL-Single Gable-MBS!F39.

    F39 is a plain VALUE of 1 in the workbook, NOT a formula driven by
    Table 1 / PN (F35) / design life (F14). `design_life` is accepted here
    only for interface symmetry with the other risk-coefficient inputs; it
    has no effect, faithfully matching the workbook. See module docstring.
    """
    del design_life  # unused, by design -- matches the workbook exactly
    return _K1_RISK_COEFFICIENT


# ---------------------------------------------------------------------------
# Table 2 / K2 - terrain, height & structure size factor
# (WL-Single Gable-MBS!F139:F146)
# ---------------------------------------------------------------------------


def terrain_height_size_factor_k2(terrain_category: int, height_m: float) -> float:
    """K2 -- WL-Single Gable-MBS!F146:
    `=IF((F141-F140)>0,F143+(F144-F143)*(F139-F140)/(F141-F140),F143)`
    where F139=HT=height, F140=HT1=INT(HT/5)*5, F141=HT2=CEILING(HT,5),
    and F143/F144 are VLOOKUP(HT1/HT2, <category table>, 2).
    """
    if terrain_category not in TABLE2_TERRAIN_HEIGHT_SIZE_FACTOR:
        raise WindLoadError(f"Unsupported terrain category: {terrain_category}")
    table = TABLE2_TERRAIN_HEIGHT_SIZE_FACTOR[terrain_category]

    ht = height_m
    ht1 = math.floor(ht / 5.0) * 5.0  # F140 = INT(F139/5)*5
    ht2 = math.ceil(ht / 5.0) * 5.0  # F141 = CEILING(F139,5)

    k21 = _vlookup_approx(table, ht1)  # F143
    k22 = _vlookup_approx(table, ht2)  # F144

    if (ht2 - ht1) > 0:
        return _interp(ht, ht1, k21, ht2, k22)
    return k21


# ---------------------------------------------------------------------------
# K3, K4 (fixed) and Vz, Pz
# ---------------------------------------------------------------------------


def topography_factor_k3() -> float:
    """K3 - Topography factor -- WL-Single Gable-MBS!F147 (fixed at 1)."""
    return _K3_TOPOGRAPHY_FACTOR


def cyclonic_region_factor_k4() -> float:
    """K4 - Cyclonic region factor -- WL-Single Gable-MBS!F148 (fixed at 1)."""
    return _K4_CYCLONIC_FACTOR


def design_wind_speed_vz(basic_wind_speed: float, terrain_category: int, height_m: float, design_life: int) -> float:
    """Vz - Design wind speed (m/s) -- WL-Single Gable-MBS!F150:
    `=F39*F146*F147*F13*F148` = K1*K2*K3*Vb*K4.
    """
    k1 = risk_coefficient_k1(design_life)
    k2 = terrain_height_size_factor_k2(terrain_category, height_m)
    k3 = topography_factor_k3()
    k4 = cyclonic_region_factor_k4()
    return k1 * k2 * k3 * basic_wind_speed * k4


def design_wind_pressure_pz(vz: float) -> float:
    """Pz - Wind pressure (kN/m^2) -- WL-Single Gable-MBS!F151: `=0.6*F150^2/1000`."""
    return 0.6 * vz**2 / 1000.0


# ---------------------------------------------------------------------------
# Area-averaging factor Ka (WL-Single Gable-MBS!N146,N147,J147)
# ---------------------------------------------------------------------------


def area_reduction_factor_ka(area_sqm: float) -> float:
    """Ka - Area-averaging reduction factor for a single tributary area,
    from the P43:Q150 / S43:T150 breakpoint tables (both share the same
    breakpoints: <=10 sqm -> 1.0, 25 sqm -> 0.9, >=100 sqm -> 0.8, with
    linear interpolation in between).
    """
    bp = _KA_AREA_BREAKPOINTS
    if area_sqm <= bp[0][0]:
        return bp[0][1]
    if area_sqm >= bp[-1][0]:
        return bp[-1][1]
    for (x1, y1), (x2, y2) in zip(bp, bp[1:]):
        if x1 <= area_sqm <= x2:
            return _interp(area_sqm, x1, y1, x2, y2)
    return bp[-1][1]  # unreachable, kept for safety


def eaves_height_ffl(height_m: float) -> float:
    """F8 ("h2-Eaves height from FFL") -- WL-Single Gable-MBS!F8, taken as
    F10 - F7 = height - plinth_height = height - 0 (see module docstring:
    this tool has no separate plinth-height input, so F7 is pinned at 0).
    """
    return height_m - _PLINTH_HEIGHT_M


def combined_area_factor_ka(width_m: float, bay_spacing_m: float, height_m: float) -> float:
    """Ka (combined) -- WL-Single Gable-MBS!J147: `=MAX(N146,N147)`, the
    larger of the roof-area Ka (M146 = I7*F6 = bay_spacing*width) and the
    wall-area Ka (M147 = I7*F8 = bay_spacing * eaves_height_ffl(height)).
    """
    roof_area = bay_spacing_m * width_m  # M146
    wall_area = bay_spacing_m * eaves_height_ffl(height_m)  # M147
    ka_roof = area_reduction_factor_ka(roof_area)  # N146
    ka_wall = area_reduction_factor_ka(wall_area)  # N147
    return max(ka_roof, ka_wall)  # J147


# ---------------------------------------------------------------------------
# Design wind pressure Pd (WL-Single Gable-MBS!K152,F152)
# ---------------------------------------------------------------------------


def design_wind_pressure_pd(pz: float, ka: float, height_m: float) -> float:
    """Pd - Design wind pressure (kN/m^2) -- WL-Single Gable-MBS!F152 and
    K152:
      K152 = IF(F8<10, 0.8*Pz*Kd*Ka*Kc, Pz*Kd*Ka*Kc)
      F152 = IF(K152 < 0.7*Pz, 0.7*Pz, K152)   -- the "0.7 x Pz floor"

    `height_m` feeds F8 via eaves_height_ffl() (see module docstring).
    """
    kd = _KD_DIRECTIONAL_FACTOR
    kc = _KC_COMBINATION_FACTOR
    eaves_height_ffl_m = eaves_height_ffl(height_m)
    if eaves_height_ffl_m < 10:
        k152 = 0.8 * pz * kd * ka * kc
    else:
        k152 = pz * kd * ka * kc

    floor = 0.7 * pz
    return floor if k152 < floor else k152


# ---------------------------------------------------------------------------
# Table 4 - wall Cpe (WL-Single Gable-MBS!D193:H194)
# ---------------------------------------------------------------------------


def wall_pressure_coefficients(h_over_w: float, l_over_w: float) -> dict[str, dict[str, float]]:
    """Cpe for walls A/B/C/D + Local, at wind angles 0 deg (perpendicular to
    ridge) and 90 deg (parallel to ridge) -- Table No. 4, IS 875 (Part 3)
    1987. WL-Single Gable-MBS!D193:H194, replicating the nested-IF band
    selection exactly:
        h/w<=0.5: l/w<=1.5 else l/w>1.5
        0.5<h/w<=1.5: l/w<=1.5 else l/w>1.5
        1.5<h/w<=6: l/w<=1.5 else l/w>1.5
        h/w>6: l/w<=1, 1<l/w<=1.5, else l/w>1.5

    Returns {"0": {"A":.., "B":.., "C":.., "D":.., "Local":..}, "90": {...}}.
    """
    if h_over_w <= 0.5:
        band = "h<=0.5"
        sub = "l<=1.5" if l_over_w <= 1.5 else "l>1.5"
    elif h_over_w <= 1.5:
        band = "0.5<h<=1.5"
        sub = "l<=1.5" if l_over_w <= 1.5 else "l>1.5"
    elif h_over_w <= 6:
        band = "1.5<h<=6"
        sub = "l<=1.5" if l_over_w <= 1.5 else "l>1.5"
    else:
        band = "h>6"
        if l_over_w <= 1:
            sub = "l<=1"
        elif l_over_w <= 1.5:
            sub = "1<l<=1.5"
        else:
            sub = "l>1.5"

    raw = TABLE4_WALL_CPE[(band, sub)]
    result = {}
    for angle, values in raw.items():
        a, b, c, d, local = values
        result[angle] = {"A": a, "B": b, "C": c, "D": d, "Local": local}
    return result


# ---------------------------------------------------------------------------
# Table 5 - roof Cpe (WL-Single Gable-MBS!F198, C262:G264)
# ---------------------------------------------------------------------------


def roof_slope_angle_degrees(roof_slope_x: float) -> float:
    """Roof angle in degrees -- WL-Single Gable-MBS!F198:
    `=ATAN(E12/F12)*180/PI()`, where E12=1 (fixed "rise" of the 1:x slope
    notation) and F12=roof_slope_x is the "x" in 1:x.
    """
    return math.degrees(math.atan(1.0 / roof_slope_x))


def roof_pressure_coefficients(h_over_w: float, roof_slope_x: float) -> dict[str, float]:
    """Cpe for roof zones EF/GH (wind 0 deg, i.e. perpendicular to ridge)
    and EG/FH (wind 90 deg, parallel to ridge) -- Table No. 5, IS 875
    (Part 3) 1987. WL-Single Gable-MBS!C262:G264: look up the table row for
    the roof angle rounded down and up to the nearest 5 degrees, then
    linearly interpolate at the actual roof angle.

    Raises WindLoadError for h/w >= 6, matching the workbook's own
    "Out of Range" VLOOKUP branch (WL-Single Gable-MBS!D262 etc.).
    """
    if h_over_w <= 0.5:
        table = TABLE5_ROOF_CPE["h<=0.5"]
    elif h_over_w <= 1.5:
        table = TABLE5_ROOF_CPE["0.5<h<=1.5"]
    elif h_over_w < 6:
        table = TABLE5_ROOF_CPE["1.5<h<6"]
    else:
        raise WindLoadError(
            f"h/w = {h_over_w:.3f} is out of range for the Table 5 roof Cpe lookup "
            "(workbook returns 'Out of Range' for h/w >= 6)."
        )

    angle = roof_slope_angle_degrees(roof_slope_x)  # F198
    angle1 = math.floor(angle / 5.0) * 5.0  # F257 = INT(F198/5)*5
    angle2 = math.ceil(angle / 5.0) * 5.0  # F258 = CEILING(F198,5)

    angles = [row[0] for row in table]
    max_angle = angles[-1]
    angle1 = min(angle1, max_angle)
    angle2 = min(angle2, max_angle)

    row1 = dict(zip(("EF", "GH", "EG", "FH"), _vlookup_row(table, angle1)))
    row2 = dict(zip(("EF", "GH", "EG", "FH"), _vlookup_row(table, angle2)))

    result = {}
    for key in ("EF", "GH", "EG", "FH"):
        if angle2 - angle1 > 0:
            result[key] = _interp(angle, angle1, row1[key], angle2, row2[key])
        else:
            result[key] = row1[key]
    return result


def _vlookup_row(table: list[tuple[float, tuple]], key: float) -> tuple:
    """VLOOKUP(key, table, col, TRUE) row-tuple variant for the Table 5
    (angle -> (EF,GH,EG,FH)) tables.
    """
    best = table[0][1]
    for breakpoint, values in table:
        if breakpoint <= key:
            best = values
        else:
            break
    return best


# ---------------------------------------------------------------------------
# Sheet1 - per-zone pressure buildup
# ---------------------------------------------------------------------------


@dataclass
class ZoneForces:
    """Per-metre-run zone forces (kN/m) feeding the final STAAD output,
    replicating Sheet1!B59:B107 (only the rows actually referenced by
    'For STAAD' -- the "Front/Back End Wall" rows, which depend on the
    fixed I10 tributary-endwall constant, are never referenced by the
    final output and are intentionally omitted; see module docstring).
    """

    # Wind pressure, left to right (Sheet1!B59:B62)
    pressure_l2r_near_sidewall: float
    pressure_l2r_back_sidewall: float
    pressure_l2r_near_roof: float
    pressure_l2r_back_roof: float

    # Wind suction, left to right (Sheet1!B68:B71)
    suction_l2r_near_sidewall: float
    suction_l2r_back_sidewall: float
    suction_l2r_near_roof: float
    suction_l2r_back_roof: float

    # Wind pressure, right to left (Sheet1!B77:B80)
    pressure_r2l_near_sidewall: float
    pressure_r2l_back_sidewall: float
    pressure_r2l_near_roof: float
    pressure_r2l_back_roof: float

    # Wind suction, right to left (Sheet1!B86:B89)
    suction_r2l_near_sidewall: float
    suction_r2l_back_sidewall: float
    suction_r2l_near_roof: float
    suction_r2l_back_roof: float

    # Wind parallel to ridge, pressure (Sheet1!B95:B98)
    parallel_pressure_near_sidewall: float
    parallel_pressure_back_sidewall: float
    parallel_pressure_near_roof: float
    parallel_pressure_back_roof: float

    # Wind parallel to ridge, suction (Sheet1!B104:B107)
    parallel_suction_near_sidewall: float
    parallel_suction_back_sidewall: float
    parallel_suction_near_roof: float
    parallel_suction_back_roof: float


def zone_forces(params: Is875WindParameters) -> ZoneForces:
    """Compute all per-metre-run zone forces feeding the STAAD output,
    replicating Sheet1's formula chain end to end.
    """
    height = params.height
    width = params.width
    length = params.building_length
    bay = params.bay_spacing
    cpi = params.cpi()  # Sheet1!B27 = WL!F270

    k2 = terrain_height_size_factor_k2(params.terrain_category, height)
    vz = design_wind_speed_vz(params.basic_wind_speed, params.terrain_category, height, params.design_life)
    pz = design_wind_pressure_pz(vz)
    ka = combined_area_factor_ka(width, bay, height)
    pd = design_wind_pressure_pd(pz, ka, height)  # Sheet1!B14 = WL!F152

    h_over_w = height / width  # WL!F161 = Sheet1!B31
    l_over_w = length / width  # WL!F162 = Sheet1!B32

    walls = wall_pressure_coefficients(h_over_w, l_over_w)
    wall_a0, wall_b0 = walls["0"]["A"], walls["0"]["B"]  # Sheet1!B36, B37
    wall_a90, wall_b90 = walls["90"]["A"], walls["90"]["B"]  # Sheet1!B40, B41

    roof = roof_pressure_coefficients(h_over_w, params.roof_slope_x)
    # Sheet1!B46 = ROUNDUP(WL!D264, 3); B47 = WL!E264 (no rounding)
    roof_near_0 = _excel_roundup(roof["EF"], 3)
    roof_far_0 = roof["GH"]
    # Sheet1!B50 = WL!F264; B51 = WL!G264 (G264/FH is computed but unused
    # downstream, see module docstring)
    roof_near_90 = roof["EG"]

    def cp_force(cpe: float, tributary: float) -> float:
        return (cpe - cpi) * tributary * pd

    def cp_force_suction(cpe: float, tributary: float) -> float:
        return (cpe + cpi) * tributary * pd

    return ZoneForces(
        pressure_l2r_near_sidewall=cp_force(wall_a0, bay),  # B59
        pressure_l2r_back_sidewall=cp_force(wall_b0, bay),  # B60
        pressure_l2r_near_roof=cp_force(roof_near_0, bay),  # B61
        pressure_l2r_back_roof=cp_force(roof_far_0, bay),  # B62
        suction_l2r_near_sidewall=cp_force_suction(wall_a0, bay),  # B68
        suction_l2r_back_sidewall=cp_force_suction(wall_b0, bay),  # B69
        suction_l2r_near_roof=cp_force_suction(roof_near_0, bay),  # B70
        suction_l2r_back_roof=cp_force_suction(roof_far_0, bay),  # B71
        pressure_r2l_near_sidewall=cp_force(wall_b0, bay),  # B77 (swap A/B)
        pressure_r2l_back_sidewall=cp_force(wall_a0, bay),  # B78
        pressure_r2l_near_roof=cp_force(roof_far_0, bay),  # B79 (swap near/far)
        pressure_r2l_back_roof=cp_force(roof_near_0, bay),  # B80
        suction_r2l_near_sidewall=cp_force_suction(wall_b0, bay),  # B86
        suction_r2l_back_sidewall=cp_force_suction(wall_a0, bay),  # B87
        suction_r2l_near_roof=cp_force_suction(roof_far_0, bay),  # B88
        suction_r2l_back_roof=cp_force_suction(roof_near_0, bay),  # B89
        parallel_pressure_near_sidewall=cp_force(wall_a90, bay),  # B95
        parallel_pressure_back_sidewall=cp_force(wall_b90, bay),  # B96
        parallel_pressure_near_roof=cp_force(roof_near_90, bay),  # B97
        parallel_pressure_back_roof=cp_force(roof_near_90, bay),  # B98 (== B97, quirk)
        parallel_suction_near_sidewall=cp_force_suction(wall_a90, bay),  # B104
        parallel_suction_back_sidewall=cp_force_suction(wall_b90, bay),  # B105
        parallel_suction_near_roof=cp_force_suction(roof_near_90, bay),  # B106
        parallel_suction_back_roof=cp_force_suction(roof_near_90, bay),  # B107 (== B106, quirk)
    )


# ---------------------------------------------------------------------------
# 'For STAAD' output assembly
# ---------------------------------------------------------------------------


def _format_cell(value) -> str:
    """Matches the old Excel-COM driver's cell-to-text formatting: floats
    are formatted to 3 decimal places, with trailing zeros (and a trailing
    '.') stripped when the value is a whole number.
    """
    if isinstance(value, float):
        return f"{value:.3f}".rstrip("0").rstrip(".") if value == int(value) else f"{value:.3f}"
    return str(value).strip()


def _member_load_line(member_ids: list[int], axis: str, value: float) -> str:
    ids = " ".join(str(i) for i in member_ids)
    return f"{ids} UNI {axis} {_format_cell(value)}"


def generate_is875_wind_load_lines(
    params: Is875WindParameters,
    load_number: int,
    left_column_beams: list[int],
    left_rafter_beams: list[int],
    right_rafter_beams: list[int],
    right_column_beams: list[int],
) -> list[str]:
    """Compute the IS 875 Part 3 wind loads and return the STAAD command
    lines (8 wind load cases) that 'For STAAD'!B13:F84 used to produce via
    Excel COM automation. Same signature/return type as the old
    Excel-driven implementation, so callers need no changes.
    """
    if not left_column_beams or not right_column_beams:
        raise WindLoadError(
            "Wind load requires at least one left and right column member — "
            "check the frame geometry before generating wind load."
        )

    z = zone_forces(params)

    lines: list[str] = []

    def wind_case(header: str, wl_index: int, near_sw: float, back_sw: float, near_roof: float, back_roof: float) -> None:
        n = load_number + (wl_index - 1)
        lines.append(header)
        lines.append(f"LOAD  {n} LOADTYPE Wind  TITLE WL{wl_index}")
        lines.append("MEMBER LOAD")
        lines.append(_member_load_line(left_column_beams, "GX", near_sw))
        lines.append(_member_load_line(left_rafter_beams, "Y", -near_roof))
        lines.append(_member_load_line(right_rafter_beams, "Y", -back_roof))
        lines.append(_member_load_line(right_column_beams, "GX", -back_sw))

    wind_case(
        "************WIND PRESSURE LEFT TO RIGHT*********", 1,
        z.pressure_l2r_near_sidewall, z.pressure_l2r_back_sidewall,
        z.pressure_l2r_near_roof, z.pressure_l2r_back_roof,
    )
    wind_case(
        "************WIND PRESSURE RIGHT TO LEFT*********", 2,
        z.pressure_r2l_near_sidewall, z.pressure_r2l_back_sidewall,
        z.pressure_r2l_near_roof, z.pressure_r2l_back_roof,
    )
    wind_case(
        "************WIND SUCTION LEFT TO RIGHT*********", 3,
        z.suction_l2r_near_sidewall, z.suction_l2r_back_sidewall,
        z.suction_l2r_near_roof, z.suction_l2r_back_roof,
    )
    wind_case(
        "************WIND SUCTION RIGHT TO LEFT*********", 4,
        z.suction_r2l_near_sidewall, z.suction_r2l_back_sidewall,
        z.suction_r2l_near_roof, z.suction_r2l_back_roof,
    )
    wind_case(
        "************WIND PARALLEL PRESSURE GABLE-1*********", 5,
        z.parallel_pressure_near_sidewall, z.parallel_pressure_back_sidewall,
        z.parallel_pressure_near_roof, z.parallel_pressure_back_roof,
    )
    # NOTE: the workbook's own header literal for GABLE-2 has one fewer
    # trailing '*' than GABLE-1 (WL-Single Gable-MBS!B58/B76 are hard-typed
    # VALUE cells, not formulas) -- replicated exactly, not "fixed".
    wind_case(
        "************WIND PARALLEL PRESSURE GABLE-2********", 6,
        z.parallel_pressure_near_sidewall, z.parallel_pressure_back_sidewall,
        z.parallel_pressure_near_roof, z.parallel_pressure_back_roof,
    )
    wind_case(
        "************WIND PARALLEL SUCTION GABLE-1*********", 7,
        z.parallel_suction_near_sidewall, z.parallel_suction_back_sidewall,
        z.parallel_suction_near_roof, z.parallel_suction_back_roof,
    )
    wind_case(
        "************WIND PARALLEL SUCTION GABLE-2********", 8,
        z.parallel_suction_near_sidewall, z.parallel_suction_back_sidewall,
        z.parallel_suction_near_roof, z.parallel_suction_back_roof,
    )

    return lines
