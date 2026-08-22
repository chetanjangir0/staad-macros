"""Verify the pure-Python IS 875 Part 3 wind load port
(src/staad_ext/macros/is875_wind.py) against golden-master fixtures
captured from the validated Excel workbook via Excel COM automation.

Fixtures live in tests/fixtures/is875_wind_golden.json and are regenerated
with scripts/generate_wind_golden_fixtures.py whenever the source .xlsx
changes.

If a fixture fails, the intermediate values captured alongside it
(WL-Single Gable-MBS!F36/F37/F146/F150/F151/F152/K152/F161/F162/F270 and
the full Sheet1!B column) pinpoint exactly which step of the calculation
chain diverges -- do not loosen the tolerance to make a failure disappear.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from staad_ext.macros.is875_wind import (
    Is875WindParameters,
    area_reduction_factor_ka,
    combined_area_factor_ka,
    design_wind_pressure_pd,
    design_wind_pressure_pz,
    design_wind_speed_vz,
    generate_is875_wind_load_lines,
    terrain_height_size_factor_k2,
    zone_forces,
)

FIXTURES_PATH = Path(__file__).parent / "fixtures" / "is875_wind_golden.json"

TOLERANCE_KN_PER_M = 0.01  # kN/m — matches the 3-decimal-place STAAD output format
TOLERANCE_FACTOR = 1e-6  # for dimensionless intermediate factors (K2, Ka, Cpi, ratios)
TOLERANCE_PRESSURE = 1e-4  # kN/m^2 for Vz/Pz/Pd


def _load_fixtures():
    with open(FIXTURES_PATH, encoding="utf-8") as f:
        return json.load(f)


FIXTURES = _load_fixtures()


def _params_from_fixture(fixture) -> Is875WindParameters:
    p = fixture["params"]
    return Is875WindParameters(
        building_length=p["building_length"],
        width=p["width"],
        height=p["height"],
        roof_slope_x=p["roof_slope_x"],
        basic_wind_speed=p["basic_wind_speed"],
        design_life=p["design_life"],
        terrain_category=p["terrain_category"],
        bay_spacing=p["bay_spacing"],
        opening=p["opening"],
    )


def test_fixture_file_has_enough_cases():
    assert len(FIXTURES) >= 300, "Expected at least 300 golden fixtures"


@pytest.mark.parametrize("fixture", FIXTURES, ids=lambda fx: fx["label"])
def test_intermediate_values_match(fixture):
    params = _params_from_fixture(fixture)
    inter = fixture["intermediates"]

    k2 = terrain_height_size_factor_k2(params.terrain_category, params.height)
    assert k2 == pytest.approx(inter["K2_146"], abs=TOLERANCE_FACTOR), "K2 (terrain/height/size factor) mismatch"

    vz = design_wind_speed_vz(params.basic_wind_speed, params.terrain_category, params.height, params.design_life)
    assert vz == pytest.approx(inter["Vz_150"], abs=TOLERANCE_PRESSURE), "Vz (design wind speed) mismatch"

    pz = design_wind_pressure_pz(vz)
    assert pz == pytest.approx(inter["Pz_151"], abs=TOLERANCE_PRESSURE), "Pz (wind pressure) mismatch"

    ka = combined_area_factor_ka(params.width, params.bay_spacing, params.height)
    pd = design_wind_pressure_pd(pz, ka, params.height)
    assert pd == pytest.approx(inter["Pd_152"], abs=TOLERANCE_PRESSURE), "Pd (design wind pressure) mismatch"

    h_over_w = params.height / params.width
    l_over_w = params.building_length / params.width
    assert h_over_w == pytest.approx(inter["h_over_w_161"], abs=TOLERANCE_FACTOR)
    assert l_over_w == pytest.approx(inter["l_over_w_162"], abs=TOLERANCE_FACTOR)

    assert params.cpi() == pytest.approx(inter["Cpi_270"], abs=TOLERANCE_FACTOR), "Cpi mismatch"


@pytest.mark.parametrize("fixture", FIXTURES, ids=lambda fx: fx["label"])
def test_sheet1_zone_forces_match(fixture):
    params = _params_from_fixture(fixture)
    z = zone_forces(params)
    b = fixture["sheet1_b_column"]

    mapping = {
        "59": z.pressure_l2r_near_sidewall,
        "60": z.pressure_l2r_back_sidewall,
        "61": z.pressure_l2r_near_roof,
        "62": z.pressure_l2r_back_roof,
        "68": z.suction_l2r_near_sidewall,
        "69": z.suction_l2r_back_sidewall,
        "70": z.suction_l2r_near_roof,
        "71": z.suction_l2r_back_roof,
        "77": z.pressure_r2l_near_sidewall,
        "78": z.pressure_r2l_back_sidewall,
        "79": z.pressure_r2l_near_roof,
        "80": z.pressure_r2l_back_roof,
        "86": z.suction_r2l_near_sidewall,
        "87": z.suction_r2l_back_sidewall,
        "88": z.suction_r2l_near_roof,
        "89": z.suction_r2l_back_roof,
        "95": z.parallel_pressure_near_sidewall,
        "96": z.parallel_pressure_back_sidewall,
        "97": z.parallel_pressure_near_roof,
        "98": z.parallel_pressure_back_roof,
        "104": z.parallel_suction_near_sidewall,
        "105": z.parallel_suction_back_sidewall,
        "106": z.parallel_suction_near_roof,
        "107": z.parallel_suction_back_roof,
    }
    for row, python_value in mapping.items():
        if row not in b:
            continue
        assert python_value == pytest.approx(b[row], abs=TOLERANCE_KN_PER_M), f"Sheet1!B{row} mismatch"


@pytest.mark.parametrize("fixture", FIXTURES, ids=lambda fx: fx["label"])
def test_staad_output_lines_match(fixture):
    params = _params_from_fixture(fixture)
    lines = generate_is875_wind_load_lines(
        params,
        fixture["load_number"],
        fixture["left_column_beams"],
        fixture["left_rafter_beams"],
        fixture["right_rafter_beams"],
        fixture["right_column_beams"],
    )
    expected = fixture["output_lines"]

    assert len(lines) == len(expected), (
        f"Line count mismatch for {fixture['label']}: got {len(lines)}, expected {len(expected)}\n"
        f"Got: {lines}\nExpected: {expected}"
    )

    for i, (actual, exp) in enumerate(zip(lines, expected)):
        if actual == exp:
            continue
        # Numeric lines like "1 5 UNI GX 2.372" — compare the trailing
        # number with a tolerance instead of exact string equality, since
        # both sides already round to 3 decimals but float rounding can
        # differ by 1 ULP at the boundary.
        actual_parts = actual.rsplit(" ", 1)
        exp_parts = exp.rsplit(" ", 1)
        assert len(actual_parts) == 2 and len(exp_parts) == 2 and actual_parts[0] == exp_parts[0], (
            f"Line {i} mismatch for {fixture['label']}:\n  got:      {actual!r}\n  expected: {exp!r}"
        )
        try:
            actual_num = float(actual_parts[1])
            exp_num = float(exp_parts[1])
        except ValueError:
            pytest.fail(f"Line {i} mismatch for {fixture['label']}:\n  got:      {actual!r}\n  expected: {exp!r}")
        assert actual_num == pytest.approx(exp_num, abs=TOLERANCE_KN_PER_M), (
            f"Line {i} numeric mismatch for {fixture['label']}:\n  got:      {actual!r}\n  expected: {exp!r}"
        )
