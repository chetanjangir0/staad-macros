"""Regenerate tests/fixtures/is875_wind_golden.json from the validated Excel
workbook `src/staad_ext/wind_profiles/IS_Mid Frame_2015wind.xlsx`.

This is the golden-master fixture generator for the IS 875 Part 3 wind load
port (src/staad_ext/macros/is875_wind.py). It drives the workbook live via
Excel COM automation (comtypes) -- the same mechanism the old
`generate_is875_wind_load_lines` used -- writes a matrix of representative
inputs, forces a full recalculation, and captures both the final STAAD
command lines ('For STAAD'!B13:F84) and a set of key intermediate values
(from 'WL-Single Gable-MBS' and 'Sheet1') for debugging the pure-Python port.

Run with: .venv/Scripts/python.exe scripts/generate_wind_golden_fixtures.py

Re-run this whenever the source .xlsx is revised.
"""
from __future__ import annotations

import itertools
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SRC = REPO_ROOT / "src"
sys.path.insert(0, str(SRC))

from staad_ext.macros.wind_load import (  # noqa: E402
    BASIC_WIND_SPEEDS,
    DESIGN_LIVES,
    OPENING_OPTIONS,
    TERRAIN_CATEGORIES,
    Is875WindParameters,
)

WORKBOOK_PATH = SRC / "staad_ext" / "wind_profiles" / "IS_Mid Frame_2015wind.xlsx"
INPUT_SHEET = "WL-Single Gable-MBS"
OUTPUT_SHEET = "For STAAD"
OUTPUT_RANGE = "B13:F84"

FIXTURES_PATH = REPO_ROOT / "tests" / "fixtures" / "is875_wind_golden.json"

# Fixed member-id groups used for every fixture -- chosen to also exercise
# multi-id space-joined formatting on the column groups.
LOAD_NUMBER = 7
LEFT_COLUMN_BEAMS = [1, 5]
LEFT_RAFTER_BEAMS = [3]
RIGHT_RAFTER_BEAMS = [4]
RIGHT_COLUMN_BEAMS = [2, 6]

# Representative geometries for the wind-parameter sweep (basic_wind_speed x
# terrain_category x design_life x opening, all discrete values).
GEOM_SWEEP_BASES = [
    # (building_length, width, height, roof_slope_x, bay_spacing)
    (72.0, 22.015, 28.5, 10.0, 6.0),   # the workbook's own shipped example
    (30.0, 15.0, 6.0, 6.0, 6.0),        # smaller, low-rise, h/w < 1
]

# Representative geometries for the geometry sweep (small/medium/large spans,
# shallow/steep roof slopes, h/w above and below 1), at one fixed
# wind-speed/terrain/life/opening combination.
GEOM_ONLY_CASES = [
    # (building_length, width, height, roof_slope_x, bay_spacing)
    (20.0, 10.0, 5.0, 20.0, 5.0),    # small span, very shallow roof, h/w<1
    (20.0, 10.0, 5.0, 3.0, 5.0),     # small span, steep roof
    (40.0, 20.0, 8.0, 10.0, 6.0),    # medium span
    (40.0, 20.0, 8.0, 2.0, 6.0),     # medium span, steep roof (h/w<1 still, w large)
    (60.0, 18.0, 22.0, 10.0, 7.5),   # tall narrow building, h/w > 1
    (100.0, 30.0, 10.0, 12.0, 8.0),  # large span, wide, h/w < 1
    (100.0, 12.0, 20.0, 8.0, 6.0),   # large length, narrow width, h/w > 1.5
    (25.0, 25.0, 25.0, 10.0, 5.0),   # square-ish, h/w ~1
    (50.0, 8.0, 15.0, 6.0, 5.0),     # very narrow width, h/w > 1.5 (possibly > 6)
    (36.0, 24.0, 6.0, 5.0, 6.0),     # h/w << 1
    (18.0, 16.0, 24.0, 15.0, 4.0),   # h/w > 1, shallow roof
    (48.0, 20.0, 10.0, 10.0, 6.0),   # h/w ~ 0.5 boundary
]
FIXED_WIND_PARAMS_FOR_GEOM_SWEEP = dict(
    basic_wind_speed=39.0, design_life=50, terrain_category=2, opening="<5%"
)

# Cells to capture from 'WL-Single Gable-MBS' for debugging intermediate
# values (see the docstring above).
WL_CELLS = {
    "A_36": "F36",
    "B_37": "F37",
    "K2_146": "F146",
    "Vz_150": "F150",
    "Pz_151": "F151",
    "Pd_152": "F152",
    "K152_local": "K152",
    "h_over_w_161": "F161",
    "l_over_w_162": "F162",
    "Cpi_270": "F270",
}


def _format_cell(value):
    if isinstance(value, float):
        return f"{value:.3f}".rstrip("0").rstrip(".") if value == int(value) else f"{value:.3f}"
    return str(value).strip()


def build_param_matrix():
    """Yield (label, Is875WindParameters) pairs for every fixture to generate."""
    # 1) Wind-parameter sweep: ALL discrete values of basic_wind_speed x
    #    terrain_category x design_life x opening, at each representative
    #    base geometry.
    for geom_idx, geom in enumerate(GEOM_SWEEP_BASES):
        length, width, height, slope, bay = geom
        for speed, terrain, life, opening in itertools.product(
            BASIC_WIND_SPEEDS, TERRAIN_CATEGORIES, DESIGN_LIVES, OPENING_OPTIONS
        ):
            label = f"windsweep_geom{geom_idx}_v{speed}_t{terrain}_n{life}_op{opening}"
            params = Is875WindParameters(
                building_length=length,
                width=width,
                height=height,
                roof_slope_x=slope,
                basic_wind_speed=speed,
                design_life=life,
                terrain_category=terrain,
                bay_spacing=bay,
                opening=opening,
            )
            yield label, params

    # 2) Geometry sweep at one fixed wind-parameter combination.
    for geom_idx, geom in enumerate(GEOM_ONLY_CASES):
        length, width, height, slope, bay = geom
        label = f"geomsweep_{geom_idx}_L{length}_W{width}_H{height}_S{slope}_B{bay}"
        params = Is875WindParameters(
            building_length=length,
            width=width,
            height=height,
            roof_slope_x=slope,
            bay_spacing=bay,
            **FIXED_WIND_PARAMS_FOR_GEOM_SWEEP,
        )
        yield label, params


def main():
    from comtypes.client import CreateObject

    if not WORKBOOK_PATH.exists():
        raise SystemExit(f"Workbook not found: {WORKBOOK_PATH}")

    excel = CreateObject("Excel.Application", dynamic=True)
    excel.Visible = False
    excel.DisplayAlerts = False

    fixtures = []
    try:
        workbook = excel.Workbooks.Open(str(WORKBOOK_PATH), False, False)
        try:
            input_sheet = workbook.Sheets(INPUT_SHEET)
            output_sheet = workbook.Sheets(OUTPUT_SHEET)
            sheet1 = workbook.Sheets("Sheet1")

            matrix = list(build_param_matrix())
            total = len(matrix)
            print(f"Generating {total} fixtures...")

            for idx, (label, params) in enumerate(matrix, start=1):
                input_sheet.Range("F5").Value2 = params.building_length
                input_sheet.Range("F6").Value2 = params.width
                input_sheet.Range("F7").Value2 = 0  # h1-Height of plinth: no separate UI field, always 0
                input_sheet.Range("F8").Value2 = params.height  # h2-Eaves height from FFL: since F7=0,
                # F10 (=F7+F8, eaves height from FGL) must equal the user's `height` input, so F8 must too.
                # The original 9-cell COM driver never wrote F7/F8, silently leaving F8 stuck at the
                # template's shipped demo value (28.5 m) regardless of the user's actual height -- a real
                # bug (F8 feeds the Ka wall-area lookup and the K152 "<10m" branch), not a faithful quirk.
                input_sheet.Range("F10").Value2 = params.height
                input_sheet.Range("F12").Value2 = params.roof_slope_x
                input_sheet.Range("F13").Value2 = params.basic_wind_speed
                input_sheet.Range("F14").Value2 = params.design_life
                input_sheet.Range("F15").Value2 = params.terrain_category
                input_sheet.Range("I7").Value2 = params.bay_spacing
                input_sheet.Range("F270").Value2 = params.cpi()

                output_sheet.Range("E6").Value2 = LOAD_NUMBER
                output_sheet.Range("E7").Value2 = " ".join(str(i) for i in LEFT_COLUMN_BEAMS)
                output_sheet.Range("E8").Value2 = " ".join(str(i) for i in LEFT_RAFTER_BEAMS)
                output_sheet.Range("E9").Value2 = " ".join(str(i) for i in RIGHT_RAFTER_BEAMS)
                output_sheet.Range("E10").Value2 = " ".join(str(i) for i in RIGHT_COLUMN_BEAMS)

                excel.CalculateFullRebuild()

                raw_rows = output_sheet.Range(OUTPUT_RANGE).Value2
                lines = []
                for row in raw_rows:
                    parts = [_format_cell(cell) for cell in row if cell is not None and str(cell).strip()]
                    if parts:
                        lines.append(" ".join(parts))

                intermediates = {
                    key: input_sheet.Range(addr).Value2 for key, addr in WL_CELLS.items()
                }
                sheet1_b = {}
                for row in range(17, 111):
                    val = sheet1.Range(f"B{row}").Value2
                    if val is not None:
                        sheet1_b[str(row)] = val

                fixtures.append(
                    {
                        "label": label,
                        "params": {
                            "building_length": params.building_length,
                            "width": params.width,
                            "height": params.height,
                            "roof_slope_x": params.roof_slope_x,
                            "basic_wind_speed": params.basic_wind_speed,
                            "design_life": params.design_life,
                            "terrain_category": params.terrain_category,
                            "bay_spacing": params.bay_spacing,
                            "opening": params.opening,
                        },
                        "load_number": LOAD_NUMBER,
                        "left_column_beams": LEFT_COLUMN_BEAMS,
                        "left_rafter_beams": LEFT_RAFTER_BEAMS,
                        "right_rafter_beams": RIGHT_RAFTER_BEAMS,
                        "right_column_beams": RIGHT_COLUMN_BEAMS,
                        "intermediates": intermediates,
                        "sheet1_b_column": sheet1_b,
                        "output_lines": lines,
                    }
                )

                if idx % 25 == 0 or idx == total:
                    print(f"  {idx}/{total} ({label})")
        finally:
            workbook.Close(False)
    finally:
        excel.Quit()

    FIXTURES_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(FIXTURES_PATH, "w", encoding="utf-8") as f:
        json.dump(fixtures, f, indent=1)

    print(f"Wrote {len(fixtures)} fixtures to {FIXTURES_PATH}")


if __name__ == "__main__":
    main()
