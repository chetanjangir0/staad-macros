from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

# IS 875 Part 3 dropdown options exposed by the UI.
BASIC_WIND_SPEEDS = (33.0, 39.0, 44.0, 47.0, 50.0, 55.0)
DESIGN_LIVES = (5, 25, 50, 100)
TERRAIN_CATEGORIES = (1, 2, 3, 4)
OPENING_OPTIONS = ("<5%", "5-20%", ">20%")

_CPI_BY_OPENING = {"<5%": 0.2, "5-20%": 0.5, ">20%": 0.7}

_IS875_WORKBOOK = "IS_Mid Frame_2015wind.xlsx"
_IS875_INPUT_SHEET = "WL-Single Gable-MBS"
_IS875_OUTPUT_SHEET = "For STAAD"
_IS875_OUTPUT_RANGE = "B13:F84"


class WindLoadError(RuntimeError):
    """Raised when the wind load workbook cannot be driven to produce results."""


@dataclass
class Is875WindParameters:
    """User-facing inputs for the IS 875 Part 3 'WL-Single Gable-MBS' workbook."""

    building_length: float  # m (F5)
    width: float  # m (F6)
    height: float  # m, from FGL (F10, overwrites the F7+F8 formula)
    roof_slope_x: float  # x in 1:x (F12)
    basic_wind_speed: float  # m/s, one of BASIC_WIND_SPEEDS (F13)
    design_life: int  # years, one of DESIGN_LIVES (F14)
    terrain_category: int  # 1-4 (F15)
    bay_spacing: float  # m (I7)
    opening: str  # one of OPENING_OPTIONS

    def cpi(self) -> float:
        try:
            return _CPI_BY_OPENING[self.opening]
        except KeyError as exc:
            raise ValueError(f"Unknown opening percentage option: {self.opening!r}") from exc


def _wind_profiles_dir() -> Path:
    """Locate the wind_profiles resource folder in both source and frozen (.exe) layouts."""
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).parent
        for candidate in (base / "_internal" / "staad_ext" / "wind_profiles", base / "staad_ext" / "wind_profiles"):
            if candidate.exists():
                return candidate
        return candidate  # let the caller raise a clear "not found" error
    return Path(__file__).resolve().parent.parent / "wind_profiles"


def _is875_workbook_path() -> Path:
    path = _wind_profiles_dir() / _IS875_WORKBOOK
    if not path.exists():
        raise WindLoadError(f"IS 875 Part 3 wind load workbook not found: {path}")
    return path


def _format_cell(value: Any) -> str:
    if isinstance(value, float):
        return f"{value:.3f}".rstrip("0").rstrip(".") if value == int(value) else f"{value:.3f}"
    return str(value).strip()


def generate_is875_wind_load_lines(
    params: Is875WindParameters,
    load_number: int,
    left_column_beams: list[int],
    left_rafter_beams: list[int],
    right_rafter_beams: list[int],
    right_column_beams: list[int],
) -> list[str]:
    """Drive the IS 875 Part 3 workbook via Excel COM and return the generated
    STAAD command lines (4 wind load cases) from its 'For STAAD' sheet.

    Requires Microsoft Excel to be installed. Opens the bundled template
    invisibly, writes the inputs, forces a full recalculation, reads back the
    computed output range, and closes without saving any changes to the
    template file.
    """
    if not left_column_beams or not right_column_beams:
        raise WindLoadError(
            "Wind load requires at least one left and right column member — "
            "check the frame geometry before generating wind load."
        )

    from comtypes.client import CreateObject

    workbook_path = _is875_workbook_path()

    try:
        excel = CreateObject("Excel.Application", dynamic=True)
    except OSError as exc:
        raise WindLoadError(
            "Could not start Microsoft Excel via COM automation. "
            "Excel must be installed to compute IS 875 Part 3 wind loads."
        ) from exc

    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Open(str(workbook_path), False, False)
        try:
            input_sheet = workbook.Sheets(_IS875_INPUT_SHEET)
            output_sheet = workbook.Sheets(_IS875_OUTPUT_SHEET)

            input_sheet.Range("F5").Value2 = params.building_length
            input_sheet.Range("F6").Value2 = params.width
            input_sheet.Range("F10").Value2 = params.height
            input_sheet.Range("F12").Value2 = params.roof_slope_x
            input_sheet.Range("F13").Value2 = params.basic_wind_speed
            input_sheet.Range("F14").Value2 = params.design_life
            input_sheet.Range("F15").Value2 = params.terrain_category
            input_sheet.Range("I7").Value2 = params.bay_spacing
            input_sheet.Range("F270").Value2 = params.cpi()

            output_sheet.Range("E6").Value2 = load_number
            output_sheet.Range("E7").Value2 = " ".join(str(i) for i in left_column_beams)
            output_sheet.Range("E8").Value2 = " ".join(str(i) for i in left_rafter_beams)
            output_sheet.Range("E9").Value2 = " ".join(str(i) for i in right_rafter_beams)
            output_sheet.Range("E10").Value2 = " ".join(str(i) for i in right_column_beams)

            excel.CalculateFullRebuild()

            raw_rows = output_sheet.Range(_IS875_OUTPUT_RANGE).Value2
            lines: list[str] = []
            for row in raw_rows:
                parts = [_format_cell(cell) for cell in row if cell is not None and str(cell).strip()]
                if parts:
                    lines.append(" ".join(parts))
            return lines
        finally:
            workbook.Close(False)
    finally:
        excel.Quit()
