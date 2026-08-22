"""Public entry point for IS 875 Part 3 wind load generation.

This used to drive the bundled 'IS_Mid Frame_2015wind.xlsx' workbook live
via Excel COM automation. It has been replaced by a pure-Python port
(src/staad_ext/macros/is875_wind.py), verified against golden-master
fixtures captured from that same workbook (tests/fixtures/
is875_wind_golden.json, tests/test_is875_wind_golden.py) -- see
is875_wind.py's module docstring for the full calculation chain and the
workbook quirks it faithfully replicates.

`Is875WindParameters`, `WindLoadError` and `generate_is875_wind_load_lines`
are re-exported here unchanged so existing imports
(`from staad_ext.macros.wind_load import ...`, e.g. in frame_generator.py)
continue to work without modification. Excel is no longer required at
runtime for IS 875 Part 3 wind loads.
"""
from __future__ import annotations

from staad_ext.macros.is875_wind import (
    BASIC_WIND_SPEEDS,
    DESIGN_LIVES,
    OPENING_OPTIONS,
    TERRAIN_CATEGORIES,
    Is875WindParameters,
    WindLoadError,
    generate_is875_wind_load_lines,
)

__all__ = [
    "BASIC_WIND_SPEEDS",
    "DESIGN_LIVES",
    "OPENING_OPTIONS",
    "TERRAIN_CATEGORIES",
    "Is875WindParameters",
    "WindLoadError",
    "generate_is875_wind_load_lines",
]
