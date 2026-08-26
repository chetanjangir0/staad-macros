from __future__ import annotations

import argparse

from staad_ext.application import launch
from staad_ext.workflows import (
    run_plate_summary, run_std_to_dxf, run_std_to_ga_dxf, run_std_to_ifc,
    run_taper_optimizer,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="staad-ext")
    commands = parser.add_subparsers(dest="command")
    commands.add_parser("std-to-dxf", help="Export selected STAAD members to an R12 DXF")
    commands.add_parser(
        "std-to-ga-dxf",
        help="Export selected STAAD members as a GA drawing with a member size schedule",
    )
    commands.add_parser("std-to-ifc", help="Export the whole STAAD model to a 3D IFC file")
    commands.add_parser("plate-summary", help="Show a plate summary for selected STAAD members")
    commands.add_parser(
        "optimize-tapers",
        help="Optimize the tapered I sections of the open 2D STAAD frame",
    )
    args = parser.parse_args(argv)
    if args.command == "std-to-dxf":
        return 0 if run_std_to_dxf() else 1
    if args.command == "std-to-ga-dxf":
        return 0 if run_std_to_ga_dxf() else 1
    if args.command == "std-to-ifc":
        return 0 if run_std_to_ifc() else 1
    if args.command == "plate-summary":
        return 0 if run_plate_summary() else 1
    if args.command == "optimize-tapers":
        return 0 if run_taper_optimizer() else 1
    return launch()
