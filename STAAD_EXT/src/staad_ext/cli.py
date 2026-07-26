from __future__ import annotations

import argparse

from staad_ext.application import launch
from staad_ext.workflows import run_std_to_dxf


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="staad-ext")
    commands = parser.add_subparsers(dest="command")
    commands.add_parser("std-to-dxf", help="Export selected STAAD members to an R12 DXF")
    args = parser.parse_args(argv)
    if args.command == "std-to-dxf":
        return 0 if run_std_to_dxf() else 1
    return launch()
