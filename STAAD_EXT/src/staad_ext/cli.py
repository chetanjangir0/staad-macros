from __future__ import annotations

import argparse
from pathlib import Path
from tkinter import messagebox

from staad_ext.macros.std_to_dxf import export_selected_members
from staad_ext.openstaad import OpenStaad, OpenStaadError
from staad_ext.ui import ask_export_settings


def _std_to_dxf() -> int:
    try:
        staad = OpenStaad.connect()
        try:
            model = staad.model_path()
        except OpenStaadError:
            # Some STAAD.Pro 2025 COM registrations execute GetSTAADFile but
            # do not marshal its output VARIANT back through late-bound
            # comtypes. The model path is only used to suggest an output name;
            # all actual model data comes from Geometry and Property.
            model = Path.cwd() / "STAAD_Model.std"
        requested = ask_export_settings(model.with_name(f"{model.stem}_Geometry_Sections.dxf"))
        if requested is None:
            return 0
        output, settings = requested
        count = export_selected_members(staad, output, settings)
        if count:
            messagebox.showinfo("STAAD_EXT", f"DXF created successfully:\n{output}\n\nMembers exported: {count}")
        else:
            messagebox.showwarning("STAAD_EXT", f"DXF created, but no selected members were exported.\n\n{output}")
        return 0
    except (OpenStaadError, OSError) as exc:
        messagebox.showerror("STAAD_EXT", str(exc))
        return 1


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="staad-ext")
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("std-to-dxf", help="Export selected STAAD members to an R12 DXF")
    args = parser.parse_args(argv)
    return _std_to_dxf() if args.command == "std-to-dxf" else 2
