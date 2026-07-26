from __future__ import annotations

from pathlib import Path
from tkinter import Misc, messagebox

from staad_ext.macros.std_to_dxf import export_selected_members
from staad_ext.openstaad import OpenStaad, OpenStaadError
from staad_ext.ui import ask_export_settings


def run_std_to_dxf(parent: Misc | None = None) -> bool:
    """Run the selected-member DXF export workflow."""
    try:
        staad = OpenStaad.connect()
        try:
            model = staad.model_path()
        except OpenStaadError:
            # Some STAAD.Pro 2025 COM registrations do not marshal the
            # GetSTAADFile output string through late-bound comtypes. The path
            # is only needed to suggest an output filename.
            model = Path.cwd() / "STAAD_Model.std"

        default_output = model.with_name(f"{model.stem}_Geometry_Sections.dxf")
        requested = ask_export_settings(default_output, parent=parent)
        if requested is None:
            return False

        output, settings = requested
        count = export_selected_members(staad, output, settings)
        if count:
            messagebox.showinfo(
                "STAAD_EXT",
                f"DXF created successfully:\n{output}\n\nMembers exported: {count}",
                parent=parent,
            )
            return True

        messagebox.showwarning(
            "STAAD_EXT",
            "DXF created, but no selected members were exported.\n\n"
            "Select one or more analytical beam members in STAAD.Pro and retry.\n\n"
            f"{output}",
            parent=parent,
        )
    except (OpenStaadError, OSError, TypeError, ValueError) as exc:
        messagebox.showerror("STAAD_EXT", str(exc), parent=parent)
    return False
