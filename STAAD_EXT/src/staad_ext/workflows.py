from __future__ import annotations

from pathlib import Path
from tkinter import Misc, messagebox

from staad_ext.macros.plate_summary import selected_member_plate_summary
from staad_ext.macros.std_to_dxf import export_selected_members
from staad_ext.macros.std_to_ga_dxf import export_ga_drawing
from staad_ext.macros.std_to_ifc import export_structure
from staad_ext.models import GaExportSettings, IfcExportSettings
from staad_ext.openstaad import OpenStaad, OpenStaadError
from staad_ext.ui import ask_export_settings, show_plate_summary


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


def run_std_to_ga_dxf(parent: Misc | None = None) -> bool:
    """Run the selected-member GA drawing export with default settings."""
    try:
        staad = OpenStaad.connect()
        try:
            model = staad.model_path()
        except OpenStaadError:
            model = Path.cwd() / "STAAD_Model.std"

        output = model.with_name(f"{model.stem}_GA_Drawing.dxf")
        count = export_ga_drawing(staad, output, GaExportSettings())
        if count:
            messagebox.showinfo(
                "STAAD_EXT",
                f"GA drawing created successfully:\n{output}\n\nMembers exported: {count}",
                parent=parent,
            )
            return True

        messagebox.showwarning(
            "STAAD_EXT",
            "No members exported.\n\n"
            "Select one or more analytical beam members in STAAD.Pro and retry.",
            parent=parent,
        )
    except (OpenStaadError, OSError, TypeError, ValueError) as exc:
        messagebox.showerror("STAAD_EXT", str(exc), parent=parent)
    return False


def run_std_to_ifc(parent: Misc | None = None) -> bool:
    """Run the whole-structure IFC export workflow."""
    try:
        staad = OpenStaad.connect()
        try:
            model = staad.model_path()
        except OpenStaadError:
            model = Path.cwd() / "STAAD_Model.std"

        output = model.with_name(f"{model.stem}_3D_Structure.ifc")
        count = export_structure(staad, output, IfcExportSettings())
        if count:
            messagebox.showinfo(
                "STAAD_EXT",
                f"IFC created successfully:\n{output}\n\nMembers exported: {count}",
                parent=parent,
            )
            return True

        messagebox.showwarning(
            "STAAD_EXT",
            f"IFC created, but the model has no analytical beam members.\n\n{output}",
            parent=parent,
        )
    except (OpenStaadError, OSError, TypeError, ValueError) as exc:
        messagebox.showerror("STAAD_EXT", str(exc), parent=parent)
    return False


def run_plate_summary(parent: Misc | None = None) -> bool:
    """Show the plate/whole-section schedule for selected beam members."""
    try:
        staad = OpenStaad.connect()
        selected_count = len(staad.selected_beams())
        if selected_count == 0:
            messagebox.showwarning(
                "STAAD_EXT",
                "Select one or more analytical beam members in STAAD.Pro and retry.",
                parent=parent,
            )
            return False
        rows = selected_member_plate_summary(staad)
        show_plate_summary(rows, selected_count, parent=parent)
        return True
    except (OpenStaadError, OSError, TypeError, ValueError) as exc:
        messagebox.showerror("STAAD_EXT", str(exc), parent=parent)
        return False
