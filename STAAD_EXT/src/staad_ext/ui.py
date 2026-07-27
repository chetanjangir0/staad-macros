from __future__ import annotations

from pathlib import Path
from tkinter import (
    BooleanVar,
    DoubleVar,
    Misc,
    StringVar,
    Tk,
    Toplevel,
    filedialog,
    messagebox,
    ttk,
)

from staad_ext.models import ExportSettings, ViewPlane


def ask_export_settings(
    default_path: Path,
    parent: Misc | None = None,
) -> tuple[Path, ExportSettings] | None:
    """Show the STD-to-DXF settings dialog."""
    window = Toplevel(parent) if parent is not None else Tk()
    window.title("Export STAAD Members to DXF")
    window.resizable(False, False)
    if parent is not None:
        window.transient(parent)
        window.grab_set()

    path = StringVar(master=window, value=str(default_path))
    plane = StringVar(master=window, value=ViewPlane.XY.value)
    labels = BooleanVar(master=window, value=True)
    peb_joins = BooleanVar(master=window, value=False)
    scale = DoubleVar(master=window, value=1.0)
    result: list[tuple[Path, ExportSettings]] = []

    def browse() -> None:
        selected = filedialog.asksaveasfilename(
            parent=window,
            title="Save DXF",
            initialdir=default_path.parent,
            initialfile=default_path.name,
            defaultextension=".dxf",
            filetypes=(("DXF files", "*.dxf"),),
        )
        if selected:
            path.set(selected)

    def submit() -> None:
        try:
            target = Path(path.get().strip()).with_suffix(".dxf")
            settings = ExportSettings(
                ViewPlane(plane.get()),
                labels.get(),
                float(scale.get()),
                peb_joins.get(),
            )
            if not target.name:
                raise ValueError("Choose an output file.")
        except ValueError as exc:
            messagebox.showerror("Invalid settings", str(exc), parent=window)
            return
        result.append((target, settings))
        window.destroy()

    frame = ttk.Frame(window, padding=16)
    frame.grid()
    ttk.Label(frame, text="Output DXF").grid(row=0, column=0, sticky="w")
    ttk.Entry(frame, textvariable=path, width=55).grid(
        row=1, column=0, padx=(0, 8)
    )
    ttk.Button(frame, text="Browse...", command=browse).grid(row=1, column=1)

    ttk.Label(frame, text="Projection plane").grid(
        row=2, column=0, pady=(14, 2), sticky="w"
    )
    plane_frame = ttk.Frame(frame)
    plane_frame.grid(row=3, column=0, sticky="w")
    for value in ViewPlane:
        ttk.Radiobutton(
            plane_frame,
            text=value.value,
            value=value.value,
            variable=plane,
        ).pack(side="left", padx=(0, 12))

    ttk.Checkbutton(frame, text="Write text labels", variable=labels).grid(
        row=4, column=0, pady=(14, 2), sticky="w"
    )
    ttk.Checkbutton(
        frame,
        text="Use PEB-style corner joins",
        variable=peb_joins,
    ).grid(row=5, column=0, pady=(6, 2), sticky="w")
    ttk.Label(frame, text="Text size scale (0.1-10.0)").grid(
        row=6, column=0, sticky="w"
    )
    ttk.Entry(frame, textvariable=scale, width=12).grid(
        row=7, column=0, sticky="w"
    )

    actions = ttk.Frame(frame)
    actions.grid(row=8, column=0, columnspan=2, pady=(18, 0), sticky="e")
    ttk.Button(actions, text="Cancel", command=window.destroy).pack(
        side="left", padx=4
    )
    ttk.Button(actions, text="Export", command=submit).pack(side="left")

    window.protocol("WM_DELETE_WINDOW", window.destroy)
    if parent is None:
        window.mainloop()
    else:
        window.wait_window()
    return result[0] if result else None
