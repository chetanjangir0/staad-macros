from __future__ import annotations

from pathlib import Path
from tkinter import BooleanVar, DoubleVar, StringVar, Tk, filedialog, messagebox, ttk

from staad_ext.models import ExportSettings, ViewPlane


def ask_export_settings(default_path: Path) -> tuple[Path, ExportSettings] | None:
    root = Tk()
    root.title("Export STAAD Members to DXF")
    root.resizable(False, False)
    path, plane = StringVar(value=str(default_path)), StringVar(value=ViewPlane.XY.value)
    labels, scale = BooleanVar(value=True), DoubleVar(value=1.0)
    result: list[tuple[Path, ExportSettings]] = []

    def browse() -> None:
        selected = filedialog.asksaveasfilename(parent=root, title="Save DXF", initialdir=default_path.parent,
            initialfile=default_path.name, defaultextension=".dxf", filetypes=(("DXF files", "*.dxf"),))
        if selected:
            path.set(selected)

    def submit() -> None:
        try:
            target = Path(path.get().strip()).with_suffix(".dxf")
            settings = ExportSettings(ViewPlane(plane.get()), labels.get(), float(scale.get()))
            if not target.name:
                raise ValueError("Choose an output file.")
        except ValueError as exc:
            messagebox.showerror("Invalid settings", str(exc), parent=root)
            return
        result.append((target, settings))
        root.destroy()

    frame = ttk.Frame(root, padding=16)
    frame.grid()
    ttk.Label(frame, text="Output DXF").grid(row=0, column=0, sticky="w")
    ttk.Entry(frame, textvariable=path, width=55).grid(row=1, column=0, padx=(0, 8))
    ttk.Button(frame, text="Browse…", command=browse).grid(row=1, column=1)
    ttk.Label(frame, text="Projection plane").grid(row=2, column=0, pady=(14, 2), sticky="w")
    plane_frame = ttk.Frame(frame)
    plane_frame.grid(row=3, column=0, sticky="w")
    for value in ViewPlane:
        ttk.Radiobutton(plane_frame, text=value.value, value=value.value, variable=plane).pack(side="left", padx=(0, 12))
    ttk.Checkbutton(frame, text="Write text labels", variable=labels).grid(row=4, column=0, pady=(14, 2), sticky="w")
    ttk.Label(frame, text="Text size scale (0.1–10.0)").grid(row=5, column=0, sticky="w")
    ttk.Entry(frame, textvariable=scale, width=12).grid(row=6, column=0, sticky="w")
    actions = ttk.Frame(frame)
    actions.grid(row=7, column=0, columnspan=2, pady=(18, 0), sticky="e")
    ttk.Button(actions, text="Cancel", command=root.destroy).pack(side="left", padx=4)
    ttk.Button(actions, text="Export", command=submit).pack(side="left")
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()
    return result[0] if result else None
