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

from staad_ext.models import (
    DeflectionLimits, ExportSettings, TaperOptimizerSettings, ViewPlane,
)
from staad_ext.macros.plate_summary import PlateSummaryRow
from staad_ext.macros.support_reactions import parse_load_cases
from staad_ext.macros.taper_optimizer import OptimizationResult


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
    connection_faces = BooleanVar(master=window, value=False)
    color_by_section = BooleanVar(master=window, value=True)
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
                connection_faces.get(),
                color_by_section.get(),
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
    ttk.Checkbutton(
        frame,
        text="Show GA connection details",
        variable=connection_faces,
    ).grid(row=6, column=0, pady=(6, 2), sticky="w")
    ttk.Checkbutton(
        frame,
        text="Color members by section size",
        variable=color_by_section,
    ).grid(row=7, column=0, pady=(6, 2), sticky="w")
    ttk.Label(frame, text="Text size scale (0.1-10.0)").grid(
        row=8, column=0, sticky="w"
    )
    ttk.Entry(frame, textvariable=scale, width=12).grid(
        row=9, column=0, sticky="w"
    )

    actions = ttk.Frame(frame)
    actions.grid(row=10, column=0, columnspan=2, pady=(18, 0), sticky="e")
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


def ask_taper_optimizer_settings(
    parent: Misc | None = None,
) -> TaperOptimizerSettings | None:
    """Show the tapered-section optimizer settings dialog."""
    window = Toplevel(parent) if parent is not None else Tk()
    window.title("Optimize Tapered Sections")
    window.resizable(False, False)
    if parent is not None:
        window.transient(parent)
        window.grab_set()

    ceiling = StringVar(master=window, value="0.95")
    vertical = StringVar(master=window, value="240")
    horizontal = StringVar(master=window, value="150")
    cases = StringVar(master=window, value="")
    budget = StringVar(master=window, value="40")
    tie_knees = BooleanVar(master=window, value=False)
    straight_columns = BooleanVar(master=window, value=False)
    apply_now = BooleanVar(master=window, value=False)
    result: list[TaperOptimizerSettings] = []

    def submit() -> None:
        try:
            settings = TaperOptimizerSettings(
                deflection=DeflectionLimits(
                    vertical_span_ratio=float(vertical.get()),
                    horizontal_height_ratio=float(horizontal.get()),
                    load_cases=tuple(parse_load_cases(cases.get())),
                ),
                utilisation_ceiling=float(ceiling.get()),
                tie_depths_at_all_shared_nodes=tie_knees.get(),
                prismatic_columns=straight_columns.get(),
                analysis_budget=int(budget.get()),
                apply_to_model=apply_now.get(),
            )
        except ValueError as exc:
            messagebox.showerror("Invalid settings", str(exc), parent=window)
            return
        result.append(settings)
        window.destroy()

    frame = ttk.Frame(window, padding=16)
    frame.grid()
    row = 0
    for label, variable, hint in (
        ("Target utilisation ceiling (0.1 - 1.0)", ceiling,
         "Sections are reduced only while every design ratio stays at or below this."),
        ("Vertical deflection limit  span /", vertical,
         "Rafter sag is measured against the straight chord between the run's ends."),
        ("Horizontal deflection limit  height /", horizontal,
         "Column drift is checked at the top of each column against its height."),
        ("Deflection load combinations", cases,
         "Examples: 101, 102, 105 or 101-105. Use the serviceability combinations."),
        ("Analysis budget (runs)", budget,
         "Each run is a full STAAD.Pro analysis. Most of the saving lands in the first ten."),
    ):
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w",
                                          pady=(0 if row == 0 else 12, 2))
        ttk.Entry(frame, textvariable=variable, width=46).grid(
            row=row + 1, column=0, sticky="w")
        ttk.Label(frame, text=hint, foreground="#555555").grid(
            row=row + 2, column=0, sticky="w", pady=(3, 0))
        row += 3

    ttk.Checkbutton(
        frame, text="Also match depths where different members meet (knee joints)",
        variable=tie_knees,
    ).grid(row=row, column=0, sticky="w", pady=(14, 2))
    ttk.Checkbutton(
        frame, text="Keep columns straight (one depth, no taper)",
        variable=straight_columns,
    ).grid(row=row + 1, column=0, sticky="w", pady=(4, 2))
    ttk.Label(
        frame,
        text=("One depth from base to eave for each column, even where the column "
              "is split into several members. The flange plates are still sized "
              "member by member. Rafters taper either way."),
        foreground="#555555", wraplength=380, justify="left",
    ).grid(row=row + 2, column=0, sticky="w")
    ttk.Checkbutton(
        frame, text="Assign the optimized sections to the model",
        variable=apply_now,
    ).grid(row=row + 3, column=0, sticky="w", pady=(12, 2))
    ttk.Label(
        frame,
        text=("Leave unticked for a dry run: every member goes back on the section "
              "it started with and only the report is produced. The model is saved "
              "either way, because each analysis writes the section it is testing "
              "into the .STD file. The candidate sections the search tried stay in "
              "the model's property table."),
        foreground="#555555", wraplength=380, justify="left",
    ).grid(row=row + 4, column=0, sticky="w")

    actions = ttk.Frame(frame)
    actions.grid(row=row + 5, column=0, pady=(18, 0), sticky="e")
    ttk.Button(actions, text="Cancel", command=window.destroy).pack(
        side="left", padx=4)
    ttk.Button(actions, text="Optimize", command=submit).pack(side="left")

    window.protocol("WM_DELETE_WINDOW", window.destroy)
    if parent is None:
        window.mainloop()
    else:
        window.wait_window()
    return result[0] if result else None


def show_taper_optimization(result: OptimizationResult,
                            parent: Misc | None = None) -> None:
    """Show the before/after tapered section schedule for an optimizer run."""
    window = Toplevel(parent) if parent is not None else Tk()
    window.title("Tapered Section Optimization")
    window.geometry("1240x620")
    window.minsize(960, 460)
    if parent is not None:
        window.transient(parent)
    frame = ttk.Frame(window, padding=18)
    frame.pack(fill="both", expand=True)
    ttk.Label(frame, text="Tapered Section Optimization",
              font=("Segoe UI", 18, "bold")).pack(anchor="w")
    changed = sum(1 for change in result.changes if change.changed)
    state = "assigned to the model" if result.applied else "not assigned (dry run)"
    ttk.Label(frame, text=(f"{len(result.changes)} tapered member(s) • "
                           f"{changed} resized • {result.analyses_used} analysis run(s) • "
                           f"{state}"),
              foreground="#555555").pack(anchor="w", pady=(2, 14))

    table_frame = ttk.Frame(frame)
    table_frame.pack(fill="both", expand=True)
    columns = ("member", "length", "before", "after", "ratio",
               "weight_before", "weight_after", "saved")
    headings = ("Member", "Length (m)", "Existing section", "Optimized section",
                "Ratio", "Before (kg)", "After (kg)", "Saved (kg)")
    widths = (70, 85, 300, 300, 65, 95, 95, 95)
    numeric = {"length", "ratio", "weight_before", "weight_after", "saved"}
    table = ttk.Treeview(table_frame, columns=columns, show="headings")
    for column, heading, width in zip(columns, headings, widths):
        table.heading(column, text=heading)
        table.column(column, width=width, minwidth=50,
                     anchor="e" if column in numeric else "w")
    table.tag_configure("resized", background="#eef6ff")
    for change in result.changes:
        saved = change.weight_before_kg - change.weight_after_kg
        table.insert("", "end", values=(
            change.number, f"{change.length_m:.3f}",
            change.before.describe(), change.after.describe(),
            f"{change.ratio:.2f}" if change.ratio is not None else "—",
            f"{change.weight_before_kg:.1f}", f"{change.weight_after_kg:.1f}",
            f"{saved:.1f}",
        ), tags=("resized",) if change.changed else ())
    vertical = ttk.Scrollbar(table_frame, orient="vertical", command=table.yview)
    horizontal = ttk.Scrollbar(table_frame, orient="horizontal", command=table.xview)
    table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
    table.grid(row=0, column=0, sticky="nsew")
    vertical.grid(row=0, column=1, sticky="ns")
    horizontal.grid(row=1, column=0, sticky="ew")
    table_frame.rowconfigure(0, weight=1)
    table_frame.columnconfigure(0, weight=1)

    for note in result.notes:
        ttk.Label(frame, text=f"Note: {note}", foreground="#8a5a00",
                  wraplength=1150).pack(anchor="w", pady=(10, 0))
    ttk.Label(frame, text=(f"Steel weight {result.weight_before_kg:,.1f} kg → "
                           f"{result.weight_after_kg:,.1f} kg  "
                           f"(saved {result.saved_kg:,.1f} kg, "
                           f"{result.saved_percent:.1f}%)"),
              font=("Segoe UI", 10, "bold")).pack(side="left", pady=(12, 0))
    ttk.Button(frame, text="Close", command=window.destroy).pack(
        side="right", pady=(12, 0))
    if parent is None:
        window.mainloop()


def show_plate_summary(rows: list[PlateSummaryRow], selected_count: int,
                       parent: Misc | None = None) -> None:
    """Show the selected-member plate summary in a scrollable table."""
    window = Toplevel(parent) if parent is not None else Tk()
    window.title("Selected Member Plate Summary")
    window.geometry("1180x600")
    window.minsize(900, 440)
    if parent is not None:
        window.transient(parent)
    frame = ttk.Frame(window, padding=18)
    frame.pack(fill="both", expand=True)
    ttk.Label(frame, text="Selected Member Plate Summary",
              font=("Segoe UI", 18, "bold")).pack(anchor="w")
    plate_count = sum(1 for row in rows if row.category != "Whole section")
    ttk.Label(frame, text=(f"{selected_count} selected member(s) • "
                           f"{len(rows)} summary row(s) • "
                           f"{plate_count} fabricated plate row(s)"),
              foreground="#555555").pack(anchor="w", pady=(2, 14))
    table_frame = ttk.Frame(frame)
    table_frame.pack(fill="both", expand=True)
    columns = ("category", "description", "size", "members", "quantity",
               "each_length", "total_length", "area", "weight")
    headings = ("Item", "Description", "Section / plate size", "Members", "Qty",
                "Each length (m)", "Total length (m)", "Plate area (m²)",
                "Weight (kg)")
    widths = (105, 125, 210, 120, 55, 105, 110, 105, 100)
    numeric = {"quantity", "each_length", "total_length", "area", "weight"}
    table = ttk.Treeview(table_frame, columns=columns, show="headings")
    for column, heading, width in zip(columns, headings, widths):
        table.heading(column, text=heading)
        table.column(column, width=width, minwidth=50,
                     anchor="e" if column in numeric else "w")
    table.tag_configure("plate", background="#eef6ff")
    for row in rows:
        table.insert("", "end", values=(
            row.category, row.description, row.size,
            ", ".join(str(member) for member in row.members), row.quantity,
            f"{row.length_each_m:.3f}", f"{row.total_length_m:.3f}",
            f"{row.plate_area_m2:.3f}" if row.plate_area_m2 is not None else "—",
            f"{row.weight_kg:.1f}" if row.weight_kg is not None else "—",
        ), tags=("" if row.category == "Whole section" else "plate",))
    vertical = ttk.Scrollbar(table_frame, orient="vertical", command=table.yview)
    horizontal = ttk.Scrollbar(table_frame, orient="horizontal", command=table.xview)
    table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
    table.grid(row=0, column=0, sticky="nsew")
    vertical.grid(row=0, column=1, sticky="ns")
    horizontal.grid(row=1, column=0, sticky="ew")
    table_frame.rowconfigure(0, weight=1)
    table_frame.columnconfigure(0, weight=1)
    total_weight = sum(row.weight_kg or 0.0 for row in rows)
    ttk.Label(frame, text=f"Estimated steel weight: {total_weight:,.1f} kg",
              font=("Segoe UI", 10, "bold")).pack(side="left", pady=(12, 0))
    ttk.Button(frame, text="Close", command=window.destroy).pack(
        side="right", pady=(12, 0))
    if parent is None:
        window.mainloop()
