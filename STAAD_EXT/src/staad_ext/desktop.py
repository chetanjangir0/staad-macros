from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, ttk
from typing import Callable

from staad_ext.macros.frame_generator import (
    FrameParameters,
    build_model_in_openstaad,
    compute_frame_geometry,
    generate_std_file_content,
)
from staad_ext.macros.load_combinations import (
    LOAD_CATEGORIES,
    CombinationPreset,
    ComboRule,
    GeneratedCombo,
    PrimaryLoadCase,
    delete_custom_preset,
    fetch_primary_cases_from_openstaad,
    format_staad_combo_text,
    generate_combinations,
    load_all_presets,
    push_combos_to_openstaad,
    save_custom_preset,
)
from staad_ext.macros.plate_summary import PlateSummaryRow, selected_member_plate_summary
from staad_ext.macros.support_reactions import (
    SupportReaction,
    get_support_reactions,
    parse_load_cases,
    reaction_envelopes,
)
from staad_ext.macros.std_to_dxf import export_selected_members
from staad_ext.macros.std_to_ifc import export_structure
from staad_ext.models import ExportSettings, IfcExportSettings, ViewPlane
from staad_ext.openstaad import OpenStaad, OpenStaadError


@dataclass(frozen=True, slots=True)
class UtilityView:
    key: str
    title: str
    short_title: str
    description: str
    builder_name: str


UTILITY_VIEWS = (
    UtilityView(
        "frame_generator",
        "2D Frame Generator",
        "2D Frame Gen",
        "Generate 2D portal and gabled frame STAAD models with live canvas geometry preview.",
        "_build_frame_generator_view",
    ),
    UtilityView(
        "load_combinations",
        "Load Combinations",
        "Load Combos",
        "Auto-generate load combinations with preset creation, custom rules, and aggregate/separate options.",
        "_build_load_combinations_view",
    ),
    UtilityView(
        "std_to_dxf",
        "STD to DXF",
        "STD to DXF",
        "Export selected analytical members and section envelopes to AutoCAD.",
        "_build_dxf_view",
    ),
    UtilityView(
        "std_to_ifc",
        "STD to IFC (3D)",
        "STD to IFC",
        "Export the whole analytical model as a 3D IFC file for BIM viewers.",
        "_build_ifc_view",
    ),
    UtilityView(
        "plate_summary",
        "Plate Summary",
        "Plate Summary",
        "Create a fabrication summary from the members selected in STAAD.Pro.",
        "_build_plate_summary_view",
    ),
    UtilityView(
        "support_reactions",
        "Support Reactions",
        "Support Reactions",
        "Review detailed support reactions and min/max envelopes for chosen load combinations.",
        "_build_support_reactions_view",
    ),
)


class StaadExtApplication:
    """Single-window desktop shell for STAAD_EXT utilities."""

    BG = "#090e1a"
    SIDEBAR = "#0d1424"
    PANEL = "#111a2d"
    PANEL_ALT = "#162138"
    BORDER = "#263550"
    TEXT = "#f1f5f9"
    MUTED = "#94a3b8"
    ACCENT = "#3b82f6"
    ACCENT_HOVER = "#2563eb"
    SUCCESS = "#34d399"
    WARNING = "#fbbf24"
    ERROR = "#fb7185"

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("STAAD_EXT")
        self.root.geometry("1220x760")
        self.root.minsize(980, 620)
        self.root.configure(bg=self.BG)
        self._nav_buttons: dict[str, tk.Button] = {}
        self._current_view = "dashboard"
        self._status_text = tk.StringVar(value="Ready")
        self._status_kind = "muted"
        self._configure_styles()
        self._build_shell()
        self.show_view("dashboard")

        self.root.bind(
            "<Map>",
            lambda _event: self.root.after(50, self._enable_dark_title_bar),
            add="+",
        )
        self.root.after_idle(self._enable_dark_title_bar)

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(
            "Dark.Treeview",
            background=self.PANEL,
            fieldbackground=self.PANEL,
            foreground=self.TEXT,
            bordercolor=self.BORDER,
            lightcolor=self.BORDER,
            darkcolor=self.BORDER,
            rowheight=31,
            font=("Segoe UI", 9),
        )
        style.configure(
            "Dark.Treeview.Heading",
            background=self.PANEL_ALT,
            foreground=self.TEXT,
            relief="flat",
            padding=(8, 9),
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Dark.Treeview",
            background=[("selected", self.ACCENT)],
            foreground=[("selected", "#ffffff")],
        )
        style.map("Dark.Treeview.Heading", background=[("active", self.BORDER)])
        style.configure(
            "Dark.Vertical.TScrollbar",
            background=self.PANEL_ALT,
            troughcolor=self.BG,
            bordercolor=self.BG,
            arrowcolor=self.MUTED,
        )
        style.configure(
            "Dark.Horizontal.TScrollbar",
            background=self.PANEL_ALT,
            troughcolor=self.BG,
            bordercolor=self.BG,
            arrowcolor=self.MUTED,
        )
    def _enable_dark_title_bar(self) -> None:
        """Ask Windows DWM to render the native caption in immersive dark mode."""
        try:
            from ctypes import byref, c_int, c_void_p, sizeof, windll
            from ctypes import wintypes

            get_parent = windll.user32.GetParent
            get_parent.argtypes = [wintypes.HWND]
            get_parent.restype = wintypes.HWND
            hwnd = get_parent(self.root.winfo_id()) or self.root.winfo_id()

            set_attribute = windll.dwmapi.DwmSetWindowAttribute
            set_attribute.argtypes = [
                wintypes.HWND, wintypes.DWORD, c_void_p, wintypes.DWORD
            ]
            set_attribute.restype = wintypes.HRESULT
            enabled = c_int(1)
            # Attribute 20 is current; 19 supports earlier Windows 10 builds.
            for attribute in (20, 19):
                result = set_attribute(
                    hwnd, attribute, byref(enabled), sizeof(enabled)
                )
                if result == 0:

                    break
            def colorref(value: str) -> c_int:
                red = int(value[1:3], 16)
                green = int(value[3:5], 16)
                blue = int(value[5:7], 16)
                return c_int(red | (green << 8) | (blue << 16))

            # Windows 11 supports explicit native caption colors. These also
            # prevent the caption reverting to white after restore/maximize.
            for attribute, color in (
                (34, colorref(self.BORDER)),
                (35, colorref(self.SIDEBAR)),
                (36, colorref(self.TEXT)),
            ):
                set_attribute(hwnd, attribute, byref(color), sizeof(color))
        except (AttributeError, OSError):
            # Non-Windows systems and older DWM versions keep their native caption.
            pass


    def _build_shell(self) -> None:
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

        sidebar = tk.Frame(self.root, bg=self.SIDEBAR, width=238)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        content_column = tk.Frame(self.root, bg=self.BG)
        content_column.grid(row=0, column=1, sticky="nsew")
        content_column.grid_rowconfigure(0, weight=1)
        content_column.grid_columnconfigure(0, weight=1)

        brand = tk.Frame(sidebar, bg=self.SIDEBAR)
        brand.pack(fill="x", padx=22, pady=(25, 30))
        tk.Label(
            brand, text="S", bg=self.ACCENT, fg="white", width=3, height=1,
            font=("Segoe UI", 14, "bold"),
        ).pack(side="left")
        brand_text = tk.Frame(brand, bg=self.SIDEBAR)
        brand_text.pack(side="left", padx=(11, 0))
        tk.Label(
            brand_text, text="STAAD_EXT", bg=self.SIDEBAR, fg=self.TEXT,
            font=("Segoe UI", 14, "bold"),
        ).pack(anchor="w")
        tk.Label(
            brand_text, text="Engineering tools", bg=self.SIDEBAR, fg=self.MUTED,
            font=("Segoe UI", 8),
        ).pack(anchor="w")

        self._add_nav_button(sidebar, "dashboard", "Overview")
        tk.Label(
            sidebar, text="UTILITIES", bg=self.SIDEBAR, fg="#64748b",
            font=("Segoe UI", 8, "bold"),
        ).pack(fill="x", padx=24, pady=(22, 8), anchor="w")
        for index, utility in enumerate(UTILITY_VIEWS, start=1):
            self._add_nav_button(
                sidebar, utility.key, utility.short_title, f"{index:02d}"
            )

        footer = tk.Frame(sidebar, bg=self.SIDEBAR)
        footer.pack(side="bottom", fill="x", padx=20, pady=20)
        tk.Frame(footer, bg=self.BORDER, height=1).pack(fill="x", pady=(0, 15))
        tk.Label(
            footer, text="STAAD.Pro 2025", bg=self.SIDEBAR, fg=self.MUTED,
            font=("Segoe UI", 9),
        ).pack(anchor="w")
        tk.Label(
            footer, text="Select members before running a utility",
            bg=self.SIDEBAR, fg="#64748b", font=("Segoe UI", 8),
            wraplength=190, justify="left",
        ).pack(anchor="w", pady=(4, 0))

        self.content = tk.Frame(content_column, bg=self.BG)
        self.content.grid(row=0, column=0, sticky="nsew", padx=34, pady=(28, 18))
        self.status_bar = tk.Frame(content_column, bg=self.SIDEBAR, height=38)
        self.status_bar.grid(row=1, column=0, sticky="ew")
        self.status_dot = tk.Label(
            self.status_bar, text="●", bg=self.SIDEBAR, fg=self.MUTED,
            font=("Segoe UI", 8),
        )
        self.status_dot.pack(side="left", padx=(20, 8))
        tk.Label(
            self.status_bar, textvariable=self._status_text, bg=self.SIDEBAR,
            fg=self.MUTED, font=("Segoe UI", 9),
        ).pack(side="left")

    def _add_nav_button(
        self, parent: tk.Widget, key: str, label: str, number: str | None = None
    ) -> None:
        button = tk.Button(
            parent,
            text=f"  {number}     {label}" if number else f"           {label}",
            command=lambda selected=key: self.show_view(selected),
            bg=self.SIDEBAR,
            fg=self.MUTED,
            activebackground=self.PANEL_ALT,
            activeforeground=self.TEXT,
            relief="flat",
            bd=0,
            anchor="w",
            padx=16,
            pady=11,
            font=("Segoe UI", 10),
            cursor="hand2",
        )
        button.pack(fill="x", padx=12, pady=2)
        self._nav_buttons[key] = button

    def show_view(self, key: str) -> None:
        self._current_view = key
        for child in self.content.winfo_children():
            child.destroy()
        for nav_key, button in self._nav_buttons.items():
            active = nav_key == key
            button.configure(
                bg=self.PANEL_ALT if active else self.SIDEBAR,
                fg=self.TEXT if active else self.MUTED,
                font=("Segoe UI", 10, "bold" if active else "normal"),
            )
        if key == "dashboard":
            self._build_dashboard()
            return
        utility = next(item for item in UTILITY_VIEWS if item.key == key)
        builder: Callable[[UtilityView], None] = getattr(self, utility.builder_name)
        builder(utility)

    def _page_header(self, title: str, description: str) -> None:
        tk.Label(
            self.content, text=title, bg=self.BG, fg=self.TEXT,
            font=("Segoe UI", 24, "bold"),
        ).pack(anchor="w")
        tk.Label(
            self.content, text=description, bg=self.BG, fg=self.MUTED,
            font=("Segoe UI", 10), wraplength=850, justify="left",
        ).pack(anchor="w", pady=(5, 23))

    def _panel(self, parent: tk.Widget, padding: int = 22) -> tk.Frame:
        panel = tk.Frame(
            parent, bg=self.PANEL, highlightbackground=self.BORDER,
            highlightthickness=1, padx=padding, pady=padding,
        )
        return panel

    def _primary_button(
        self, parent: tk.Widget, text: str, command: Callable[[], None]
    ) -> tk.Button:
        return tk.Button(
            parent, text=text, command=command, bg=self.ACCENT, fg="white",
            activebackground=self.ACCENT_HOVER, activeforeground="white",
            relief="flat", bd=0, padx=18, pady=9,
            font=("Segoe UI", 9, "bold"), cursor="hand2",
        )

    def _secondary_button(
        self, parent: tk.Widget, text: str, command: Callable[[], None]
    ) -> tk.Button:
        return tk.Button(
            parent, text=text, command=command, bg=self.PANEL_ALT, fg=self.TEXT,
            activebackground=self.BORDER, activeforeground=self.TEXT,
            relief="flat", bd=0, padx=14, pady=8,
            font=("Segoe UI", 9), cursor="hand2",
        )

    def _build_dashboard(self) -> None:
        self._page_header(
            "Engineering workspace",
            "Run focused STAAD.Pro utilities from one consistent workspace.",
        )
        hero = self._panel(self.content, 25)
        hero.pack(fill="x", pady=(0, 18))
        tk.Label(
            hero, text="STAAD.Pro utilities, Made by Chetan Jangir out of the goodness of his heart.", bg=self.PANEL,
            fg=self.TEXT, font=("Segoe UI", 17, "bold"),
        ).pack(anchor="w")
        tk.Label(
            hero,
            text="(Bro deserves a nice promotion and a hefty salary)",
            bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 10),
            wraplength=760, justify="left",
        ).pack(anchor="w", pady=(8, 0))

        cards = tk.Frame(self.content, bg=self.BG)
        cards.pack(fill="both", expand=True)
        for column, utility in enumerate(UTILITY_VIEWS):
            cards.grid_columnconfigure(column, weight=1, uniform="utilities")
            card = self._panel(cards, 22)
            card.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 9,
                                                                9 if column == 0 else 0))
            tk.Label(
                card, text=f"{column + 1:02d}", bg=self.PANEL, fg=self.ACCENT,
                font=("Consolas", 10, "bold"),
            ).pack(anchor="w")
            tk.Label(
                card, text=utility.title, bg=self.PANEL, fg=self.TEXT,
                font=("Segoe UI", 15, "bold"),
            ).pack(anchor="w", pady=(12, 7))
            tk.Label(
                card, text=utility.description, bg=self.PANEL, fg=self.MUTED,
                font=("Segoe UI", 9), wraplength=330, justify="left",
            ).pack(anchor="w")
            self._secondary_button(
                card, "Open utility", lambda selected=utility.key: self.show_view(selected)
            ).pack(anchor="w", pady=(20, 0))

    def _build_dxf_view(self, utility: UtilityView) -> None:
        self._page_header(utility.title, utility.description)
        panel = self._panel(self.content, 24)
        panel.pack(fill="x")
        panel.grid_columnconfigure(0, weight=1)

        self.dxf_path = tk.StringVar(value=str(Path.cwd() / "STAAD_Model_Geometry_Sections.dxf"))
        self.dxf_plane = tk.StringVar(value=ViewPlane.XY.value)
        self.dxf_labels = tk.BooleanVar(value=True)
        self.dxf_peb_joins = tk.BooleanVar(value=False)
        self.dxf_connection_faces = tk.BooleanVar(value=False)
        self.dxf_color_by_section = tk.BooleanVar(value=True)
        self.dxf_scale = tk.StringVar(value="1.0")

        self._field_label(panel, "OUTPUT FILE", 0)
        path_row = tk.Frame(panel, bg=self.PANEL)
        path_row.grid(row=1, column=0, sticky="ew", pady=(7, 19))
        path_row.grid_columnconfigure(0, weight=1)
        self._dark_entry(path_row, self.dxf_path).grid(row=0, column=0, sticky="ew")
        self._secondary_button(path_row, "Browse", self._browse_dxf).grid(
            row=0, column=1, padx=(10, 0))

        options = tk.Frame(panel, bg=self.PANEL)
        options.grid(row=2, column=0, sticky="ew")
        options.grid_columnconfigure((0, 1), weight=1, uniform="options")
        plane_panel = tk.Frame(options, bg=self.PANEL_ALT, padx=16, pady=14)
        plane_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 7))
        tk.Label(plane_panel, text="PROJECTION PLANE", bg=self.PANEL_ALT,
                 fg=self.MUTED, font=("Segoe UI", 8, "bold")).pack(anchor="w")
        plane_row = tk.Frame(plane_panel, bg=self.PANEL_ALT)
        plane_row.pack(anchor="w", pady=(10, 0))
        for plane in ViewPlane:
            self._dark_radio(plane_row, plane.value, self.dxf_plane, plane.value).pack(
                side="left", padx=(0, 18))

        scale_panel = tk.Frame(options, bg=self.PANEL_ALT, padx=16, pady=14)
        scale_panel.grid(row=0, column=1, sticky="nsew", padx=(7, 0))
        tk.Label(scale_panel, text="TEXT SIZE SCALE", bg=self.PANEL_ALT,
                 fg=self.MUTED, font=("Segoe UI", 8, "bold")).pack(anchor="w")
        self._dark_entry(scale_panel, self.dxf_scale, width=14).pack(
            anchor="w", pady=(8, 0))

        checks = tk.Frame(panel, bg=self.PANEL)
        checks.grid(row=3, column=0, sticky="w", pady=(19, 21))
        self._dark_check(checks, "Write text labels", self.dxf_labels).pack(anchor="w")
        self._dark_check(checks, "Use PEB-style corner joins", self.dxf_peb_joins).pack(
            anchor="w", pady=5)
        self._dark_check(
            checks, "Show GA connection details", self.dxf_connection_faces
        ).pack(anchor="w", pady=(0, 5))
        self._dark_check(
            checks, "Color members by section size", self.dxf_color_by_section
        ).pack(anchor="w")
        self._primary_button(panel, "Export selected members", self._export_dxf).grid(
            row=4, column=0, sticky="w")

    def _field_label(self, parent: tk.Widget, text: str, row: int) -> None:
        tk.Label(parent, text=text, bg=self.PANEL, fg=self.MUTED,
                 font=("Segoe UI", 8, "bold")).grid(row=row, column=0, sticky="w")

    def _dark_entry(
        self, parent: tk.Widget, variable: tk.Variable, width: int | None = None
    ) -> tk.Entry:
        return tk.Entry(
            parent, textvariable=variable, width=width, bg="#0b1220", fg=self.TEXT,
            insertbackground=self.TEXT, selectbackground=self.ACCENT,
            relief="flat", bd=0, highlightthickness=1,
            highlightbackground=self.BORDER, highlightcolor=self.ACCENT,
            font=("Segoe UI", 10),
        )

    def _dark_check(
        self, parent: tk.Widget, text: str, variable: tk.BooleanVar
    ) -> tk.Checkbutton:
        return tk.Checkbutton(
            parent, text=text, variable=variable, bg=self.PANEL, fg=self.TEXT,
            activebackground=self.PANEL, activeforeground=self.TEXT,
            selectcolor=self.PANEL_ALT, font=("Segoe UI", 9),
        )

    def _dark_radio(
        self, parent: tk.Widget, text: str, variable: tk.StringVar, value: str
    ) -> tk.Radiobutton:
        return tk.Radiobutton(
            parent, text=text, variable=variable, value=value,
            bg=self.PANEL_ALT, fg=self.TEXT, activebackground=self.PANEL_ALT,
            activeforeground=self.TEXT, selectcolor=self.PANEL, font=("Segoe UI", 9),
        )

    def _browse_dxf(self) -> None:
        current = Path(self.dxf_path.get())
        selected = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save DXF",
            initialdir=current.parent,
            initialfile=current.name,
            defaultextension=".dxf",
            filetypes=(("DXF files", "*.dxf"),),
        )
        if selected:
            self.dxf_path.set(selected)

    def _export_dxf(self) -> None:
        self._set_status("Connecting to STAAD.Pro…", "muted")
        self.root.update_idletasks()
        try:
            output = Path(self.dxf_path.get().strip()).with_suffix(".dxf")
            if not output.name:
                raise ValueError("Choose an output file.")
            settings = ExportSettings(
                ViewPlane(self.dxf_plane.get()),
                self.dxf_labels.get(),
                float(self.dxf_scale.get()),
                self.dxf_peb_joins.get(),
                self.dxf_connection_faces.get(),
                self.dxf_color_by_section.get(),
            )
            count = export_selected_members(OpenStaad.connect(), output, settings)
            if count:
                self._set_status(
                    f"Exported {count} member(s) to {output}", "success"
                )
            else:
                self._set_status(
                    "No members exported. Select analytical members in STAAD.Pro and retry.",
                    "warning",
                )
        except (OpenStaadError, OSError, TypeError, ValueError) as exc:
            self._set_status(str(exc), "error")

    def _build_ifc_view(self, utility: UtilityView) -> None:
        self._page_header(utility.title, utility.description)
        panel = self._panel(self.content, 24)
        panel.pack(fill="x")
        panel.grid_columnconfigure(0, weight=1)

        self.ifc_path = tk.StringVar(value=str(Path.cwd() / "STAAD_Model_3D_Structure.ifc"))
        self.ifc_selected_only = tk.BooleanVar(value=False)

        self._field_label(panel, "OUTPUT FILE", 0)
        path_row = tk.Frame(panel, bg=self.PANEL)
        path_row.grid(row=1, column=0, sticky="ew", pady=(7, 19))
        path_row.grid_columnconfigure(0, weight=1)
        self._dark_entry(path_row, self.ifc_path).grid(row=0, column=0, sticky="ew")
        self._secondary_button(path_row, "Browse", self._browse_ifc).grid(
            row=0, column=1, padx=(10, 0))

        self._dark_check(
            panel, "Export selected members only (default: whole model)",
            self.ifc_selected_only,
        ).grid(row=2, column=0, sticky="w", pady=(0, 19))

        self._primary_button(panel, "Export structure to IFC", self._export_ifc).grid(
            row=3, column=0, sticky="w")

    def _browse_ifc(self) -> None:
        current = Path(self.ifc_path.get())
        selected = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save IFC",
            initialdir=current.parent,
            initialfile=current.name,
            defaultextension=".ifc",
            filetypes=(("IFC files", "*.ifc"),),
        )
        if selected:
            self.ifc_path.set(selected)

    def _export_ifc(self) -> None:
        self._set_status("Connecting to STAAD.Pro…", "muted")
        self.root.update_idletasks()
        try:
            output = Path(self.ifc_path.get().strip()).with_suffix(".ifc")
            if not output.name:
                raise ValueError("Choose an output file.")
            settings = IfcExportSettings(self.ifc_selected_only.get())
            count = export_structure(OpenStaad.connect(), output, settings)
            if count:
                self._set_status(
                    f"Exported {count} member(s) to {output}", "success"
                )
            else:
                self._set_status(
                    "No members exported. The model has no analytical beam members "
                    "(or none are selected).",
                    "warning",
                )
        except (OpenStaadError, OSError, TypeError, ValueError) as exc:
            self._set_status(str(exc), "error")

    def _build_plate_summary_view(self, utility: UtilityView) -> None:
        self._page_header(utility.title, utility.description)
        toolbar = tk.Frame(self.content, bg=self.BG)
        toolbar.pack(fill="x", pady=(0, 16))
        self._primary_button(toolbar, "Refresh from selection", self._load_plate_summary).pack(
            side="left")
        tk.Label(
            toolbar, text="Uses the analytical beam selection in the active model",
            bg=self.BG, fg=self.MUTED, font=("Segoe UI", 9),
        ).pack(side="left", padx=(14, 0))

        metrics = tk.Frame(self.content, bg=self.BG)
        metrics.pack(fill="x", pady=(0, 14))
        self.summary_metrics: dict[str, tk.Label] = {}
        for index, (key, title) in enumerate((
            ("members", "SELECTED MEMBERS"),
            ("rows", "SUMMARY ROWS"),
            ("weight", "EST. STEEL WEIGHT"),
        )):
            metrics.grid_columnconfigure(index, weight=1, uniform="metric")
            card = self._panel(metrics, 15)
            card.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 7,
                                                              0 if index == 2 else 7))
            tk.Label(card, text=title, bg=self.PANEL, fg=self.MUTED,
                     font=("Segoe UI", 8, "bold")).pack(anchor="w")
            value = tk.Label(card, text="—", bg=self.PANEL, fg=self.TEXT,
                             font=("Segoe UI", 17, "bold"))
            value.pack(anchor="w", pady=(5, 0))
            self.summary_metrics[key] = value

        table_panel = self._panel(self.content, 0)
        table_panel.pack(fill="both", expand=True)
        table_panel.grid_rowconfigure(0, weight=1)
        table_panel.grid_columnconfigure(0, weight=1)
        columns = ("category", "description", "size", "members", "quantity",
                   "each_length", "total_length", "area", "weight")
        headings = ("Item", "Description", "Section / plate size", "Members", "Qty",
                    "Each length (m)", "Total length (m)", "Plate area (m²)",
                    "Weight (kg)")
        widths = (105, 120, 205, 110, 50, 105, 110, 100, 95)
        numeric = {"quantity", "each_length", "total_length", "area", "weight"}
        self.summary_table = ttk.Treeview(
            table_panel, columns=columns, show="headings", style="Dark.Treeview"
        )
        for column, heading, width in zip(columns, headings, widths):
            self.summary_table.heading(column, text=heading)
            self.summary_table.column(
                column, width=width, minwidth=45,
                anchor="e" if column in numeric else "w",
            )
        self.summary_table.tag_configure("plate", background="#12243d")
        vertical = ttk.Scrollbar(
            table_panel, orient="vertical", command=self.summary_table.yview,
            style="Dark.Vertical.TScrollbar",
        )
        horizontal = ttk.Scrollbar(
            table_panel, orient="horizontal", command=self.summary_table.xview,
            style="Dark.Horizontal.TScrollbar",
        )
        self.summary_table.configure(
            yscrollcommand=vertical.set, xscrollcommand=horizontal.set
        )
        self.summary_table.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")

    def _load_plate_summary(self) -> None:
        self._set_status("Reading selected members from STAAD.Pro…", "muted")
        self.root.update_idletasks()
        try:
            staad = OpenStaad.connect()
            rows = selected_member_plate_summary(staad)
            if not rows:
                self._set_status(
                    "No analytical members selected in STAAD.Pro.", "warning"
                )
                return
            self._render_summary_rows(rows)
            members = {member for row in rows for member in row.members}
            total_weight = sum(row.weight_kg or 0.0 for row in rows)
            self.summary_metrics["members"].configure(text=f"{len(members):,}")
            self.summary_metrics["rows"].configure(text=f"{len(rows):,}")
            self.summary_metrics["weight"].configure(text=f"{total_weight:,.1f} kg")
            self._set_status(
                f"Plate summary refreshed for {len(members)} member(s).", "success"
            )
        except (OpenStaadError, OSError, TypeError, ValueError) as exc:
            self._set_status(str(exc), "error")

    def _render_summary_rows(self, rows: list[PlateSummaryRow]) -> None:
        self.summary_table.delete(*self.summary_table.get_children())
        for row in rows:
            self.summary_table.insert(
                "", "end",
                values=(
                    row.category, row.description, row.size,
                    ", ".join(str(member) for member in row.members), row.quantity,
                    f"{row.length_each_m:.3f}", f"{row.total_length_m:.3f}",
                    f"{row.plate_area_m2:.3f}" if row.plate_area_m2 is not None else "—",
                    f"{row.weight_kg:.1f}" if row.weight_kg is not None else "—",
                ),
                tags=("plate",) if row.category != "Whole section" else (),
            )

    def _build_support_reactions_view(self, utility: UtilityView) -> None:
        self._page_header(utility.title, utility.description)
        controls = self._panel(self.content, 16)
        controls.pack(fill="x", pady=(0, 14))
        controls.grid_columnconfigure(0, weight=1)
        self.reaction_load_cases = tk.StringVar()
        tk.Label(controls, text="LOAD COMBINATION NUMBERS", bg=self.PANEL,
                 fg=self.MUTED, font=("Segoe UI", 8, "bold")).grid(
            row=0, column=0, sticky="w")
        self._dark_entry(controls, self.reaction_load_cases).grid(
            row=1, column=0, sticky="ew", pady=(7, 0))
        self._primary_button(controls, "Get reactions",
                             self._load_support_reactions).grid(
            row=1, column=1, padx=(12, 0), pady=(7, 0))
        tk.Label(
            controls,
            text="Examples: 101, 102, 105 or 101-105. Results use global axes.",
            bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8),
        ).grid(row=2, column=0, sticky="w", pady=(7, 0))

        results = tk.PanedWindow(self.content, orient="vertical", bg=self.BG,
                                 sashwidth=7, sashrelief="flat", bd=0)
        results.pack(fill="both", expand=True)
        detail_panel = self._panel(results, 0)
        envelope_panel = self._panel(results, 0)
        results.add(detail_panel, minsize=180, stretch="always")
        results.add(envelope_panel, minsize=150, stretch="always")
        self.reaction_table = self._reaction_tree(
            detail_panel, "DETAILED REACTIONS",
            ("node", "case", "fx", "fy", "fz", "mx", "my", "mz"),
            ("Node", "Load combination", "FX", "FY", "FZ", "MX", "MY", "MZ"),
        )
        self.reaction_envelope_table = self._reaction_tree(
            envelope_panel, "MIN / MAX SUMMARY",
            ("component", "minimum", "min_node", "min_case",
             "maximum", "max_node", "max_case"),
            ("Component", "Minimum", "Node", "Load combination",
             "Maximum", "Node", "Load combination"),
        )

    def _reaction_tree(self, parent: tk.Frame, title: str,
                       columns: tuple[str, ...], headings: tuple[str, ...]
                       ) -> ttk.Treeview:
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        tk.Label(parent, text=title, bg=self.PANEL, fg=self.MUTED,
                 font=("Segoe UI", 8, "bold"), padx=12, pady=8).grid(
            row=0, column=0, sticky="w")
        table = ttk.Treeview(parent, columns=columns, show="headings",
                             style="Dark.Treeview")
        for column, heading in zip(columns, headings):
            table.heading(column, text=heading)
            table.column(column, width=120, minwidth=55,
                         anchor="w" if column == "component" else "e")
        vertical = ttk.Scrollbar(parent, orient="vertical", command=table.yview,
                                 style="Dark.Vertical.TScrollbar")
        horizontal = ttk.Scrollbar(parent, orient="horizontal", command=table.xview,
                                   style="Dark.Horizontal.TScrollbar")
        table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        table.grid(row=1, column=0, sticky="nsew")
        vertical.grid(row=1, column=1, sticky="ns")
        horizontal.grid(row=2, column=0, sticky="ew")
        return table

    def _load_support_reactions(self) -> None:
        self._set_status("Reading support reactions from STAAD.Pro…", "muted")
        self.root.update_idletasks()
        try:
            load_cases = parse_load_cases(self.reaction_load_cases.get())
            rows = get_support_reactions(OpenStaad.connect(), load_cases)
            self._render_support_reactions(rows)
            node_count = len({row.node for row in rows})
            self._set_status(
                f"Loaded {len(rows):,} reactions for {node_count} support(s) "
                f"and {len(load_cases)} load combination(s).", "success")
        except (OpenStaadError, OSError, TypeError, ValueError) as exc:
            self._set_status(str(exc), "error")

    def _render_support_reactions(self, rows: list[SupportReaction]) -> None:
        self.reaction_table.delete(*self.reaction_table.get_children())
        self.reaction_envelope_table.delete(
            *self.reaction_envelope_table.get_children())
        for row in rows:
            self.reaction_table.insert(
                "", "end", values=(row.node, row.load_case, *(
                    f"{value:,.3f}" for value in row.values)))
        for row in reaction_envelopes(rows):
            self.reaction_envelope_table.insert(
                "", "end", values=(
                    row.component, f"{row.minimum:,.3f}", row.minimum_node,
                    row.minimum_load_case, f"{row.maximum:,.3f}",
                    row.maximum_node, row.maximum_load_case))

    def _build_frame_generator_view(self, utility: UtilityView) -> None:
        self._page_header(utility.title, utility.description)

        container = tk.Frame(self.content, bg=self.BG)
        container.pack(fill="both", expand=True)
        container.grid_columnconfigure(0, weight=4)
        container.grid_columnconfigure(1, weight=5)
        container.grid_rowconfigure(0, weight=1)

        left_panel = self._panel(container, 14)
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        form_canvas = tk.Canvas(left_panel, bg=self.PANEL, highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(left_panel, orient="vertical", command=form_canvas.yview, style="Dark.Vertical.TScrollbar")
        form_frame = tk.Frame(form_canvas, bg=self.PANEL)

        form_frame.bind(
            "<Configure>",
            lambda e: form_canvas.configure(scrollregion=form_canvas.bbox("all"))
        )
        form_window = form_canvas.create_window((0, 0), window=form_frame, anchor="nw")
        form_canvas.bind("<Configure>", lambda e: form_canvas.itemconfig(form_window, width=e.width))
        form_canvas.configure(yscrollcommand=scrollbar.set)

        form_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.fg_width = tk.StringVar(value="20.0")
        self.fg_eave_height = tk.StringVar(value="7.0")
        self.fg_ridge_distance = tk.StringVar(value="10.0")
        self.fg_slope = tk.StringVar(value="5.0")
        self.fg_brick_wall_height = tk.StringVar(value="0.0")
        self.fg_col_mode = tk.StringVar(value="count")
        self.fg_col_input = tk.StringVar(value="1")
        self.fg_mezzanine_enabled = tk.BooleanVar(value=False)
        self.fg_mezzanine_height = tk.StringVar(value="3.5")
        self.fg_mezzanine_start_x = tk.StringVar(value="0.0")
        self.fg_mezzanine_end_x = tk.StringVar(value="10.0")
        self.fg_bay_spacing = tk.StringVar(value="6.0")
        self.fg_left_support = tk.StringVar(value="Fixed")
        self.fg_right_support = tk.StringVar(value="Fixed")
        self.fg_int_support = tk.StringVar(value="Fixed")
        self.fg_basic_wind_speed = tk.StringVar(value="39.0")
        self.fg_wind_standard = tk.StringVar(value="IS 875 Part 3")
        self.fg_wind_building_length = tk.StringVar(value="30.0")
        self.fg_wind_design_life = tk.StringVar(value="50")
        self.fg_wind_terrain_category = tk.StringVar(value="2")
        self.fg_wind_opening = tk.StringVar(value="<5%")
        self.fg_seismic_zone = tk.StringVar(value="Zone III (0.16)")
        self.fg_response_reduction_factor = tk.StringVar(value="5")
        self.fg_importance_factor = tk.StringVar(value="1")
        self.fg_soil_type = tk.StringVar(value="Medium")
        self.fg_dead_load = tk.StringVar(value="0.15")
        self.fg_roof_live_load = tk.StringVar(value="0.75")
        self.fg_collateral_load = tk.StringVar(value="0.10")
        self.fg_mezzanine_live_load = tk.StringVar(value="3.0")
        self.fg_mezzanine_dead_load = tk.StringVar(value="1.5")
        self.fg_design_code = tk.StringVar(value="IS 800:2007")

        all_vars: list[tk.Variable] = [
            self.fg_width, self.fg_eave_height, self.fg_ridge_distance, self.fg_slope,
            self.fg_brick_wall_height, self.fg_col_mode, self.fg_col_input,
            self.fg_mezzanine_enabled, self.fg_mezzanine_height, self.fg_mezzanine_start_x,
            self.fg_mezzanine_end_x, self.fg_bay_spacing, self.fg_left_support,
            self.fg_right_support, self.fg_int_support, self.fg_basic_wind_speed,
            self.fg_seismic_zone, self.fg_response_reduction_factor, self.fg_importance_factor,
            self.fg_soil_type, self.fg_dead_load, self.fg_roof_live_load,
            self.fg_collateral_load, self.fg_mezzanine_live_load, self.fg_mezzanine_dead_load,
            self.fg_design_code
        ]
        for v in all_vars:
            v.trace_add("write", lambda *_: self._redraw_frame_canvas())

        row = 0
        self._section_header(form_frame, "FRAME GEOMETRY", row)
        row += 1
        row = self._form_row(form_frame, "Width (m):", self.fg_width, row)
        row = self._form_row(form_frame, "Eave Height (m):", self.fg_eave_height, row)
        row = self._form_row(form_frame, "Ridge Distance (m):", self.fg_ridge_distance, row)
        row = self._form_row(form_frame, "Roof Slope (1:x):", self.fg_slope, row)
        row = self._form_row(form_frame, "Brick Wall Ht (m):", self.fg_brick_wall_height, row)

        self._section_header(form_frame, "INTERIOR COLUMNS", row)
        row += 1
        mode_frame = tk.Frame(form_frame, bg=self.PANEL)
        mode_frame.grid(row=row, column=0, columnspan=2, sticky="w", pady=(2, 6))
        self._dark_radio(mode_frame, "By Count", self.fg_col_mode, "count").pack(side="left", padx=(0, 15))
        self._dark_radio(mode_frame, "By Spacing (e.g. 5, 5)", self.fg_col_mode, "spacing").pack(side="left")
        row += 1
        row = self._form_row(form_frame, "Count / Spacing:", self.fg_col_input, row)

        self._section_header(form_frame, "MEZZANINE FLOOR", row)
        row += 1
        self._dark_check(form_frame, "Enable Mezzanine Floor", self.fg_mezzanine_enabled).grid(
            row=row, column=0, columnspan=2, sticky="w", pady=(2, 6)
        )
        row += 1
        row = self._form_row(form_frame, "Mezzanine Height (m):", self.fg_mezzanine_height, row)
        row = self._form_row(form_frame, "Start X (m):", self.fg_mezzanine_start_x, row)
        row = self._form_row(form_frame, "End X (m):", self.fg_mezzanine_end_x, row)

        self._section_header(form_frame, "SUPPORTS & BAY SPACING", row)
        row += 1
        row = self._form_row(form_frame, "Bay Spacing (m):", self.fg_bay_spacing, row)
        row = self._form_choice_row(form_frame, "Left Support:", self.fg_left_support, ["Fixed", "Pinned"], row)
        row = self._form_choice_row(form_frame, "Right Support:", self.fg_right_support, ["Fixed", "Pinned"], row)
        row = self._form_choice_row(form_frame, "Interior Support:", self.fg_int_support, ["Fixed", "Pinned"], row)

        self._section_header(form_frame, "LOADS & ENVIRONMENT", row)
        row += 1
        row = self._form_choice_row(
            form_frame, "Seismic Zone:", self.fg_seismic_zone,
            ["Zone II (0.10)", "Zone III (0.16)", "Zone IV (0.24)", "Zone V (0.36)"], row
        )
        row = self._form_row(form_frame, "Response Reduction Factor (RF):", self.fg_response_reduction_factor, row)
        row = self._form_row(form_frame, "Importance Factor (I):", self.fg_importance_factor, row)
        row = self._form_choice_row(
            form_frame, "Soil Type:", self.fg_soil_type,
            ["Hard", "Medium", "Soft"], row
        )
        row = self._form_row(form_frame, "Dead Load (kN/m²):", self.fg_dead_load, row)
        row = self._form_row(form_frame, "Roof Live Load (kN/m²):", self.fg_roof_live_load, row)
        row = self._form_row(form_frame, "Collateral Load (kN/m²):", self.fg_collateral_load, row)
        row = self._form_row(form_frame, "Mezzanine Live (kN/m²):", self.fg_mezzanine_live_load, row)
        row = self._form_row(form_frame, "Mezzanine Dead (kN/m²):", self.fg_mezzanine_dead_load, row)

        self._section_header(form_frame, "WIND LOAD", row)
        row += 1
        row = self._form_choice_row(
            form_frame, "Wind Code:", self.fg_wind_standard,
            ["IS 875 Part 3", "MBMA 12"], row
        )
        self.fg_wind_frame = tk.Frame(form_frame, bg=self.PANEL)
        self.fg_wind_frame.grid(row=row, column=0, columnspan=2, sticky="ew")
        row += 1
        wrow = 0
        wrow = self._form_row(self.fg_wind_frame, "Length (m):", self.fg_wind_building_length, wrow)
        wrow = self._form_choice_row(
            self.fg_wind_frame, "Basic Wind Speed (m/s):", self.fg_basic_wind_speed,
            ["33.0", "39.0", "44.0", "47.0", "50.0", "55.0"], wrow
        )
        wrow = self._form_choice_row(
            self.fg_wind_frame, "Design Life (years):", self.fg_wind_design_life,
            ["5", "25", "50", "100"], wrow
        )
        wrow = self._form_choice_row(
            self.fg_wind_frame, "Terrain Category:", self.fg_wind_terrain_category,
            ["1", "2", "3", "4"], wrow
        )
        wrow = self._form_choice_row(
            self.fg_wind_frame, "Openings:", self.fg_wind_opening,
            ["<5%", "5-20%", ">20%"], wrow
        )
        tk.Label(
            self.fg_wind_frame,
            text="Width, height, roof slope, and bay spacing are taken from Frame Geometry above.",
            bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8), wraplength=280, justify="left",
        ).grid(row=wrow, column=0, columnspan=2, sticky="w", pady=(4, 0))

        def _sync_wind_frame(*_args: object) -> None:
            if self.fg_wind_standard.get() == "IS 875 Part 3":
                self.fg_wind_frame.grid()
            else:
                self.fg_wind_frame.grid_remove()

        self.fg_wind_standard.trace_add("write", _sync_wind_frame)
        _sync_wind_frame()

        self._section_header(form_frame, "DESIGN CODE", row)
        row += 1
        row = self._form_choice_row(
            form_frame, "Steel Design Code:", self.fg_design_code,
            ["IS 800:2007", "AISC 360-16 (LRFD)", "AISC 360-16 (ASD)"], row
        )

        right_panel = self._panel(container, 18)
        right_panel.grid(row=0, column=1, sticky="nsew")
        right_panel.grid_rowconfigure(1, weight=1)
        right_panel.grid_columnconfigure(0, weight=1)

        header_right = tk.Frame(right_panel, bg=self.PANEL)
        header_right.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        tk.Label(
            header_right, text="2D GEOMETRY PREVIEW", bg=self.PANEL, fg=self.TEXT,
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left")
        tk.Label(
            header_right, text="● Real-time sync", bg=self.PANEL, fg=self.SUCCESS,
            font=("Segoe UI", 8, "bold"),
        ).pack(side="right")

        self.fg_canvas = tk.Canvas(
            right_panel, bg="#0b1220", highlightbackground=self.BORDER,
            highlightthickness=1, bd=0
        )
        self.fg_canvas.grid(row=1, column=0, sticky="nsew", pady=(0, 14))
        self.fg_canvas.bind("<Configure>", lambda _e: self._redraw_frame_canvas())

        self.fg_load_info = tk.Label(
            right_panel, text="", bg=self.PANEL_ALT, fg=self.TEXT,
            font=("Consolas", 8), padx=10, pady=6, anchor="w", justify="left"
        )
        self.fg_load_info.grid(row=2, column=0, sticky="ew", pady=(0, 14))

        action_frame = tk.Frame(right_panel, bg=self.PANEL)
        action_frame.grid(row=3, column=0, sticky="ew")
        self._primary_button(action_frame, "Send to Active STAAD.Pro", self._send_frame_to_openstaad).pack(
            side="left", padx=(0, 10)
        )
        self._secondary_button(action_frame, "Save .STD File...", self._save_frame_std_file).pack(
            side="left"
        )

        self._redraw_frame_canvas()

    def _section_header(self, parent: tk.Widget, text: str, row: int) -> None:
        tk.Label(
            parent, text=text, bg=self.PANEL, fg=self.ACCENT,
            font=("Segoe UI", 8, "bold"),
        ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(14, 6))

    def _form_row(self, parent: tk.Widget, label_text: str, variable: tk.Variable, row: int) -> int:
        tk.Label(
            parent, text=label_text, bg=self.PANEL, fg=self.MUTED,
            font=("Segoe UI", 9),
        ).grid(row=row, column=0, sticky="w", pady=3)
        entry = self._dark_entry(parent, variable, width=16)
        entry.grid(row=row, column=1, sticky="e", pady=3, padx=(10, 0))
        entry.bind("<KeyRelease>", lambda _e: self._redraw_frame_canvas())
        return row + 1

    def _form_choice_row(self, parent: tk.Widget, label_text: str, variable: tk.StringVar, options: list[str], row: int) -> int:
        tk.Label(
            parent, text=label_text, bg=self.PANEL, fg=self.MUTED,
            font=("Segoe UI", 9),
        ).grid(row=row, column=0, sticky="w", pady=3)
        combo = ttk.Combobox(parent, textvariable=variable, values=options, state="readonly", width=16)
        combo.grid(row=row, column=1, sticky="e", pady=3, padx=(10, 0))
        combo.bind("<<ComboboxSelected>>", lambda _e: self._redraw_frame_canvas())
        return row + 1

    def _get_frame_params(self) -> FrameParameters:
        return FrameParameters(
            width=float(self.fg_width.get().strip()),
            eave_height=float(self.fg_eave_height.get().strip()),
            ridge_distance=float(self.fg_ridge_distance.get().strip()),
            slope=float(self.fg_slope.get().strip()),
            col_mode=self.fg_col_mode.get(),
            col_input=self.fg_col_input.get().strip(),
            brick_wall_height=float(self.fg_brick_wall_height.get().strip() or "0"),
            mezzanine_enabled=self.fg_mezzanine_enabled.get(),
            mezzanine_height=float(self.fg_mezzanine_height.get().strip() or "0"),
            mezzanine_start_x=float(self.fg_mezzanine_start_x.get().strip() or "0"),
            mezzanine_end_x=float(self.fg_mezzanine_end_x.get().strip() or "0"),
            bay_spacing=float(self.fg_bay_spacing.get().strip() or "6"),
            left_support=self.fg_left_support.get(),
            right_support=self.fg_right_support.get(),
            int_support=self.fg_int_support.get(),
            basic_wind_speed=float(self.fg_basic_wind_speed.get().strip() or "39"),
            wind_standard=self.fg_wind_standard.get(),
            wind_building_length=float(self.fg_wind_building_length.get().strip() or "30"),
            wind_design_life=int(self.fg_wind_design_life.get().strip() or "50"),
            wind_terrain_category=int(self.fg_wind_terrain_category.get().strip() or "2"),
            wind_opening=self.fg_wind_opening.get(),
            seismic_zone=self.fg_seismic_zone.get(),
            response_reduction_factor=float(self.fg_response_reduction_factor.get().strip() or "5"),
            importance_factor=float(self.fg_importance_factor.get().strip() or "1"),
            soil_type=self.fg_soil_type.get(),
            dead_load=float(self.fg_dead_load.get().strip() or "0"),
            roof_live_load=float(self.fg_roof_live_load.get().strip() or "0"),
            collateral_load=float(self.fg_collateral_load.get().strip() or "0"),
            mezzanine_live_load=float(self.fg_mezzanine_live_load.get().strip() or "0"),
            mezzanine_dead_load=float(self.fg_mezzanine_dead_load.get().strip() or "0"),
            design_code=self.fg_design_code.get(),
        )

    def _redraw_frame_canvas(self) -> None:
        if not hasattr(self, "fg_canvas"):
            return
        c = self.fg_canvas
        c.delete("all")

        try:
            params = self._get_frame_params()
            geom = compute_frame_geometry(params)
        except Exception as exc:
            c.create_text(
                20, 20, text=f"Preview unavailable: {exc}",
                anchor="nw", fill=self.ERROR, font=("Segoe UI", 9, "bold")
            )
            return

        cw = c.winfo_width() or 480
        ch = c.winfo_height() or 360
        margin = 45

        max_x = params.width
        max_y = max(geom.ridge_height, params.eave_height) * 1.15

        if max_x <= 0 or max_y <= 0:
            return

        scale_x = (cw - 2 * margin) / max_x
        scale_y = (ch - 2 * margin) / max_y
        scale = min(scale_x, scale_y)

        ox = margin + ((cw - 2 * margin) - max_x * scale) / 2
        oy = ch - margin

        def to_c(x: float, y: float) -> tuple[float, float]:
            return (ox + x * scale, oy - y * scale)

        gx1, gy1 = to_c(-params.width * 0.05, 0)
        gx2, gy2 = to_c(params.width * 1.05, 0)
        c.create_line(gx1, gy1, gx2, gy2, fill="#334155", width=2)
        for hx in range(int(gx1), int(gx2), 15):
            c.create_line(hx, gy1, hx - 6, gy1 + 8, fill="#1e293b", width=1)

        node_map = {n.id: n for n in geom.nodes}
        for m in geom.members:
            n1 = node_map[m.start_node]
            n2 = node_map[m.end_node]
            x1, y1 = to_c(n1.x, n1.y)
            x2, y2 = to_c(n2.x, n2.y)

            if m.group_type == "outer_column":
                c.create_line(x1, y1, x2, y2, fill="#3b82f6", width=3)
            elif m.group_type == "rafter":
                c.create_line(x1, y1, x2, y2, fill="#38bdf8", width=3)
            elif m.group_type == "interior_column":
                c.create_line(x1, y1, x2, y2, fill="#818cf8", width=2, dash=(4, 2))
            elif m.group_type == "mezzanine":
                c.create_line(x1, y1, x2, y2, fill="#f59e0b", width=3)

        if params.brick_wall_height > 0.001:
            bw_y = params.brick_wall_height
            bx1, by1 = to_c(0, bw_y)
            bx2, by2 = to_c(params.width, bw_y)
            c.create_line(bx1 - 5, by1, bx1 + 5, by1, fill="#f43f5e", width=2)
            c.create_line(bx2 - 5, by2, bx2 + 5, by2, fill="#f43f5e", width=2)
            c.create_text(bx1 - 12, by1, text=f"BW {bw_y}m", fill="#f43f5e", font=("Segoe UI", 7), anchor="e")

        for n in geom.nodes:
            nx, ny = to_c(n.x, n.y)
            color = "#38bdf8" if n.tag == "ridge" else ("#f59e0b" if n.tag == "mezzanine" else "#64748b")
            r = 3.5 if n.tag in ("ridge", "eave", "base") else 2.5
            c.create_oval(nx - r, ny - r, nx + r, ny + r, fill=color, outline="#0b1220")

        def draw_support(nid: int, supp_type: str):
            n = node_map[nid]
            sx, sy = to_c(n.x, n.y)
            if supp_type == "Pinned":
                c.create_polygon(sx, sy, sx - 8, sy + 12, sx + 8, sy + 12, fill="", outline="#fbbf24", width=2)
            else:
                c.create_rectangle(sx - 10, sy, sx + 10, sy + 6, fill="#fbbf24", outline="")

        draw_support(geom.left_support_node, params.left_support)
        draw_support(geom.right_support_node, params.right_support)
        for inid in geom.interior_support_nodes:
            draw_support(inid, params.int_support)

        dim_y = oy + 25
        cx0, _ = to_c(0, 0)
        cxw, _ = to_c(params.width, 0)
        c.create_line(cx0, dim_y, cxw, dim_y, fill=self.MUTED, arrow="both", arrowshape=(6, 8, 3))
        c.create_text((cx0 + cxw) / 2, dim_y + 12, text=f"Width = {params.width:.1f} m", fill=self.TEXT, font=("Segoe UI", 8))

        dim_x_l = ox - 25
        _, cy_eave = to_c(0, params.eave_height)
        c.create_line(dim_x_l, oy, dim_x_l, cy_eave, fill=self.MUTED, arrow="both", arrowshape=(6, 8, 3))
        c.create_text(dim_x_l - 6, (oy + cy_eave) / 2, text=f"Eave {params.eave_height:.1f}m", fill=self.TEXT, font=("Segoe UI", 8), anchor="e")

        cx_r, cy_r = to_c(params.ridge_distance, geom.ridge_height)
        c.create_text(cx_r, cy_r - 12, text=f"Ridge {geom.ridge_height:.2f}m", fill="#38bdf8", font=("Segoe UI", 8, "bold"))

        w_dl = params.dead_load * params.bay_spacing
        w_rll = params.roof_live_load * params.bay_spacing
        w_cl = params.collateral_load * params.bay_spacing
        info = f"LOAD CONVERSIONS (Bay Spacing = {params.bay_spacing:.1f} m):\n"
        info += f"• Rafter Dead Load: {w_dl:.2f} kN/m  |  Rafter Roof Live Load: {w_rll:.2f} kN/m  |  Collateral Load: {w_cl:.2f} kN/m\n"
        if params.mezzanine_enabled:
            w_mll = params.mezzanine_live_load * params.bay_spacing
            w_mdl = params.mezzanine_dead_load * params.bay_spacing
            info += f"• Mezzanine Live Load: {w_mll:.2f} kN/m  |  Mezzanine Dead Load: {w_mdl:.2f} kN/m"
        else:
            info += "• Mezzanine: Disabled"

        self.fg_load_info.configure(text=info)

    def _send_frame_to_openstaad(self) -> None:
        self._set_status("Connecting to STAAD.Pro and constructing 2D Frame…", "muted")
        self.root.update_idletasks()
        try:
            params = self._get_frame_params()
            staad = OpenStaad.connect()
            build_model_in_openstaad(staad, params)
            self._set_status("2D Frame successfully constructed in active STAAD.Pro model!", "success")
        except (OpenStaadError, OSError, TypeError, ValueError) as exc:
            self._set_status(str(exc), "error")

    def _save_frame_std_file(self) -> None:
        try:
            params = self._get_frame_params()
            std_content = generate_std_file_content(params)

            default_filename = f"2D_Frame_{params.width}x{params.eave_height}m.std"
            selected = filedialog.asksaveasfilename(
                parent=self.root,
                title="Save STAAD.Pro Model (.STD)",
                initialdir=Path.cwd(),
                initialfile=default_filename,
                defaultextension=".std",
                filetypes=(("STAAD Files", "*.std"), ("All Files", "*.*")),
            )
            if selected:
                Path(selected).write_text(std_content, encoding="utf-8")
                self._set_status(f"Saved STAAD file to {selected}", "success")
        except (OSError, TypeError, ValueError) as exc:
            self._set_status(str(exc), "error")

    def _build_load_combinations_view(self, utility: UtilityView) -> None:
        self._page_header(utility.title, utility.description)

        toolbar = tk.Frame(self.content, bg=self.BG)
        toolbar.pack(fill="x", pady=(0, 14))

        self._primary_button(toolbar, "Fetch Load Cases from STAAD", self._fetch_lc_from_staad).pack(side="left")
        self._secondary_button(toolbar, "+ Add Primary Case Row", self._add_manual_lc_row).pack(side="left", padx=(10, 0))

        controls_panel = self._panel(self.content, 14)
        controls_panel.pack(fill="x", pady=(0, 14))
        controls_panel.grid_columnconfigure(1, weight=1)

        tk.Label(controls_panel, text="PRESET:", bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8, "bold")).grid(
            row=0, column=0, sticky="w"
        )

        self.lc_preset_name = tk.StringVar()
        self.lc_presets_list = load_all_presets()
        preset_names = [p.name for p in self.lc_presets_list]
        self.lc_preset_combo = ttk.Combobox(
            controls_panel, textvariable=self.lc_preset_name, values=preset_names, state="readonly", width=32
        )
        if preset_names:
            self.lc_preset_name.set(preset_names[0])
        self.lc_preset_combo.grid(row=0, column=1, sticky="w", padx=(10, 10))
        self.lc_preset_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_preset_changed())

        preset_btn_frame = tk.Frame(controls_panel, bg=self.PANEL)
        preset_btn_frame.grid(row=0, column=2, sticky="e")
        self._secondary_button(preset_btn_frame, "Edit / Create Factor Preset...", self._show_edit_preset_factors_dialog).pack(side="left", padx=(0, 6))
        self._secondary_button(preset_btn_frame, "Delete Custom Preset", self._delete_current_preset).pack(side="left")

        opt_frame = tk.Frame(controls_panel, bg=self.PANEL)
        opt_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(10, 0))

        self.lc_aggregate_var = tk.BooleanVar(value=False)
        self._dark_check(
            opt_frame, "Aggregate load cases of same category (e.g. DL1 + DL2 together)", self.lc_aggregate_var
        ).pack(side="left", padx=(0, 25))
        self.lc_aggregate_var.trace_add("write", lambda *_: self._recalculate_combinations())

        tk.Label(opt_frame, text="Start ULS #:", bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 9)).pack(side="left")
        self.lc_start_uls = tk.StringVar(value="101")
        e_uls = self._dark_entry(opt_frame, self.lc_start_uls, width=6)
        e_uls.pack(side="left", padx=(5, 15))
        e_uls.bind("<KeyRelease>", lambda _e: self._recalculate_combinations())

        tk.Label(opt_frame, text="Start SLS #:", bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 9)).pack(side="left")
        self.lc_start_sls = tk.StringVar(value="201")
        e_sls = self._dark_entry(opt_frame, self.lc_start_sls, width=6)
        e_sls.pack(side="left", padx=(5, 0))
        e_sls.bind("<KeyRelease>", lambda _e: self._recalculate_combinations())

        paned = tk.PanedWindow(self.content, orient="vertical", bg=self.BG, sashwidth=7, sashrelief="flat", bd=0)
        paned.pack(fill="both", expand=True)

        primary_panel = self._panel(paned, 0)
        preview_panel = self._panel(paned, 0)

        paned.add(primary_panel, minsize=160, stretch="always")
        paned.add(preview_panel, minsize=220, stretch="always")

        primary_panel.grid_rowconfigure(1, weight=1)
        primary_panel.grid_columnconfigure(0, weight=1)
        tk.Label(
            primary_panel, text="PRIMARY LOAD CASES MAPPING (Double click a row to change Category)", bg=self.PANEL, fg=self.MUTED,
            font=("Segoe UI", 8, "bold"), padx=12, pady=8
        ).grid(row=0, column=0, sticky="w")

        cols_p = ("id", "title", "category")
        heads_p = ("Load Case #", "Title", "Assigned Category (Double-click to edit)")
        widths_p = (100, 300, 400)
        self.lc_primary_table = ttk.Treeview(primary_panel, columns=cols_p, show="headings", style="Dark.Treeview")
        for col, head, w in zip(cols_p, heads_p, widths_p):
            self.lc_primary_table.heading(col, text=head)
            self.lc_primary_table.column(col, width=w, minwidth=60, anchor="center" if col in ("id", "category") else "w")

        v_p = ttk.Scrollbar(primary_panel, orient="vertical", command=self.lc_primary_table.yview, style="Dark.Vertical.TScrollbar")
        self.lc_primary_table.configure(yscrollcommand=v_p.set)
        self.lc_primary_table.grid(row=1, column=0, sticky="nsew")
        v_p.grid(row=1, column=1, sticky="ns")

        self.lc_primary_table.bind("<Double-1>", self._on_edit_primary_category)

        preview_panel.grid_rowconfigure(1, weight=1)
        preview_panel.grid_columnconfigure(0, weight=1)

        prev_header = tk.Frame(preview_panel, bg=self.PANEL)
        prev_header.grid(row=0, column=0, sticky="ew", padx=12, pady=8)
        tk.Label(
            prev_header, text="GENERATED LOAD COMBINATIONS PREVIEW", bg=self.PANEL, fg=self.MUTED,
            font=("Segoe UI", 8, "bold")
        ).pack(side="left")
        self.lc_combo_count_lbl = tk.Label(
            prev_header, text="0 combinations", bg=self.PANEL, fg=self.SUCCESS,
            font=("Segoe UI", 8, "bold")
        )
        self.lc_combo_count_lbl.pack(side="right")

        cols_c = ("number", "title", "type", "formula")
        heads_c = ("Combo #", "Combination Name", "Type", "Constituent Load Cases & Factors")
        widths_c = (90, 320, 80, 420)
        self.lc_combo_table = ttk.Treeview(preview_panel, columns=cols_c, show="headings", style="Dark.Treeview")
        for col, head, w in zip(cols_c, heads_c, widths_c):
            self.lc_combo_table.heading(col, text=head)
            self.lc_combo_table.column(col, width=w, minwidth=60, anchor="center" if col in ("number", "type") else "w")

        v_c = ttk.Scrollbar(preview_panel, orient="vertical", command=self.lc_combo_table.yview, style="Dark.Vertical.TScrollbar")
        self.lc_combo_table.configure(yscrollcommand=v_c.set)
        self.lc_combo_table.grid(row=1, column=0, sticky="nsew")
        v_c.grid(row=1, column=1, sticky="ns")

        action_bar = tk.Frame(self.content, bg=self.BG)
        action_bar.pack(fill="x", pady=(12, 0))
        self._primary_button(action_bar, "Send Combinations to Active STAAD.Pro", self._push_combos_to_staad).pack(
            side="left", padx=(0, 10)
        )
        self._secondary_button(action_bar, "Copy STAAD Commands", self._copy_staad_combo_commands).pack(
            side="left"
        )

        self.primary_cases_data: list[PrimaryLoadCase] = [
            PrimaryLoadCase(1, "DEAD LOAD", "DL"),
            PrimaryLoadCase(2, "COLLATERAL LOAD", "DL"),
            PrimaryLoadCase(3, "LIVE LOAD", "LL"),
            PrimaryLoadCase(4, "WIND +X", "WL"),
            PrimaryLoadCase(5, "WIND -X", "WL"),
        ]
        self.generated_combos_data: list[GeneratedCombo] = []
        self._render_primary_cases_table()
        self._recalculate_combinations()

    def _render_primary_cases_table(self) -> None:
        self.lc_primary_table.delete(*self.lc_primary_table.get_children())
        for plc in self.primary_cases_data:
            self.lc_primary_table.insert("", "end", iid=str(plc.id), values=(plc.id, plc.title, plc.category))

    def _on_edit_primary_category(self, _event: Any) -> None:
        sel = self.lc_primary_table.selection()
        if not sel:
            return
        lc_id = int(sel[0])
        plc = next((c for c in self.primary_cases_data if c.id == lc_id), None)
        if not plc:
            return

        menu = tk.Menu(self.root, tearoff=0, bg=self.PANEL_ALT, fg=self.TEXT, activebackground=self.ACCENT)
        for cat in LOAD_CATEGORIES:
            menu.add_command(
                label=cat,
                command=lambda selected_cat=cat: self._set_primary_category(plc, selected_cat),
            )
        try:
            menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
        finally:
            menu.grab_release()

    def _set_primary_category(self, plc: PrimaryLoadCase, category: str) -> None:
        plc.category = category
        self._render_primary_cases_table()
        self._recalculate_combinations()

    def _fetch_lc_from_staad(self) -> None:
        self._set_status("Fetching primary load cases from STAAD.Pro…", "muted")
        self.root.update_idletasks()
        try:
            staad = OpenStaad.connect()
            cases = fetch_primary_cases_from_openstaad(staad)
            if not cases:
                self._set_status("No primary load cases found in active model.", "warning")
                return
            self.primary_cases_data = cases
            self._render_primary_cases_table()
            self._recalculate_combinations()
            self._set_status(f"Loaded {len(cases)} primary load case(s) from STAAD.Pro.", "success")
        except (OpenStaadError, OSError, TypeError, ValueError) as exc:
            self._set_status(str(exc), "error")

    def _add_manual_lc_row(self) -> None:
        next_id = max([c.id for c in self.primary_cases_data], default=0) + 1
        self.primary_cases_data.append(PrimaryLoadCase(next_id, f"LOAD CASE {next_id}", "DL"))
        self._render_primary_cases_table()
        self._recalculate_combinations()

    def _on_preset_changed(self) -> None:
        self._recalculate_combinations()

    def _get_selected_preset(self) -> CombinationPreset:
        name = self.lc_preset_name.get()
        return next((p for p in self.lc_presets_list if p.name == name), self.lc_presets_list[0])

    def _recalculate_combinations(self) -> None:
        if not hasattr(self, "lc_combo_table"):
            return
        preset = self._get_selected_preset()
        try:
            uls = int(self.lc_start_uls.get().strip() or "101")
            sls = int(self.lc_start_sls.get().strip() or "201")
        except ValueError:
            uls, sls = 101, 201

        combos = generate_combinations(
            primary_cases=self.primary_cases_data,
            preset=preset,
            aggregate_same_type=self.lc_aggregate_var.get(),
            start_uls=uls,
            start_sls=sls,
        )
        self.generated_combos_data = combos
        self.lc_combo_table.delete(*self.lc_combo_table.get_children())
        for c in combos:
            formula = " + ".join(f"{factor:g}×[LC{lc_id}]" for lc_id, factor in c.factors)
            self.lc_combo_table.insert("", "end", values=(c.number, c.title, c.combo_type, formula))

        self.lc_combo_count_lbl.configure(text=f"{len(combos)} combinations generated")

    def _show_edit_preset_factors_dialog(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("Create / Edit Load Factor Preset")
        win.geometry("920x560")
        win.configure(bg=self.PANEL)
        win.transient(self.root)
        win.grab_set()

        curr_preset = self._get_selected_preset()

        top_frame = tk.Frame(win, bg=self.PANEL)
        top_frame.pack(fill="x", padx=16, pady=(16, 10))

        tk.Label(top_frame, text="Preset Name:", bg=self.PANEL, fg=self.TEXT, font=("Segoe UI", 9, "bold")).pack(side="left")
        name_var = tk.StringVar(value=f"{curr_preset.name} (Copy)" if curr_preset.is_builtin else curr_preset.name)
        self._dark_entry(top_frame, name_var, width=28).pack(side="left", padx=(8, 20))

        tk.Label(top_frame, text="Description:", bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 9)).pack(side="left")
        desc_var = tk.StringVar(value=curr_preset.description)
        self._dark_entry(top_frame, desc_var, width=32).pack(side="left", padx=(8, 0))

        table_container = tk.Frame(win, bg=self.PANEL_ALT, highlightbackground=self.BORDER, highlightthickness=1)
        table_container.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        canvas = tk.Canvas(table_container, bg=self.PANEL_ALT, highlightthickness=0, bd=0)
        vbar = ttk.Scrollbar(table_container, orient="vertical", command=canvas.yview, style="Dark.Vertical.TScrollbar")
        rules_frame = tk.Frame(canvas, bg=self.PANEL_ALT)

        rules_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        window_id = canvas.create_window((0, 0), window=rules_frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(window_id, width=e.width))
        canvas.configure(yscrollcommand=vbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")

        cats = ("DL", "LL", "RLL", "CL", "WL", "EQ", "CRANE", "TEMP")
        headers = ["Rule Name / Formula", "Type"] + list(cats)
        widths = [20, 7] + [5] * len(cats)

        hdr_frame = tk.Frame(rules_frame, bg=self.PANEL)
        hdr_frame.pack(fill="x", padx=4, pady=(4, 6))
        for col_idx, (h, w) in enumerate(zip(headers, widths)):
            tk.Label(
                hdr_frame, text=h, bg=self.PANEL, fg=self.MUTED,
                font=("Segoe UI", 8, "bold"), width=w, anchor="center" if col_idx > 0 else "w"
            ).pack(side="left", padx=2)

        rule_rows_data: list[dict[str, Any]] = []

        def add_rule_row(rule: ComboRule | None = None) -> None:
            r_frame = tk.Frame(rules_frame, bg=self.PANEL_ALT)
            r_frame.pack(fill="x", padx=4, pady=2)

            name_v = tk.StringVar(value=rule.name_template if rule else "1.5(DL + LL)")
            type_v = tk.StringVar(value=rule.combo_type if rule else "ULS")
            factor_vars: dict[str, tk.StringVar] = {}

            e_name = self._dark_entry(r_frame, name_v, width=20)
            e_name.pack(side="left", padx=2)

            combo_type = ttk.Combobox(r_frame, textvariable=type_v, values=["ULS", "SLS"], state="readonly", width=6)
            combo_type.pack(side="left", padx=2)

            rule_factors = rule.factors if rule else {}
            for cat in cats:
                init_val = str(rule_factors.get(cat, 0.0)) if rule else ("1.5" if cat in ("DL", "LL") else "0")
                f_var = tk.StringVar(value=init_val)
                factor_vars[cat] = f_var
                e_f = self._dark_entry(r_frame, f_var, width=5)
                e_f.pack(side="left", padx=2)

            del_btn = tk.Button(
                r_frame, text="✕", bg=self.PANEL_ALT, fg=self.ERROR,
                activebackground=self.BORDER, activeforeground=self.ERROR,
                relief="flat", bd=0, font=("Segoe UI", 8, "bold"), cursor="hand2",
                command=lambda rf=r_frame, rd=rule_rows_data: remove_rule_row(rf, rd)
            )
            del_btn.pack(side="left", padx=(4, 0))

            rule_rows_data.append({
                "frame": r_frame,
                "name": name_v,
                "type": type_v,
                "factors": factor_vars,
            })

        def remove_rule_row(r_frame: tk.Frame, rd: list[dict[str, Any]]) -> None:
            r_frame.destroy()
            rd[:] = [item for item in rd if item["frame"] is r_frame]

        for r in curr_preset.rules:
            add_rule_row(r)

        btn_toolbar = tk.Frame(win, bg=self.PANEL)
        btn_toolbar.pack(fill="x", padx=16, pady=(0, 16))

        self._secondary_button(btn_toolbar, "+ Add Factor Rule Row", lambda: add_rule_row()).pack(side="left")

        def save_preset_factors() -> None:
            pname = name_var.get().strip()
            if not pname:
                return

            new_rules: list[ComboRule] = []
            for row in rule_rows_data:
                try:
                    r_name = row["name"].get().strip()
                    r_type = row["type"].get().strip()
                    if not r_name:
                        continue
                    factors_dict: dict[str, float] = {}
                    for cat, fvar in row["factors"].items():
                        try:
                            val = float(fvar.get().strip())
                            if abs(val) > 1e-4:
                                factors_dict[cat] = val
                        except ValueError:
                            pass
                    if factors_dict:
                        new_rules.append(ComboRule(name_template=r_name, combo_type=r_type, factors=factors_dict))
                except Exception:
                    pass

            if not new_rules:
                return

            new_preset = CombinationPreset(
                name=pname,
                description=desc_var.get().strip(),
                is_builtin=False,
                rules=new_rules,
            )
            save_custom_preset(new_preset)

            self.lc_presets_list = load_all_presets()
            preset_names = [p.name for p in self.lc_presets_list]
            self.lc_preset_combo.configure(values=preset_names)
            self.lc_preset_name.set(pname)
            win.destroy()
            self._recalculate_combinations()
            self._set_status(f"Saved custom factor preset '{pname}' with {len(new_rules)} rules.", "success")

        self._primary_button(btn_toolbar, "Save Factor Preset", save_preset_factors).pack(side="right", padx=(10, 0))
        self._secondary_button(btn_toolbar, "Cancel", win.destroy).pack(side="right")

    def _delete_current_preset(self) -> None:
        preset = self._get_selected_preset()
        if preset.is_builtin:
            self._set_status("Cannot delete built-in standard presets.", "warning")
            return
        if delete_custom_preset(preset.name):
            self.lc_presets_list = load_all_presets()
            preset_names = [p.name for p in self.lc_presets_list]
            self.lc_preset_combo.configure(values=preset_names)
            if preset_names:
                self.lc_preset_name.set(preset_names[0])
            self._recalculate_combinations()
            self._set_status(f"Deleted custom preset '{preset.name}'.", "success")

    def _push_combos_to_staad(self) -> None:
        if not self.generated_combos_data:
            self._set_status("No combinations generated to send.", "warning")
            return
        self._set_status("Connecting to STAAD.Pro and creating load combinations…", "muted")
        self.root.update_idletasks()
        try:
            staad = OpenStaad.connect()
            count = push_combos_to_openstaad(staad, self.generated_combos_data)
            self._set_status(f"Successfully created {count} load combination(s) in active STAAD.Pro model!", "success")
        except (OpenStaadError, OSError, TypeError, ValueError) as exc:
            self._set_status(str(exc), "error")

    def _copy_staad_combo_commands(self) -> None:
        if not self.generated_combos_data:
            self._set_status("No combinations generated to copy.", "warning")
            return
        text = format_staad_combo_text(self.generated_combos_data)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self._set_status(f"Copied STAAD text for {len(self.generated_combos_data)} combinations to clipboard!", "success")

    def _set_status(self, message: str, kind: str = "muted") -> None:
        colors = {
            "muted": self.MUTED,
            "success": self.SUCCESS,
            "warning": self.WARNING,
            "error": self.ERROR,
        }
        self._status_text.set(message)
        self.status_dot.configure(fg=colors.get(kind, self.MUTED))

    def run(self) -> int:
        self.root.mainloop()
        return 0
