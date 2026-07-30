from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, ttk
from typing import Callable

from staad_ext.macros.plate_summary import PlateSummaryRow, selected_member_plate_summary
from staad_ext.macros.support_reactions import (
    SupportReaction,
    get_support_reactions,
    parse_load_cases,
    reaction_envelopes,
)
from staad_ext.macros.std_to_dxf import export_selected_members
from staad_ext.models import ExportSettings, ViewPlane
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
        "std_to_dxf",
        "STD to DXF",
        "STD to DXF",
        "Export selected analytical members and section envelopes to AutoCAD.",
        "_build_dxf_view",
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
            hero, text="STAAD.Pro utilities, kept together.", bg=self.PANEL,
            fg=self.TEXT, font=("Segoe UI", 17, "bold"),
        ).pack(anchor="w")
        tk.Label(
            hero,
            text=("Select analytical members in your open model, choose a utility "
                  "from the sidebar, and review the result without leaving this window."),
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
