from __future__ import annotations

from dataclasses import dataclass
from tkinter import Tk, ttk
from typing import Callable

from staad_ext.workflows import run_plate_summary, run_std_to_dxf


@dataclass(frozen=True, slots=True)
class UtilityDefinition:
    """Metadata used to render one utility on the home screen."""

    name: str
    description: str
    action: Callable[[Tk], bool]


UTILITIES = (
    UtilityDefinition(
        name="STD to DXF",
        description=(
            "Export selected analytical beam members, section envelopes, and "
            "labels to an AutoCAD R12 DXF file."
        ),
        action=run_std_to_dxf,
    ),
    UtilityDefinition(
        name="Plate Summary",
        description=(
            "Summarize selected tapered I-members as web and flange plates, "
            "and list every other member as a complete section."
        ),
        action=run_plate_summary,
    ),
)


class StaadExtApplication:
    """Desktop home screen and utility launcher."""

    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("STAAD_EXT")
        self.root.geometry("760x480")
        self.root.minsize(680, 420)
        self._configure_styles()
        self._build_layout()

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        style.configure("Title.TLabel", font=("Segoe UI", 24, "bold"))
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10), foreground="#555555")
        style.configure("CardTitle.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("CardText.TLabel", font=("Segoe UI", 10), foreground="#444444")
        style.configure("Card.TFrame", relief="solid", borderwidth=1)
        style.configure("Launch.TButton", font=("Segoe UI", 10, "bold"), padding=(16, 8))

    def _build_layout(self) -> None:
        container = ttk.Frame(self.root, padding=(32, 26))
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="STAAD_EXT", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            container,
            text="Useful engineering utilities for STAAD.Pro 2025",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 24))

        ttk.Label(container, text="Utilities", style="CardTitle.TLabel").pack(
            anchor="w", pady=(0, 10)
        )
        utilities_frame = ttk.Frame(container)
        utilities_frame.pack(fill="x")
        for utility in UTILITIES:
            self._build_utility_card(utilities_frame, utility)

        ttk.Separator(container).pack(fill="x", side="bottom", pady=(16, 8))
        ttk.Label(
            container,
            text="Open and save a STAAD model before running a utility.",
            style="Subtitle.TLabel",
        ).pack(side="bottom", anchor="w")

    def _build_utility_card(
        self, parent: ttk.Frame, utility: UtilityDefinition
    ) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        card.pack(fill="x", pady=(0, 12))
        card.columnconfigure(0, weight=1)

        ttk.Label(card, text=utility.name, style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            card,
            text=utility.description,
            style="CardText.TLabel",
            wraplength=520,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Button(
            card,
            text="Open utility",
            style="Launch.TButton",
            command=lambda: utility.action(self.root),
        ).grid(row=0, column=1, rowspan=2, padx=(24, 0))

    def run(self) -> int:
        self.root.mainloop()
        return 0


def launch() -> int:
    return StaadExtApplication().run()
