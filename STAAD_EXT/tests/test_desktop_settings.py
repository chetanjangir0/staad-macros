"""The desktop panels have to hand every setting they show to the macro.

Each utility is offered twice: as a panel in the desktop application and as a
standalone dialog. Both build the same settings object, so an option added to
one and forgotten in the other is invisible to whoever uses the other -- the
control is on screen but the run ignores it, or it never appears at all. These
tests call the builders with stand-in variables rather than real Tk ones, so
they need no display.
"""

from __future__ import annotations

from types import SimpleNamespace

from staad_ext.desktop import StaadExtApplication


class Var:
    """The read half of a tkinter variable."""

    def __init__(self, value: object) -> None:
        self._value = value

    def get(self) -> object:
        return self._value


def test_the_desktop_taper_panel_passes_every_option_it_shows() -> None:
    panel = SimpleNamespace(
        taper_vertical=Var("300"),
        taper_horizontal=Var("200"),
        taper_cases=Var("1002, 1003"),
        taper_ceiling=Var("0.9"),
        taper_tie_knees=Var(True),
        taper_straight_columns=Var(True),
        taper_budget=Var("25"),
        taper_apply=Var(True),
    )
    settings = StaadExtApplication._taper_settings(panel)

    assert settings.deflection.vertical_span_ratio == 300.0
    assert settings.deflection.horizontal_height_ratio == 200.0
    assert settings.deflection.load_cases == (1002, 1003)
    assert settings.utilisation_ceiling == 0.9
    assert settings.tie_depths_at_all_shared_nodes is True
    assert settings.prismatic_columns is True
    assert settings.analysis_budget == 25
    assert settings.apply_to_model is True


def test_the_taper_checkboxes_are_off_unless_the_panel_ticks_them() -> None:
    # Every checkbox defaults to off, so an unticked panel has to produce a
    # settings object with all of them off -- not one that quietly inverts.
    panel = SimpleNamespace(
        taper_vertical=Var("240"),
        taper_horizontal=Var("150"),
        taper_cases=Var("101"),
        taper_ceiling=Var("0.95"),
        taper_tie_knees=Var(False),
        taper_straight_columns=Var(False),
        taper_budget=Var("40"),
        taper_apply=Var(False),
    )
    settings = StaadExtApplication._taper_settings(panel)

    assert settings.tie_depths_at_all_shared_nodes is False
    assert settings.prismatic_columns is False
    assert settings.apply_to_model is False
