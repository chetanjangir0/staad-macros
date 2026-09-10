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


def test_deliverable_record_and_recent_outputs_tracking() -> None:
    from pathlib import Path
    from staad_ext.desktop import DeliverableRecord

    app = SimpleNamespace(recent_outputs=[])
    StaadExtApplication._record_output(app, "STD to DXF", Path("sample.dxf"))

    assert len(app.recent_outputs) == 1
    rec = app.recent_outputs[0]
    assert isinstance(rec, DeliverableRecord)
    assert rec.label == "STD to DXF"
    assert rec.file_path == Path("sample.dxf")
    assert rec.status == "Completed"


def test_model_telemetry_disconnected_graceful_state() -> None:
    from staad_ext.desktop import ModelTelemetry

    telemetry = ModelTelemetry(
        connected=False,
        model_name="No Active Model",
        model_path=None,
        base_unit_str="Unknown",
        member_count=0,
        node_count=0,
        support_count=0,
        primary_cases=0,
        combo_cases=0,
        results_available=False,
        selected_count=0,
        selected_length=0.0,
        selected_sections={},
    )
    assert telemetry.connected is False
    assert telemetry.member_count == 0
    assert telemetry.selected_count == 0


def test_query_model_telemetry_graceful_when_staad_absent() -> None:
    app = SimpleNamespace()
    telemetry = StaadExtApplication._query_model_telemetry(app)
    assert isinstance(telemetry.connected, bool)
    if not telemetry.connected:
        assert telemetry.model_path is None
        assert telemetry.member_count == 0
        assert telemetry.selected_count == 0


