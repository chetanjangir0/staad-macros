from io import StringIO
from math import cos, hypot, radians, sin

import pytest

from staad_ext.dxf import DxfWriter
from staad_ext.framing import FramingModel, Member, move
from staad_ext.macros.std_to_ga_dxf import (
    GA_LAYERS, MARK_LEADER_RADII, build_schedule, describe_section, export_ga_drawing,
    label_rotation, mark_anchor, mark_direction, mark_radius, member_grade,
    member_length_label, write_inner_faces, write_mark_bubble, write_schedule,
)
from staad_ext.models import (
    GaExportSettings, Point3D, ScheduleCorner, SectionEnvelope, ViewPlane,
)


def make_member(number=1, property_type=675, values=None, name="TAPERED",
                start=Point3D(0, 0), end=Point3D(0, 4), half_width=0.3, length=4.0,
                thickness=0.0):
    envelope = SectionEnvelope(half_width, half_width, property_type, thickness)
    return Member(
        number=number, start=start, end=end, incidence=(number, number + 1),
        length=length, envelope=envelope, name=name,
        property_values=list(values or []),
        outline=[Point3D(-half_width, 0), Point3D(-half_width, 4),
                 Point3D(half_width, 0), Point3D(half_width, 4)],
    )


def values_for(property_type):
    v = [0.0] * 24
    if property_type == 675:                    # tapered built-up I
        v[1], v[4], v[5], v[6], v[7] = 0.005, 0.382, 0.632, 0.150, 0.008
    elif property_type == 650:                  # RHS
        v[1], v[2], v[3] = 0.400, 0.300, 0.016
    elif property_type == 660:                  # CHS by OD + wall
        v[1], v[2] = 0.1397, 0.0036
    elif property_type == 695:                  # CHS by OD + ID
        v[1], v[2] = 0.1651, 0.1579
    return v


class FakeStaad:
    """Only the surface build_schedule touches."""

    def __init__(self, grades=None):
        self.grades = grades or {}
        self.material_calls = 0

    def beam_material_name(self, beam_no):
        return self.grades.get(beam_no, "STEEL")

    def material_yield_strength(self, name):
        self.material_calls += 1
        return {"STEEL": 250.0, "STEEL_E355": 355.0}.get(name)


def model_of(*members):
    return FramingModel({member.number: member for member in members}, set(), 0.0)


def test_rhs_description_matches_schedule_convention() -> None:
    member = make_member(property_type=650, values=values_for(650), name="TUB40030016")
    assert describe_section(member) == "400x300x16.0  (RHS)"


def test_chs_description_keeps_one_decimal_diameter() -> None:
    member = make_member(property_type=660, values=values_for(660), name="PIPE")
    assert describe_section(member) == "139.7x3.6  (CHS)"


def test_chs_given_as_outside_and_inside_diameter_reports_wall_thickness() -> None:
    member = make_member(property_type=695, values=values_for(695), name="PIPE")
    assert describe_section(member) == "165.1x3.6  (CHS)"


def test_tapered_built_up_reports_web_plate_depths_largest_first() -> None:
    member = make_member(property_type=675, values=values_for(675))
    # 632 and 382 overall, less two 8mm flanges, is a 616~366 web plate.
    assert describe_section(member) == "W=616~366x5/F=150x8"


def test_prismatic_built_up_collapses_to_a_single_web_depth() -> None:
    values = values_for(675)
    values[4] = values[5]
    member = make_member(property_type=675, values=values)
    assert describe_section(member) == "W=616x5/F=150x8"


def test_description_falls_back_to_the_staad_section_name() -> None:
    assert describe_section(make_member(property_type=0, values=[], name="ISMB400")) == "ISMB400"


def test_description_survives_a_short_property_value_array() -> None:
    member = make_member(property_type=675, values=[0.0, 0.005], name="TAPERED")
    assert describe_section(member) == "TAPERED"


def test_members_of_one_size_and_grade_share_a_mark() -> None:
    members = [make_member(number=n, property_type=650, values=values_for(650),
                           name="TUB40030016") for n in (1, 2, 3)]
    marks, entries = build_schedule(FakeStaad(), model_of(*members))
    assert marks == {1: 1, 2: 1, 3: 1}
    assert len(entries) == 1
    assert entries[0].beams == (1, 2, 3)


def test_grade_column_is_left_blank_and_does_not_split_a_mark() -> None:
    # The GRADE cell is filled in by hand for now, so differing STAAD materials
    # must not hand two members of one size separate marks.
    members = [make_member(number=n, property_type=650, values=values_for(650),
                           name="TUB40030016") for n in (1, 2)]
    staad = FakeStaad(grades={2: "STEEL_E355"})
    marks, entries = build_schedule(staad, model_of(*members))
    assert marks == {1: 1, 2: 1}
    assert [entry.grade for entry in entries] == [""]
    assert staad.material_calls == 0


def test_differently_named_sections_of_one_size_share_a_mark() -> None:
    # A schedule row is keyed on what it prints, so two rows that would read
    # identically must not be handed separate marks.
    members = [make_member(number=1, property_type=650, values=values_for(650), name="COL-RHS"),
               make_member(number=2, property_type=650, values=values_for(650), name="RAF-RHS")]
    marks, entries = build_schedule(FakeStaad(), model_of(*members))
    assert marks == {1: 1, 2: 1}
    assert len(entries) == 1


def test_marks_are_numbered_by_lowest_member_number() -> None:
    members = [make_member(number=9, property_type=650, values=values_for(650), name="A"),
               make_member(number=2, property_type=660, values=values_for(660), name="B")]
    marks, entries = build_schedule(FakeStaad(), model_of(*members))
    assert marks == {2: 1, 9: 2}
    assert [entry.mark for entry in entries] == [1, 2]


def test_grade_is_blank_and_never_reads_the_model() -> None:
    staad = FakeStaad()
    assert member_grade(staad, 1) == ""
    assert staad.material_calls == 0


def test_schedule_has_a_row_for_every_mark_plus_spares() -> None:
    stream = StringIO()
    writer = DxfWriter(stream)
    writer.header(GA_LAYERS)
    members = [make_member(number=1, property_type=650, values=values_for(650), name="A"),
               make_member(number=2, property_type=660, values=values_for(660), name="B")]
    model = model_of(*members)
    _, entries = build_schedule(FakeStaad(), model)
    write_schedule(writer, model, entries, GaExportSettings(blank_rows=4))
    value = stream.getvalue()
    # title + heading + 2 entries + 4 spares = 8 rows, so 9 horizontal rules,
    # plus two full-height and two part-height verticals.
    assert value.count("8\nSCHEDULE_GRID") == 9 + 4
    assert "MEMBER SIZE SCHEDULE" in value
    assert "400x300x16.0  (RHS)" in value
    assert "139.7x3.6  (CHS)" in value
    # One bubble per scheduled mark, none for the spare rows.
    assert value.count("0\nCIRCLE") == 2


def test_schedule_corner_moves_the_table_across_the_frame() -> None:
    members = [make_member(number=1, property_type=650, values=values_for(650), name="A")]
    model = model_of(*members)
    _, entries = build_schedule(FakeStaad(), model)

    def first_grid_x(corner):
        stream = StringIO()
        writer = DxfWriter(stream)
        write_schedule(writer, model, entries, GaExportSettings(schedule_corner=corner))
        line = stream.getvalue().split("8\nSCHEDULE_GRID\n")[1]
        return float(line.split("10\n")[1].split("\n")[0])

    assert first_grid_x(ScheduleCorner.TOP_LEFT) < 0
    assert first_grid_x(ScheduleCorner.TOP_RIGHT) > 0


def test_mark_bubble_draws_a_circle_a_leader_and_an_arrow_head() -> None:
    stream = StringIO()
    writer = DxfWriter(stream)
    write_mark_bubble(writer, make_member(), 7, 0.3)
    value = stream.getvalue()
    assert value.count("0\nCIRCLE") == 1
    assert value.count("8\nMARK_LEADERS") == 3       # one leader, two arrow barbs
    assert "\n1\n7\n" in value


def texts_of(dxf: str) -> dict[str, tuple[Point3D, float, float]]:
    """Map each TEXT entity's string to its (insertion point, height, rotation)."""
    found = {}
    for chunk in dxf.split("0\nTEXT\n")[1:]:
        pairs = chunk.split("\n0\n")[0].split("\n")
        codes = {pairs[index]: pairs[index + 1] for index in range(0, len(pairs) - 1, 2)}
        found[codes["1"]] = (Point3D(float(codes["10"]), float(codes["20"])),
                             float(codes["40"]), float(codes["50"]))
    return found


def test_the_mark_carries_the_members_true_length_in_metres() -> None:
    assert member_length_label(make_member(length=6.0)) == "L=6.00m"
    assert member_length_label(make_member(length=6.1234)) == "L=6.12m"


def test_the_length_sits_past_the_far_side_of_the_bubble() -> None:
    # Everything else -- the member, its leader, the neighbouring marks -- is on
    # the near side, so the label is placed clear of all of it.
    member, radius = make_member(), 0.3
    stream = StringIO()
    write_mark_bubble(DxfWriter(stream), member, 7, radius)
    point, height, _ = texts_of(stream.getvalue())["L=4.00m"]

    unit = mark_direction(member)
    center = move(mark_anchor(member), unit, radius * MARK_LEADER_RADII)
    outward = (point.x - center.x) * unit.x + (point.y - center.y) * unit.y
    assert outward > radius
    assert hypot(point.x - center.x, point.y - center.y) == pytest.approx(outward)
    assert height == pytest.approx(radius * 0.8)


@pytest.mark.parametrize("end", [Point3D(0, 4), Point3D(0, -4), Point3D(4, 0),
                                 Point3D(-4, 0), Point3D(-3, -3), Point3D(3, -3)])
def test_the_length_reads_left_to_right_and_grows_away_from_the_member(end) -> None:
    member = make_member(start=Point3D(0, 0), end=end)
    rotation = label_rotation(member)
    assert -90.0 <= rotation <= 90.0
    # Text grows along its own "up" (rotation + 90), which must be the side the
    # bubble is on, or the label would be laid back over the member.
    up = Point3D(-sin(radians(rotation)), cos(radians(rotation)))
    unit = mark_direction(member)
    assert up.x * unit.x + up.y * unit.y == pytest.approx(1.0)


def inner_faces_dxf(property_type, name, thickness) -> str:
    stream = StringIO()
    write_inner_faces(DxfWriter(stream), make_member(
        property_type=property_type, name=name, thickness=thickness,
    ))
    return stream.getvalue()


def test_a_flange_is_drawn_as_a_visible_face_inside_the_outline() -> None:
    value = inner_faces_dxf(612, "ISMB400", 0.016)
    assert value.count("8\nMEMBER_OUTLINE") == 2
    assert "HIDDEN" not in value
    # The outline runs to +/-0.3, so a 16mm flange sits just inside it.
    assert "10\n-0.284000" in value and "10\n0.284000" in value


def test_a_tubes_walls_are_drawn_hidden_behind_its_front_face() -> None:
    assert inner_faces_dxf(650, "TUB40030016", 0.016).count("6\nHIDDEN") == 2


def test_a_member_without_a_thickness_stays_a_plain_outline() -> None:
    assert inner_faces_dxf(675, "TAPERED", 0.0) == ""


def test_export_writes_a_complete_dxf(tmp_path) -> None:
    class ExportStaad(FakeStaad):
        def selected_beams(self):
            return [1, 2]

        def member_incidence(self, beam_no):
            return (1, 2) if beam_no == 1 else (2, 3)

        def node_coordinates(self, node_no):
            return {1: Point3D(0, 0, 0), 2: Point3D(0, 6, 0), 3: Point3D(10, 8, 0)}[node_no]

        def beam_length(self, beam_no):
            return 6.0 if beam_no == 1 else 10.2

        def section_name(self, beam_no):
            return "TAPERED"

        def beam_property_all(self, beam_no):
            return (0.15, 0.632, 0.01, 0.005, 0.005, 1e-4, 1e-4, 1e-4, 0.008, 0.005)

        def section_property_values(self, beam_no):
            return 675, values_for(675)

    output = tmp_path / "ga.dxf"
    assert export_ga_drawing(ExportStaad(), output, GaExportSettings()) == 2
    value = output.read_text()
    assert value.endswith("0\nEOF\n")
    assert "SECTION\n2\nENTITIES" in value
    for layer, _, _ in GA_LAYERS:
        assert f"0\nLAYER\n2\n{layer}" in value
    assert "W=616~366x5/F=150x8" in value
    # Both members are one size and grade, so the schedule carries a single mark.
    assert value.count("MEMBER SIZE SCHEDULE") == 1

    # A GA drawing is dimensioned in millimetres by default: the 6m column
    # reaches y=6000, while its length label still reads in metres.
    assert "9\n$INSUNITS\n70\n4" in value
    assert "20\n6000.000000" in value
    assert "L=6.00m" in value

    metres = tmp_path / "metres.dxf"
    export_ga_drawing(ExportStaad(), metres, GaExportSettings(millimetre_units=False))
    in_metres = metres.read_text()
    assert "9\n$INSUNITS\n70\n6" in in_metres
    assert "20\n6.000000" in in_metres
    assert "L=6.00m" in in_metres

    class NoThicknessStaad(ExportStaad):
        def beam_property_all(self, beam_no):
            raise OSError("no property table")

        def section_property_values(self, beam_no):
            return 0, []

    plain = tmp_path / "plain.dxf"
    export_ga_drawing(NoThicknessStaad(), plain, GaExportSettings())
    # Two flange faces per member, on top of whatever the outlines already draw.
    assert (value.count("8\nMEMBER_OUTLINE")
            == plain.read_text().count("8\nMEMBER_OUTLINE") + 4)


def test_export_returns_zero_without_a_selection(tmp_path) -> None:
    class EmptyStaad(FakeStaad):
        def selected_beams(self):
            return []

    output = tmp_path / "ga.dxf"
    assert export_ga_drawing(EmptyStaad(), output, GaExportSettings()) == 0
    assert not output.exists()


def spaced_frame(pitch, span=24.0):
    """A cluster of vertical members at ``pitch`` centres over the left half of
    a fixed span, plus one far member that pins the drawing extent.

    Holding the extent constant means only the member spacing varies between
    cases, which is what the bubble-sizing cap is supposed to react to.
    """
    positions = []
    x = 0.0
    while x <= span / 2 and len(positions) < 40:
        positions.append(x)
        x += pitch
    positions.append(span)
    return model_of(*[
        make_member(number=index + 1, property_type=650, values=values_for(650),
                    name="WEB", start=Point3D(position, 0), end=Point3D(position, 4),
                    half_width=0.06)
        for index, position in enumerate(positions)
    ])


def closest_bubble_gap(model, settings):
    radius = mark_radius(model, settings)
    centres = [move(mark_anchor(member), mark_direction(member), radius * MARK_LEADER_RADII)
               for member in model.members.values()]
    gap = min(hypot(b.x - a.x, b.y - a.y)
              for index, a in enumerate(centres) for b in centres[index + 1:])
    return radius, gap


@pytest.mark.parametrize("pitch", [4.0, 2.0, 1.0, 0.5, 0.25])
def test_mark_bubbles_never_overlap_at_any_member_spacing(pitch) -> None:
    radius, gap = closest_bubble_gap(spaced_frame(pitch), GaExportSettings())
    assert gap > 2 * radius, f"bubbles overlap at {pitch}m spacing"


def test_crowded_framing_shrinks_the_bubbles() -> None:
    settings = GaExportSettings()
    assert mark_radius(spaced_frame(0.25), settings) < mark_radius(spaced_frame(4.0), settings)


def test_bubble_size_is_capped_by_spacing_not_just_drawing_extent() -> None:
    # Same overall extent, tighter members: the spacing cap must bind.
    settings = GaExportSettings()
    loose = mark_radius(spaced_frame(4.0), settings)
    tight = mark_radius(spaced_frame(0.2), settings)
    assert tight < loose * 0.9


def test_mark_scale_resizes_bubbles_without_touching_schedule_text() -> None:
    model = spaced_frame(4.0)
    base = mark_radius(model, GaExportSettings())
    assert mark_radius(model, GaExportSettings(mark_scale=0.5)) == pytest.approx(base * 0.5)
    # text_scale drives the schedule, so it must leave the bubbles alone.
    assert mark_radius(model, GaExportSettings(text_scale=3.0)) == pytest.approx(base)


@pytest.mark.parametrize("field, value", [("text_scale", 0.0), ("text_scale", 11.0),
                                          ("mark_scale", 0.0), ("mark_scale", 11.0),
                                          ("blank_rows", -1), ("blank_rows", 41)])
def test_settings_reject_out_of_range_values(field, value) -> None:
    with pytest.raises(ValueError):
        GaExportSettings(**{field: value})


def test_settings_default_to_the_xy_plane_and_a_top_right_schedule() -> None:
    settings = GaExportSettings()
    assert settings.plane is ViewPlane.XY
    assert settings.schedule_corner is ScheduleCorner.TOP_RIGHT
