from io import StringIO
from math import hypot

import pytest

from staad_ext.dxf import DxfWriter
from staad_ext.framing import (
    apply_peb_corner_joins, inner_face_lines, is_tapered, member_envelope, project,
)
from staad_ext.macros.std_to_dxf import (
    write_connection_face_lines, write_member_envelope,
)
from staad_ext.models import ExportSettings, Point3D, SectionEnvelope, ViewPlane


def test_projection_planes() -> None:
    point = Point3D(1, 2, 3)
    assert project(point, ViewPlane.XY) == Point3D(1, 2, 0)
    assert project(point, ViewPlane.YZ) == Point3D(3, 2, 0)
    assert project(point, ViewPlane.ZX) == Point3D(3, 1, 0)


def test_property_type_675_is_tapered() -> None:
    assert is_tapered(SectionEnvelope(0.2, 0.2, 675), "")


def test_dxf_has_valid_sections_and_eof() -> None:
    stream = StringIO()
    writer = DxfWriter(stream)
    writer.header()
    writer.line("MEMBER_CENTERLINE", Point3D(0, 0), Point3D(1, 0), "DASHED")
    writer.footer()
    value = stream.getvalue()
    assert "SECTION\n2\nENTITIES" in value
    assert value.endswith("0\nEOF\n")


def test_a_millimetre_drawing_scales_every_length_but_not_the_angles() -> None:
    stream = StringIO()
    writer = DxfWriter(stream, ExportSettings().scale)
    writer.header()
    writer.line("MEMBER_CENTERLINE", Point3D(0, 0), Point3D(6, 0), "DASHED")
    writer.circle("MEMBER_MARKS", Point3D(1.5, 0), 0.25)
    writer.text("MEMBER_LABELS", Point3D(1.5, 0), 0.1, 45.0, "L=6.00m", 7)
    value = stream.getvalue()
    assert "9\n$INSUNITS\n70\n4" in value
    assert "11\n6000.000000" in value          # the line's far end
    assert "40\n250.000000" in value           # the circle radius
    assert "40\n100.000000" in value           # the text height
    assert "50\n45.000000" in value            # the rotation, left in degrees
    assert "49\n500.000000" in value           # the DASHED pattern's dash


def test_a_metre_drawing_leaves_the_model_units_alone() -> None:
    stream = StringIO()
    writer = DxfWriter(stream, ExportSettings(millimetre_units=False).scale)
    writer.header()
    writer.line("MEMBER_CENTERLINE", Point3D(0, 0), Point3D(6, 0), "DASHED")
    value = stream.getvalue()
    assert "9\n$INSUNITS\n70\n6" in value
    assert "11\n6.000000" in value
    assert "49\n0.500000" in value


def test_peb_corner_join_extends_rafter_edges_to_column_flange_lines() -> None:
    outlines = {
        1: [
            Point3D(-0.2, 0), Point3D(-0.2, 4),
            Point3D(0.2, 0), Point3D(0.2, 4),
        ],
        2: [
            Point3D(-0.1, 4.2), Point3D(4.9, 5.2),
            Point3D(0.1, 3.8), Point3D(5.1, 4.8),
        ],
    }
    open_ends = apply_peb_corner_joins(
        outlines,
        {1: (1, 2), 2: (2, 3)},
        {
            1: (Point3D(0, 0), Point3D(0, 4)),
            2: (Point3D(0, 4), Point3D(5, 5)),
        },
    )

    # Top rafter line reaches the outside flange; its bottom stops at the inside flange.
    assert outlines[2][0].x == -0.2
    assert round(outlines[2][0].y, 6) == 4.18
    assert outlines[2][2].x == 0.2
    assert round(outlines[2][2].y, 6) == 3.82
    # The inside column flange has priority and continues up to the top rafter line.
    assert round(outlines[1][3].x, 6) == 0.2
    assert round(outlines[1][3].y, 6) == 4.26
    assert open_ends == {(1, 1), (2, 0)}

def test_peb_corner_join_preserves_caps_at_same_slope_section_changes() -> None:
    outlines = {
        1: [
            Point3D(0.04, -0.2), Point3D(5.04, 0.8),
            Point3D(-0.04, 0.2), Point3D(4.96, 1.2),
        ],
        2: [
            Point3D(4.94, 1.3), Point3D(9.94, 2.3),
            Point3D(5.06, 0.7), Point3D(10.06, 1.7),
        ],
    }
    original = {beam: points.copy() for beam, points in outlines.items()}

    open_ends = apply_peb_corner_joins(
        outlines,
        {1: (1, 2), 2: (2, 3)},
        {
            1: (Point3D(0, 0), Point3D(5, 1)),
            2: (Point3D(5, 1), Point3D(10, 2)),
        },
    )

    assert open_ends == set()
    assert outlines == original


def test_peb_ridge_join_creates_one_vertical_cap() -> None:
    outlines = {1: [Point3D(0, 3.8), Point3D(5, 4.8), Point3D(0, 4.2), Point3D(5, 5.2)],
                2: [Point3D(5, 5.2), Point3D(10, 4.2), Point3D(5, 4.8), Point3D(10, 3.8)]}
    open_ends = apply_peb_corner_joins(outlines, {1: (1, 2), 2: (2, 3)},
        {1: (Point3D(0, 4), Point3D(5, 5)), 2: (Point3D(5, 5), Point3D(10, 4))})
    assert outlines[1][3] == outlines[2][0] == Point3D(5, 5.2)
    assert outlines[1][1] == outlines[2][2] == Point3D(5, 4.8)
    assert open_ends == {(2, 0)}


def test_peb_ridge_column_stops_at_rafter_bottom_chords() -> None:
    outlines = {1: [Point3D(0, 3.8), Point3D(5, 4.8), Point3D(0, 4.2), Point3D(5, 5.2)],
                2: [Point3D(5, 5.2), Point3D(10, 4.2), Point3D(5, 4.8), Point3D(10, 3.8)],
                3: [Point3D(4.8, 0), Point3D(4.8, 5), Point3D(5.2, 0), Point3D(5.2, 5)]}
    open_ends = apply_peb_corner_joins(outlines, {1: (1, 2), 2: (2, 3), 3: (4, 2)},
        {1: (Point3D(0, 4), Point3D(5, 5)), 2: (Point3D(5, 5), Point3D(10, 4)),
         3: (Point3D(5, 0), Point3D(5, 5))})
    assert round(outlines[3][1].y, 6) == 4.76
    assert round(outlines[3][3].y, 6) == 4.76
    assert open_ends == {(2, 0), (3, 1)}

def test_peb_column_priority_applies_to_multiple_non_ridge_members() -> None:
    outlines = {
        1: [Point3D(-0.2, 0), Point3D(-0.2, 4), Point3D(0.2, 0), Point3D(0.2, 4)],
        2: [Point3D(-0.1, 4.2), Point3D(4.9, 5.2), Point3D(0.1, 3.8), Point3D(5.1, 4.8)],
        3: [Point3D(0.1, 4.2), Point3D(-4.9, 5.7), Point3D(-0.1, 3.8), Point3D(-5.1, 5.3)],
    }
    open_ends = apply_peb_corner_joins(
        outlines, {1: (1, 2), 2: (2, 3), 3: (2, 4)},
        {1: (Point3D(0, 0), Point3D(0, 4)),
         2: (Point3D(0, 4), Point3D(5, 5)),
         3: (Point3D(0, 4), Point3D(-5, 5.5))},
    )
    assert round(outlines[2][0].x, 6) == -0.2
    assert round(outlines[2][2].x, 6) == 0.2
    assert round(outlines[3][0].x, 6) == 0.2
    assert round(outlines[3][2].x, 6) == -0.2
    assert open_ends == {(1, 1), (2, 0), (3, 0)}

def test_peb_horizontal_beam_stops_at_continuous_column_flange() -> None:
    outlines = {
        1: [Point3D(-0.2, 0), Point3D(-0.2, 2), Point3D(0.2, 0), Point3D(0.2, 2)],
        2: [Point3D(-0.2, 2), Point3D(-0.2, 4), Point3D(0.2, 2), Point3D(0.2, 4)],
        3: [Point3D(0, 2.1), Point3D(5, 2.1), Point3D(0, 1.9), Point3D(5, 1.9)],
    }
    open_ends = apply_peb_corner_joins(
        outlines, {1: (1, 2), 2: (2, 3), 3: (2, 4)},
        {1: (Point3D(0, 0), Point3D(0, 2)),
         2: (Point3D(0, 2), Point3D(0, 4)),
         3: (Point3D(0, 2), Point3D(5, 2))},
    )
    assert outlines[3][0] == Point3D(0.2, 2.1)
    assert outlines[3][2] == Point3D(0.2, 1.9)
    assert open_ends == {(1, 1), (2, 0), (3, 0)}

def test_connection_details_follow_section_dividing_face() -> None:
    stream = StringIO()
    writer = DxfWriter(stream)
    writer.header()
    write_connection_face_lines(
        writer,
        {1: [Point3D(0, 0.2), Point3D(5, 0.2),
             Point3D(0, -0.2), Point3D(5, -0.2)]},
        {1: (Point3D(0, 0), Point3D(5, 0))},
        {(1, 0)},
    )
    writer.footer()
    value = stream.getvalue()
    assert 'CONNECTION_DETAILS' in value
    assert value.count('8\nCONNECTION_DETAILS') == 1
    assert value.count('8\nCONNECTION_PLATES') == 2
    assert value.count('8\nCONNECTION_CLEATS') == 8
    assert value.count('8\nCONNECTION_BOLTS') == 28
    assert '0\nCIRCLE' not in value

def test_connection_face_ignores_analytical_split_in_continuous_column() -> None:
    stream = StringIO()
    writer = DxfWriter(stream)
    writer.header()
    write_connection_face_lines(
        writer,
        {
            1: [Point3D(-0.2, 0), Point3D(-0.2, 2), Point3D(0.2, 0), Point3D(0.2, 2)],
            2: [Point3D(-0.2, 2), Point3D(-0.2, 4), Point3D(0.2, 2), Point3D(0.2, 4)],
            3: [Point3D(0.2, 2.1), Point3D(5, 2.1), Point3D(0.2, 1.9), Point3D(5, 1.9)],
        },
        {
            1: (Point3D(0, 0), Point3D(0, 2)),
            2: (Point3D(0, 2), Point3D(0, 4)),
            3: (Point3D(0, 2), Point3D(5, 2)),
        },
        {(1, 1), (2, 0), (3, 0)},
    )
    writer.footer()
    value = stream.getvalue()
    assert value.count('8\nCONNECTION_DETAILS') == 1
    assert '10\n0.200000\n20\n2.100000' in value

def test_connection_face_uses_inner_flange_when_top_chord_is_extended() -> None:
    outlines = {
        1: [Point3D(-0.2, 0), Point3D(-0.2, 4.18),
            Point3D(0.2, 0), Point3D(0.2, 4.26)],
        2: [Point3D(-0.2, 4.18), Point3D(4.9, 5.2),
            Point3D(0.2, 3.82), Point3D(5.1, 4.8)],
    }
    stream = StringIO()
    writer = DxfWriter(stream)
    writer.header()
    write_connection_face_lines(
        writer, outlines,
        {1: (Point3D(0, 0), Point3D(0, 4)),
         2: (Point3D(0, 4), Point3D(5, 5))},
        {(1, 1), (2, 0)},
    )
    writer.footer()
    value = stream.getvalue()
    assert value.count('8\nCONNECTION_DETAILS') == 1
    assert '10\n0.200000\n20\n4.260000' in value
    assert '11\n0.200000\n21\n3.820000' in value

def test_ridge_connection_face_stays_vertical_with_ridge_column() -> None:
    stream = StringIO()
    writer = DxfWriter(stream)
    writer.header()
    write_connection_face_lines(
        writer,
        {
            1: [Point3D(0, 3.8), Point3D(5, 4.8), Point3D(0, 4.2), Point3D(5, 5.2)],
            2: [Point3D(5, 5.2), Point3D(10, 4.2), Point3D(5, 4.8), Point3D(10, 3.8)],
            3: [Point3D(4.8, 0), Point3D(4.8, 4.76), Point3D(5.2, 0), Point3D(5.2, 4.76)],
        },
        {
            1: (Point3D(0, 4), Point3D(5, 5)),
            2: (Point3D(5, 5), Point3D(10, 4)),
            3: (Point3D(5, 0), Point3D(5, 5)),
        },
        {(2, 0), (3, 1)},
    )
    writer.footer()
    value = stream.getvalue()
    assert value.count('8\nCONNECTION_DETAILS') == 1
    assert '10\n5.000000\n20\n4.800000' in value
    assert '11\n5.000000\n21\n5.200000' in value

def test_two_beams_at_column_use_both_flange_connection_faces() -> None:
    stream = StringIO()
    writer = DxfWriter(stream)
    writer.header()
    write_connection_face_lines(
        writer,
        {
            1: [Point3D(-0.2, 0), Point3D(-0.2, 4.3), Point3D(0.2, 0), Point3D(0.2, 4.3)],
            2: [Point3D(-5, 3.6), Point3D(-0.2, 4.1), Point3D(-5, 3.2), Point3D(-0.2, 3.7)],
            3: [Point3D(0.2, 4.1), Point3D(5, 3.6), Point3D(0.2, 3.7), Point3D(5, 3.2)],
        },
        {
            1: (Point3D(0, 0), Point3D(0, 4)),
            2: (Point3D(-5, 3.4), Point3D(0, 4)),
            3: (Point3D(0, 4), Point3D(5, 3.4)),
        },
        {(1, 1), (2, 1), (3, 0)},
    )
    writer.footer()
    value = stream.getvalue()
    assert value.count('8\nCONNECTION_DETAILS') == 2
    assert '10\n-0.200000' in value
    assert '10\n0.200000' in value


PRISMATIC = [Point3D(-0.3, 0), Point3D(-0.3, 4), Point3D(0.3, 0), Point3D(0.3, 4)]
# A member 600 deep at the start and 400 at the end, sloping on one face only.
TAPERED = [Point3D(-0.3, 0), Point3D(-0.3, 4), Point3D(0.3, 0), Point3D(0.1, 4)]


def test_the_faces_are_inset_by_the_thickness_towards_each_other() -> None:
    first, second = inner_face_lines(PRISMATIC, 0.05)
    assert first == (Point3D(-0.25, 0), Point3D(-0.25, 4))
    assert second == (Point3D(0.25, 0), Point3D(0.25, 4))


def test_a_tapered_flange_keeps_one_thickness_along_the_sloping_face() -> None:
    # Insetting by a share of the depth would wedge the flange open towards the
    # deep end; it is offset from its own face instead, so it stays parallel.
    outer_start, outer_end = TAPERED[2], TAPERED[3]
    (_, _), (inner_start, inner_end) = inner_face_lines(TAPERED, 0.05)
    shifts = [(inner_start.x - outer_start.x, inner_start.y - outer_start.y),
              (inner_end.x - outer_end.x, inner_end.y - outer_end.y)]
    assert shifts[0] == pytest.approx(shifts[1])
    assert hypot(*shifts[0]) == pytest.approx(0.05)
    assert inner_start.x < outer_start.x        # inset towards the far face


@pytest.mark.parametrize("thickness", [0.0, -0.01, 0.25])
def test_a_thickness_that_does_not_fit_the_drawn_depth_is_dropped(thickness) -> None:
    # 0.25 of a 0.6 deep section leaves no web between the two faces, so it is
    # read as a thickness that belongs to a view this drawing is not showing.
    assert inner_face_lines(PRISMATIC, thickness) == []


def envelope_dxf(property_type, name, thickness) -> str:
    stream = StringIO()
    write_member_envelope(
        DxfWriter(stream), PRISMATIC,
        SectionEnvelope(0.3, 0.3, property_type, thickness), name,
    )
    return stream.getvalue()


def test_an_i_section_draws_both_flanges_as_visible_faces() -> None:
    value = envelope_dxf(612, "ISMB400", 0.016)
    assert value.count("0\nLINE") == 6          # 2 faces + 2 caps + 2 flanges
    assert value.count("8\nMEMBER_SECTION") == 6
    assert "HIDDEN" not in value


def test_a_tube_draws_its_walls_hidden_behind_the_front_face() -> None:
    value = envelope_dxf(650, "TUB40030016", 0.016)
    assert value.count("0\nLINE") == 6
    assert value.count("6\nHIDDEN") == 2


def test_a_tapered_member_without_a_thickness_stays_a_plain_outline() -> None:
    value = envelope_dxf(675, "TAPERED", 0.0)
    assert value.count("0\nLINE") == 4
    assert value.count("8\nTAPERED_SECTION") == 4


class ThicknessStaad:
    """Only the surface member_envelope touches."""

    def __init__(self, section=(0, []), table=(0.15, 0.4, 0.0, 0.0, 0.0,
                                               0.0, 0.0, 0.0, 0.012, 0.007)):
        self.section, self.table = section, table

    def beam_property_all(self, beam_no):
        if self.table is None:
            raise OSError("no property table")
        return self.table

    def section_property_values(self, beam_no):
        return self.section


def test_the_thickness_falls_back_to_the_property_tables_flange() -> None:
    assert member_envelope(ThicknessStaad(), 1, 4.0).wall_thickness == 0.012


def test_a_built_up_section_takes_its_flange_thickness_from_its_values() -> None:
    values = [0.0] * 24
    values[1], values[4], values[5], values[6], values[7] = 0.005, 0.382, 0.632, 0.150, 0.008
    envelope = member_envelope(ThicknessStaad(section=(675, values)), 1, 4.0)
    assert envelope.wall_thickness == 0.008


def test_a_hollow_section_takes_its_wall_thickness_from_its_values() -> None:
    values = [0.0] * 24
    values[1], values[2], values[3] = 0.400, 0.300, 0.016
    envelope = member_envelope(ThicknessStaad(section=(650, values)), 1, 4.0)
    assert envelope.wall_thickness == 0.016


def test_an_unreadable_section_leaves_no_thickness() -> None:
    assert member_envelope(ThicknessStaad(table=None), 1, 4.0).wall_thickness == 0.0