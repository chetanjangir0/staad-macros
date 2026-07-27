from io import StringIO

from staad_ext.dxf import DxfWriter
from staad_ext.macros.std_to_dxf import (
    apply_peb_corner_joins, is_tapered, project, write_connection_face_lines,
)
from staad_ext.models import Point3D, SectionEnvelope, ViewPlane


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