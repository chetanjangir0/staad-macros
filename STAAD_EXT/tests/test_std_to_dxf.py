from io import StringIO

from staad_ext.dxf import DxfWriter
from staad_ext.macros.std_to_dxf import apply_peb_corner_joins, is_tapered, project
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
    apply_peb_corner_joins(
        outlines,
        {1: (1, 2), 2: (2, 3)},
        {
            1: (Point3D(0, 0), Point3D(0, 4)),
            2: (Point3D(0, 4), Point3D(5, 5)),
        },
    )

    rafter_join = {outlines[2][0], outlines[2][2]}
    assert {round(point.x, 6) for point in rafter_join} == {-0.2, 0.2}
    assert outlines[1][1] in rafter_join
    assert outlines[1][3] in rafter_join