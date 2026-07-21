from io import StringIO

from staad_ext.dxf import DxfWriter
from staad_ext.macros.std_to_dxf import is_tapered, project
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
