from io import StringIO

from staad_ext.ifc import IfcWriter, new_guid
from staad_ext.macros.std_to_ifc import (
    CircularHollowProfile,
    _circular_hollow_profile,
    is_column,
    member_axes,
)
from staad_ext.models import Point3D

_GUID_ALPHABET = set("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$")


def test_guid_is_22_chars_from_ifc_alphabet() -> None:
    guid = new_guid()
    assert len(guid) == 22
    assert set(guid) <= _GUID_ALPHABET


def test_guids_are_unique() -> None:
    assert len({new_guid() for _ in range(50)}) == 50


def test_is_column_detects_vertical_member() -> None:
    assert is_column(Point3D(0, 0, 0), Point3D(0, 4, 0))
    assert not is_column(Point3D(0, 0, 0), Point3D(5, 0.1, 0))


def test_member_axes_perpendicular_to_member_and_unit_length() -> None:
    axis, width = member_axes(Point3D(0, 0, 0), Point3D(5, 0, 0), 0.0)
    dot = axis[0] * width[0] + axis[1] * width[1] + axis[2] * width[2]
    assert abs(dot) < 1e-9
    assert abs(sum(c * c for c in axis) - 1.0) < 1e-9
    assert abs(sum(c * c for c in width) - 1.0) < 1e-9


def test_member_axes_vertical_member_uses_global_x_reference() -> None:
    axis, width = member_axes(Point3D(0, 0, 0), Point3D(0, 4, 0), 0.0)
    assert axis == (0.0, 1.0, 0.0)
    # Width direction must still be perpendicular and horizontal for a vertical member.
    assert abs(width[1]) < 1e-9


def test_member_axes_beta_rotation_keeps_unit_length_and_orthogonality() -> None:
    axis, width = member_axes(Point3D(0, 0, 0), Point3D(5, 0, 0), 45.0)
    assert abs(sum(c * c for c in width) - 1.0) < 1e-9
    dot = axis[0] * width[0] + axis[1] * width[1] + axis[2] * width[2]
    assert abs(dot) < 1e-9


def test_circular_hollow_profile_for_pipe_property_type() -> None:
    profile = _circular_hollow_profile(2, [0.2, 0.18], "PIPE 200X180")
    assert isinstance(profile, CircularHollowProfile)
    assert round(profile.outer_radius, 4) == 0.1


def test_circular_hollow_profile_none_for_non_tube_type() -> None:
    assert _circular_hollow_profile(1, [0.2, 0.1], "W12X26") is None


def test_ifc_writer_produces_valid_step_wrapper() -> None:
    stream = StringIO()
    writer = IfcWriter(stream)
    writer.header("Test Model")
    writer.write("test.ifc")
    value = stream.getvalue()
    assert value.startswith("ISO-10303-21;\n")
    assert "FILE_SCHEMA(('IFC4'));" in value
    assert "IFCPROJECT(" in value
    assert "IFCBUILDINGSTOREY(" in value
    assert value.rstrip().endswith("END-ISO-10303-21;")
