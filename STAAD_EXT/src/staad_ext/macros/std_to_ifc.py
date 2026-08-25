from __future__ import annotations

from dataclasses import dataclass
from math import cos, hypot, radians, sin
from pathlib import Path

from staad_ext.ifc import IfcWriter
from staad_ext.framing import is_tube_or_pipe
from staad_ext.models import IfcExportSettings, Point3D
from staad_ext.openstaad import OpenStaad

MIN_PROFILE_DIM = 0.05
Vector = tuple[float, float, float]


def _sub(a: Point3D, b: Point3D) -> Vector:
    return (a.x - b.x, a.y - b.y, a.z - b.z)


def _length(v: Vector) -> float:
    return hypot(hypot(v[0], v[1]), v[2])


def _normalize(v: Vector) -> Vector:
    length = _length(v)
    if length < 1e-9:
        return (1.0, 0.0, 0.0)
    return (v[0] / length, v[1] / length, v[2] / length)


def _cross(a: Vector, b: Vector) -> Vector:
    return (a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0])


def _dot(a: Vector, b: Vector) -> float:
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _rotate_about_axis(vector: Vector, axis: Vector, angle_rad: float) -> Vector:
    """Rotate `vector` about a unit `axis` by `angle_rad` (Rodrigues' rotation formula)."""
    cos_a, sin_a = cos(angle_rad), sin(angle_rad)
    dot_av = _dot(axis, vector)
    cross_av = _cross(axis, vector)
    return tuple(
        vector[i] * cos_a + cross_av[i] * sin_a + axis[i] * dot_av * (1 - cos_a)
        for i in range(3)
    )


def member_axes(start: Point3D, end: Point3D, beta_deg: float) -> tuple[Vector, Vector]:
    """Return (member_axis, width_direction) local axes for a member, following the
    same default local-axis convention STAAD.Pro uses: the local x-y plane contains
    the global Y axis (or, for vertical members, the global X axis), rotated about
    the member axis by the assigned beta angle.
    """
    axis = _normalize(_sub(end, start))
    reference = (1.0, 0.0, 0.0) if abs(axis[1]) > 0.999 else (0.0, 1.0, 0.0)
    width = _normalize(_cross(reference, axis))
    if abs(beta_deg) > 1e-9:
        width = _rotate_about_axis(width, axis, radians(beta_deg))
    return axis, width


def is_column(start: Point3D, end: Point3D) -> bool:
    dx, dy, dz = end.x - start.x, end.y - start.y, end.z - start.z
    horizontal = hypot(dx, dz)
    return abs(dy) > horizontal * 1.5


@dataclass(frozen=True, slots=True)
class RectangularProfile:
    width: float
    depth: float


@dataclass(frozen=True, slots=True)
class CircularHollowProfile:
    outer_radius: float
    wall_thickness: float


Profile = RectangularProfile | CircularHollowProfile


def _circular_hollow_profile(property_type: int, values: list[float], name: str) -> CircularHollowProfile | None:
    upper = name.upper()
    v = [abs(value) for value in values]
    try:
        if property_type == 2 and "PIPE" in upper and len(v) > 1:
            return CircularHollowProfile(v[0] / 2, max((v[0] - v[1]) / 2, 0.001))
        if property_type in {660, 655} and len(v) > 2:
            return CircularHollowProfile(v[1] / 2, v[2])
        if property_type == 695 and len(v) > 2:
            return CircularHollowProfile(v[1] / 2, max((v[1] - v[2]) / 2, 0.001))
    except IndexError:
        return None
    return None


def member_profile(staad: OpenStaad, beam_no: int, name: str) -> Profile:
    width = depth = 0.0
    try:
        width, depth, *_ = staad.beam_property_all(beam_no)
        width, depth = abs(width), abs(depth)
    except (OSError, TypeError, ValueError):
        pass

    property_type = 0
    values: list[float] = []
    try:
        property_type, values = staad.section_property_values(beam_no)
    except (OSError, TypeError, ValueError):
        pass

    if is_tube_or_pipe(property_type, name):
        circular = _circular_hollow_profile(property_type, values, name)
        if circular is not None and circular.outer_radius > 0:
            return circular

    if width <= 0 or depth <= 0:
        default = max(width, depth, MIN_PROFILE_DIM)
        width, depth = default, default
    return RectangularProfile(max(width, MIN_PROFILE_DIM), max(depth, MIN_PROFILE_DIM))


def export_structure(staad: OpenStaad, output: Path, settings: IfcExportSettings) -> int:
    beams = staad.selected_beams() if settings.selected_only else staad.all_beams()
    if not beams:
        return 0

    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="ascii", newline="\r\n") as stream:
        writer = IfcWriter(stream)
        writer.header(output.stem)

        exported = 0
        related: list[int] = []
        for beam in beams:
            start_node, end_node = staad.member_incidence(beam)
            if start_node <= 0 or end_node <= 0:
                continue
            start, end = staad.node_coordinates(start_node), staad.node_coordinates(end_node)
            length = _length(_sub(end, start))
            if length <= 1e-6:
                continue

            name = staad.section_name(beam)
            profile = member_profile(staad, beam, name)
            beta = staad.beta_angle(beam)
            axis, width_dir = member_axes(start, end, beta)

            location = writer.point(start.x, start.y, start.z)
            axis_dir = writer.direction(*axis)
            ref_dir = writer.direction(*width_dir)
            placement = writer.local_placement(
                writer.storey_id, writer.axis2placement3d(location, axis_dir, ref_dir)
            )

            if isinstance(profile, CircularHollowProfile):
                profile_id = writer.circle_hollow_profile(
                    name, profile.outer_radius, profile.wall_thickness
                )
            else:
                profile_id = writer.rectangle_profile(name, profile.width, profile.depth)

            solid = writer.extruded_area_solid(profile_id, length)
            shape = writer.product_definition_shape(writer.shape_representation(solid))
            member_id = writer.member(
                is_column(start, end), f"M{beam} {name}".strip(), placement, shape
            )
            related.append(member_id)
            exported += 1

        writer.contained_in_storey(related)
        writer.write(output.name)
    return exported
