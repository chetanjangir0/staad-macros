"""Analytical DXF renderer: section envelopes, per-member labels and GA joints.

The model read and geometry solve live in :mod:`staad_ext.framing`; everything
here turns a solved :class:`~staad_ext.framing.FramingModel` into DXF entities.
"""

from __future__ import annotations

from math import atan2, degrees, dist, hypot
from pathlib import Path

from staad_ext.dxf import DxfWriter, dxf_document
from staad_ext.framing import (
    FramingModel, Member, build_framing_model, index_groups, is_column,
    is_tapered, is_tube_or_pipe, line_intersection, move, offset_vector,
    rafter_rises_to_joint,
)
from staad_ext.models import ExportSettings, Point3D, SectionEnvelope

LABEL_HEIGHT_FACTOR = 0.035
MIN_LABEL_HEIGHT = 0.05
LABEL_WIDTH_FACTOR = 0.7
LABEL_MAX_SPAN_FACTOR = 0.8


def _build_section_color_palette() -> tuple[int, ...]:
    """Build a large, visually distinct AutoCAD Color Index (ACI) palette.

    ACI indices 10-249 fall into 24 hue "bands" of 10 shades each
    (band base, base+1, ... base+9), so cycling through every band's base
    shade first gives 24 maximally distinct hues; once those are used up,
    later shade offsets are layered in (still same 24 hues, but lighter or
    darker) so colors keep spreading out instead of repeating outright.
    True-color (DXF group 420) was tried and rejected: AutoCAD refuses to
    open an R12-header (AC1009) file containing it, so only plain ACI
    (group 62) values -- valid at any DXF version -- are used here.
    """
    basics = (1, 2, 3, 4, 5, 6, 8, 9)
    hue_bases = tuple(range(10, 250, 10))
    shade_offsets = (0, 5, 2, 7, 4, 9, 1, 6, 3, 8)
    banded = tuple(
        base + offset
        for offset in shade_offsets
        for base in hue_bases
        if base + offset <= 249
    )
    return basics + banded


SECTION_COLOR_PALETTE = _build_section_color_palette()


def assign_section_colors(model: FramingModel) -> dict[int, int]:
    """Assign a stable DXF (ACI) color index per beam based on its section size."""
    return {number: SECTION_COLOR_PALETTE[index % len(SECTION_COLOR_PALETTE)]
            for number, index in index_groups(model).items()}


def _label(member: Member) -> str:
    if member.property_type in {675, 680} and member.property_values:
        v = member.property_values
        try:
            if member.property_type == 675:
                d1, d2, tw, bf, tf = abs(v[4]), abs(v[5]), abs(v[1]), abs(v[6]), abs(v[7])
            else:
                d1, d2, tw = abs(v[0]), abs(v[2]), abs(v[1])
                bf, tf = max(abs(v[3]), abs(v[5])), max(abs(v[4]), abs(v[6]))
            return (f"W({round((d2 - 2 * tf) * 1000)}~{round((d1 - 2 * tf) * 1000)}"
                    f"x{round(tw * 1000)})\\P2F({round(bf * 1000)}x{round(tf * 1000)}); "
                    f"({member.length:.3f}M)")
        except (IndexError, TypeError, ValueError):
            pass
    return (f"{member.name}\\P({member.length:.3f}M)" if member.name
            else f"({member.length:.3f}M)")


def _line(writer: DxfWriter, layer: str, a: Point3D, b: Point3D, kind: str = "CONTINUOUS",
          color: int | None = None) -> None:
    writer.line(layer, a, b, kind, color=color)


def write_member_envelope(writer: DxfWriter, outline: list[Point3D], envelope: SectionEnvelope,
                          name: str, open_start: bool = False, open_end: bool = False,
                          color: int | None = None) -> None:
    """Draw one member's section envelope on the layer its section type earns."""
    tube = is_tube_or_pipe(envelope.property_type, name)
    tapered = is_tapered(envelope, name)
    layer = "TAPERED_SECTION" if tapered else ("TUBE_PIPE_SECTION" if tube else "MEMBER_SECTION")
    writer.envelope(
        layer, outline, open_start, open_end, color,
        hidden_fractions=(0.175, 0.825) if tube and not tapered else (),
        hidden_layer="TUBE_PIPE_SECTION",
    )


def write_connection_face_lines(
    writer: DxfWriter,
    outlines: dict[int, list[Point3D]],
    centerlines: dict[int, tuple[Point3D, Point3D]],
    open_ends: set[tuple[int, int]],
) -> None:
    """Draw each physical section-dividing face with schematic GA bolt holes."""
    by_joint: dict[tuple[float, float], list[tuple[int, int]]] = {}
    for beam, end in open_ends:
        joint = centerlines[beam][end]
        by_joint.setdefault((round(joint.x, 6), round(joint.y, 6)), []).append((beam, end))

    written: set[tuple[tuple[float, float], tuple[float, float]]] = set()
    for joint_key, connections in by_joint.items():
        joint = Point3D(*joint_key)
        all_connections: list[tuple[int, int]] = []
        for beam, centerline in centerlines.items():
            for end, point in enumerate(centerline):
                if (round(point.x, 6), round(point.y, 6)) == joint_key:
                    all_connections.append((beam, end))
        ridge_rafters = [
            connection for connection in all_connections
            if not is_column(centerlines[connection[0]])
            and rafter_rises_to_joint(
                centerlines[connection[0]], connection[1], joint
            )
        ]
        if len(ridge_rafters) == 2:
            far_x = [centerlines[beam][1 - end].x for beam, end in ridge_rafters]
            cap_points: list[tuple[Point3D, Point3D]] = []
            for beam, end in ridge_rafters:
                outline = outlines[beam]
                indices = (0, 2) if end == 0 else (1, 3)
                cap_points.append((outline[indices[0]], outline[indices[1]]))
            caps_on_ridge = all(
                abs(point.x - joint.x) <= 1e-6
                for cap in cap_points for point in cap
            )
            if caps_on_ridge and (far_x[0] - joint.x) * (far_x[1] - joint.x) < 0:
                first, second = cap_points[0]
                if hypot(second.x - first.x, second.y - first.y) > 1e-9:
                    _write_connection_detail(writer, first, second)
                continue

        columns = [
            connection for connection in connections
            if is_column(centerlines[connection[0]])
        ]
        non_columns = [connection for connection in connections if connection not in columns]
        candidates = non_columns or connections
        column_lines: list[tuple[Point3D, Point3D]] = []
        for column_beam, _ in columns:
            column = outlines[column_beam]
            column_lines.extend(((column[0], column[1]), (column[2], column[3])))

        for beam, end in candidates:
            outline = outlines[beam]
            first_index, second_index = ((0, 2) if end == 0 else (1, 3))
            first, second = outline[first_index], outline[second_index]
            if column_lines and not is_column(centerlines[beam]):
                far = centerlines[beam][1 - end]
                receiving_flange = min(
                    column_lines,
                    key=lambda line: abs((line[0].x + line[1].x) / 2 - far.x),
                )
                member_lines = ((outline[0], outline[1]), (outline[2], outline[3]))
                intersections = [
                    line_intersection(*member_line, *receiving_flange)
                    for member_line in member_lines
                ]
                if intersections[0] is None or intersections[1] is None:
                    continue
                first, second = intersections[0], intersections[1]
            if hypot(second.x - first.x, second.y - first.y) <= 1e-9:
                continue
            endpoints = sorted(((round(first.x, 6), round(first.y, 6)),
                                (round(second.x, 6), round(second.y, 6))))
            key = (endpoints[0], endpoints[1])
            if key in written:
                continue
            written.add(key)
            _write_connection_detail(writer, first, second)


def _detail_point(
    origin: Point3D,
    along: Point3D,
    across: Point3D,
    along_distance: float,
    across_distance: float,
) -> Point3D:
    return Point3D(
        origin.x + along.x * along_distance + across.x * across_distance,
        origin.y + along.y * along_distance + across.y * across_distance,
        origin.z,
    )


def _write_connection_detail(
    writer: DxfWriter,
    first: Point3D,
    second: Point3D,
) -> None:
    """Draw a schematic bolted splice assembly about the true dividing face."""
    _line(writer, "CONNECTION_DETAILS", first, second)
    length = hypot(second.x - first.x, second.y - first.y)
    if length <= 1e-9:
        return

    along = Point3D((second.x - first.x) / length, (second.y - first.y) / length)
    across = Point3D(-along.y, along.x)
    plate_half_width = max(min(length * 0.045, 0.04), 0.012)
    end_extension = length * 0.06
    rail_start = _detail_point(first, along, across, -end_extension, 0)
    rail_end = _detail_point(first, along, across, length + end_extension, 0)
    for side in (-1.0, 1.0):
        _line(
            writer,
            "CONNECTION_PLATES",
            _detail_point(rail_start, along, across, 0, side * plate_half_width),
            _detail_point(rail_end, along, across, 0, side * plate_half_width),
        )

    cleat_reach = max(length * 0.18, plate_half_width * 3.5)
    bolt_half_length = max(length * 0.025, 0.01)
    bolt_half_width = plate_half_width * 1.35
    for position in (0.14, 0.38, 0.62, 0.86):
        center_distance = length * position
        _line(
            writer,
            "CONNECTION_CLEATS",
            _detail_point(first, along, across, center_distance, -cleat_reach),
            _detail_point(first, along, across, center_distance, cleat_reach),
        )
        corners = [
            _detail_point(first, along, across, center_distance - bolt_half_length, -bolt_half_width),
            _detail_point(first, along, across, center_distance + bolt_half_length, -bolt_half_width),
            _detail_point(first, along, across, center_distance + bolt_half_length, bolt_half_width),
            _detail_point(first, along, across, center_distance - bolt_half_length, bolt_half_width),
        ]
        for index in range(4):
            _line(writer, "CONNECTION_BOLTS", corners[index], corners[(index + 1) % 4])
        for fraction in (-0.5, 0.0, 0.5):
            offset = bolt_half_width * fraction
            _line(
                writer,
                "CONNECTION_BOLTS",
                _detail_point(first, along, across, center_distance - bolt_half_length, offset - bolt_half_length),
                _detail_point(first, along, across, center_distance + bolt_half_length, offset + bolt_half_length),
            )

    triangle_depth = length * 0.14
    triangle_half_width = max(length * 0.22, plate_half_width * 3.5)
    for end_origin, direction in ((first, -1.0), (second, 1.0)):
        apex = _detail_point(end_origin, along, across, direction * triangle_depth, 0)
        for side in (-1.0, 1.0):
            shoulder = _detail_point(
                end_origin,
                along,
                across,
                -direction * triangle_depth * 0.35,
                side * triangle_half_width,
            )
            _line(writer, "CONNECTION_CLEATS", apex, shoulder)


def _write_label(writer: DxfWriter, start: Point3D, end: Point3D, value: str,
                 half_width: float, settings: ExportSettings, color: int = 7) -> None:
    if end.x < start.x:
        start, end = end, start
    length = dist((start.x, start.y, start.z), (end.x, end.y, end.z))
    height = max(length * LABEL_HEIGHT_FACTOR, half_width * 0.45) * settings.text_scale
    height = max(height, MIN_LABEL_HEIGHT)
    chars = len(value.replace(r"\P", ""))
    if chars and length > 0:
        height = max(min(height, length * LABEL_MAX_SPAN_FACTOR / (chars * LABEL_WIDTH_FACTOR)), MIN_LABEL_HEIGHT)
    unit = offset_vector(start, end)
    if unit.y < 0:
        unit = Point3D(-unit.x, -unit.y, -unit.z)
    middle = Point3D((start.x + end.x) / 2, (start.y + end.y) / 2, (start.z + end.z) / 2)
    location = move(middle, unit, half_width + height * 1.5)
    writer.colored_label("MEMBER_LABELS", location, height,
                         degrees(atan2(end.y - start.y, end.x - start.x)), value, unit, color)


def export_selected_members(staad, output: Path, settings: ExportSettings) -> int:
    model = build_framing_model(
        staad, settings.plane,
        settings.peb_corner_joins or settings.connection_face_lines,
    )
    if not model:
        return 0
    section_colors = assign_section_colors(model) if settings.color_by_section else {}
    with dxf_document(output) as writer:
        for member in model.members.values():
            color = section_colors.get(member.number)
            writer.line("MEMBER_CENTERLINE", member.start, member.end, "DASHED")
            write_member_envelope(
                writer, member.outline, member.envelope, member.name,
                (member.number, 0) in model.open_ends,
                (member.number, 1) in model.open_ends,
                color=color,
            )
            if settings.write_labels:
                _write_label(writer, member.start, member.end, _label(member),
                             member.half_width, settings, color or 7)
        if settings.connection_face_lines:
            write_connection_face_lines(
                writer, model.outlines(), model.centerlines(), model.open_ends
            )
    return len(model.members)
