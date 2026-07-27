from __future__ import annotations

from math import atan2, degrees, dist, hypot
from pathlib import Path
from typing import Iterable

from staad_ext.dxf import DxfWriter
from staad_ext.models import ExportSettings, Point3D, SectionEnvelope, ViewPlane
from staad_ext.openstaad import OpenStaad, OpenStaadError

MIN_SECTION_HALF_WIDTH = 0.05
LABEL_HEIGHT_FACTOR = 0.035
MIN_LABEL_HEIGHT = 0.05
LABEL_WIDTH_FACTOR = 0.7
LABEL_MAX_SPAN_FACTOR = 0.8
TUBE_PIPE_TYPES = {650, 654, 655, 660, 675, 695, 696}


def project(point: Point3D, plane: ViewPlane) -> Point3D:
    if plane is ViewPlane.YZ:
        return Point3D(point.z, point.y)
    if plane is ViewPlane.ZX:
        return Point3D(point.z, point.x)
    return Point3D(point.x, point.y)


def offset_vector(start: Point3D, end: Point3D) -> Point3D:
    dx, dy = end.x - start.x, end.y - start.y
    length = hypot(dx, dy)
    return Point3D(-dy / length, dx / length) if length > 1e-6 else Point3D(1.0, 0.0)


def is_tube_or_pipe(property_type: int, name: str) -> bool:
    upper = name.upper()
    return property_type in TUBE_PIPE_TYPES or any(token in upper for token in ("TUBE", "PIPE", "CHS", "RHS", "SHS"))


def is_tapered(envelope: SectionEnvelope, name: str) -> bool:
    maximum = max(envelope.start_half_width, envelope.end_half_width)
    unequal = maximum > 0 and abs(envelope.start_half_width - envelope.end_half_width) > maximum * 0.1
    return "TAPER" in name.upper() or envelope.property_type in {675, 680} or unequal


def _envelope_from_values(property_type: int, values: list[float], default: float) -> tuple[float, float]:
    v = [abs(value) for value in values]
    maximum = 0.0
    if property_type in {610, 611, 612, 613, 614, 615, 630, 631, 633, 620, 656,
                         640, 641, 642, 643, 644, 645, 646, 650, 654, 662, 663, 664, 666}:
        maximum = max(v[1], v[2])
    elif property_type == 616:
        maximum = max(v[0], v[1])
    elif property_type in {660, 655, 668, 695}:
        maximum = v[1]
    elif property_type == 675:
        if v[5] > 0 and v[4] > 0:
            return max(v[5] / 2, MIN_SECTION_HALF_WIDTH), max(v[4] / 2, MIN_SECTION_HALF_WIDTH)
    elif property_type == 680:
        if v[0] > 0 and v[2] > 0:
            return max(v[0] / 2, MIN_SECTION_HALF_WIDTH), max(v[2] / 2, MIN_SECTION_HALF_WIDTH)
    elif property_type == 671:
        maximum = v[4]
    elif property_type in {672, 674, 673, 699}:
        maximum = max(v[4], v[5])
    elif property_type == 676:
        maximum = max(v[6], v[7])
    elif property_type in {690, 691, 694, 696, 697}:
        maximum = max(v[1], v[3])
    elif property_type in {692, 693}:
        maximum = max(v[1], v[2])
    elif property_type == 698:
        maximum = max(v[:7])
    half = max(maximum / 2, MIN_SECTION_HALF_WIDTH) if maximum > 0 else default
    return half, half


def member_envelope(staad: OpenStaad, beam_no: int, length: float) -> SectionEnvelope:
    default = max(length * 0.0125, MIN_SECTION_HALF_WIDTH)
    start = end = default
    try:
        width, depth, *_ = staad.beam_property_all(beam_no)
        candidate = max(width, depth) / 2
        if candidate > 0:
            start = end = max(candidate, MIN_SECTION_HALF_WIDTH)
    except (OSError, TypeError, ValueError):
        pass
    property_type = 0
    try:
        property_type, values = staad.section_property_values(beam_no)
        start, end = _envelope_from_values(property_type, values, start)
    except (OSError, TypeError, ValueError):
        pass
    return SectionEnvelope(start, end, property_type)


def _mm(value: float) -> str:
    return f"{abs(value) * 1000:.3f}".rstrip("0").rstrip(".")


def _tube_pipe_name(staad: OpenStaad, beam_no: int, envelope: SectionEnvelope, name: str) -> str:
    if not is_tube_or_pipe(envelope.property_type, name):
        return name
    upper = name.upper()
    try:
        pt, v = staad.section_property_values(beam_no)
        if pt == 2 and "PIPE" in upper:
            return f"PIPE (OD {_mm(v[0])} ID {_mm(max(v[1], 0))})"
        if pt == 2 and any(x in upper for x in ("TUBE", "RHS", "SHS")):
            return f"TUBE ({_mm(v[2])}x{_mm(v[1])}x{_mm(v[0])})"
        if pt in {650, 654, 696}:
            return f"TUBE ({_mm(v[1])}x{_mm(v[2])}x{_mm(v[3])})"
        if pt in {660, 655}:
            return f"PIPE (OD {_mm(v[1])} ID {_mm(max(v[1] - 2 * v[2], 0))})"
        if pt == 695:
            return f"PIPE (OD {_mm(v[1])} ID {_mm(max(v[2], 0))})"
    except (OSError, TypeError, ValueError):
        pass
    try:
        width, depth, *_, tf, tw = staad.beam_property_all(beam_no)
        thickness = max(abs(tf), abs(tw))
        if any(x in upper for x in ("TUBE", "RHS", "SHS")) and min(abs(depth), abs(width), thickness) > 0:
            return f"TUBE ({_mm(depth)}x{_mm(width)}x{_mm(thickness)})"
        if any(x in upper for x in ("PIPE", "CHS")):
            od = max(abs(depth), abs(width))
            if od > 0:
                return f"PIPE (OD {_mm(od)} ID {_mm(max(od - 2 * thickness, 0))})"
    except (OSError, TypeError, ValueError):
        pass
    return name


def _label(staad: OpenStaad, beam_no: int, name: str, length: float, property_type: int) -> str:
    if property_type in {675, 680}:
        try:
            pt, v = staad.section_property_values(beam_no)
            if pt == 675:
                d1, d2, tw, bf, tf = abs(v[4]), abs(v[5]), abs(v[1]), abs(v[6]), abs(v[7])
            else:
                d1, d2, tw = abs(v[0]), abs(v[2]), abs(v[1])
                bf, tf = max(abs(v[3]), abs(v[5])), max(abs(v[4]), abs(v[6]))
            return f"W({round((d2 - 2 * tf) * 1000)}~{round((d1 - 2 * tf) * 1000)}x{round(tw * 1000)})\\P2F({round(bf * 1000)}x{round(tf * 1000)}); ({length:.3f}M)"
        except (OSError, TypeError, ValueError):
            pass
    return f"{name}\\P({length:.3f}M)" if name else f"({length:.3f}M)"


def _line(writer: DxfWriter, layer: str, a: Point3D, b: Point3D, kind: str = "CONTINUOUS") -> None:
    writer.line(layer, a, b, kind)


def _envelope_points(start: Point3D, end: Point3D, envelope: SectionEnvelope,
                     name: str, fixed_width: float, center_x: float) -> list[Point3D]:
    unit = offset_vector(start, end)
    if is_tapered(envelope, name):
        dx, dy = end.x - start.x, end.y - start.y
        if abs(dy) > abs(dx) * 1.5:
            desired_x = -1 if (start.x + end.x) / 2 < center_x else 1
            sign = 1 if unit.x * desired_x >= 0 else -1
        else:
            sign = 1 if unit.y >= 0 else -1
        offsets = (sign * fixed_width, sign * fixed_width - sign * envelope.start_half_width * 2,
                   sign * fixed_width - sign * envelope.end_half_width * 2)
        sf, st, et = offsets
        p1, p2 = _move(start, unit, sf), _move(end, unit, sf)
        p3, p4 = _move(start, unit, st), _move(end, unit, et)
    else:
        p1, p2 = _move(start, unit, envelope.start_half_width), _move(end, unit, envelope.end_half_width)
        p3, p4 = _move(start, unit, -envelope.start_half_width), _move(end, unit, -envelope.end_half_width)
    return [p1, p2, p3, p4]


def _write_envelope(writer: DxfWriter, outline: list[Point3D], envelope: SectionEnvelope,
                    name: str, open_start: bool = False, open_end: bool = False) -> None:
    p1, p2, p3, p4 = outline
    layer = "TUBE_PIPE_SECTION" if is_tube_or_pipe(envelope.property_type, name) else "MEMBER_SECTION"
    if is_tapered(envelope, name):
        layer = "TAPERED_SECTION"
    edges = [(p1, p2), (p3, p4)]
    if not open_start:
        edges.append((p1, p3))
    if not open_end:
        edges.append((p2, p4))
    for a, b in edges:
        _line(writer, layer, a, b)
    if is_tube_or_pipe(envelope.property_type, name) and not is_tapered(envelope, name):
        for fraction in (0.175, 0.825):
            hidden_start = Point3D(p1.x + (p3.x - p1.x) * fraction,
                                   p1.y + (p3.y - p1.y) * fraction)
            hidden_end = Point3D(p2.x + (p4.x - p2.x) * fraction,
                                 p2.y + (p4.y - p2.y) * fraction)
            _line(writer, "TUBE_PIPE_SECTION", hidden_start, hidden_end, "HIDDEN")


def _line_intersection(a: Point3D, b: Point3D, c: Point3D, d: Point3D) -> Point3D | None:
    """Return the intersection of two infinite 2D lines."""
    abx, aby = b.x - a.x, b.y - a.y
    cdx, cdy = d.x - c.x, d.y - c.y
    denominator = abx * cdy - aby * cdx
    if abs(denominator) < 1e-9:
        return None
    t = ((c.x - a.x) * cdy - (c.y - a.y) * cdx) / denominator
    return Point3D(a.x + t * abx, a.y + t * aby)


def _is_column(centerline: tuple[Point3D, Point3D]) -> bool:
    start, end = centerline
    return abs(end.y - start.y) > abs(end.x - start.x) * 1.5


def _centerlines_parallel(
    first: tuple[Point3D, Point3D], second: tuple[Point3D, Point3D]
) -> bool:
    a, b = first
    c, d = second
    abx, aby = b.x - a.x, b.y - a.y
    cdx, cdy = d.x - c.x, d.y - c.y
    lengths = hypot(abx, aby) * hypot(cdx, cdy)
    return lengths <= 1e-12 or abs(abx * cdx + aby * cdy) / lengths >= 0.9999


def _apply_column_rafter_join(
    column: list[Point3D],
    rafter: list[Point3D],
    column_end: int,
    rafter_end: int,
    rafter_centerline: tuple[Point3D, Point3D],
) -> bool:
    """Apply the flange priorities used at a PEB eave joint."""
    column_lines = ((column[0], column[1]), (column[2], column[3]))
    rafter_lines = ((rafter[0], rafter[1]), (rafter[2], rafter[3]))
    top_rafter = max(range(2), key=lambda side: (rafter_lines[side][0].y + rafter_lines[side][1].y) / 2)
    bottom_rafter = 1 - top_rafter

    far_rafter_x = rafter_centerline[1 - rafter_end].x
    inner_column = min(
        range(2),
        key=lambda side: abs((column_lines[side][0].x + column_lines[side][1].x) / 2 - far_rafter_x),
    )
    outer_column = 1 - inner_column

    top_outer = _line_intersection(*rafter_lines[top_rafter], *column_lines[outer_column])
    top_inner = _line_intersection(*rafter_lines[top_rafter], *column_lines[inner_column])
    bottom_inner = _line_intersection(*rafter_lines[bottom_rafter], *column_lines[inner_column])
    if top_outer is None or top_inner is None or bottom_inner is None:
        return False

    joint = rafter_centerline[rafter_end]
    column_width = hypot(column[0].x - column[2].x, column[0].y - column[2].y)
    if max(hypot(point.x - joint.x, point.y - joint.y)
           for point in (top_outer, top_inner, bottom_inner)) > column_width * 20:
        return False

    column_indices = (0, 2) if column_end == 0 else (1, 3)
    rafter_indices = (0, 2) if rafter_end == 0 else (1, 3)
    column[column_indices[outer_column]] = top_outer
    column[column_indices[inner_column]] = top_inner
    rafter[rafter_indices[top_rafter]] = top_outer
    rafter[rafter_indices[bottom_rafter]] = bottom_inner
    return True


def _rafter_rises_to_joint(centerline: tuple[Point3D, Point3D], end: int,
                            joint: Point3D) -> bool:
    far = centerline[1 - end]
    return far.y < joint.y - 1e-9 and abs(far.x - joint.x) > 1e-9


def _apply_rafter_ridge_join(outlines: dict[int, list[Point3D]],
                              centerlines: dict[int, tuple[Point3D, Point3D]],
                              first: tuple[int, int], second: tuple[int, int],
                              column: tuple[int, int] | None = None) -> set[tuple[int, int]]:
    """Create one vertical ridge cap and stop an optional column at the soffit."""
    beam_a, end_a = first
    beam_b, end_b = second
    joint = centerlines[beam_a][end_a]
    vertical = (Point3D(joint.x, joint.y - 1), Point3D(joint.x, joint.y + 1))
    ridge_points: list[list[Point3D]] = []
    for beam in (beam_a, beam_b):
        outline = outlines[beam]
        a = _line_intersection(outline[0], outline[1], *vertical)
        b = _line_intersection(outline[2], outline[3], *vertical)
        if a is None or b is None:
            return set()
        ridge_points.append([a, b])
    top = Point3D(joint.x, max(p.y for pair in ridge_points for p in pair))
    bottom = Point3D(joint.x, min(p.y for pair in ridge_points for p in pair))
    for (beam, end), points in zip((first, second), ridge_points):
        indices = (0, 2) if end == 0 else (1, 3)
        top_side = 0 if points[0].y >= points[1].y else 1
        outlines[beam][indices[top_side]] = top
        outlines[beam][indices[1 - top_side]] = bottom
    open_ends = {(beam_b, end_b)}
    if column is None:
        return open_ends
    column_beam, column_end = column
    column_outline = outlines[column_beam]
    column_indices = (0, 2) if column_end == 0 else (1, 3)
    for side, column_index in enumerate(column_indices):
        column_line = ((column_outline[0], column_outline[1]) if side == 0
                       else (column_outline[2], column_outline[3]))
        flange_x = column_outline[column_index].x
        rafter_beam, rafter_end = min(
            (first, second),
            key=lambda item: abs(centerlines[item[0]][1 - item[1]].x - flange_x),
        )
        rafter = outlines[rafter_beam]
        rafter_lines = ((rafter[0], rafter[1]), (rafter[2], rafter[3]))
        bottom_line = min(rafter_lines, key=lambda line: (line[0].y + line[1].y) / 2)
        intersection = _line_intersection(*column_line, *bottom_line)
        if intersection is not None:
            column_outline[column_index] = intersection
    open_ends.add((column_beam, column_end))
    return open_ends

def _apply_continuous_column_join(
    outlines: dict[int, list[Point3D]],
    centerlines: dict[int, tuple[Point3D, Point3D]],
    columns: list[tuple[int, int]],
    members: list[tuple[int, int]],
) -> set[tuple[int, int]]:
    """Stop framing members at the near flange of a continuous column."""
    open_ends = set(columns)
    column_lines = []
    for column_beam, _ in columns:
        outline = outlines[column_beam]
        column_lines.extend(((outline[0], outline[1]), (outline[2], outline[3])))

    for member_beam, member_end in members:
        member = outlines[member_beam]
        member_indices = (0, 2) if member_end == 0 else (1, 3)
        far = centerlines[member_beam][1 - member_end]
        near_flange = min(
            column_lines,
            key=lambda line: abs((line[0].x + line[1].x) / 2 - far.x),
        )
        member_lines = ((member[0], member[1]), (member[2], member[3]))
        intersections = [_line_intersection(*line, *near_flange) for line in member_lines]
        if any(point is None for point in intersections):
            continue
        for index, point in zip(member_indices, intersections):
            member[index] = point
        open_ends.add((member_beam, member_end))
    return open_ends

def apply_peb_corner_joins(
    outlines: dict[int, list[Point3D]],
    incidences: dict[int, tuple[int, int]],
    centerlines: dict[int, tuple[Point3D, Point3D]],
) -> set[tuple[int, int]]:
    """Join envelope edges at simple, non-parallel PEB frame corners."""
    open_ends: set[tuple[int, int]] = set()
    by_node: dict[int, list[tuple[int, int]]] = {}
    for beam, nodes in incidences.items():
        by_node.setdefault(nodes[0], []).append((beam, 0))
        by_node.setdefault(nodes[1], []).append((beam, 1))

    for connections in by_node.values():
        if len(connections) < 2:
            continue
        joint = centerlines[connections[0][0]][connections[0][1]]
        rafters = [item for item in connections
                   if not _is_column(centerlines[item[0]])
                   and _rafter_rises_to_joint(centerlines[item[0]], item[1], joint)]
        columns = [item for item in connections if _is_column(centerlines[item[0]])]
        if len(rafters) == 2 and len(columns) == len(connections) - 2:
            far_x = [centerlines[beam][1 - end].x for beam, end in rafters]
            if (far_x[0] - joint.x) * (far_x[1] - joint.x) < 0:
                open_ends.update(_apply_rafter_ridge_join(
                    outlines, centerlines, rafters[0], rafters[1],
                    columns[0] if columns else None,
                ))
                continue
        non_columns = [item for item in connections if item not in columns]
        if len(columns) >= 2 and non_columns:
            column_far_points = [centerlines[beam][1 - end] for beam, end in columns]
            spans_joint = any(
                (first.y - joint.y) * (second.y - joint.y) < 0
                for index, first in enumerate(column_far_points)
                for second in column_far_points[index + 1:]
            )
            if spans_joint:
                open_ends.update(_apply_continuous_column_join(
                    outlines, centerlines, columns, non_columns,
                ))
                continue
        if len(columns) == 1 and non_columns:
            column_beam, column_end = columns[0]
            column_outline = outlines[column_beam]
            original_column = column_outline.copy()
            column_indices = (0, 2) if column_end == 0 else (1, 3)
            candidates: list[list[Point3D]] = []
            for member_beam, member_end in non_columns:
                column_outline[:] = original_column
                if _apply_column_rafter_join(
                    column_outline, outlines[member_beam], column_end, member_end,
                    centerlines[member_beam],
                ):
                    candidates.append([column_outline[index] for index in column_indices])
                    open_ends.add((member_beam, member_end))
            column_outline[:] = original_column
            if candidates:
                far = centerlines[column_beam][1 - column_end]
                for side, index in enumerate(column_indices):
                    column_outline[index] = max(
                        (candidate[side] for candidate in candidates),
                        key=lambda point: hypot(point.x - far.x, point.y - far.y),
                    )
                open_ends.add((column_beam, column_end))
            continue

        if len(connections) != 2:
            continue
        (beam_a, end_a), (beam_b, end_b) = connections
        if _centerlines_parallel(centerlines[beam_a], centerlines[beam_b]):
            continue

        a, b = outlines[beam_a], outlines[beam_b]
        a_lines = ((a[0], a[1]), (a[2], a[3]))
        b_lines = ((b[0], b[1]), (b[2], b[3]))
        intersections = [
            [_line_intersection(*a_line, *b_line) for b_line in b_lines]
            for a_line in a_lines
        ]
        if any(point is None for row in intersections for point in row):
            continue

        joint = centerlines[beam_a][end_a]
        direct = sum(hypot(point.x - joint.x, point.y - joint.y)
                     for point in (intersections[0][0], intersections[1][1]))
        crossed = sum(hypot(point.x - joint.x, point.y - joint.y)
                      for point in (intersections[0][1], intersections[1][0]))
        pairing = (0, 1) if direct <= crossed else (1, 0)
        joined = [intersections[0][pairing[0]], intersections[1][pairing[1]]]

        widths = [hypot(a[0].x - a[2].x, a[0].y - a[2].y),
                  hypot(b[0].x - b[2].x, b[0].y - b[2].y)]
        if max(hypot(point.x - joint.x, point.y - joint.y) for point in joined) > max(widths) * 20:
            continue

        a_indices = (0, 2) if end_a == 0 else (1, 3)
        b_indices = (0, 2) if end_b == 0 else (1, 3)
        for side in range(2):
            a[a_indices[side]] = joined[side]
            b[b_indices[pairing[side]]] = joined[side]
        open_ends.update(((beam_a, end_a), (beam_b, end_b)))
    return open_ends
def write_connection_face_lines(
    writer: DxfWriter,
    outlines: dict[int, list[Point3D]],
    centerlines: dict[int, tuple[Point3D, Point3D]],
    open_ends: set[tuple[int, int]],
) -> None:
    """Draw each physical section-dividing cap once as the connection face."""
    by_joint: dict[tuple[float, float], list[tuple[int, int]]] = {}
    for beam, end in open_ends:
        joint = centerlines[beam][end]
        by_joint.setdefault((round(joint.x, 6), round(joint.y, 6)), []).append((beam, end))

    written: set[tuple[tuple[float, float], tuple[float, float]]] = set()
    for connections in by_joint.values():
        columns = [
            connection for connection in connections
            if _is_column(centerlines[connection[0]])
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
            if column_lines and not _is_column(centerlines[beam]):
                far = centerlines[beam][1 - end]
                receiving_flange = min(
                    column_lines,
                    key=lambda line: abs((line[0].x + line[1].x) / 2 - far.x),
                )
                member_lines = ((outline[0], outline[1]), (outline[2], outline[3]))
                intersections = [
                    _line_intersection(*member_line, *receiving_flange)
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
            _line(writer, "CONNECTION_DETAILS", first, second)

def _move(point: Point3D, vector: Point3D, amount: float) -> Point3D:
    return Point3D(point.x + vector.x * amount, point.y + vector.y * amount, point.z + vector.z * amount)


def _write_label(writer: DxfWriter, start: Point3D, end: Point3D, value: str,
                 half_width: float, settings: ExportSettings) -> None:
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
    location = _move(middle, unit, half_width + height * 1.5)
    writer.colored_label("MEMBER_LABELS", location, height,
                         degrees(atan2(end.y - start.y, end.x - start.x)), value, unit)


def _connected_fixed_width(staad: OpenStaad, beam_numbers: list[int], target: int,
                           envelopes: dict[int, SectionEnvelope], names: dict[int, str],
                           points: dict[int, tuple[Point3D, Point3D]]) -> float:
    target_env = envelopes[target]
    result = max(target_env.start_half_width, target_env.end_half_width)
    if not is_tapered(target_env, names[target]):
        return result
    included = {target}
    changed = True
    incidences = {beam: staad.member_incidence(beam) for beam in beam_numbers}
    while changed:
        changed = False
        for beam in beam_numbers:
            if beam in included:
                continue
            for member in tuple(included):
                if set(incidences[beam]).isdisjoint(incidences[member]):
                    continue
                a, b = points[beam]
                c, d = points[member]
                len1, len2 = hypot(b.x-a.x, b.y-a.y), hypot(d.x-c.x, d.y-c.y)
                parallel = len1 > 1e-6 and len2 > 1e-6 and abs(((b.x-a.x)*(d.x-c.x)+(b.y-a.y)*(d.y-c.y))/(len1*len2)) >= 0.9999
                if parallel and is_tapered(envelopes[beam], names[beam]):
                    included.add(beam)
                    changed = True
                    break
    return max(max(envelopes[b].start_half_width, envelopes[b].end_half_width) for b in included)


def export_selected_members(staad: OpenStaad, output: Path, settings: ExportSettings) -> int:
    beams = staad.selected_beams()
    if not beams:
        return 0
    points: dict[int, tuple[Point3D, Point3D]] = {}
    incidences: dict[int, tuple[int, int]] = {}
    lengths: dict[int, float] = {}
    envelopes: dict[int, SectionEnvelope] = {}
    names: dict[int, str] = {}
    for beam in beams:
        start_node, end_node = staad.member_incidence(beam)
        if start_node <= 0 or end_node <= 0:
            continue
        raw_start, raw_end = staad.node_coordinates(start_node), staad.node_coordinates(end_node)
        points[beam] = project(raw_start, settings.plane), project(raw_end, settings.plane)
        incidences[beam] = (start_node, end_node)
        try:
            lengths[beam] = staad.beam_length(beam)
            if lengths[beam] <= 0:
                raise ValueError
        except (OSError, TypeError, ValueError):
            lengths[beam] = dist((raw_start.x, raw_start.y, raw_start.z), (raw_end.x, raw_end.y, raw_end.z))
        envelopes[beam] = member_envelope(staad, beam, lengths[beam])
        names[beam] = _tube_pipe_name(staad, beam, envelopes[beam], staad.section_name(beam))
    valid = list(points)
    if not valid:
        raise OpenStaadError(
            "STAAD.Pro reported selected members "
            f"{beams}, but returned no valid start/end node incidences for them."
        )
    center_x = sum((points[b][0].x + points[b][1].x) / 2 for b in valid) / len(valid) if valid else 0.0
    outlines: dict[int, list[Point3D]] = {}
    for beam in valid:
        start, end = points[beam]
        fixed = _connected_fixed_width(staad, valid, beam, envelopes, names, points)
        outlines[beam] = _envelope_points(
            start, end, envelopes[beam], names[beam], fixed, center_x
        )
    open_ends: set[tuple[int, int]] = set()
    if settings.peb_corner_joins or settings.connection_face_lines:
        open_ends = apply_peb_corner_joins(outlines, incidences, points)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="ascii", newline="\n") as stream:
        writer = DxfWriter(stream)
        writer.header()
        for beam in valid:
            start, end = points[beam]
            envelope, name = envelopes[beam], names[beam]
            writer.line("MEMBER_CENTERLINE", start, end, "DASHED")
            _write_envelope(
                writer, outlines[beam], envelope, name,
                (beam, 0) in open_ends, (beam, 1) in open_ends,
            )
            if settings.write_labels:
                _write_label(writer, start, end, _label(staad, beam, name, lengths[beam], envelope.property_type),
                             max(envelope.start_half_width, envelope.end_half_width), settings)
        if settings.connection_face_lines:
            write_connection_face_lines(writer, outlines, points, open_ends)
        writer.footer()
    return len(valid)
