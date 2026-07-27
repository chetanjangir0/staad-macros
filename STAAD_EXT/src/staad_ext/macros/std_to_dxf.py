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
                    name: str) -> None:
    p1, p2, p3, p4 = outline
    layer = "TUBE_PIPE_SECTION" if is_tube_or_pipe(envelope.property_type, name) else "MEMBER_SECTION"
    if is_tapered(envelope, name):
        layer = "TAPERED_SECTION"
    for a, b in ((p1, p2), (p3, p4), (p1, p3), (p2, p4)):
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


def apply_peb_corner_joins(
    outlines: dict[int, list[Point3D]],
    incidences: dict[int, tuple[int, int]],
    centerlines: dict[int, tuple[Point3D, Point3D]],
) -> None:
    """Join envelope edges at simple, non-parallel PEB frame corners."""
    by_node: dict[int, list[tuple[int, int]]] = {}
    for beam, nodes in incidences.items():
        by_node.setdefault(nodes[0], []).append((beam, 0))
        by_node.setdefault(nodes[1], []).append((beam, 1))

    for connections in by_node.values():
        if len(connections) != 2:
            continue
        (beam_a, end_a), (beam_b, end_b) = connections
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
    if settings.peb_corner_joins:
        apply_peb_corner_joins(outlines, incidences, points)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="ascii", newline="\n") as stream:
        writer = DxfWriter(stream)
        writer.header()
        for beam in valid:
            start, end = points[beam]
            envelope, name = envelopes[beam], names[beam]
            writer.line("MEMBER_CENTERLINE", start, end, "DASHED")
            _write_envelope(writer, outlines[beam], envelope, name)
            if settings.write_labels:
                _write_label(writer, start, end, _label(staad, beam, name, lengths[beam], envelope.property_type),
                             max(envelope.start_half_width, envelope.end_half_width), settings)
        writer.footer()
    return len(valid)
