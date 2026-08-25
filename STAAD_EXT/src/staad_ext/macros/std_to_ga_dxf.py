"""General-arrangement DXF renderer.

Where the analytical export labels every member in place, a GA drawing marks
each member with a bubbled number and collects the section descriptions into one
MEMBER SIZE SCHEDULE beside the frame. Members sharing a section size and a
material grade share a mark.

The model read and the geometry solve are shared with the analytical exporter --
see :mod:`staad_ext.framing`.
"""

from __future__ import annotations

from dataclasses import dataclass
from math import hypot
from pathlib import Path
from typing import Any

from staad_ext.dxf import DxfWriter, dxf_document
from staad_ext.framing import (
    FramingModel, Member, build_framing_model, index_groups, move, offset_vector,
)
from staad_ext.models import GaExportSettings, Point3D, ScheduleCorner

GA_LAYERS: tuple[tuple[str, int, str], ...] = (
    ("MEMBER_CENTERLINE", 8, "DASHED"),
    ("MEMBER_OUTLINE", 3, "CONTINUOUS"),
    ("MEMBER_MARKS", 7, "CONTINUOUS"),
    ("MARK_LEADERS", 8, "CONTINUOUS"),
    ("SCHEDULE_GRID", 7, "CONTINUOUS"),
    ("SCHEDULE_HEADING", 1, "CONTINUOUS"),
    ("SCHEDULE_TEXT", 6, "CONTINUOUS"),
)

MARK_COLOR = 1        # red -- mark numbers, schedule title and headings
DESCRIPTION_COLOR = 6  # magenta -- schedule descriptions and grades
GRID_COLOR = 7        # white -- schedule rules and mark bubbles

# Mark bubbles are drawn one size for the whole drawing, from the frame's
# larger extent -- but that alone oversizes them on densely framed models, so
# the radius is also held below a share of the closest gap between two members'
# mark anchors. MARK_CLEARANCE_FACTOR below 0.5 is what keeps two neighbouring
# bubbles from touching.
MARK_RADIUS_FACTOR = 0.009
MARK_CLEARANCE_FACTOR = 0.34
MIN_MARK_RADIUS_FACTOR = 0.0025
MARK_LEADER_RADII = 2.2

SCHEDULE_TITLE = "MEMBER SIZE SCHEDULE"
SCHEDULE_HEADINGS = ("MARK", "DESCRIPTION", "GRADE")
# MARK / DESCRIPTION / GRADE column widths as fractions of the table width.
COLUMN_FRACTIONS = (0.19, 0.54, 0.27)


@dataclass(frozen=True, slots=True)
class ScheduleEntry:
    mark: int
    description: str
    grade: str
    beams: tuple[int, ...]


def _dim(value: float) -> str:
    """Format a length in mm: whole where it is whole, else one decimal."""
    millimetres = abs(value) * 1000
    return (f"{millimetres:.0f}" if abs(millimetres - round(millimetres)) < 0.05
            else f"{millimetres:.1f}")


def _thickness(value: float) -> str:
    """Format a plate/wall thickness in mm, always to one decimal (16.0, 3.6)."""
    return f"{abs(value) * 1000:.1f}"


def _hollow(depth: float, width: float, thickness: float) -> str:
    return f"{_dim(depth)}x{_dim(width)}x{_thickness(thickness)}  (RHS)"


def _round_hollow(diameter: float, thickness: float) -> str:
    return f"{_dim(diameter)}x{_thickness(thickness)}  (CHS)"


def _built_up(first_depth: float, second_depth: float, web_thickness: float,
              flange_width: float, flange_thickness: float) -> str:
    """Describe a built-up I as its plate sizes, e.g. W=600~350x5/F=150x8.

    STAAD reports the overall section depth, so the flanges are taken off to
    leave the web plate depth a fabricator would actually order.
    """
    webs = sorted(
        (abs(first_depth) - 2 * abs(flange_thickness),
         abs(second_depth) - 2 * abs(flange_thickness)),
        reverse=True,
    )
    web = (_dim(webs[0]) if abs(webs[0] - webs[1]) < 5e-4
           else f"{_dim(webs[0])}~{_dim(webs[1])}")
    return (f"W={web}x{_dim(web_thickness)}"
            f"/F={_dim(flange_width)}x{_dim(flange_thickness)}")


def describe_section(member: Member) -> str:
    """Build the MEMBER SIZE SCHEDULE description for a member's section."""
    values, property_type = member.property_values, member.property_type
    upper = member.name.upper()
    if values:
        try:
            v = [abs(value) for value in values]
            if property_type in {650, 654, 696}:
                return _hollow(v[1], v[2], v[3])
            if property_type in {660, 655}:
                return _round_hollow(v[1], v[2])
            if property_type == 695:
                return _round_hollow(v[1], (v[1] - v[2]) / 2)
            if property_type == 675:
                return _built_up(v[4], v[5], v[1], v[6], v[7])
            if property_type == 680:
                return _built_up(v[0], v[2], v[1], max(v[3], v[5]), max(v[4], v[6]))
            if property_type == 2 and any(x in upper for x in ("TUBE", "RHS", "SHS")):
                return _hollow(v[2], v[1], v[0])
            if property_type == 2 and any(x in upper for x in ("PIPE", "CHS")):
                return _round_hollow(v[0], (v[0] - v[1]) / 2)
        except (IndexError, TypeError, ValueError):
            pass
    return member.name


def member_grade(staad: Any, beam_no: int) -> str:
    """Return a member's steel grade for the schedule.

    Blank for now. The column stays in the schedule so it can be filled in by
    hand; the yield strength read back through ``beam_material_name`` and
    ``material_yield_strength`` is left unused until it can be trusted.
    """
    return ""


def build_schedule(staad: Any, model: FramingModel) -> tuple[dict[int, int], list[ScheduleEntry]]:
    """Group members into marks and build the schedule rows.

    Members are grouped on the description and grade the schedule will print,
    not on the STAAD section name: two members whose rows would read
    identically are the same physical size and must share one mark. Grades are
    blank for now, so in practice the description alone decides the mark.
    """
    grades = {number: member_grade(staad, number) for number in model.members}
    descriptions = {number: describe_section(member)
                    for number, member in model.members.items()}
    indices = index_groups(
        model,
        key=lambda member: (descriptions[member.number], grades[member.number]),
        order_by_appearance=True,
    )
    marks = {number: index + 1 for number, index in indices.items()}
    entries = []
    for mark in sorted(set(marks.values())):
        beams = tuple(sorted(number for number, value in marks.items() if value == mark))
        entries.append(ScheduleEntry(
            mark, descriptions[beams[0]], grades[beams[0]], beams,
        ))
    return marks, entries


def _bubble(writer: DxfWriter, layer: str, center: Point3D, radius: float,
            value: str, color: int = MARK_COLOR) -> None:
    """Draw a circled number centred on ``center``."""
    writer.circle(layer, center, radius)
    height = radius * (1.15 if len(value) < 2 else 0.85)
    writer.text(layer, Point3D(center.x, center.y - height * 0.5), height, 0.0,
                value, color)


def mark_anchor(member: Member) -> Point3D:
    """Return the point on a member's face that its mark leader points at."""
    unit = mark_direction(member)
    start, end = member.start, member.end
    middle = Point3D((start.x + end.x) / 2, (start.y + end.y) / 2, (start.z + end.z) / 2)
    return move(middle, unit, member.half_width)


def mark_direction(member: Member) -> Point3D:
    """Return the outward unit normal a member's mark is offset along."""
    unit = offset_vector(member.start, member.end)
    return Point3D(-unit.x, -unit.y, -unit.z) if unit.y < 0 else unit


def mark_radius(model: FramingModel, settings: GaExportSettings) -> float:
    """Size mark bubbles so neighbouring members' bubbles cannot collide."""
    min_x, min_y, max_x, max_y = model.bounds()
    extent = max(max_x - min_x, max_y - min_y, 1.0)
    radius = extent * MARK_RADIUS_FACTOR
    anchors = [mark_anchor(member) for member in model.members.values()]
    gaps = [hypot(second.x - first.x, second.y - first.y)
            for index, first in enumerate(anchors) for second in anchors[index + 1:]]
    if gaps:
        radius = min(radius, min(gaps) * MARK_CLEARANCE_FACTOR)
    return max(radius, extent * MIN_MARK_RADIUS_FACTOR) * settings.mark_scale


def write_mark_bubble(writer: DxfWriter, member: Member, mark: int,
                      radius: float) -> None:
    """Draw a member's mark bubble, offset clear of the section with a leader."""
    unit = mark_direction(member)
    anchor = mark_anchor(member)
    center = move(anchor, unit, radius * MARK_LEADER_RADII)

    writer.line("MARK_LEADERS", move(center, unit, -radius), anchor)
    head = radius * 0.5
    back = move(anchor, unit, head)
    side = Point3D(-unit.y, unit.x)
    for direction in (-1.0, 1.0):
        writer.line("MARK_LEADERS", anchor, move(back, side, direction * head * 0.35))
    _bubble(writer, "MEMBER_MARKS", center, radius, str(mark))


def _schedule_origin(model: FramingModel, settings: GaExportSettings,
                     width: float, height: float) -> Point3D:
    """Return the table's top-left corner, placed clear of the framing."""
    min_x, min_y, max_x, max_y = model.bounds()
    # Wide enough to clear the mark bubbles, which sit outside the outlines
    # that bounds() measures.
    gap = max(max_x - min_x, max_y - min_y, 1.0) * 0.12
    left = (max_x + gap if settings.schedule_corner in
            (ScheduleCorner.TOP_RIGHT, ScheduleCorner.BOTTOM_RIGHT)
            else min_x - gap - width)
    top = (max_y if settings.schedule_corner in
           (ScheduleCorner.TOP_LEFT, ScheduleCorner.TOP_RIGHT)
           else min_y + height)
    return Point3D(left, top)


def write_schedule(writer: DxfWriter, model: FramingModel, entries: list[ScheduleEntry],
                   settings: GaExportSettings) -> None:
    """Draw the MEMBER SIZE SCHEDULE table beside the framing."""
    min_x, min_y, max_x, max_y = model.bounds()
    extent = max(max_x - min_x, max_y - min_y, 1.0)
    width = extent * 0.42 * settings.text_scale
    row_height = width * 0.052
    text_height = row_height * 0.42
    total_rows = 2 + len(entries) + settings.blank_rows
    origin = _schedule_origin(model, settings, width, total_rows * row_height)

    edges = [origin.x + width * sum(COLUMN_FRACTIONS[:index])
             for index in range(len(COLUMN_FRACTIONS) + 1)]

    def row_top(index: int) -> float:
        return origin.y - index * row_height

    def centered(index: int, left: float, right: float, value: str,
                 layer: str, color: int) -> None:
        writer.text(layer, Point3D((left + right) / 2,
                                   row_top(index) - row_height / 2 - text_height * 0.35),
                    text_height, 0.0, value, color)

    for index in range(total_rows + 1):
        writer.line("SCHEDULE_GRID", Point3D(origin.x, row_top(index)),
                    Point3D(origin.x + width, row_top(index)))
    bottom = row_top(total_rows)
    for edge in (edges[0], edges[-1]):
        writer.line("SCHEDULE_GRID", Point3D(edge, origin.y), Point3D(edge, bottom))
    # The title spans the full width, so the column rules start one row down.
    for edge in edges[1:-1]:
        writer.line("SCHEDULE_GRID", Point3D(edge, row_top(1)), Point3D(edge, bottom))

    centered(0, edges[0], edges[-1], SCHEDULE_TITLE, "SCHEDULE_HEADING", MARK_COLOR)
    for index, heading in enumerate(SCHEDULE_HEADINGS):
        centered(1, edges[index], edges[index + 1], heading, "SCHEDULE_HEADING", MARK_COLOR)

    padding = width * 0.02
    for offset, entry in enumerate(entries):
        index = offset + 2
        middle_y = row_top(index) - row_height / 2
        _bubble(writer, "SCHEDULE_HEADING",
                Point3D((edges[0] + edges[1]) / 2, middle_y),
                row_height * 0.32, str(entry.mark))
        writer.text("SCHEDULE_TEXT", Point3D(edges[1] + padding, middle_y - text_height * 0.35),
                    text_height, 0.0, entry.description, DESCRIPTION_COLOR, halign=0)
        if entry.grade:  # an empty GRADE cell is left for the user to fill in
            centered(index, edges[2], edges[3], entry.grade, "SCHEDULE_TEXT",
                     DESCRIPTION_COLOR)


def export_ga_drawing(staad: Any, output: Path, settings: GaExportSettings) -> int:
    """Export the selected members as a GA drawing with a member size schedule."""
    model = build_framing_model(staad, settings.plane, True)
    if not model:
        return 0
    marks, entries = build_schedule(staad, model)
    radius = mark_radius(model, settings)

    with dxf_document(output, GA_LAYERS) as writer:
        for member in model.members.values():
            if settings.write_centerlines:
                writer.line("MEMBER_CENTERLINE", member.start, member.end, "DASHED")
            writer.envelope(
                "MEMBER_OUTLINE", member.outline,
                (member.number, 0) in model.open_ends,
                (member.number, 1) in model.open_ends,
            )
            if settings.write_marks:
                write_mark_bubble(writer, member, marks[member.number], radius)
        write_schedule(writer, model, entries, settings)
    return len(model.members)
