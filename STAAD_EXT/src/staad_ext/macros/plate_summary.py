from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from staad_ext.openstaad import OpenStaad, OpenStaadError

STEEL_DENSITY_KG_M3 = 7850.0
TAPERED_I_PROPERTY_TYPE = 680


@dataclass(frozen=True, slots=True)
class PlateSummaryRow:
    category: str
    description: str
    size: str
    members: tuple[int, ...]
    quantity: int
    length_each_m: float
    total_length_m: float
    plate_area_m2: float | None
    weight_kg: float | None


@dataclass(frozen=True, slots=True)
class _SummaryItem:
    category: str
    description: str
    size: str
    member: int
    quantity: int
    length_m: float
    plate_area_m2: float | None
    weight_kg: float | None


def _mm(value_m: float) -> str:
    return f"{abs(value_m) * 1000.0:.2f}".rstrip("0").rstrip(".")


def _tapered_i_items(beam_no: int, length_m: float, values: list[float]) -> list[_SummaryItem]:
    """Split documented property type 680 into web and flange plates."""
    if len(values) < 7:
        raise OpenStaadError(f"Member {beam_no}: incomplete tapered section data.")
    depth_1, web_thickness, depth_2 = (abs(values[index]) for index in range(3))
    top_width, top_thickness = abs(values[3]), abs(values[4])
    # Zero F6/F7 means the bottom flange is identical to the top flange.
    # Inherit F4/F5 instead of treating the tapered property as invalid.
    bottom_width = abs(values[5]) or top_width
    bottom_thickness = abs(values[6]) or top_thickness
    dimensions = (depth_1, web_thickness, depth_2, top_width, top_thickness,
                  bottom_width, bottom_thickness, length_m)
    if any(value <= 0 for value in dimensions):
        raise OpenStaadError(f"Member {beam_no}: tapered section dimensions are incomplete or zero.")

    web_1 = max(depth_1 - top_thickness - bottom_thickness, 0.0)
    web_2 = max(depth_2 - top_thickness - bottom_thickness, 0.0)
    web_area = (web_1 + web_2) * 0.5 * length_m
    items = [_SummaryItem(
        "Web plate", "Tapered web",
        f"{_mm(web_1)} → {_mm(web_2)} × {_mm(web_thickness)} mm",
        beam_no, 1, length_m, web_area,
        web_area * web_thickness * STEEL_DENSITY_KG_M3,
    )]
    for description, width, thickness in (
        ("Top flange", top_width, top_thickness),
        ("Bottom flange", bottom_width, bottom_thickness),
    ):
        area = width * length_m
        items.append(_SummaryItem(
            "Flange plate", description, f"{_mm(width)} × {_mm(thickness)} mm",
            beam_no, 1, length_m, area,
            area * thickness * STEEL_DENSITY_KG_M3,
        ))
    return items


def _whole_section_item(staad: OpenStaad, beam_no: int, length_m: float,
                        section_name: str) -> _SummaryItem:
    weight: float | None = None
    try:
        _, _, cross_section_area, *_ = staad.beam_property_all(beam_no)
        if cross_section_area > 0:
            weight = cross_section_area * length_m * STEEL_DENSITY_KG_M3
    except (OSError, TypeError, ValueError):
        pass
    return _SummaryItem("Whole section", section_name, section_name, beam_no,
                        1, length_m, None, weight)


def _aggregate(items: Iterable[_SummaryItem]) -> list[PlateSummaryRow]:
    groups: dict[tuple[object, ...], list[_SummaryItem]] = {}
    for item in items:
        key = (item.category, item.description, item.size, round(item.length_m, 6))
        groups.setdefault(key, []).append(item)
    rows = []
    for grouped in groups.values():
        first = grouped[0]
        rows.append(PlateSummaryRow(
            first.category, first.description, first.size,
            tuple(sorted(item.member for item in grouped)),
            sum(item.quantity for item in grouped), first.length_m,
            sum(item.length_m * item.quantity for item in grouped),
            sum(item.plate_area_m2 or 0.0 for item in grouped)
            if all(item.plate_area_m2 is not None for item in grouped) else None,
            sum(item.weight_kg or 0.0 for item in grouped)
            if all(item.weight_kg is not None for item in grouped) else None,
        ))
    return sorted(rows, key=lambda row: (row.category, row.size, row.members))


def selected_member_plate_summary(staad: OpenStaad) -> list[PlateSummaryRow]:
    """Build a plate/section summary for selected analytical beam members."""
    items: list[_SummaryItem] = []
    for beam_no in staad.selected_beams():
        length_m = staad.beam_length(beam_no)
        section_name = staad.section_name(beam_no)
        property_type, values = staad.section_property_values(beam_no)
        if property_type == TAPERED_I_PROPERTY_TYPE:
            items.extend(_tapered_i_items(beam_no, length_m, values))
        else:
            items.append(_whole_section_item(staad, beam_no, length_m, section_name))
    return _aggregate(items)
