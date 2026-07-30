from staad_ext.macros.plate_summary import selected_member_plate_summary


class FakeStaad:
    def selected_beams(self) -> list[int]:
        return [1, 2, 3]

    def beam_length(self, beam_no: int) -> float:
        return {1: 5.0, 2: 5.0, 3: 3.0}[beam_no]

    def section_name(self, beam_no: int) -> str:
        return {1: "TAPERED", 2: "TAPERED", 3: "ISMB300"}[beam_no]

    def section_property_values(self, beam_no: int) -> tuple[int, list[float]]:
        if beam_no in {1, 2}:
            return 680, [0.5, 0.008, 0.7, 0.2, 0.012, 0.18, 0.01] + [0.0] * 17
        return 610, [0.0] * 24

    def beam_property_all(self, beam_no: int) -> tuple[float, ...]:
        return (0.14, 0.3, 0.00562, 0.0, 0.0, 0.0, 0.0, 0.0, 0.012, 0.007)


def test_tapered_members_split_into_plates_and_other_sections_stay_whole() -> None:
    rows = selected_member_plate_summary(FakeStaad())  # type: ignore[arg-type]
    assert len(rows) == 4
    web = next(row for row in rows if row.description == "Tapered web")
    assert web.members == (1, 2)
    assert web.quantity == 2
    assert web.size == "478 → 678 × 8 mm"
    assert web.total_length_m == 10.0
    assert round(web.plate_area_m2 or 0.0, 3) == 5.78
    top = next(row for row in rows if row.description == "Top flange")
    assert top.size == "200 × 12 mm"
    assert top.quantity == 2
    section = next(row for row in rows if row.category == "Whole section")
    assert section.description == "ISMB300"
    assert section.members == (3,)
    assert section.plate_area_m2 is None
    assert round(section.weight_kg or 0.0, 1) == 132.4


def test_tapered_tube_property_is_kept_as_a_whole_section() -> None:
    class TaperedTube(FakeStaad):
        def selected_beams(self) -> list[int]:
            return [1]

        def section_property_values(self, beam_no: int) -> tuple[int, list[float]]:
            return 675, [0.0] * 24

        def section_name(self, beam_no: int) -> str:
            return "TAPERED TUBE"

    rows = selected_member_plate_summary(TaperedTube())  # type: ignore[arg-type]
    assert len(rows) == 1
    assert rows[0].category == "Whole section"


def test_zero_bottom_flange_values_inherit_top_flange_dimensions() -> None:
    class EqualFlanges(FakeStaad):
        def selected_beams(self) -> list[int]:
            return [1]

        def section_property_values(self, beam_no: int) -> tuple[int, list[float]]:
            return 680, [0.5, 0.008, 0.7, 0.2, 0.012, 0.0, 0.0] + [0.0] * 17

    rows = selected_member_plate_summary(EqualFlanges())  # type: ignore[arg-type]
    top = next(row for row in rows if row.description == "Top flange")
    bottom = next(row for row in rows if row.description == "Bottom flange")
    assert top.size == bottom.size == "200 × 12 mm"
    assert top.weight_kg == bottom.weight_kg
    web = next(row for row in rows if row.description == "Tapered web")
    assert web.size == "476 → 676 × 8 mm"
