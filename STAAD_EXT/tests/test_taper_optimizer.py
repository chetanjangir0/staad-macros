from __future__ import annotations

import pytest

from staad_ext.models import DeflectionLimits, Point3D, TaperOptimizerSettings
from staad_ext.openstaad import OpenStaadError
from staad_ext.macros.taper_optimizer import (
    DesignState, Evaluation, PLATE_THICKNESS_MM, TaperedMember, TaperedSection,
    build_frame, deflection_checks, ladder, optimize, optimize_tapered_sections,
    prepare_seed, read_tapered_frame, snap_to_ladders, snap_up,
)


def settings(**overrides: object) -> TaperOptimizerSettings:
    values: dict[str, object] = {
        "deflection": DeflectionLimits(load_cases=(101,)),
        "utilisation_ceiling": 0.95,
        "analysis_budget": 60,
    }
    values.update(overrides)
    return TaperOptimizerSettings(**values)  # type: ignore[arg-type]


def section(start: float, end: float, web: float = 6.0,
            width: float = 200.0, flange: float = 10.0) -> TaperedSection:
    return TaperedSection(start, end, web, width, flange)


def portal_frame(tie_all: bool = False):
    """A two-member rafter continuing through node 2, plus a column below it.

    Nodes 1-2-3 run up the slope; node 0 sits at the base of the column, which
    meets the rafter at node 1 at right angles.
    """
    points = {
        0: Point3D(0.0, 0.0), 1: Point3D(0.0, 6.0),
        2: Point3D(6.0, 7.0), 3: Point3D(12.0, 8.0),
    }
    members = [
        TaperedMember(1, 0, 1, 6.0, points[0], points[1], section(400.0, 600.0)),
        TaperedMember(2, 1, 2, 6.083, points[1], points[2], section(600.0, 450.0)),
        TaperedMember(3, 2, 3, 6.083, points[2], points[3], section(450.0, 300.0)),
    ]
    return build_frame(members, points, 0.0, tie_all)


def split_column_frame(prismatic_columns: bool):
    """A column spliced at node 1, carrying a rafter away from its top.

    The splice is what the prismatic option has to see through: holding each
    member straight on its own would still allow a step in the web at node 1.
    """
    points = {
        0: Point3D(0.0, 0.0), 1: Point3D(0.0, 3.0),
        2: Point3D(0.0, 6.0), 3: Point3D(6.0, 7.0),
    }
    members = [
        TaperedMember(1, 0, 1, 3.0, points[0], points[1], section(400.0, 500.0)),
        TaperedMember(2, 1, 2, 3.0, points[1], points[2], section(500.0, 600.0)),
        TaperedMember(3, 2, 3, 6.083, points[2], points[3], section(600.0, 400.0)),
    ]
    return build_frame(members, points, 0.0, tie_all=False,
                       prismatic_columns=prismatic_columns)


# --------------------------------------------------------------------------
# Ladders
# --------------------------------------------------------------------------

def test_ladders_start_and_step_as_the_fabrication_rules_require() -> None:
    depths = ladder(250.0, 50.0, 500.0)
    assert depths == (250.0, 300.0, 350.0, 400.0, 450.0, 500.0)
    widths = ladder(150.0, 25.0, 250.0)
    assert widths == (150.0, 175.0, 200.0, 225.0, 250.0)


def test_snap_up_never_returns_a_smaller_size() -> None:
    assert snap_up(263.0, ladder(250.0, 50.0, 500.0)) == 300.0
    assert snap_up(300.0, ladder(250.0, 50.0, 500.0)) == 300.0
    assert snap_up(7.0, PLATE_THICKNESS_MM) == 8.0


def test_snapping_rounds_every_variable_up_onto_its_ladder() -> None:
    frame = portal_frame()
    odd = DesignState(
        depth_by_group={group: 413.0 for group in set(frame.group_of_end.values())},
        web_by_member={1: 7.0, 2: 7.0, 3: 7.0},
        width_by_member={1: 190.0, 2: 190.0, 3: 190.0},
        flange_by_member={1: 9.0, 2: 9.0, 3: 9.0},
    )
    snapped = snap_to_ladders(odd, frame, settings())
    assert set(snapped.depth_by_group.values()) == {450.0}
    assert set(snapped.web_by_member.values()) == {8.0}
    assert set(snapped.width_by_member.values()) == {200.0}
    assert set(snapped.flange_by_member.values()) == {10.0}


# --------------------------------------------------------------------------
# Depth continuity
# --------------------------------------------------------------------------

def test_collinear_members_share_one_depth_at_the_node_they_meet() -> None:
    frame = portal_frame()
    # Members 2 and 3 continue one another through node 2.
    assert frame.group_of_end[(2, 1)] == frame.group_of_end[(3, 0)]


def test_a_knee_joint_is_not_tied_by_default_but_can_be() -> None:
    # Member 1 (the column) meets member 2 (the rafter) at node 1 at an angle.
    assert portal_frame().group_of_end[(1, 1)] != portal_frame().group_of_end[(2, 0)]
    tied = portal_frame(tie_all=True)
    assert tied.group_of_end[(1, 1)] == tied.group_of_end[(2, 0)]


def test_the_seed_depth_at_a_shared_node_takes_the_deeper_of_the_two() -> None:
    points = {1: Point3D(0.0, 0.0), 2: Point3D(6.0, 0.0), 3: Point3D(12.0, 0.0)}
    frame = build_frame([
        TaperedMember(1, 1, 2, 6.0, points[1], points[2], section(400.0, 500.0)),
        TaperedMember(2, 2, 3, 6.0, points[2], points[3], section(650.0, 400.0)),
    ], points, 0.0, tie_all=False)
    state = frame.initial_state()
    assert state.depth_by_group[frame.group_of_end[(1, 1)]] == 650.0


def test_a_column_run_shares_one_depth_only_when_it_is_held_straight() -> None:
    tapered = split_column_frame(prismatic_columns=False)
    assert tapered.group_of_end[(1, 0)] != tapered.group_of_end[(2, 1)]
    straight = split_column_frame(prismatic_columns=True)
    # Base to eave is one depth, across the splice at node 1 as well.
    assert len({straight.group_of_end[(member, end)]
                for member in (1, 2) for end in (0, 1)}) == 1
    # ...and the rafter is left free to taper either way.
    assert straight.group_of_end[(3, 0)] != straight.group_of_end[(3, 1)]


def test_a_straight_column_is_seeded_from_the_deepest_section_it_carried() -> None:
    # Collapsing four ends onto one depth has to round up, not average, or the
    # first analysis would start from something weaker than the engineer drew.
    frame = split_column_frame(prismatic_columns=True)
    state = frame.initial_state()
    for member in (1, 2):
        assert frame.section_for(member, state) == section(600.0, 600.0)


def test_a_column_held_straight_stays_straight_through_the_search() -> None:
    frame = split_column_frame(prismatic_columns=True)
    result = optimize(frame, settings(), lambda state: Evaluation(
        {number: 0.4 for number in frame.numbers}))
    after = {change.number: change.after for change in result.changes}
    assert after[1].is_prismatic and after[2].is_prismatic
    assert after[1].start_depth_mm == after[2].start_depth_mm


def test_every_optimized_section_keeps_a_continuous_web_through_the_node() -> None:
    frame = portal_frame()
    result = optimize(frame, settings(), lambda state: Evaluation(
        {number: 0.4 for number in frame.numbers}))
    after = {change.number: change.after for change in result.changes}
    assert after[2].end_depth_mm == after[3].start_depth_mm


# --------------------------------------------------------------------------
# Fabrication constraints
# --------------------------------------------------------------------------

def test_the_flange_stays_thicker_than_the_web_everywhere() -> None:
    frame = portal_frame()
    result = optimize(frame, settings(), lambda state: Evaluation(
        {number: 0.2 for number in frame.numbers}))
    for change in result.changes:
        assert change.after.flange_thickness_mm > change.after.web_thickness_mm


def test_optimized_sizes_stay_on_the_fabrication_ladders() -> None:
    frame = portal_frame()
    result = optimize(frame, settings(), lambda state: Evaluation(
        {number: 0.3 for number in frame.numbers}))
    for change in result.changes:
        after = change.after
        for depth in (after.start_depth_mm, after.end_depth_mm):
            assert depth >= 250.0 and (depth - 250.0) % 50.0 == 0
        assert after.flange_width_mm >= 150.0
        assert (after.flange_width_mm - 150.0) % 25.0 == 0
        assert after.web_thickness_mm in PLATE_THICKNESS_MM
        assert after.flange_thickness_mm in PLATE_THICKNESS_MM


def slender_frame():
    """One member whose 900mm web on 5mm plate is d/t ~ 175."""
    points = {1: Point3D(0.0, 0.0), 2: Point3D(6.0, 0.0)}
    return build_frame([TaperedMember(
        1, 1, 2, 6.0, points[1], points[2],
        TaperedSection(900.0, 900.0, 5.0, 200.0, 10.0))], points, 0.0, False)


def test_a_seed_web_too_slender_for_the_cap_is_thickened_before_optimizing() -> None:
    frame = slender_frame()
    state, notes = prepare_seed(frame, settings(max_web_slenderness=100.0))
    assert state.web_by_member[1] >= 9.0     # 880 clear / 100 needs 8.8mm plate
    assert any("thickened" in note for note in notes)


def test_an_illegal_seed_does_not_survive_the_search() -> None:
    # Regression: every candidate step is judged against the whole state, so an
    # illegal starting point used to have every step rejected and was handed
    # straight back unchanged.
    frame = slender_frame()
    limit = 100.0
    result = optimize(frame, settings(max_web_slenderness=limit),
                      lambda state: Evaluation({1: 0.9}))
    after = result.changes[0].after
    clear = after.start_depth_mm - 2 * after.flange_thickness_mm
    assert clear / after.web_thickness_mm <= limit + 1e-9


def test_a_web_that_cannot_meet_the_cap_at_any_thickness_is_reported() -> None:
    points = {1: Point3D(0.0, 0.0), 2: Point3D(6.0, 0.0)}
    frame = build_frame([TaperedMember(
        1, 1, 2, 6.0, points[1], points[2],
        TaperedSection(2400.0, 2400.0, 5.0, 200.0, 10.0))], points, 0.0, False)
    with pytest.raises(ValueError, match="room for a thicker flange"):
        prepare_seed(frame, settings(max_web_slenderness=25.0))


def test_a_seed_deeper_than_the_maximum_is_reported_not_silently_clamped() -> None:
    points = {1: Point3D(0.0, 0.0), 2: Point3D(6.0, 0.0)}
    frame = build_frame([TaperedMember(
        1, 1, 2, 6.0, points[1], points[2],
        TaperedSection(1400.0, 1400.0, 12.0, 200.0, 16.0))], points, 0.0, False)
    with pytest.raises(ValueError, match="deeper than the 1000mm maximum"):
        optimize(frame, settings(max_depth_mm=1000.0),
                 lambda state: Evaluation({1: 0.5}))


def test_a_seed_wider_than_the_maximum_is_reported() -> None:
    points = {1: Point3D(0.0, 0.0), 2: Point3D(6.0, 0.0)}
    frame = build_frame([TaperedMember(
        1, 1, 2, 6.0, points[1], points[2],
        TaperedSection(600.0, 600.0, 12.0, 700.0, 16.0))], points, 0.0, False)
    with pytest.raises(ValueError, match="wider than the 500mm maximum"):
        optimize(frame, settings(max_flange_width_mm=500.0),
                 lambda state: Evaluation({1: 0.5}))


def test_the_web_slenderness_cap_is_respected() -> None:
    frame = portal_frame()
    limit = 120.0
    result = optimize(frame, settings(max_web_slenderness=limit),
                      lambda state: Evaluation(
                          {number: 0.2 for number in frame.numbers}))
    for change in result.changes:
        after = change.after
        clear = max(after.start_depth_mm, after.end_depth_mm) - 2 * after.flange_thickness_mm
        assert clear / after.web_thickness_mm <= limit + 1e-9


# --------------------------------------------------------------------------
# The search
# --------------------------------------------------------------------------

def test_a_lightly_loaded_frame_is_driven_down_to_the_smallest_sizes() -> None:
    frame = portal_frame()
    result = optimize(frame, settings(), lambda state: Evaluation(
        {number: 0.05 for number in frame.numbers}))
    assert result.feasible
    assert result.saved_kg > 0
    for change in result.changes:
        assert change.after.start_depth_mm == 250.0
        assert change.after.flange_width_mm == 150.0
        assert change.after.web_thickness_mm == 5.0
        assert change.after.flange_thickness_mm == 6.0


def test_sections_are_left_alone_when_the_frame_is_already_fully_stressed() -> None:
    frame = portal_frame()
    start = snap_to_ladders(frame.initial_state(), frame, settings())
    weight = frame.weight_kg(start)

    def evaluate(state: DesignState) -> Evaluation:
        # Anything lighter than the starting frame fails the code check.
        ratio = 0.9 if frame.weight_kg(state) >= weight - 1e-9 else 1.4
        return Evaluation({number: ratio for number in frame.numbers})

    result = optimize(frame, settings(), evaluate)
    assert result.feasible
    assert result.saved_kg == pytest.approx(0.0)


def test_an_overstressed_frame_is_strengthened_until_it_passes() -> None:
    frame = portal_frame()
    start = snap_to_ladders(frame.initial_state(), frame, settings())
    weight = frame.weight_kg(start)

    def evaluate(state: DesignState) -> Evaluation:
        # Only a frame at least 30% heavier than the seed survives.
        heavy = frame.weight_kg(state) >= weight * 1.3
        return Evaluation({number: 0.8 if heavy else 1.6 for number in frame.numbers})

    result = optimize(frame, settings(), evaluate)
    assert result.feasible
    assert frame.weight_kg(result.state) >= weight * 1.3


def test_a_frame_that_cannot_be_made_to_pass_reports_infeasible() -> None:
    frame = portal_frame()
    result = optimize(frame, settings(analysis_budget=8), lambda state: Evaluation(
        {number: 5.0 for number in frame.numbers}))
    assert not result.feasible
    assert result.evaluation is not None


def test_the_analysis_budget_is_never_exceeded() -> None:
    frame = portal_frame()
    calls = 0

    def evaluate(state: DesignState) -> Evaluation:
        nonlocal calls
        calls += 1
        return Evaluation({number: 0.1 for number in frame.numbers})

    result = optimize(frame, settings(analysis_budget=5), evaluate)
    assert calls <= 5
    assert result.analyses_used == calls


def test_a_failing_deflection_blocks_a_reduction_the_ratios_would_allow() -> None:
    frame = portal_frame()
    start = snap_to_ladders(frame.initial_state(), frame, settings())
    weight = frame.weight_kg(start)

    def evaluate(state: DesignState) -> Evaluation:
        # Every member passes its code check, but sag runs away as soon as the
        # frame gets lighter than it started.
        sag = 30.0 if frame.weight_kg(state) < weight - 1e-9 else 10.0
        return Evaluation(
            {number: 0.3 for number in frame.numbers},
            (DeflectionCheckStub("Vertical", 2, 101, sag, 25.0, 12.0),),
        )

    result = optimize(frame, settings(), evaluate)
    assert result.saved_kg == pytest.approx(0.0)


class DeflectionCheckStub:
    """A stand-in with the two attributes the optimizer reads off a check."""

    def __init__(self, kind: str, node: int, load_case: int, actual: float,
                 allowed: float, span: float) -> None:
        self.kind, self.node, self.load_case = kind, node, load_case
        self.actual_mm, self.allowed_mm, self.span_m = actual, allowed, span

    @property
    def passes(self) -> bool:
        return self.actual_mm <= self.allowed_mm


# --------------------------------------------------------------------------
# Deflection checks
# --------------------------------------------------------------------------

def test_rafter_sag_is_measured_against_the_chord_not_the_ground() -> None:
    frame = portal_frame()
    rafter = next(chain for chain in frame.chains if not chain.is_column)
    # Both rafter ends drop 20mm and the middle drops 40mm: the sag relative to
    # the chord is 20mm, not the 40mm a ground-referenced check would report.
    vertical = {(1, 101): -0.020, (2, 101): -0.040, (3, 101): -0.020}
    horizontal = {node: 0.0 for node in vertical}
    checks = deflection_checks(frame, horizontal, vertical, (101,), 240.0, 150.0)
    sag = next(check for check in checks if check.kind == "Vertical")
    assert sag.node == 2
    assert sag.actual_mm == pytest.approx(20.0)
    assert sag.allowed_mm == pytest.approx(rafter.span_m / 240.0 * 1000.0)


def test_column_drift_is_checked_at_the_top_against_its_height() -> None:
    frame = portal_frame()
    horizontal = {(0, 101): 0.0, (1, 101): 0.030, (2, 101): 0.0, (3, 101): 0.0}
    vertical = {node: 0.0 for node in horizontal}
    checks = deflection_checks(frame, horizontal, vertical, (101,), 240.0, 150.0)
    drift = next(check for check in checks if check.kind == "Horizontal")
    assert drift.node == 1
    assert drift.actual_mm == pytest.approx(30.0)
    assert drift.allowed_mm == pytest.approx(6.0 / 150.0 * 1000.0)
    assert drift.passes


def test_column_drift_beyond_the_height_limit_fails() -> None:
    frame = portal_frame()
    horizontal = {(0, 101): 0.0, (1, 101): 0.055, (2, 101): 0.0, (3, 101): 0.0}
    vertical = {node: 0.0 for node in horizontal}
    checks = deflection_checks(frame, horizontal, vertical, (101,), 240.0, 150.0)
    drift = next(check for check in checks if check.kind == "Horizontal")
    assert drift.actual_mm == pytest.approx(55.0)
    assert not drift.passes


# --------------------------------------------------------------------------
# Weight and reporting
# --------------------------------------------------------------------------

def test_section_weight_counts_the_web_between_the_flanges() -> None:
    # 500mm deep, 8mm web, 200x12 flanges, 1m long.
    weight = TaperedSection(500.0, 500.0, 8.0, 200.0, 12.0).weight_kg(1.0)
    web = (0.500 - 2 * 0.012) * 0.008
    flanges = 2 * 0.200 * 0.012
    assert weight == pytest.approx((web + flanges) * 7850.0)


# --------------------------------------------------------------------------
# The STAAD.Pro round trip
# --------------------------------------------------------------------------

class FakeStaad:
    """A stand-in for the OpenStaad facade, in meters like the real one."""

    # Node 1 is the column base; 2 is the eaves; 3 the apex. Member 4 is a
    # rolled section that must survive the run untouched.
    POINTS = {1: (0.0, 0.0, 0.0), 2: (0.0, 6.0, 0.0), 3: (10.0, 7.5, 0.0),
              4: (10.0, 0.0, 0.0)}
    INCIDENCE = {1: (1, 2), 2: (2, 3), 3: (3, 4), 4: (1, 4)}

    def __init__(self, selected: list[int] | None = None,
                 ratio: float = 0.4) -> None:
        self.selected = selected or []
        self.ratio = ratio
        self.sections: dict[int, tuple[float, ...]] = {
            1: (0.400, 0.008, 0.700, 0.250, 0.016, 0.250, 0.016),
            2: (0.700, 0.008, 0.500, 0.250, 0.016, 0.0, 0.0),
            3: (0.500, 0.008, 0.400, 0.250, 0.016, 0.250, 0.016),
        }
        # Each tapered member starts on its own existing property, numbered
        # 101-103, the way a real model would.
        self.properties: dict[int, tuple[float, ...]] = {
            100 + beam: values for beam, values in self.sections.items()
        }
        self.assigned: dict[int, int] = {beam: 100 + beam for beam in self.sections}
        self.material: dict[int, str] = {beam: "STEEL" for beam in (1, 2, 3, 4)}
        self.analyses = 0
        self.saves = 0
        self.silent = False
        # The sections each analysis actually saw, so a run that graded
        # something other than the candidate it assigned is visible.
        self.analysed: list[dict[int, tuple[float, ...]]] = []

    def beam_property_ref(self, beam: int) -> int:
        return self.assigned.get(beam, 0)

    def beam_material_name(self, beam: int) -> str:
        return self.material.get(beam, "")

    # -- reads ------------------------------------------------------------
    def selected_beams(self) -> list[int]:
        return list(self.selected)

    def all_beams(self) -> list[int]:
        return [1, 2, 3, 4]

    def section_property_values(self, beam: int) -> tuple[int, list[float]]:
        if beam == 4:                       # a plain rolled section
            return 610, [0.0] * 24
        return 680, list(self.sections[beam]) + [0.0] * 17

    def member_incidence(self, beam: int) -> tuple[int, int]:
        return self.INCIDENCE[beam]

    def node_coordinates(self, node: int) -> Point3D:
        return Point3D(*self.POINTS[node])

    def beam_length(self, beam: int) -> float:
        start, end = (self.POINTS[node] for node in self.INCIDENCE[beam])
        return sum((a - b) ** 2 for a, b in zip(start, end)) ** 0.5

    def support_nodes(self) -> list[int]:
        return [1, 4]

    def load_combination_cases(self) -> list[int]:
        return [101, 102]

    def steel_design_ratio(self, beam: int) -> float | None:
        # A member with no material is not designed, so losing the material
        # assignment shows up here rather than as a wrong answer.
        if not self.material.get(beam):
            return None
        return self.ratio

    def node_displacements(self, node: int, load_case: int) -> tuple[float, ...]:
        return (0.001, -0.002, 0.0, 0.0, 0.0, 0.0)

    # -- writes -----------------------------------------------------------
    def create_tapered_i_property(self, values_m) -> int:
        number = max(self.properties, default=0) + 1
        self.properties[number] = tuple(values_m)
        return number

    def assign_beam_property(self, beam: int, property_no: int) -> None:
        self.assigned[beam] = property_no
        # STAAD.Pro drops the member's material when its property changes.
        self.material.pop(beam, None)

    def assign_material_to_beam(self, material_name: str, beam: int) -> None:
        if material_name:
            self.material[beam] = material_name

    def set_silent_mode(self, enabled: bool = True) -> None:
        self.silent = enabled

    def save_model(self) -> None:
        self.saves += 1

    def update_structure(self) -> None:
        # STAAD.Pro's UpdateStructure restores the model from its .STD file:
        # it deletes the properties created since the last save, drops the
        # members back onto the sections the file gives them, and clears the
        # results. Calling it between assigning a candidate and analysing it
        # therefore throws the candidate away -- so the fake refuses rather
        # than letting the mistake come back unnoticed.
        raise AssertionError(
            "UpdateStructure reloads the model from disk and discards the "
            "sections just assigned. The analysis saves the model itself; "
            "nothing needs reloading before it."
        )

    def analyze(self) -> None:
        assert self.silent, "silent mode has to be on before STAAD.Pro is driven"
        self.analyses += 1
        self.analysed.append(
            {beam: self.properties[self.assigned[beam]] for beam in self.sections}
        )

    def assigned_section(self, beam: int) -> tuple[float, ...]:
        return self.properties[self.assigned[beam]]


def test_only_tapered_members_are_collected_from_the_model() -> None:
    frame = read_tapered_frame(FakeStaad(), settings())
    assert frame.numbers == (1, 2, 3)       # member 4 is rolled, so it is skipped


def test_the_selection_wins_over_the_whole_model() -> None:
    frame = read_tapered_frame(FakeStaad(selected=[2, 3]), settings())
    assert frame.numbers == (2, 3)


def test_a_model_without_a_design_block_is_refused() -> None:
    # A model whose analysis never reaches a CHECK CODE designs nothing, which
    # STAAD reports member by member rather than as a missing block.
    class Undesigned(FakeStaad):
        def steel_design_ratio(self, beam: int) -> float | None:
            return None

    with pytest.raises(OpenStaadError, match="designed none of the tapered members"):
        optimize_tapered_sections(Undesigned(), settings())


def test_every_property_write_puts_the_members_material_back() -> None:
    # Assigning a property clears the material, which in a model whose
    # materials come from one CONSTANTS block removes the block and leaves the
    # members undesigned. Both the search and the restore have to repair it.
    for apply_to_model in (False, True):
        staad = FakeStaad()
        optimize_tapered_sections(
            staad, settings(analysis_budget=12, apply_to_model=apply_to_model))
        assert staad.material == {1: "STEEL", 2: "STEEL", 3: "STEEL", 4: "STEEL"}


def test_the_straight_column_setting_reaches_the_frame_it_shapes() -> None:
    # Members 1 and 3 of the fake model are the columns; 2 is the rafter.
    frame = read_tapered_frame(FakeStaad(), settings(prismatic_columns=True))
    for column in (1, 3):
        assert frame.group_of_end[(column, 0)] == frame.group_of_end[(column, 1)]
    assert frame.group_of_end[(2, 0)] != frame.group_of_end[(2, 1)]
    loose = read_tapered_frame(FakeStaad(), settings())
    assert loose.group_of_end[(1, 0)] != loose.group_of_end[(1, 1)]


def test_every_analysis_grades_the_candidate_that_was_just_assigned() -> None:
    # The optimizer only learns anything if STAAD judges the sections it just
    # wrote. Reloading the model in between (UpdateStructure) puts the members
    # back on the sections they already had, so every candidate would come
    # back with the same ratios and the search would be reasoning about the
    # model it started with.
    staad = FakeStaad()
    optimize_tapered_sections(staad, settings(analysis_budget=12))
    assert len(staad.analysed) > 1
    assert any(seen != staad.analysed[0] for seen in staad.analysed[1:])


def test_the_run_saves_so_the_file_matches_the_model_it_leaves_behind() -> None:
    # Each analysis writes the candidate it is about to judge into the .STD
    # file, so a run that put the originals back only in memory would leave
    # the file holding the last section tried.
    for apply_to_model in (False, True):
        staad = FakeStaad()
        optimize_tapered_sections(
            staad, settings(analysis_budget=12, apply_to_model=apply_to_model))
        assert staad.saves == 1


def test_a_deflection_combination_missing_from_the_model_is_refused() -> None:
    missing = settings(deflection=DeflectionLimits(load_cases=(999,)))
    with pytest.raises(OpenStaadError, match="not found in the active model"):
        optimize_tapered_sections(FakeStaad(), missing)


def test_a_dry_run_puts_every_member_back_on_its_original_property() -> None:
    staad = FakeStaad()
    result = optimize_tapered_sections(staad, settings(analysis_budget=12))
    assert result.feasible and not result.applied
    assert result.saved_kg > 0               # it did find something lighter
    # ...but the model is left holding the very properties it started on, not
    # freshly created copies of the same dimensions.
    assert staad.assigned == {1: 101, 2: 102, 3: 103}
    for beam in (1, 2, 3):
        assert staad.assigned_section(beam) == pytest.approx(staad.sections[beam])


def test_applying_writes_the_optimized_sections_to_the_model() -> None:
    staad = FakeStaad()
    result = optimize_tapered_sections(
        staad, settings(analysis_budget=12, apply_to_model=True))
    assert result.feasible and result.applied
    optimized = {change.number: change.after for change in result.changes}
    assert staad.assigned_section(2) == pytest.approx(optimized[2].values_m())
    assert 4 not in staad.assigned          # the rolled member is never touched


def test_identical_dimensions_reuse_one_property_instead_of_piling_up() -> None:
    staad = FakeStaad(ratio=0.05)
    optimize_tapered_sections(staad, settings(analysis_budget=12))
    # Three members driven to the same minimum size share one property per
    # distinct set of dimensions, rather than one per member per analysis.
    assert len(staad.properties) < staad.analyses * 3


def test_written_sections_always_carry_two_identical_flanges() -> None:
    staad = FakeStaad()
    existing = set(staad.properties)
    optimize_tapered_sections(staad, settings(analysis_budget=12,
                                              apply_to_model=True))
    written = [values for number, values in staad.properties.items()
               if number not in existing]
    assert written
    for values in written:
        assert values[3] == values[5]        # F4 width  == F6 width
        assert values[4] == values[6]        # F5 thick  == F7 thick


def test_an_undesigned_member_stops_the_run_with_a_clear_message() -> None:
    class Undesigned(FakeStaad):
        def steel_design_ratio(self, beam: int) -> float | None:
            return None if beam == 2 else 0.4

    with pytest.raises(OpenStaadError, match="no steel design ratio"):
        optimize_tapered_sections(Undesigned(), settings())


def test_a_model_with_no_tapered_sections_says_so() -> None:
    class Rolled(FakeStaad):
        def section_property_values(self, beam: int) -> tuple[int, list[float]]:
            return 610, [0.0] * 24

    with pytest.raises(OpenStaadError, match="No tapered I sections"):
        optimize_tapered_sections(Rolled(), settings())


def test_a_frame_spanning_all_three_axes_is_refused() -> None:
    class Solid(FakeStaad):
        POINTS = {1: (0.0, 0.0, 0.0), 2: (0.0, 6.0, 3.0), 3: (10.0, 7.5, 0.0),
                  4: (10.0, 0.0, 6.0)}

    with pytest.raises(OpenStaadError, match="2D frames only"):
        optimize_tapered_sections(Solid(), settings())


def test_a_frame_drawn_in_the_yz_plane_is_handled_like_one_in_xy() -> None:
    class SidePlane(FakeStaad):
        POINTS = {1: (0.0, 0.0, 0.0), 2: (0.0, 6.0, 0.0), 3: (0.0, 7.5, 10.0),
                  4: (0.0, 0.0, 10.0)}

    frame = read_tapered_frame(SidePlane(), settings())
    rafter = next(chain for chain in frame.chains if not chain.is_column)
    assert rafter.span_m == pytest.approx(10.11, abs=0.01)


def test_settings_reject_limits_that_make_no_sense() -> None:
    with pytest.raises(ValueError, match="utilisation ceiling"):
        settings(utilisation_ceiling=1.5)
    with pytest.raises(ValueError, match="at least one load combination"):
        DeflectionLimits(load_cases=())
    with pytest.raises(ValueError, match="deflection limit"):
        DeflectionLimits(vertical_span_ratio=0.0, load_cases=(1,))
