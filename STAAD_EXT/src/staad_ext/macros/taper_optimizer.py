"""Automatic optimizer for the tapered I sections of a 2D STAAD.Pro frame.

Only the dimensions of tapered I sections (property type 680) change. Geometry,
supports, loading, load combinations and every non-tapered section are read but
never written, so the optimized model is the same structure with lighter plates.

Two rules shape the search:

* **Web depth is a property of the node, not of the member.** Members that
  continue one another through a node share a single depth variable, so a rafter
  built from several tapered members can never come back with a step in its web
  at a splice. Depths are solved per connected end group; a knee joint between
  two different members is only tied when the caller asks for it.
* **The two flanges are identical**, so each member carries one flange width and
  one flange thickness rather than separate top and bottom values.

``prismatic_columns`` extends the first rule to a whole run: every end of every
member in a column shares one depth, so the column comes back straight rather
than tapered. Its flange plates still vary member by member.

Sizes come off fixed fabrication ladders (see the module constants), and each
candidate is judged by STAAD.Pro itself: the caller supplies an ``evaluate``
callback that assigns the sections, re-runs the analysis, and reports back the
steel design ratios and the node displacements. That keeps the code check in
STAAD's hands -- whatever code the model is already set up for is the code the
optimizer respects.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from math import hypot
from typing import Any, Callable, Iterable, Mapping, Sequence

from staad_ext.models import Point3D, TaperOptimizerSettings
from staad_ext.openstaad import OpenStaad, OpenStaadError

TAPERED_I_PROPERTY_TYPE = 680
STEEL_DENSITY_KG_M3 = 7850.0

# Fabrication ladders. Web depth starts at 250mm and climbs in 50mm steps;
# flange width starts at 150mm and climbs in 25mm steps; both plate thicknesses
# come off the stocked plate list, and the flange is always thicker than the web.
DEPTH_START_MM = 250.0
DEPTH_STEP_MM = 50.0
FLANGE_WIDTH_START_MM = 150.0
FLANGE_WIDTH_STEP_MM = 25.0
PLATE_THICKNESS_MM: tuple[float, ...] = (
    5.0, 6.0, 8.0, 10.0, 12.0, 14.0, 16.0, 20.0, 25.0, 30.0, 32.0, 35.0, 40.0,
    45.0, 50.0,
)
# The flange must be thicker than the web, so the web can never reach the
# thickest stocked plate -- the one below it is the real ceiling.
MAX_WEB_THICKNESS_MM = PLATE_THICKNESS_MM[-2]

# Two members count as one continuous run when their axes are within about 2.5
# degrees of parallel.
_COLLINEAR_TOLERANCE = 0.999
# A frame is planar when every node sits within 1mm of the same plane.
_PLANAR_TOLERANCE_M = 0.001


# --------------------------------------------------------------------------
# Ladders
# --------------------------------------------------------------------------

def ladder(start: float, step: float, maximum: float) -> tuple[float, ...]:
    """Return the ascending fabrication sizes from ``start`` up to ``maximum``."""
    if step <= 0:
        raise ValueError("A ladder step must be greater than zero.")
    values = []
    value = start
    while value <= maximum + 1e-9:
        values.append(round(value, 6))
        value += step
    return tuple(values) or (start,)


def snap_up(value: float, options: Sequence[float]) -> float:
    """Return the smallest option at or above ``value`` (the largest if none is)."""
    return next((option for option in options if option >= value - 1e-9), options[-1])


def _neighbour(value: float, options: Sequence[float], upwards: bool) -> float | None:
    """Return the option one rung above or below ``value``, or None at the end."""
    if upwards:
        return next((option for option in options if option > value + 1e-9), None)
    below = [option for option in options if option < value - 1e-9]
    return below[-1] if below else None


def _thicknesses_above(value: float) -> tuple[float, ...]:
    return tuple(option for option in PLATE_THICKNESS_MM if option > value + 1e-9)


# --------------------------------------------------------------------------
# Sections and members
# --------------------------------------------------------------------------

@dataclass(frozen=True, slots=True)
class TaperedSection:
    """One tapered I section with identical top and bottom flanges."""

    start_depth_mm: float
    end_depth_mm: float
    web_thickness_mm: float
    flange_width_mm: float
    flange_thickness_mm: float

    @classmethod
    def from_property_values(cls, values: Sequence[float]) -> TaperedSection:
        """Build a section from the meter values STAAD reports for type 680.

        A zero F6/F7 means "the bottom flange repeats the top one", the same
        convention the plate summary follows.
        """
        if len(values) < 7:
            raise ValueError("A tapered I section needs 7 property values.")
        depth_1, web, depth_2, top_width, top_thickness = (
            abs(values[index]) * 1000.0 for index in range(5)
        )
        return cls(depth_1, depth_2, web, top_width, top_thickness)

    def values_m(self) -> tuple[float, ...]:
        """Return the seven F-values CreateTaperedIProperty expects, in meters."""
        width = self.flange_width_mm / 1000.0
        thickness = self.flange_thickness_mm / 1000.0
        return (self.start_depth_mm / 1000.0, self.web_thickness_mm / 1000.0,
                self.end_depth_mm / 1000.0, width, thickness, width, thickness)

    @property
    def is_prismatic(self) -> bool:
        return abs(self.start_depth_mm - self.end_depth_mm) < 1e-6

    def describe(self) -> str:
        depth = (f"{self.start_depth_mm:g}" if self.is_prismatic
                 else f"{self.start_depth_mm:g}→{self.end_depth_mm:g}")
        return (f"WEB {depth}×{self.web_thickness_mm:g} / "
                f"FLG {self.flange_width_mm:g}×{self.flange_thickness_mm:g}")

    def weight_kg(self, length_m: float) -> float:
        """Return the plate weight of this section over ``length_m``."""
        flange = self.flange_thickness_mm / 1000.0
        web_height = [max(depth / 1000.0 - 2 * flange, 0.0)
                      for depth in (self.start_depth_mm, self.end_depth_mm)]
        web_volume = (sum(web_height) / 2) * (self.web_thickness_mm / 1000.0) * length_m
        flange_volume = 2 * (self.flange_width_mm / 1000.0) * flange * length_m
        return (web_volume + flange_volume) * STEEL_DENSITY_KG_M3


@dataclass(frozen=True, slots=True)
class TaperedMember:
    """One tapered beam, with the geometry the optimizer needs to reason about."""

    number: int
    start_node: int
    end_node: int
    length_m: float
    start: Point3D
    end: Point3D
    section: TaperedSection

    @property
    def nodes(self) -> tuple[int, int]:
        return self.start_node, self.end_node

    def direction(self) -> tuple[float, float]:
        dx, dy = self.end.x - self.start.x, self.end.y - self.start.y
        length = hypot(dx, dy)
        return (dx / length, dy / length) if length > 1e-9 else (1.0, 0.0)


@dataclass(frozen=True, slots=True)
class Chain:
    """A run of collinear tapered members -- one rafter, or one column."""

    members: tuple[int, ...]
    nodes: tuple[int, ...]
    ends: tuple[int, int]
    is_column: bool
    span_m: float

    @property
    def interior_nodes(self) -> tuple[int, ...]:
        return tuple(node for node in self.nodes if node not in self.ends)


# --------------------------------------------------------------------------
# Design state
# --------------------------------------------------------------------------

@dataclass(slots=True)
class DesignState:
    """The free variables: one depth per tied end group, the rest per member."""

    depth_by_group: dict[int, float]
    web_by_member: dict[int, float]
    width_by_member: dict[int, float]
    flange_by_member: dict[int, float]

    def copy(self) -> DesignState:
        return DesignState(dict(self.depth_by_group), dict(self.web_by_member),
                           dict(self.width_by_member), dict(self.flange_by_member))


@dataclass(frozen=True, slots=True)
class TaperFrame:
    """The tapered members of one planar frame, with their tie and chain layout."""

    members: tuple[TaperedMember, ...]
    group_of_end: Mapping[tuple[int, int], int]
    chains: tuple[Chain, ...]
    points: Mapping[int, Point3D]
    base_level_m: float
    # The property each member started on, so a dry run can hand the model back
    # its own properties rather than equivalent copies.
    original_property: Mapping[int, int] = field(default_factory=dict)
    # The material each member started on. Assigning a property clears it, so
    # every write has to put it back.
    original_material: Mapping[int, str] = field(default_factory=dict)

    @property
    def numbers(self) -> tuple[int, ...]:
        return tuple(member.number for member in self.members)

    def members_in_group(self, group: int) -> tuple[int, ...]:
        return tuple(sorted({number for (number, _end), value
                             in self.group_of_end.items() if value == group}))

    def section_for(self, number: int, state: DesignState) -> TaperedSection:
        return TaperedSection(
            start_depth_mm=state.depth_by_group[self.group_of_end[(number, 0)]],
            end_depth_mm=state.depth_by_group[self.group_of_end[(number, 1)]],
            web_thickness_mm=state.web_by_member[number],
            flange_width_mm=state.width_by_member[number],
            flange_thickness_mm=state.flange_by_member[number],
        )

    def weight_kg(self, state: DesignState) -> float:
        return sum(self.section_for(member.number, state).weight_kg(member.length_m)
                   for member in self.members)

    def initial_state(self) -> DesignState:
        """Seed the variables from the sections the model already carries.

        Where several members are tied at a node their existing depths can
        disagree; the deepest one wins so the seed is never weaker than what the
        engineer drew.
        """
        depths: dict[int, float] = {}
        for member in self.members:
            for end, depth in enumerate(
                (member.section.start_depth_mm, member.section.end_depth_mm)
            ):
                group = self.group_of_end[(member.number, end)]
                depths[group] = max(depths.get(group, 0.0), depth)
        return DesignState(
            depths,
            {m.number: m.section.web_thickness_mm for m in self.members},
            {m.number: m.section.flange_width_mm for m in self.members},
            {m.number: m.section.flange_thickness_mm for m in self.members},
        )


# --------------------------------------------------------------------------
# Frame assembly
# --------------------------------------------------------------------------

class _UnionFind:
    def __init__(self) -> None:
        self._parent: dict[Any, Any] = {}

    def find(self, item: Any) -> Any:
        self._parent.setdefault(item, item)
        while self._parent[item] != item:
            self._parent[item] = self._parent[self._parent[item]]
            item = self._parent[item]
        return item

    def union(self, left: Any, right: Any) -> None:
        left_root, right_root = self.find(left), self.find(right)
        if left_root != right_root:
            self._parent[left_root] = right_root


def _is_collinear(first: TaperedMember, second: TaperedMember) -> bool:
    (ax, ay), (bx, by) = first.direction(), second.direction()
    return abs(ax * bx + ay * by) >= _COLLINEAR_TOLERANCE


def _end_groups(members: Sequence[TaperedMember], tie_all: bool,
                prismatic: Sequence[Chain] = ()) -> dict[tuple[int, int], int]:
    """Tie member ends that must share one web depth into numbered groups."""
    union = _UnionFind()
    for member in members:
        for end in (0, 1):
            union.find((member.number, end))
    for index, first in enumerate(members):
        for second in members[index + 1:]:
            if not (tie_all or _is_collinear(first, second)):
                continue
            for first_end, first_node in enumerate(first.nodes):
                for second_end, second_node in enumerate(second.nodes):
                    if first_node == second_node:
                        union.union((first.number, first_end),
                                    (second.number, second_end))
    # A run held prismatic is one depth from end to end, so every end of every
    # member in it joins a single group. Tying the run rather than each of its
    # members is what stops a column split at an intermediate node from coming
    # back as two straight pieces of different depth.
    for chain in prismatic:
        ends = [(number, end) for number in chain.members for end in (0, 1)]
        for other in ends[1:]:
            union.union(ends[0], other)
    roots = {}
    groups: dict[tuple[int, int], int] = {}
    for member in members:
        for end in (0, 1):
            root = union.find((member.number, end))
            groups[(member.number, end)] = roots.setdefault(root, len(roots))
    return groups


def _build_chains(members: Sequence[TaperedMember]) -> tuple[Chain, ...]:
    """Group collinear connected members into rafter/column runs."""
    union = _UnionFind()
    for member in members:
        union.find(member.number)
    for index, first in enumerate(members):
        for second in members[index + 1:]:
            if set(first.nodes) & set(second.nodes) and _is_collinear(first, second):
                union.union(first.number, second.number)

    by_root: dict[Any, list[TaperedMember]] = {}
    for member in members:
        by_root.setdefault(union.find(member.number), []).append(member)

    chains = []
    for grouped in by_root.values():
        counts: dict[int, int] = {}
        points: dict[int, Point3D] = {}
        for member in grouped:
            for node, point in zip(member.nodes, (member.start, member.end)):
                counts[node] = counts.get(node, 0) + 1
                points[node] = point
        # A run's two extreme nodes are the ones only one member touches; a
        # closed loop has none, in which case the run gets no chord to measure
        # deflection against and is left out of the serviceability checks.
        extremes = sorted(node for node, count in counts.items() if count == 1)
        if len(extremes) != 2:
            continue
        first_point, second_point = points[extremes[0]], points[extremes[1]]
        dx = second_point.x - first_point.x
        dy = second_point.y - first_point.y
        chains.append(Chain(
            members=tuple(sorted(member.number for member in grouped)),
            nodes=tuple(sorted(counts)),
            ends=(extremes[0], extremes[1]),
            is_column=abs(dy) > abs(dx),
            span_m=hypot(dx, dy),
        ))
    return tuple(sorted(chains, key=lambda chain: chain.members))


def build_frame(members: Sequence[TaperedMember], points: Mapping[int, Point3D],
                base_level_m: float, tie_all: bool,
                original_property: Mapping[int, int] | None = None,
                original_material: Mapping[int, str] | None = None,
                prismatic_columns: bool = False) -> TaperFrame:
    ordered = tuple(sorted(members, key=lambda member: member.number))
    chains = _build_chains(ordered)
    prismatic = tuple(chain for chain in chains
                      if chain.is_column) if prismatic_columns else ()
    return TaperFrame(
        members=ordered,
        group_of_end=_end_groups(ordered, tie_all, prismatic),
        chains=chains,
        points=dict(points),
        base_level_m=base_level_m,
        original_property=dict(original_property or {}),
        original_material=dict(original_material or {}),
    )


# --------------------------------------------------------------------------
# Feasibility of a candidate set of dimensions
# --------------------------------------------------------------------------

def _section_is_valid(section: TaperedSection,
                      settings: TaperOptimizerSettings) -> bool:
    """Check the rules a set of plate sizes must obey regardless of the analysis."""
    if section.flange_thickness_mm <= section.web_thickness_mm:
        return False
    if max(section.start_depth_mm, section.end_depth_mm) > settings.max_depth_mm:
        return False
    if section.flange_width_mm > settings.max_flange_width_mm:
        return False
    clear = max(section.start_depth_mm, section.end_depth_mm) - 2 * section.flange_thickness_mm
    return clear / section.web_thickness_mm <= settings.max_web_slenderness


def _state_is_valid(state: DesignState, frame: TaperFrame,
                    settings: TaperOptimizerSettings) -> bool:
    return all(_section_is_valid(frame.section_for(member.number, state), settings)
               for member in frame.members)


def _repair_flanges(state: DesignState, frame: TaperFrame) -> bool:
    """Lift any flange that is no longer thicker than its web. False if impossible."""
    for member in frame.members:
        web = state.web_by_member[member.number]
        if state.flange_by_member[member.number] > web + 1e-9:
            continue
        options = _thicknesses_above(web)
        if not options:
            return False
        state.flange_by_member[member.number] = options[0]
    return True


def snap_to_ladders(state: DesignState, frame: TaperFrame,
                    settings: TaperOptimizerSettings) -> DesignState:
    """Round every variable up onto its fabrication ladder.

    Rounding up rather than to nearest keeps the starting point at least as
    strong as the section the engineer drew, so the first analysis is expected
    to pass and the search only ever has to work downwards.
    """
    depths = ladder(DEPTH_START_MM, DEPTH_STEP_MM, settings.max_depth_mm)
    widths = ladder(FLANGE_WIDTH_START_MM, FLANGE_WIDTH_STEP_MM,
                    settings.max_flange_width_mm)
    snapped = DesignState(
        {group: snap_up(value, depths)
         for group, value in state.depth_by_group.items()},
        {number: snap_up(value, PLATE_THICKNESS_MM)
         for number, value in state.web_by_member.items()},
        {number: snap_up(value, widths)
         for number, value in state.width_by_member.items()},
        {number: snap_up(value, PLATE_THICKNESS_MM)
         for number, value in state.flange_by_member.items()},
    )
    _repair_flanges(snapped, frame)
    return snapped


def prepare_seed(frame: TaperFrame,
                 settings: TaperOptimizerSettings) -> tuple[DesignState, tuple[str, ...]]:
    """Snap the model's own sections onto the ladders and make them legal.

    A model can arrive carrying sections that break the rules the optimizer
    works to -- most often a web too slender for the configured d/t cap. Every
    candidate step is judged against the whole state, so an illegal starting
    point would reject every step and hand the illegal sections straight back.
    The seed is therefore repaired up front: webs are thickened until they
    comply, and a section that cannot be made legal at all stops the run.
    """
    for member in frame.members:
        deepest = max(member.section.start_depth_mm, member.section.end_depth_mm)
        if deepest > settings.max_depth_mm:
            raise ValueError(
                f"Member {member.number} already has a {deepest:g}mm web, deeper "
                f"than the {settings.max_depth_mm:g}mm maximum. Raise the maximum "
                "depth so the optimizer can start from the section that is there."
            )
        if member.section.flange_width_mm > settings.max_flange_width_mm:
            raise ValueError(
                f"Member {member.number} already has a "
                f"{member.section.flange_width_mm:g}mm flange, wider than the "
                f"{settings.max_flange_width_mm:g}mm maximum. Raise the maximum "
                "flange width so the optimizer can start from the section that "
                "is there."
            )

    state = snap_to_ladders(frame.initial_state(), frame, settings)
    thickened = _repair_slenderness(state, frame, settings)
    if not thickened:
        return state, ()
    members = ", ".join(str(number) for number in sorted(thickened))
    return state, (
        f"The web of member(s) {members} was thickened before optimizing: the "
        f"section in the model exceeds the d/t limit of "
        f"{settings.max_web_slenderness:g}.",
    )


def _repair_slenderness(state: DesignState, frame: TaperFrame,
                        settings: TaperOptimizerSettings) -> set[int]:
    """Thicken any web too slender for the cap. Returns the members changed."""
    thickened: set[int] = set()
    for member in frame.members:
        # Thickening the web can push it past the flange, and lifting the flange
        # changes the clear web depth, so the two settle against each other over
        # a few passes rather than in one. Slenderness is tested before the
        # flange is lifted so that a web which simply cannot meet the cap is
        # reported as such, rather than as the flange failure it causes.
        for _attempt in range(len(PLATE_THICKNESS_MM) + 1):
            section = frame.section_for(member.number, state)
            deepest = max(section.start_depth_mm, section.end_depth_mm)
            clear = deepest - 2 * section.flange_thickness_mm
            if clear / section.web_thickness_mm <= settings.max_web_slenderness:
                break
            thicker = _neighbour(section.web_thickness_mm, PLATE_THICKNESS_MM, True)
            # The flange has to sit above the web, so the thickest usable web is
            # the second-thickest stocked plate, not the thickest.
            if thicker is None or not _thicknesses_above(thicker):
                raise ValueError(
                    f"Member {member.number}: a {deepest:g}mm web cannot meet the "
                    f"d/t limit of {settings.max_web_slenderness:g} on any stocked "
                    f"plate that still leaves room for a thicker flange "
                    f"(at most {MAX_WEB_THICKNESS_MM:g}mm). Raise the slenderness "
                    "cap or lower the maximum depth."
                )
            state.web_by_member[member.number] = thicker
            thickened.add(member.number)
            _repair_flanges(state, frame)
    return thickened


# --------------------------------------------------------------------------
# Evaluation
# --------------------------------------------------------------------------

@dataclass(frozen=True, slots=True)
class DeflectionCheck:
    kind: str                       # "Vertical" or "Horizontal"
    node: int
    load_case: int
    actual_mm: float
    allowed_mm: float
    span_m: float

    @property
    def passes(self) -> bool:
        return self.actual_mm <= self.allowed_mm + 1e-9

    @property
    def utilisation(self) -> float:
        return self.actual_mm / self.allowed_mm if self.allowed_mm > 0 else 0.0


@dataclass(frozen=True, slots=True)
class Evaluation:
    """What STAAD.Pro reported back for one candidate set of sections."""

    ratios: Mapping[int, float]
    deflections: tuple[DeflectionCheck, ...] = ()

    def failed_members(self, ceiling: float) -> tuple[int, ...]:
        return tuple(sorted(number for number, ratio in self.ratios.items()
                            if ratio > ceiling + 1e-9))

    def failed_deflections(self) -> tuple[DeflectionCheck, ...]:
        return tuple(check for check in self.deflections if not check.passes)

    def is_feasible(self, ceiling: float) -> bool:
        return not self.failed_members(ceiling) and not self.failed_deflections()


def deflection_checks(frame: TaperFrame, displacements: Mapping[tuple[int, int], float],
                      vertical: Mapping[tuple[int, int], float],
                      load_cases: Iterable[int], vertical_ratio: float,
                      horizontal_ratio: float) -> tuple[DeflectionCheck, ...]:
    """Turn node displacements into span/height deflection checks.

    A column run is checked for horizontal drift at its top node against that
    node's height above the support level; a rafter run is checked for vertical
    sag at every interior node, measured against the straight chord between the
    run's two extreme nodes so the check is not fooled by the ends settling.
    """
    checks: list[DeflectionCheck] = []
    for load_case in load_cases:
        for chain in frame.chains:
            if chain.span_m <= 0:
                continue
            if chain.is_column:
                top = max(chain.ends, key=lambda node: frame.points[node].y)
                height = frame.points[top].y - frame.base_level_m
                if height <= 0:
                    continue
                checks.append(DeflectionCheck(
                    "Horizontal", top, load_case,
                    abs(displacements[(top, load_case)]) * 1000.0,
                    height / horizontal_ratio * 1000.0, height,
                ))
                continue
            first, second = chain.ends
            start, end = frame.points[first], frame.points[second]
            run = hypot(end.x - start.x, end.y - start.y)
            if run <= 0:
                continue
            for node in chain.interior_nodes:
                point = frame.points[node]
                fraction = hypot(point.x - start.x, point.y - start.y) / run
                chord = (vertical[(first, load_case)] * (1 - fraction)
                         + vertical[(second, load_case)] * fraction)
                checks.append(DeflectionCheck(
                    "Vertical", node, load_case,
                    abs(vertical[(node, load_case)] - chord) * 1000.0,
                    chain.span_m / vertical_ratio * 1000.0, chain.span_m,
                ))
    return tuple(checks)


# --------------------------------------------------------------------------
# Search
# --------------------------------------------------------------------------

@dataclass(frozen=True, slots=True)
class Variable:
    kind: str       # "depth" | "web" | "width" | "flange"
    key: int        # end-group id for "depth", member number for the rest


def _variables(frame: TaperFrame) -> tuple[Variable, ...]:
    groups = sorted(set(frame.group_of_end.values()))
    return (tuple(Variable("depth", group) for group in groups)
            + tuple(Variable(kind, member.number)
                    for member in frame.members
                    for kind in ("web", "width", "flange")))


def _affected_members(frame: TaperFrame, variable: Variable) -> tuple[int, ...]:
    if variable.kind == "depth":
        return frame.members_in_group(variable.key)
    return (variable.key,)


def _stepped(state: DesignState, variable: Variable, frame: TaperFrame,
             settings: TaperOptimizerSettings, upwards: bool) -> DesignState | None:
    """Return ``state`` with one variable moved a single rung, or None."""
    candidate = state.copy()
    if variable.kind == "depth":
        options: Sequence[float] = ladder(DEPTH_START_MM, DEPTH_STEP_MM,
                                          settings.max_depth_mm)
        current = candidate.depth_by_group[variable.key]
        target = candidate.depth_by_group
    elif variable.kind == "width":
        options = ladder(FLANGE_WIDTH_START_MM, FLANGE_WIDTH_STEP_MM,
                         settings.max_flange_width_mm)
        current = candidate.width_by_member[variable.key]
        target = candidate.width_by_member
    elif variable.kind == "web":
        options = PLATE_THICKNESS_MM
        current = candidate.web_by_member[variable.key]
        target = candidate.web_by_member
    else:
        options = PLATE_THICKNESS_MM
        current = candidate.flange_by_member[variable.key]
        target = candidate.flange_by_member

    value = _neighbour(current, options, upwards)
    if value is None:
        return None
    target[variable.key] = value
    # A thicker web can overtake its flange, and a thinner flange can drop to
    # the web; both break the flange-thicker-than-web rule, so lift the flange
    # back above the web before judging the candidate.
    if variable.kind == "web" and upwards and not _repair_flanges(candidate, frame):
        return None
    if not _state_is_valid(candidate, frame, settings):
        return None
    return candidate


def _saving_kg(state: DesignState, candidate: DesignState, frame: TaperFrame) -> float:
    return frame.weight_kg(state) - frame.weight_kg(candidate)


def _descent_options(state: DesignState, frame: TaperFrame,
                     settings: TaperOptimizerSettings, locked: set[Variable]
                     ) -> list[tuple[Variable, DesignState, float]]:
    """Every single downward step available now, heaviest saving first."""
    options = []
    for variable in _variables(frame):
        if variable in locked:
            continue
        candidate = _stepped(state, variable, frame, settings, upwards=False)
        if candidate is None:
            continue
        saving = _saving_kg(state, candidate, frame)
        if saving > 1e-9:
            options.append((variable, candidate, saving))
    return sorted(options, key=lambda option: -option[2])


def _bulk_step_down(state: DesignState, evaluation: Evaluation, frame: TaperFrame,
                    settings: TaperOptimizerSettings,
                    threshold: float) -> DesignState | None:
    """Step down every variable whose members all sit below ``threshold``.

    One analysis then tests a whole layer of reductions at once. A greedy
    one-variable-at-a-time search would need a separate analysis per step, and
    an analysis is by far the most expensive thing this optimizer does.
    """
    candidate = state
    moved = False
    for variable in _variables(frame):
        members = _affected_members(frame, variable)
        if any(evaluation.ratios.get(number, 1.0) > threshold for number in members):
            continue
        stepped = _stepped(candidate, variable, frame, settings, upwards=False)
        if stepped is not None:
            candidate, moved = stepped, True
    return candidate if moved else None


def _step_up(state: DesignState, members: Iterable[int], frame: TaperFrame,
             settings: TaperOptimizerSettings) -> DesignState | None:
    """Strengthen the given members by one rung, or None if none can grow.

    Depth is tried first because it buys the most bending capacity per kilogram
    of added steel; web thickness, flange thickness and flange width follow.
    """
    targets = set(members)
    candidate = state
    moved = False
    for kind in ("depth", "web", "flange", "width"):
        for variable in _variables(frame):
            if variable.kind != kind:
                continue
            if not targets & set(_affected_members(frame, variable)):
                continue
            stepped = _stepped(candidate, variable, frame, settings, upwards=True)
            if stepped is not None:
                candidate, moved = stepped, True
                targets -= set(_affected_members(frame, variable))
        if not targets:
            break
    return candidate if moved else None


@dataclass(frozen=True, slots=True)
class MemberChange:
    number: int
    length_m: float
    before: TaperedSection
    after: TaperedSection
    ratio: float | None

    @property
    def weight_before_kg(self) -> float:
        return self.before.weight_kg(self.length_m)

    @property
    def weight_after_kg(self) -> float:
        return self.after.weight_kg(self.length_m)

    @property
    def changed(self) -> bool:
        return self.before != self.after


@dataclass(frozen=True, slots=True)
class OptimizationResult:
    changes: tuple[MemberChange, ...]
    state: DesignState
    evaluation: Evaluation | None
    analyses_used: int
    feasible: bool
    budget_exhausted: bool = False
    notes: tuple[str, ...] = ()
    applied: bool = False

    @property
    def weight_before_kg(self) -> float:
        return sum(change.weight_before_kg for change in self.changes)

    @property
    def weight_after_kg(self) -> float:
        return sum(change.weight_after_kg for change in self.changes)

    @property
    def saved_kg(self) -> float:
        return self.weight_before_kg - self.weight_after_kg

    @property
    def saved_percent(self) -> float:
        before = self.weight_before_kg
        return (self.saved_kg / before * 100.0) if before > 0 else 0.0


def _changes(frame: TaperFrame, state: DesignState,
             evaluation: Evaluation | None) -> tuple[MemberChange, ...]:
    ratios = evaluation.ratios if evaluation else {}
    return tuple(MemberChange(
        member.number, member.length_m, member.section,
        frame.section_for(member.number, state), ratios.get(member.number),
    ) for member in frame.members)


def optimize(frame: TaperFrame, settings: TaperOptimizerSettings,
             evaluate: Callable[[DesignState], Evaluation],
             progress: Callable[[str], None] | None = None) -> OptimizationResult:
    """Search for the lightest set of tapered sections STAAD.Pro still accepts.

    ``evaluate`` assigns a candidate to the model, re-runs the analysis and
    returns the resulting design ratios and deflection checks. It is called once
    per analysis, and the ``analysis_budget`` setting caps how many times.
    """
    def report(message: str) -> None:
        if progress is not None:
            progress(message)

    ceiling = settings.utilisation_ceiling
    state, seed_notes = prepare_seed(frame, settings)
    notes: list[str] = list(seed_notes)
    used = 0

    report("Checking the starting sections…")
    evaluation = evaluate(state)
    used += 1

    # Phase 1 -- lift the frame until it passes. A model that already passes
    # skips this entirely, which is the usual case.
    while not evaluation.is_feasible(ceiling) and used < settings.analysis_budget:
        failing = set(evaluation.failed_members(ceiling))
        # A failed deflection belongs to a whole run, not to one node, so the
        # entire chain that owns the node has to stiffen.
        for check in evaluation.failed_deflections():
            for chain in frame.chains:
                if check.node in chain.nodes:
                    failing.update(chain.members)
        stronger = _step_up(state, failing or set(frame.numbers), frame, settings)
        if stronger is None:
            notes.append(
                "The sections could not be strengthened any further within the "
                "depth, width and slenderness limits."
            )
            break
        state = stronger
        report(f"Strengthening ({used} analyses so far)…")
        evaluation = evaluate(state)
        used += 1

    if not evaluation.is_feasible(ceiling):
        return OptimizationResult(
            _changes(frame, state, evaluation), state, evaluation, used,
            feasible=False, budget_exhausted=used >= settings.analysis_budget,
            notes=tuple(notes),
        )

    best_state, best_evaluation = state, evaluation

    # Phase 2 -- bulk descent. Each pass steps down every variable with headroom
    # and tests them together; tightening the threshold after a failure narrows
    # the set instead of abandoning the descent.
    for fraction in (1.0, 0.75, 0.5):
        while used < settings.analysis_budget:
            candidate = _bulk_step_down(best_state, best_evaluation, frame,
                                        settings, ceiling * fraction)
            if candidate is None:
                break
            report(f"Reducing sections ({used} analyses so far)…")
            evaluation = evaluate(candidate)
            used += 1
            if not evaluation.is_feasible(ceiling):
                break
            best_state, best_evaluation = candidate, evaluation

    # Phase 3 -- single steps, heaviest saving first, locking whatever fails.
    locked: set[Variable] = set()
    while used < settings.analysis_budget:
        options = _descent_options(best_state, frame, settings, locked)
        if not options:
            break
        variable, candidate, _saving = options[0]
        report(f"Fine-tuning ({used} analyses so far)…")
        evaluation = evaluate(candidate)
        used += 1
        if evaluation.is_feasible(ceiling):
            best_state, best_evaluation = candidate, evaluation
        else:
            locked.add(variable)

    exhausted = used >= settings.analysis_budget
    if exhausted and _descent_options(best_state, frame, settings, locked):
        notes.append(
            f"Stopped at the {settings.analysis_budget}-analysis budget with "
            "reductions still untried -- raise the budget to keep searching."
        )
    return OptimizationResult(
        _changes(frame, best_state, best_evaluation), best_state, best_evaluation,
        used, feasible=True, budget_exhausted=exhausted, notes=tuple(notes),
    )


# --------------------------------------------------------------------------
# STAAD.Pro plumbing
# --------------------------------------------------------------------------

def read_tapered_frame(staad: OpenStaad,
                       settings: TaperOptimizerSettings) -> TaperFrame:
    """Read the tapered members of the active model into a frame.

    The members selected in STAAD.Pro win; with nothing selected the whole model
    is optimized. Either way only property type 680 is collected -- every other
    section is left exactly as it is.
    """
    beams = staad.selected_beams() or staad.all_beams()
    if not beams:
        raise OpenStaadError("The active STAAD.Pro model has no beam members.")

    members: list[TaperedMember] = []
    points: dict[int, Point3D] = {}
    original_property: dict[int, int] = {}
    original_material: dict[int, str] = {}
    for beam in beams:
        property_type, values = staad.section_property_values(beam)
        if property_type != TAPERED_I_PROPERTY_TYPE:
            continue
        start_node, end_node = staad.member_incidence(beam)
        if start_node <= 0 or end_node <= 0:
            continue
        start, end = staad.node_coordinates(start_node), staad.node_coordinates(end_node)
        points[start_node], points[end_node] = start, end
        reference = staad.beam_property_ref(beam)
        if reference > 0:
            original_property[beam] = reference
        material = staad.beam_material_name(beam)
        if material:
            original_material[beam] = material
        try:
            section = TaperedSection.from_property_values(values)
        except ValueError as exc:
            raise OpenStaadError(f"Member {beam}: {exc}") from exc
        if min(section.start_depth_mm, section.end_depth_mm,
               section.web_thickness_mm, section.flange_width_mm,
               section.flange_thickness_mm) <= 0:
            raise OpenStaadError(
                f"Member {beam}: the tapered section dimensions are incomplete."
            )
        members.append(TaperedMember(beam, start_node, end_node,
                                     staad.beam_length(beam), start, end, section))

    if not members:
        raise OpenStaadError(
            "No tapered I sections were found. Select the tapered members to "
            "optimize, or open a model that uses tapered I sections."
        )

    spread = [max(point.z for point in points.values()) - min(point.z for point in points.values()),
              max(point.x for point in points.values()) - min(point.x for point in points.values())]
    if min(spread) > _PLANAR_TOLERANCE_M:
        raise OpenStaadError(
            "This optimizer handles 2D frames only, but the tapered members span "
            "all three axes. Select the members of a single frame and retry."
        )
    # A frame drawn in the YZ plane is the same problem with Z where X usually
    # is, so fold it onto X and keep Y as the vertical axis either way.
    if spread[0] > _PLANAR_TOLERANCE_M:
        points = {node: Point3D(point.z, point.y) for node, point in points.items()}
        members = [TaperedMember(m.number, m.start_node, m.end_node, m.length_m,
                                 points[m.start_node], points[m.end_node], m.section)
                   for m in members]

    try:
        supports = [node for node in staad.support_nodes() if node in points]
    except (OpenStaadError, OSError, TypeError, ValueError):
        supports = []
    levels = [points[node].y for node in supports] or [
        point.y for point in points.values()]
    return build_frame(members, points, min(levels),
                       settings.tie_depths_at_all_shared_nodes, original_property,
                       original_material, settings.prismatic_columns)


def _assign(staad: OpenStaad, frame: TaperFrame, number: int,
            property_no: int) -> None:
    """Put a member on a property, keeping the material it came with.

    STAAD.Pro drops a member's material assignment when its property is
    re-assigned. In a model whose materials come from a single `CONSTANTS`
    block that removes the block outright, and a member with no material is
    not designed -- so the code check the optimizer relies on would quietly
    stop reporting ratios.
    """
    staad.assign_beam_property(number, property_no)
    staad.assign_material_to_beam(frame.original_material.get(number, ""), number)


def _assign_sections(staad: OpenStaad, frame: TaperFrame, state: DesignState,
                     cache: dict[tuple[float, ...], int]) -> None:
    """Create and assign the candidate sections, reusing identical properties.

    CreateTaperedIProperty mints a new property every call, so an unguarded
    search would leave hundreds of orphans in the model's property table. The
    cache means one property per distinct set of dimensions per run.
    """
    for member in frame.members:
        values = frame.section_for(member.number, state).values_m()
        key = tuple(round(value, 9) for value in values)
        if key not in cache:
            cache[key] = staad.create_tapered_i_property(values)
        _assign(staad, frame, member.number, cache[key])


def _restore_original(staad: OpenStaad, frame: TaperFrame,
                      cache: dict[tuple[float, ...], int]) -> None:
    """Put every member back on the property it was assigned when the run began.

    Re-assigning the original property number leaves the model holding its own
    property objects rather than fresh copies of the same dimensions. A member
    whose property number could not be read falls back to being re-created from
    the dimensions that were measured off it.
    """
    original = frame.initial_state()
    for member in frame.members:
        reference = frame.original_property.get(member.number, 0)
        if reference > 0:
            _assign(staad, frame, member.number, reference)
            continue
        values = frame.section_for(member.number, original).values_m()
        key = tuple(round(value, 9) for value in values)
        if key not in cache:
            cache[key] = staad.create_tapered_i_property(values)
        _assign(staad, frame, member.number, cache[key])


def make_evaluator(staad: OpenStaad, frame: TaperFrame,
                   settings: TaperOptimizerSettings,
                   progress: Callable[[str], None] | None = None
                   ) -> Callable[[DesignState], Evaluation]:
    """Build the callback that asks STAAD.Pro to judge a candidate."""
    cache: dict[tuple[float, ...], int] = {}
    load_cases = settings.deflection.load_cases

    def evaluate(state: DesignState) -> Evaluation:
        # Nothing is saved or reloaded between assigning the candidate and
        # analysing it. The analysis writes the model out and runs the engine
        # over what it wrote, so the sections just assigned are the ones
        # judged. Reloading here (UpdateStructure) would restore the model from
        # the file as it stood *before* those edits: the properties created for
        # this candidate are deleted, the members drop back to the sections
        # they already had, and the results are cleared -- so the analysis then
        # either grades the wrong sections or reports nothing at all.
        _assign_sections(staad, frame, state, cache)
        if progress is not None:
            progress("Running the STAAD.Pro analysis…")
        staad.analyze()

        ratios: dict[int, float] = {}
        undesigned = []
        for member in frame.members:
            ratio = staad.steel_design_ratio(member.number)
            if ratio is None:
                undesigned.append(member.number)
            else:
                ratios[member.number] = ratio
        if undesigned:
            raise OpenStaadError(_undesigned_message(undesigned, ratios))

        horizontal: dict[tuple[int, int], float] = {}
        vertical: dict[tuple[int, int], float] = {}
        for load_case in load_cases:
            for node in frame.points:
                dx, dy, _dz, *_ = staad.node_displacements(node, load_case)
                horizontal[(node, load_case)] = dx
                vertical[(node, load_case)] = dy
        return Evaluation(ratios, deflection_checks(
            frame, horizontal, vertical, load_cases,
            settings.deflection.vertical_span_ratio,
            settings.deflection.horizontal_height_ratio,
        ))

    return evaluate


def _undesigned_message(undesigned: list[int], ratios: dict[int, float]) -> str:
    """Explain a missing design ratio.

    The optimizer sizes sections against STAAD.Pro's own code check, so it can
    only judge a member the model actually designs. STAAD reports this the same
    way whether the design block is missing entirely or simply does not list
    the member, so the two cases are told apart by whether any member at all
    came back with a ratio.
    """
    members = ", ".join(str(number) for number in undesigned)
    if not ratios:
        return (
            "STAAD.Pro designed none of the tapered members, so the optimizer has "
            "nothing to judge candidate sections by. Check that the model has a "
            "PARAMETER / CHECK CODE block covering members " + members + ", that "
            "the analysis reaches it (a CHECK CODE after FINISH, or behind a LOAD "
            "LIST selecting no cases, never runs), that those members still have a "
            "material assigned, and that the design code in use checks tapered I "
            "sections at all."
        )
    return (
        f"STAAD.Pro returned no steel design ratio for member(s) {members}. Add "
        "them to the PARAMETER / CHECK CODE block so the optimizer can judge them."
    )


def _preflight(staad: OpenStaad, settings: TaperOptimizerSettings) -> None:
    available = set(staad.load_combination_cases())
    missing = [case for case in settings.deflection.load_cases
               if case not in available]
    if missing:
        raise OpenStaadError(
            "Load combination(s) not found in the active model: "
            f"{', '.join(str(case) for case in missing)}."
        )


def optimize_tapered_sections(staad: OpenStaad, settings: TaperOptimizerSettings,
                              progress: Callable[[str], None] | None = None
                              ) -> OptimizationResult:
    """Optimize the tapered sections of the active model.

    Unless ``settings.apply_to_model`` is set, every member is put back on the
    property it started with before returning, so a dry run changes nothing
    about the structure. Searching does add the candidate section properties it
    had to assign to the model's property table; they are left unassigned.

    The model is saved on the way out either way. Every analysis writes the
    candidate it is about to judge into the .STD file, so without a final save
    the file would be left holding the last section tried rather than the one
    the run decided on.
    """
    # Before the first property write, not just before the first analysis: an
    # assignment that raises a dialog blocks until somebody clicks it.
    staad.set_silent_mode(True)
    frame = read_tapered_frame(staad, settings)
    _preflight(staad, settings)

    cache: dict[tuple[float, ...], int] = {}
    result = optimize(frame, settings, make_evaluator(staad, frame, settings, progress),
                      progress)

    applied = settings.apply_to_model and result.feasible
    if progress is not None:
        progress("Applying the optimized sections…" if applied
                 else "Restoring the original sections…")
    if applied:
        _assign_sections(staad, frame, result.state, cache)
    else:
        _restore_original(staad, frame, cache)
    staad.save_model()

    notes = result.notes
    if settings.apply_to_model and not applied:
        notes += ("The model was left unchanged because no feasible set of "
                  "sections was found.",)
    return OptimizationResult(
        result.changes, result.state, result.evaluation, result.analyses_used,
        result.feasible, result.budget_exhausted, notes, applied,
    )
