from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum


class ViewPlane(StrEnum):
    XY = "XY"
    YZ = "YZ"
    ZX = "ZX"


@dataclass(frozen=True, slots=True)
class Point3D:
    x: float
    y: float
    z: float = 0.0


@dataclass(frozen=True, slots=True)
class ExportSettings:
    plane: ViewPlane = ViewPlane.XY
    write_labels: bool = True
    text_scale: float = 1.0
    peb_corner_joins: bool = False
    connection_face_lines: bool = False
    color_by_section: bool = True

    def __post_init__(self) -> None:
        if not 0.1 <= self.text_scale <= 10.0:
            raise ValueError("text_scale must be between 0.1 and 10.0")


@dataclass(frozen=True, slots=True)
class SectionEnvelope:
    start_half_width: float
    end_half_width: float
    property_type: int


class ScheduleCorner(StrEnum):
    TOP_LEFT = "Top left"
    TOP_RIGHT = "Top right"
    BOTTOM_LEFT = "Bottom left"
    BOTTOM_RIGHT = "Bottom right"


@dataclass(frozen=True, slots=True)
class GaExportSettings:
    """Settings for the general-arrangement DXF export.

    The GA drawing marks each member with a bubbled number and moves the
    section descriptions into one MEMBER SIZE SCHEDULE beside the frame, so it
    has no per-member label or color-by-section option.
    """

    plane: ViewPlane = ViewPlane.XY
    text_scale: float = 1.0
    schedule_corner: ScheduleCorner = ScheduleCorner.TOP_RIGHT
    blank_rows: int = 4
    write_marks: bool = True
    write_centerlines: bool = True
    # Bubbles are sized from the frame and capped against member spacing, but
    # kept off text_scale: shrinking crowded bubbles must not shrink the
    # schedule along with them.
    mark_scale: float = 1.0

    def __post_init__(self) -> None:
        if not 0.1 <= self.text_scale <= 10.0:
            raise ValueError("text_scale must be between 0.1 and 10.0")
        if not 0.1 <= self.mark_scale <= 10.0:
            raise ValueError("mark_scale must be between 0.1 and 10.0")
        if not 0 <= self.blank_rows <= 40:
            raise ValueError("blank_rows must be between 0 and 40")


@dataclass(frozen=True, slots=True)
class IfcExportSettings:
    selected_only: bool = False


@dataclass(frozen=True, slots=True)
class DeflectionLimits:
    """Serviceability limits the optimized frame must still satisfy.

    Both are span/height divisors: 240 means "vertical sag must stay within
    span/240". The checks run only for ``load_cases``, which are the
    serviceability combinations -- optimizing against strength combinations
    would size the frame for a deflection that is never checked in practice.
    """

    vertical_span_ratio: float = 240.0
    horizontal_height_ratio: float = 150.0
    load_cases: tuple[int, ...] = ()

    def __post_init__(self) -> None:
        for label, ratio in (("vertical", self.vertical_span_ratio),
                             ("horizontal", self.horizontal_height_ratio)):
            if not 1.0 <= ratio <= 10000.0:
                raise ValueError(
                    f"The {label} deflection limit must be between L/1 and L/10000."
                )
        if not self.load_cases:
            raise ValueError(
                "Enter at least one load combination to check deflections against."
            )
        if any(case <= 0 for case in self.load_cases):
            raise ValueError("Load combination numbers must be greater than zero.")


@dataclass(frozen=True, slots=True)
class TaperOptimizerSettings:
    """Settings for the tapered-section optimizer.

    ``apply_to_model`` defaults to False: the optimizer rewrites the sections of
    a model the user has already analysed, so the default run is a dry run that
    reports what it would assign and leaves the model untouched.
    """

    deflection: DeflectionLimits
    utilisation_ceiling: float = 0.95
    max_depth_mm: float = 2500.0
    max_flange_width_mm: float = 500.0
    max_web_slenderness: float = 200.0
    # A knee joint is a connection between two different members, not a
    # continuation of one rafter, so by default only collinear members are
    # forced to share a depth at the node they meet.
    tie_depths_at_all_shared_nodes: bool = False
    # Columns are often stocked and spliced as straight sections, so this holds
    # every column run to one depth from base to eave -- the whole run, not
    # each of its members, so a column split at an intermediate node cannot
    # come back stepped.
    prismatic_columns: bool = False
    analysis_budget: int = 40
    apply_to_model: bool = False

    def __post_init__(self) -> None:
        if not 0.1 <= self.utilisation_ceiling <= 1.0:
            raise ValueError("The utilisation ceiling must be between 0.1 and 1.0.")
        if not 300.0 <= self.max_depth_mm <= 5000.0:
            raise ValueError("The maximum web depth must be between 300mm and 5000mm.")
        if not 175.0 <= self.max_flange_width_mm <= 2000.0:
            raise ValueError(
                "The maximum flange width must be between 175mm and 2000mm."
            )
        if not 20.0 <= self.max_web_slenderness <= 500.0:
            raise ValueError("The web slenderness cap must be between 20 and 500.")
        if not 1 <= self.analysis_budget <= 500:
            raise ValueError("The analysis budget must be between 1 and 500 runs.")
