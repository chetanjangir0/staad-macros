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

    def __post_init__(self) -> None:
        if not 0.1 <= self.text_scale <= 10.0:
            raise ValueError("text_scale must be between 0.1 and 10.0")
        if not 0 <= self.blank_rows <= 40:
            raise ValueError("blank_rows must be between 0 and 40")


@dataclass(frozen=True, slots=True)
class IfcExportSettings:
    selected_only: bool = False
