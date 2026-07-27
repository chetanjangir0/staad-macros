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

    def __post_init__(self) -> None:
        if not 0.1 <= self.text_scale <= 10.0:
            raise ValueError("text_scale must be between 0.1 and 10.0")


@dataclass(frozen=True, slots=True)
class SectionEnvelope:
    start_half_width: float
    end_half_width: float
    property_type: int
