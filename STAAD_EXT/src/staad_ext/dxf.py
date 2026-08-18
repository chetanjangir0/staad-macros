from __future__ import annotations

from io import TextIOBase
from staad_ext.models import Point3D


class DxfWriter:
    """Minimal AutoCAD R12 ASCII DXF writer used by the exporters."""

    def __init__(self, stream: TextIOBase) -> None:
        self.stream = stream

    def _pairs(self, *items: object) -> None:
        self.stream.write("\n".join(str(item) for item in items) + "\n")

    @staticmethod
    def number(value: float) -> str:
        return f"{value:.6f}"

    def header(self) -> None:
        self._pairs(0, "SECTION", 2, "HEADER", 9, "$ACADVER", 1, "AC1009", 0, "ENDSEC",
                    0, "SECTION", 2, "TABLES")
        self._pairs(0, "TABLE", 2, "LTYPE", 70, 3)
        self._linetype("CONTINUOUS", "Solid line", 0, 0.0)
        self._linetype("DASHED", "Dashed centerline", 2, 0.75)
        self._pairs(49, "0.50", 49, "-0.25")
        self._linetype("HIDDEN", "Hidden inner tube line", 2, 0.45)
        self._pairs(49, "0.25", 49, "-0.20", 0, "ENDTAB")
        layers = (("MEMBER_CENTERLINE", 8, "DASHED"), ("MEMBER_SECTION", 3, "CONTINUOUS"),
                  ("TAPERED_SECTION", 1, "CONTINUOUS"), ("TUBE_PIPE_SECTION", 5, "CONTINUOUS"),
                  ("MEMBER_LABELS", 7, "CONTINUOUS"),
                  ("CONNECTION_DETAILS", 2, "CONTINUOUS"),
                  ("CONNECTION_PLATES", 1, "CONTINUOUS"),
                  ("CONNECTION_CLEATS", 3, "CONTINUOUS"),
                  ("CONNECTION_BOLTS", 4, "CONTINUOUS"))
        self._pairs(0, "TABLE", 2, "LAYER", 70, len(layers))
        for name, color, linetype in layers:
            self._pairs(0, "LAYER", 2, name, 70, 0, 62, color, 6, linetype)
        self._pairs(0, "ENDTAB", 0, "ENDSEC", 0, "SECTION", 2, "ENTITIES")

    def _linetype(self, name: str, description: str, count: int, length: float) -> None:
        self._pairs(0, "LTYPE", 2, name, 70, 0, 3, description, 72, 65, 73, count, 40, self.number(length))

    def line(self, layer: str, start: Point3D, end: Point3D, linetype: str = "CONTINUOUS",
             color: int | None = None) -> None:
        pairs: list[object] = [0, "LINE", 8, layer]
        if color is not None:
            pairs += [62, color]
        pairs += [6, linetype, 10, self.number(start.x), 20,
                  self.number(start.y), 30, self.number(start.z), 11, self.number(end.x),
                  21, self.number(end.y), 31, self.number(end.z)]
        self._pairs(*pairs)


    def text(self, layer: str, point: Point3D, height: float, rotation: float,
             value: str, color: int) -> None:
        value = value.replace("\r", " ").replace("\n", " ")
        self._pairs(0, "TEXT", 8, layer, 62, color, 10, self.number(point.x), 20,
                    self.number(point.y), 30, self.number(point.z), 40, self.number(height),
                    1, value, 50, self.number(rotation), 72, 1, 11, self.number(point.x),
                    21, self.number(point.y), 31, self.number(point.z))

    def colored_label(self, layer: str, point: Point3D, height: float, rotation: float,
                      value: str, offset: Point3D, color: int = 7) -> None:
        lines = value.split(r"\P")
        for index, label_line in enumerate(lines):
            amount = ((len(lines) - 1) / 2 - index) * height * 1.25
            location = Point3D(point.x + offset.x * amount, point.y + offset.y * amount,
                               point.z + offset.z * amount)
            self._colored_label_line(layer, location, height, rotation, label_line, color)

    def _colored_label_line(self, layer: str, point: Point3D, height: float,
                            rotation: float, value: str, color: int) -> None:
        self.text(layer, point, height, rotation, value, color)

    def footer(self) -> None:
        self._pairs(0, "ENDSEC", 0, "EOF")
