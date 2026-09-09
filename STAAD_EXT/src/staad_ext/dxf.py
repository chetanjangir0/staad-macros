from __future__ import annotations

from contextlib import contextmanager
from io import TextIOBase
from pathlib import Path
from typing import Iterator, Sequence

from staad_ext.models import Point3D

# (layer name, ACI color, linetype) triples written into the DXF LAYER table.
DEFAULT_LAYERS: tuple[tuple[str, int, str], ...] = (
    ("MEMBER_CENTERLINE", 8, "DASHED"), ("MEMBER_SECTION", 3, "CONTINUOUS"),
    ("TAPERED_SECTION", 1, "CONTINUOUS"), ("TUBE_PIPE_SECTION", 5, "CONTINUOUS"),
    ("MEMBER_LABELS", 7, "CONTINUOUS"),
    ("CONNECTION_DETAILS", 2, "CONTINUOUS"),
    ("CONNECTION_PLATES", 1, "CONTINUOUS"),
    ("CONNECTION_CLEATS", 3, "CONTINUOUS"),
    ("CONNECTION_BOLTS", 4, "CONTINUOUS"),
)


class DxfWriter:
    """Minimal AutoCAD R12 ASCII DXF writer used by the exporters.

    STAAD reports geometry in metres, but drawings are dimensioned in
    millimetres, so ``scale`` multiplies every length written out -- the
    coordinates, text heights, radii and linetype patterns alike. Angles are
    left alone.
    """

    def __init__(self, stream: TextIOBase, scale: float = 1.0) -> None:
        self.stream = stream
        self.scale = scale

    def _pairs(self, *items: object) -> None:
        self.stream.write("\n".join(str(item) for item in items) + "\n")

    @staticmethod
    def number(value: float) -> str:
        return f"{value:.6f}"

    def length(self, value: float) -> str:
        """Format a model length in the drawing's units."""
        return self.number(value * self.scale)

    def header(self, layers: Sequence[tuple[str, int, str]] = DEFAULT_LAYERS) -> None:
        self._pairs(0, "SECTION", 2, "HEADER", 9, "$ACADVER", 1, "AC1009",
                    9, "$INSUNITS", 70, 4 if self.scale == 1000.0 else 6,
                    0, "ENDSEC", 0, "SECTION", 2, "TABLES")
        self._pairs(0, "TABLE", 2, "LTYPE", 70, 3)
        self._linetype("CONTINUOUS", "Solid line", ())
        self._linetype("DASHED", "Dashed centerline", (0.50, -0.25))
        self._linetype("HIDDEN", "Hidden inner tube line", (0.25, -0.20))
        self._pairs(0, "ENDTAB")
        self._pairs(0, "TABLE", 2, "LAYER", 70, len(layers))
        for name, color, linetype in layers:
            self._pairs(0, "LAYER", 2, name, 70, 0, 62, color, 6, linetype)
        self._pairs(0, "ENDTAB", 0, "ENDSEC", 0, "SECTION", 2, "ENTITIES")

    def _linetype(self, name: str, description: str,
                  pattern: Sequence[float]) -> None:
        """Write one LTYPE record; ``pattern`` is its dash/gap element lengths."""
        total = sum(abs(element) for element in pattern)
        self._pairs(0, "LTYPE", 2, name, 70, 0, 3, description, 72, 65,
                    73, len(pattern), 40, self.length(total))
        for element in pattern:
            self._pairs(49, self.length(element))

    def line(self, layer: str, start: Point3D, end: Point3D, linetype: str = "CONTINUOUS",
             color: int | None = None) -> None:
        pairs: list[object] = [0, "LINE", 8, layer]
        if color is not None:
            pairs += [62, color]
        pairs += [6, linetype, 10, self.length(start.x), 20,
                  self.length(start.y), 30, self.length(start.z), 11, self.length(end.x),
                  21, self.length(end.y), 31, self.length(end.z)]
        self._pairs(*pairs)

    def circle(self, layer: str, center: Point3D, radius: float,
               color: int | None = None) -> None:
        pairs: list[object] = [0, "CIRCLE", 8, layer]
        if color is not None:
            pairs += [62, color]
        pairs += [10, self.length(center.x), 20, self.length(center.y),
                  30, self.length(center.z), 40, self.length(radius)]
        self._pairs(*pairs)

    def envelope(self, layer: str, outline: Sequence[Point3D], open_start: bool = False,
                 open_end: bool = False, color: int | None = None) -> None:
        """Draw a four-corner member envelope: two long edges plus optional caps.

        ``outline`` is ordered (start-side-A, end-side-A, start-side-B, end-side-B),
        matching what the framing layer solves. The inner faces that give a
        flange or a tube wall its thickness are drawn by the renderers, from
        :func:`staad_ext.framing.inner_face_lines`.
        """
        p1, p2, p3, p4 = outline
        edges = [(p1, p2), (p3, p4)]
        if not open_start:
            edges.append((p1, p3))
        if not open_end:
            edges.append((p2, p4))
        for start, end in edges:
            self.line(layer, start, end, color=color)

    def text(self, layer: str, point: Point3D, height: float, rotation: float,
             value: str, color: int, halign: int = 1) -> None:
        value = value.replace("\r", " ").replace("\n", " ")
        self._pairs(0, "TEXT", 8, layer, 62, color, 10, self.length(point.x), 20,
                    self.length(point.y), 30, self.length(point.z), 40, self.length(height),
                    1, value, 50, self.number(rotation), 72, halign, 11, self.length(point.x),
                    21, self.length(point.y), 31, self.length(point.z))

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


@contextmanager
def dxf_document(
    output: Path, layers: Sequence[tuple[str, int, str]] = DEFAULT_LAYERS,
    scale: float = 1.0,
) -> Iterator[DxfWriter]:
    """Open ``output`` and yield a writer with the header/footer already handled."""
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="ascii", newline="\n") as stream:
        writer = DxfWriter(stream, scale)
        writer.header(layers)
        yield writer
        writer.footer()
