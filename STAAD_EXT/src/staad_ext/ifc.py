from __future__ import annotations

import time
import uuid
from io import TextIOBase

_GUID_CHARS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$"


def new_guid() -> str:
    """Return a random, syntactically valid IfcGloballyUniqueId (22 chars)."""
    bits = int.from_bytes(uuid.uuid4().bytes, "big") << 4  # 128 -> 132 bits, 22*6
    return "".join(_GUID_CHARS[(bits >> shift) & 0x3F] for shift in range(126, -1, -6))


def _escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("'", "\\'")


def s(text: str) -> str:
    """Format a STEP string literal."""
    return f"'{_escape(text)}'"


def num(value: float) -> str:
    formatted = f"{float(value):.6f}".rstrip("0")
    return formatted + "0" if formatted.endswith(".") else formatted


def ref(entity_id: int | None) -> str:
    return f"#{entity_id}" if entity_id else "$"


class IfcWriter:
    """Minimal IFC4 (STEP/SPF) writer used by the 3D structure exporter."""

    def __init__(self, stream: TextIOBase) -> None:
        self.stream = stream
        self._next_id = 1
        self._lines: list[str] = []
        self.owner_history_id = 0
        self.context_id = 0
        self.storey_id = 0

    def entity(self, name: str, *args: str) -> int:
        entity_id = self._next_id
        self._next_id += 1
        self._lines.append(f"#{entity_id}={name}({','.join(args)});")
        return entity_id

    # -- generic geometry helpers -------------------------------------------------
    def point(self, x: float, y: float, z: float = 0.0) -> int:
        return self.entity("IFCCARTESIANPOINT", f"({num(x)},{num(y)},{num(z)})")

    def direction(self, x: float, y: float, z: float | None = None) -> int:
        coords = f"({num(x)},{num(y)})" if z is None else f"({num(x)},{num(y)},{num(z)})"
        return self.entity("IFCDIRECTION", coords)

    def axis2placement3d(self, location: int, axis: int | None = None,
                         ref_direction: int | None = None) -> int:
        return self.entity("IFCAXIS2PLACEMENT3D", ref(location), ref(axis), ref(ref_direction))

    def axis2placement2d(self, location: int, ref_direction: int | None = None) -> int:
        return self.entity("IFCAXIS2PLACEMENT2D", ref(location), ref(ref_direction))

    def local_placement(self, relative_to: int | None, placement: int) -> int:
        return self.entity("IFCLOCALPLACEMENT", ref(relative_to), ref(placement))

    # -- profiles / solids ----------------------------------------------------
    def rectangle_profile(self, name: str, x_dim: float, y_dim: float) -> int:
        origin = self.axis2placement2d(self.point(0.0, 0.0))
        return self.entity("IFCRECTANGLEPROFILEDEF", ".AREA.", s(name), ref(origin),
                            num(x_dim), num(y_dim))

    def circle_hollow_profile(self, name: str, radius: float, wall_thickness: float) -> int:
        origin = self.axis2placement2d(self.point(0.0, 0.0))
        return self.entity("IFCCIRCLEHOLLOWPROFILEDEF", ".AREA.", s(name), ref(origin),
                            num(radius), num(wall_thickness))

    def extruded_area_solid(self, profile: int, depth: float) -> int:
        position = self.axis2placement3d(self.point(0.0, 0.0, 0.0), self.direction(0.0, 0.0, 1.0),
                                          self.direction(1.0, 0.0, 0.0))
        extrusion_direction = self.direction(0.0, 0.0, 1.0)
        return self.entity("IFCEXTRUDEDAREASOLID", ref(profile), ref(position),
                            ref(extrusion_direction), num(depth))

    def shape_representation(self, item: int) -> int:
        return self.entity("IFCSHAPEREPRESENTATION", ref(self.context_id), s("Body"),
                            s("SweptSolid"), f"({ref(item)})")

    def product_definition_shape(self, representation: int) -> int:
        return self.entity("IFCPRODUCTDEFINITIONSHAPE", "$", "$", f"({ref(representation)})")

    # -- structural elements ----------------------------------------------------
    def member(self, is_column: bool, name: str, placement: int, shape: int) -> int:
        kind = "IFCCOLUMN" if is_column else "IFCBEAM"
        return self.entity(kind, s(new_guid()), ref(self.owner_history_id), s(name), "$", "$",
                            ref(placement), ref(shape), "$", "$")

    def contained_in_storey(self, related: list[int]) -> None:
        if not related:
            return
        items = ",".join(ref(item) for item in related)
        self.entity("IFCRELCONTAINEDINSPATIALSTRUCTURE", s(new_guid()), ref(self.owner_history_id),
                    "$", "$", f"({items})", ref(self.storey_id))

    # -- header / footer ----------------------------------------------------
    def header(self, project_name: str, author: str = "STAAD_EXT") -> None:
        timestamp = time.strftime("%Y-%m-%dT%H:%M:%S")
        person = self.entity("IFCPERSON", "$", "$", s(author), "$", "$", "$", "$", "$")
        organization = self.entity("IFCORGANIZATION", "$", s("STAAD_EXT"), "$", "$", "$")
        person_and_org = self.entity("IFCPERSONANDORGANIZATION", ref(person), ref(organization), "$")
        application = self.entity(
            "IFCAPPLICATION", ref(organization), s("1.0"), s("STAAD_EXT"), s("STAAD_EXT")
        )
        self.owner_history_id = self.entity(
            "IFCOWNERHISTORY", ref(person_and_org), ref(application), "$", ".ADDED.",
            "$", ref(person_and_org), ref(application), str(int(time.time())),
        )

        length_unit = self.entity("IFCSIUNIT", "$", ".LENGTHUNIT.", "$", ".METRE.")
        units = self.entity("IFCUNITASSIGNMENT", f"({ref(length_unit)})")

        world_origin = self.point(0.0, 0.0, 0.0)
        world_placement = self.axis2placement3d(world_origin)
        self.context_id = self.entity(
            "IFCGEOMETRICREPRESENTATIONCONTEXT", "$", s("Model"), "3", "1.0E-5",
            ref(world_placement), "$",
        )

        project = self.entity(
            "IFCPROJECT", s(new_guid()), ref(self.owner_history_id), s(project_name), "$", "$",
            "$", "$", f"({ref(self.context_id)})", ref(units),
        )

        site_placement = self.local_placement(None, self.axis2placement3d(self.point(0.0, 0.0, 0.0)))
        site = self.entity(
            "IFCSITE", s(new_guid()), ref(self.owner_history_id), s("Site"), "$", "$",
            ref(site_placement), "$", "$", ".ELEMENT.", "$", "$", "$", "$", "$",
        )
        building_placement = self.local_placement(site_placement, self.axis2placement3d(self.point(0.0, 0.0, 0.0)))
        building = self.entity(
            "IFCBUILDING", s(new_guid()), ref(self.owner_history_id), s("Building"), "$", "$",
            ref(building_placement), "$", "$", ".ELEMENT.", "$", "$", "$",
        )
        storey_placement = self.local_placement(building_placement, self.axis2placement3d(self.point(0.0, 0.0, 0.0)))
        self.storey_id = self.entity(
            "IFCBUILDINGSTOREY", s(new_guid()), ref(self.owner_history_id), s("Storey"), "$", "$",
            ref(storey_placement), "$", "$", ".ELEMENT.", "0.0",
        )

        self.entity("IFCRELAGGREGATES", s(new_guid()), ref(self.owner_history_id), "$", "$",
                    ref(project), f"({ref(site)})")
        self.entity("IFCRELAGGREGATES", s(new_guid()), ref(self.owner_history_id), "$", "$",
                    ref(site), f"({ref(building)})")
        self.entity("IFCRELAGGREGATES", s(new_guid()), ref(self.owner_history_id), "$", "$",
                    ref(building), f"({ref(self.storey_id)})")

        self._header_timestamp = timestamp

    def write(self, filename: str) -> None:
        self.stream.write("ISO-10303-21;\n")
        self.stream.write("HEADER;\n")
        self.stream.write(
            "FILE_DESCRIPTION((''),'2;1');\n"
        )
        self.stream.write(
            f"FILE_NAME({s(filename)},{s(self._header_timestamp)},{s('')},{s('')},"
            f"{s('STAAD_EXT')},{s('STAAD_EXT')},{s('')});\n"
        )
        self.stream.write("FILE_SCHEMA(('IFC4'));\n")
        self.stream.write("ENDSEC;\n")
        self.stream.write("DATA;\n")
        for line in self._lines:
            self.stream.write(line + "\n")
        self.stream.write("ENDSEC;\n")
        self.stream.write("END-ISO-10303-21;\n")
