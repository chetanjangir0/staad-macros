from __future__ import annotations

from ctypes import byref, c_double, c_long
from pathlib import Path
from typing import Any, Iterable

from staad_ext.models import Point3D

PROG_ID = "StaadPro.OpenSTAAD"


class OpenStaadError(RuntimeError):
    """Raised when STAAD.Pro cannot provide the requested model data."""


def _flag_methods(obj: Any, names: Iterable[str]) -> None:
    """Mark late-bound OpenSTAAD members as callable COM methods."""
    flag = getattr(obj, "_FlagAsMethod", None)
    if flag is not None:
        flag(*names)


class OpenStaad:
    """Small typed facade over the documented STAAD.Pro 2025 OpenSTAAD API."""

    def __init__(self, application: Any) -> None:
        self._application = application
        self.geometry = application.Geometry
        self.property = application.Property
        _flag_methods(application, ("GetSTAADFile",))
        _flag_methods(
            self.geometry,
            ("GetNoOfSelectedBeams", "GetSelectedBeams", "GetMemberIncidence",
             "GetNodeCoordinates", "GetBeamLength"),
        )
        _flag_methods(
            self.property,
            ("GetBeamSectionDisplayName", "GetBeamSectionName",
             "GetBeamPropertyAll", "GetBeamSectionPropertyValuesEx"),
        )

    @classmethod
    def connect(cls) -> OpenStaad:
        try:
            from comtypes.client import GetActiveObject
            return cls(GetActiveObject(PROG_ID, dynamic=True))
        except (ImportError, OSError) as exc:
            raise OpenStaadError(
                "Could not attach to STAAD.Pro. Ensure STAAD.Pro 2025 is running "
                "and that 'comtypes' is installed in a matching Python architecture."
            ) from exc

    def model_path(self) -> Path:
        # STAAD.Pro 2025 documents this as:
        #   void GetSTAADFile(VARIANT &fileName, const VARIANT &bFullPath)
        # Because this is a late-bound COM object, the output VARIANT must be
        # supplied explicitly; omitting it can return the Boolean argument as
        # the call result instead of populating the model filename.
        from comtypes.automation import VARIANT

        values: list[Any] = []

        # With some STAAD.Pro/comtypes registrations, type information exposes
        # the [out] parameter as the Python return value, so no placeholder is
        # accepted. Other registrations require the explicit VARIANT used by
        # the C++ signature. Support both forms.
        for arguments in ((), (True,)):
            try:
                values.append(self._application.GetSTAADFile(*arguments))
            except (OSError, TypeError):
                pass

        file_name = VARIANT()
        try:
            result = self._application.GetSTAADFile(file_name, True)
            values.extend((file_name.value, result))
        except (OSError, TypeError):
            pass

        value = next(
            (
                candidate
                for result in values
                for candidate in self._string_results(result)
                if candidate.strip()
            ),
            "",
        )
        if not value:
            result_types = ", ".join(type(item).__name__ for item in values) or "no results"
            raise OpenStaadError(
                "STAAD.Pro did not return the current model filename "
                f"(GetSTAADFile results: {result_types}). Ensure the model has "
                "been saved, then close duplicate STAAD.Pro instances and retry."
            )
        return Path(str(value))

    @staticmethod
    def _string_results(result: Any) -> list[str]:
        """Extract filename-like strings without treating True as a path."""
        if isinstance(result, str):
            return [result]
        if isinstance(result, (tuple, list)):
            return [item for item in result if isinstance(item, str)]
        value = getattr(result, "value", None)
        return [value] if isinstance(value, str) else []

    def selected_beams(self) -> list[int]:
        from comtypes.safearray import _midlSAFEARRAY

        count = int(self.geometry.GetNoOfSelectedBeams())
        if count <= 0:
            return []
        # Passing byref(SAFEARRAY*) makes comtypes emit
        # VT_BYREF | VT_ARRAY | VT_I4. This is important: passing either a
        # ctypes array or a VARIANT containing a copied SAFEARRAY prevents
        # OpenSTAAD's output values from reaching the Python object.
        numbers = _midlSAFEARRAY(c_long).create([0] * count)
        self.geometry.GetSelectedBeams(byref(numbers), 1)
        return [int(number) for number in numbers.unpack()]

    def member_incidence(self, beam_no: int) -> tuple[int, int]:
        start, end = c_long(), c_long()
        self.geometry.GetMemberIncidence(beam_no, byref(start), byref(end))
        return int(start.value), int(end.value)

    def node_coordinates(self, node_no: int) -> Point3D:
        x, y, z = c_double(), c_double(), c_double()
        self.geometry.GetNodeCoordinates(node_no, byref(x), byref(y), byref(z))
        return Point3D(x.value, y.value, z.value)

    def beam_length(self, beam_no: int) -> float:
        return float(self.geometry.GetBeamLength(beam_no))

    def section_name(self, beam_no: int) -> str:
        try:
            name = str(self.property.GetBeamSectionDisplayName(beam_no) or "")
        except (OSError, TypeError):
            name = ""
        if not name:
            try:
                name = str(self.property.GetBeamSectionName(beam_no) or "")
            except (OSError, TypeError):
                name = "NO SECTION"
        return name or "NO SECTION"

    def beam_property_all(self, beam_no: int) -> tuple[float, ...]:
        values = [c_double() for _ in range(10)]
        self.property.GetBeamPropertyAll(beam_no, *(byref(value) for value in values))
        return tuple(value.value for value in values)

    def section_property_values(self, beam_no: int) -> tuple[int, list[float]]:
        from comtypes.safearray import _midlSAFEARRAY

        property_type = c_long()
        values = _midlSAFEARRAY(c_double).create([0.0] * 24)
        self.property.GetBeamSectionPropertyValuesEx(
            beam_no, byref(property_type), byref(values)
        )
        return int(property_type.value), [float(value) for value in values.unpack()]
