from __future__ import annotations

from ctypes import byref, c_double, c_long
from pathlib import Path
from typing import Any, Iterable

from staad_ext.models import Point3D

PROG_ID = "StaadPro.OpenSTAAD"

# STAAD.Pro OpenSTAADUI::GetBaseUnit() returns 1 for the English base unit
# system (lengths in inches) or 2 for Metric (lengths in meters). Every
# geometry/property value this facade returns is normalized to meters using
# this factor so callers never have to special-case the model's base unit.
_INCHES_TO_METERS = 0.0254


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
        self._length_scale: float | None = None
        self.geometry = application.Geometry
        self.property = application.Property
        self.support = application.Support
        self.output = application.Output
        self.load = application.Load
        _flag_methods(application, ("GetSTAADFile", "GetBaseUnit"))
        _flag_methods(
            self.geometry,
            ("GetNoOfSelectedBeams", "GetSelectedBeams", "GetMemberIncidence",
             "GetNodeCoordinates", "GetBeamLength", "GetMemberCount", "GetBeamList"),
        )
        _flag_methods(
            self.property,
            ("GetBeamSectionDisplayName", "GetBeamSectionName",
             "GetBeamPropertyAll", "GetBeamSectionPropertyValuesEx", "GetBetaAngle"),
        )
        _flag_methods(self.support, ("GetSupportCount", "GetSupportNodes"))
        _flag_methods(
            self.output, ("AreResultsAvailable", "GetSupportReactions")
        )
        _flag_methods(
            self.load,
            ("GetLoadCombinationCaseCount", "GetLoadCombinationCaseNumbers"),
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

    def base_unit(self) -> int:
        """Return the current .STD file's base unit: 1 = English, 2 = Metric."""
        return int(self._application.GetBaseUnit())

    def length_scale(self) -> float:
        """Return the factor that converts the model's base length unit to meters."""
        if self._length_scale is None:
            self._length_scale = _INCHES_TO_METERS if self.base_unit() == 1 else 1.0
        return self._length_scale

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

    def all_beams(self) -> list[int]:
        """Return every analytical beam member number in the current model."""
        from comtypes.safearray import _midlSAFEARRAY

        count = int(self.geometry.GetMemberCount())
        if count <= 0:
            return []
        numbers = _midlSAFEARRAY(c_long).create([0] * count)
        self.geometry.GetBeamList(byref(numbers))
        return [int(number) for number in numbers.unpack()]

    def beta_angle(self, beam_no: int) -> float:
        """Return the beta angle (degrees) assigned to a beam, or 0.0 if unavailable."""
        try:
            return float(self.property.GetBetaAngle(beam_no))
        except (OSError, TypeError, ValueError):
            return 0.0

    def member_incidence(self, beam_no: int) -> tuple[int, int]:
        start, end = c_long(), c_long()
        self.geometry.GetMemberIncidence(beam_no, byref(start), byref(end))
        return int(start.value), int(end.value)

    def node_coordinates(self, node_no: int) -> Point3D:
        x, y, z = c_double(), c_double(), c_double()
        self.geometry.GetNodeCoordinates(node_no, byref(x), byref(y), byref(z))
        scale = self.length_scale()
        return Point3D(x.value * scale, y.value * scale, z.value * scale)

    def beam_length(self, beam_no: int) -> float:
        return float(self.geometry.GetBeamLength(beam_no)) * self.length_scale()

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
        # Order: Width, Depth, Ax (area), Ay (area), Az (area), Ix, Iy, Iz
        # (all length**4), Tf, Tw -- see OSPropertyUI::GetBeamPropertyAll.
        values = [c_double() for _ in range(10)]
        self.property.GetBeamPropertyAll(beam_no, *(byref(value) for value in values))
        scale = self.length_scale()
        exponents = (1, 1, 2, 2, 2, 4, 4, 4, 1, 1)
        return tuple(value.value * scale ** exponent
                     for value, exponent in zip(values, exponents))

    def section_property_values(self, beam_no: int) -> tuple[int, list[float]]:
        from comtypes.safearray import _midlSAFEARRAY

        property_type = c_long()
        values = _midlSAFEARRAY(c_double).create([0.0] * 24)
        self.property.GetBeamSectionPropertyValuesEx(
            beam_no, byref(property_type), byref(values)
        )
        scale = self.length_scale()
        return int(property_type.value), [float(value) * scale for value in values.unpack()]

    def support_nodes(self) -> list[int]:
        """Return every supported node in the current model."""
        from comtypes.safearray import _midlSAFEARRAY

        count = int(self.support.GetSupportCount())
        if count <= 0:
            return []
        numbers = _midlSAFEARRAY(c_long).create([0] * count)
        result = self.support.GetSupportNodes(byref(numbers))
        if int(result) < 0:
            raise OpenStaadError("STAAD.Pro could not read the support nodes.")
        return sorted(int(number) for number in numbers.unpack())

    def load_combination_cases(self) -> list[int]:
        """Return all load-combination IDs defined in the current model."""
        from comtypes.safearray import _midlSAFEARRAY
        count = int(self.load.GetLoadCombinationCaseCount())
        if count <= 0:
            return []
        cases = _midlSAFEARRAY(c_long).create([0] * count)
        result = self.load.GetLoadCombinationCaseNumbers(byref(cases))
        if int(result) < 0:
            raise OpenStaadError(
                "STAAD.Pro could not read the load combination numbers."
            )
        return sorted(int(case) for case in cases.unpack())

    def results_available(self) -> bool:
        return bool(self.output.AreResultsAvailable())

    def support_reactions(
        self, node_no: int, load_case: int
    ) -> tuple[float, float, float, float, float, float]:
        """Return global FX, FY, FZ, MX, MY and MZ for a support node."""
        from comtypes.safearray import _midlSAFEARRAY

        values = _midlSAFEARRAY(c_double).create([0.0] * 6)
        succeeded = self.output.GetSupportReactions(
            node_no, load_case, byref(values)
        )
        if not bool(succeeded):
            raise OpenStaadError(
                f"STAAD.Pro has no support reaction result for node {node_no}, "
                f"load case {load_case}."
            )
        unpacked = tuple(float(value) for value in values.unpack())
        return unpacked  # type: ignore[return-value]
