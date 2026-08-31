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

# STAAD reports a material's yield strength in the model's base force/length**2
# unit: kip/in**2 under the English base unit, kN/m**2 under Metric. Models
# built from an imported material table can come back already in N/mm**2 (or in
# lb/in**2), so each base unit lists its likely factors most-likely-first and
# the first candidate landing in a plausible structural range wins. A blank
# GRADE cell in a schedule is better than one that is out by 1000x.
_STRESS_TO_MPA = {1: (6.894757, 0.006894757), 2: (0.001, 1.0)}
_PLAUSIBLE_MPA = (10.0, 1500.0)

# OpenSTAADUI::AnalyzeEx status codes. 2 (clean) and 3 (warnings only) are the
# two that leave usable results behind; every other code is a run that did not
# deliver any, and each says something different about why.
_ANALYSIS_SUCCEEDED = frozenset((2, 3))
_ANALYSIS_STATUS = {
    -1: "the analysis was terminated",
    0: "the analysis stopped on a general error",
    1: "the analysis is still running",
    4: "the analysis completed with errors",
    5: "the analysis was never performed",
}


def _stress_to_mpa(value: float, base_unit: int) -> float | None:
    if value <= 0:
        return None
    for factor in _STRESS_TO_MPA.get(base_unit, _STRESS_TO_MPA[2]):
        candidate = value * factor
        if _PLAUSIBLE_MPA[0] <= candidate <= _PLAUSIBLE_MPA[1]:
            return candidate
    return None


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
        self._yield_strengths: dict[str, float | None] = {}
        self.geometry = application.Geometry
        self.property = application.Property
        self.support = application.Support
        self.output = application.Output
        self.load = application.Load
        _flag_methods(application, ("GetSTAADFile", "GetBaseUnit", "SetInputUnits",
                                    "AnalyzeEx", "SetSilentMode", "SaveModel"))
        _flag_methods(
            self.geometry,
            ("GetNoOfSelectedBeams", "GetSelectedBeams", "GetMemberIncidence",
             "GetNodeCoordinates", "GetBeamLength", "GetMemberCount", "GetBeamList",
             "AddNode", "AddBeam", "GetNodeCount", "GetNodeList", "DeleteBeam", "DeleteNode"),
        )
        _flag_methods(
            self.property,
            ("GetBeamSectionDisplayName", "GetBeamSectionName",
             "GetBeamPropertyAll", "GetBeamSectionPropertyValuesEx", "GetBetaAngle",
             "GetBeamMaterialName", "GetMaterialPropertyEx",
             "GetBeamSectionPropertyRefNo", "CreateTaperedIProperty",
             "AssignBeamProperty", "AssignMaterialToMember"),
        )
        _flag_methods(
            self.support,
            ("GetSupportCount", "GetSupportNodes",
             "CreateSupportFixed", "CreateSupportPinned", "AssignSupportToNode"),
        )
        _flag_methods(
            self.output,
            ("AreResultsAvailable", "GetSupportReactions", "GetNodeDisplacements",
             "GetMemberSteelDesignRatio"),
        )
        _flag_methods(
            self.load,
            ("GetLoadCombinationCaseCount", "GetLoadCombinationCaseNumbers",
             "GetPrimaryLoadCaseCount", "GetPrimaryLoadCaseNumbers",
             "GetLoadCaseTitle", "GetLoadType"),
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

    def to_base_length(self, value_m: float) -> float:
        """Convert meters back into the model's base length unit.

        Every read on this facade normalizes to meters, so every write has to
        undo that. Writing in the base unit -- rather than calling
        SetInputUnits() first -- keeps the model's own unit settings untouched,
        which matters because changing them would silently rescale what the
        read methods return.
        """
        return value_m / self.length_scale()

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

    def all_nodes(self) -> list[int]:
        """Return every node number in the current model."""
        from comtypes.safearray import _midlSAFEARRAY

        count = int(self.geometry.GetNodeCount())
        if count <= 0:
            return []
        numbers = _midlSAFEARRAY(c_long).create([0] * count)
        self.geometry.GetNodeList(byref(numbers))
        return [int(number) for number in numbers.unpack()]

    def clear_geometry(self) -> None:
        """Delete every beam and node in the current model (beams first, since a
        node cannot be deleted while a member still references it)."""
        for beam_no in self.all_beams():
            self.geometry.DeleteBeam(beam_no)
        for node_no in self.all_nodes():
            self.geometry.DeleteNode(node_no)

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

    def beam_material_name(self, beam_no: int) -> str:
        """Return the material assigned to a beam, or "" when none is readable."""
        try:
            return str(self.property.GetBeamMaterialName(beam_no) or "").strip()
        except (OSError, TypeError, ValueError):
            return ""

    def assign_material_to_beam(self, material_name: str, beam_no: int) -> None:
        """Re-assign a material to one beam.

        Assigning a property drops the member's material assignment, which in a
        model whose materials come from a single CONSTANTS block takes the whole
        block with it. Every property write therefore has to put the material
        back, or the member silently loses its steel design.
        """
        if not material_name:
            return
        try:
            self.property.AssignMaterialToMember(material_name, int(beam_no))
        except (OSError, TypeError, ValueError) as exc:
            raise OpenStaadError(
                f"STAAD.Pro could not re-assign material {material_name!r} to "
                f"member {beam_no}."
            ) from exc

    def material_yield_strength(self, material_name: str) -> float | None:
        """Return a material's yield strength Fy in MPa, or None if unavailable.

        Results are cached per material name: a schedule asks for the same
        handful of materials once per member, and each miss is a COM round-trip.
        """
        if not material_name:
            return None
        if material_name not in self._yield_strengths:
            self._yield_strengths[material_name] = self._read_yield_strength(material_name)
        return self._yield_strengths[material_name]

    def _read_yield_strength(self, material_name: str) -> float | None:
        # STAAD.Pro 2025 documents this as:
        #   OSPropertyUI::GetMaterialPropertyEx(strMaterialName, dElasticity,
        #       dPoisson, dDensity, dAlpha, dDamp, Fy, Fu, Ry, Rt, Fcu)
        # -- ten [out] doubles after the name, with Fy sixth.
        outputs = [c_double() for _ in range(10)]
        try:
            self.property.GetMaterialPropertyEx(
                material_name, *(byref(value) for value in outputs)
            )
        except (OSError, TypeError, ValueError):
            return None
        return _stress_to_mpa(float(outputs[5].value), self.base_unit())

    def beam_property_ref(self, beam_no: int) -> int:
        """Return the section property number assigned to a beam, or 0 if none."""
        try:
            return int(self.property.GetBeamSectionPropertyRefNo(beam_no))
        except (OSError, TypeError, ValueError):
            return 0

    def create_tapered_i_property(self, values_m: Iterable[float]) -> int:
        """Create a tapered I section property and return its property number.

        STAAD.Pro 2025 documents CreateTaperedIProperty as taking one
        1-dimensional array of doubles:
          0 F1 depth at start node      1 F2 web thickness
          2 F3 depth at end node        3 F4 top flange width
          4 F5 top flange thickness     5 F6 bottom flange width
          6 F7 bottom flange thickness
        It returns the new property ID, or 0 when the library could not create
        it (-106/-108 report an unusable array argument).
        """
        from comtypes.safearray import _midlSAFEARRAY

        base = [self.to_base_length(float(value)) for value in values_m]
        if len(base) != 7:
            raise OpenStaadError(
                f"A tapered I property needs 7 dimensions, got {len(base)}."
            )
        # The parameter is [in], so the SAFEARRAY is passed by value rather
        # than byref. Different comtypes/STAAD registrations disagree about
        # whether a bare Python sequence marshals as VT_ARRAY|VT_R8, so fall
        # back to one if the typed SAFEARRAY is rejected.
        results: list[int] = []
        for argument in (_midlSAFEARRAY(c_double).create(base), tuple(base)):
            try:
                result = int(self.property.CreateTaperedIProperty(argument))
            except (OSError, TypeError, ValueError):
                continue
            if result > 0:
                return result
            results.append(result)
        raise OpenStaadError(
            "STAAD.Pro could not create a tapered I section property for "
            f"{[round(value, 6) for value in base]} (returned {results or 'no result'})."
        )

    def assign_beam_property(self, beam_no: int, property_no: int) -> None:
        """Assign an existing section property to one beam.

        AssignBeamProperty takes the beam numbers as an array and returns 0 on
        success; -3006 flags an invalid member and -6001 an invalid property.
        """
        from comtypes.safearray import _midlSAFEARRAY

        beams = _midlSAFEARRAY(c_long).create([int(beam_no)])
        try:
            result = self.property.AssignBeamProperty(beams, int(property_no))
        except (OSError, TypeError, ValueError) as exc:
            raise OpenStaadError(
                f"STAAD.Pro rejected the property assignment for member {beam_no}."
            ) from exc
        if result is not None and int(result) < 0:
            raise OpenStaadError(
                f"STAAD.Pro could not assign property {property_no} to member "
                f"{beam_no} (error {int(result)})."
            )

    def steel_design_ratio(self, beam_no: int) -> float | None:
        """Return a member's critical steel design ratio, or None if undesigned.

        STAAD.Pro documents two sentinels in place of a ratio: -1 when no
        analysis has been run, and -999 when the analysis ran but the member
        was not designed. Only the second is a fact about the member, so only
        it comes back None; the first means the caller's own analysis never
        landed, and a failed call means the query itself is broken. Reporting
        either of those as "not designed" would blame the model for a fault
        that is ours.
        """
        ratio = c_double()
        try:
            succeeded = self.output.GetMemberSteelDesignRatio(beam_no, byref(ratio))
        except (OSError, TypeError, ValueError) as exc:
            raise OpenStaadError(
                "STAAD.Pro rejected the design ratio query for member "
                f"{beam_no} ({exc})."
            ) from exc
        value = float(ratio.value)
        if value == -1.0:
            raise OpenStaadError(
                "STAAD.Pro reports that no analysis has been performed, so it "
                f"has no design ratio for member {beam_no}. The analysis the "
                "optimizer ran did not produce results."
            )
        if not bool(succeeded) or value < 0:
            return None
        return value

    def node_displacements(self, node_no: int, load_case: int) -> tuple[float, ...]:
        """Return global X, Y, Z translations (meters) and rotations (radians)."""
        from comtypes.safearray import _midlSAFEARRAY

        values = _midlSAFEARRAY(c_double).create([0.0] * 6)
        succeeded = self.output.GetNodeDisplacements(node_no, load_case, byref(values))
        if not bool(succeeded):
            raise OpenStaadError(
                f"STAAD.Pro has no displacement result for node {node_no}, "
                f"load case {load_case}."
            )
        unpacked = [float(value) for value in values.unpack()]
        scale = self.length_scale()
        return tuple(value * scale for value in unpacked[:3]) + tuple(unpacked[3:6])

    def set_silent_mode(self, enabled: bool = True) -> None:
        """Suppress the dialogs STAAD.Pro raises during automated edits.

        Property assignment, the analysis and SaveModel all put up dialogs that
        block the COM call until somebody clicks them, which in an unattended
        run reads as a hang. This has to be on before the first write, not
        only before the analysis.
        """
        try:
            self._application.SetSilentMode(1 if enabled else 0)
        except (OSError, TypeError, ValueError):
            pass

    def analyze(self) -> None:
        """Run the analysis silently and block until STAAD.Pro has finished.

        AnalyzeEx writes the in-memory model out to its .STD file and runs the
        engine over that, so the sections assigned since the last run are the
        ones analysed. Nothing needs saving or reloading first -- and reloading
        (UpdateStructure) would destroy them, since it restores the model from
        the file as it stood before those edits.
        """
        self.set_silent_mode(True)
        try:
            # AnalyzeEx(nSilent, nHidden, nWait) -- the third argument is what
            # makes this synchronous, so results are readable when it returns.
            # The arguments are documented as integers, not booleans, so they
            # are passed as 1/0 rather than as Python bools (which marshal to
            # VT_BOOL, where true is -1).
            status = self._application.AnalyzeEx(1, 1, 1)
        except (OSError, TypeError, ValueError) as exc:
            raise OpenStaadError(
                "STAAD.Pro could not run the analysis. Close the analysis window "
                "if one is open, then retry."
            ) from exc
        self._check_analysis_status(status)
        if not self.results_available():
            raise OpenStaadError(
                "STAAD.Pro finished the analysis but reports no results. Open "
                "the model in STAAD.Pro and run the analysis once by hand to "
                "see what it is objecting to."
            )

    @staticmethod
    def _check_analysis_status(status: Any) -> None:
        """Turn AnalyzeEx's status code into a message that names the fault.

        Without this the only symptom of a refused run is that no results
        turned up afterwards, which says nothing about why.
        """
        try:
            code = int(status)
        except (TypeError, ValueError):
            return  # Some registrations return nothing at all; results decide.
        if code in _ANALYSIS_SUCCEEDED:
            return
        raise OpenStaadError(
            "STAAD.Pro did not analyse the model: "
            + _ANALYSIS_STATUS.get(code, f"AnalyzeEx returned status {code}")
            + ". Open the model in STAAD.Pro and run the analysis once by hand "
            "to see what it is objecting to."
        )

    def save_model(self) -> None:
        """Write the in-memory model out to its .STD file.

        The analysis saves the model itself, so during a search the file
        already tracks whatever was last analysed. This is for the end of a
        run, to leave the file agreeing with the model that is left behind.
        """
        self.set_silent_mode(True)
        try:
            self._application.SaveModel(1)
        except (OSError, TypeError, ValueError) as exc:
            raise OpenStaadError(
                "STAAD.Pro could not save the model. The sections in the open "
                "model are correct, but the .STD file still holds the last set "
                "that was analysed -- save it in STAAD.Pro."
            ) from exc

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
