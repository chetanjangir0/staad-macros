from __future__ import annotations

from dataclasses import dataclass, field
import json
from pathlib import Path
from typing import Any, Iterable


LOAD_CATEGORIES = ("DL", "LL", "RLL", "CL", "WL", "EQ", "CRANE", "TEMP", "SKIP")


@dataclass
class PrimaryLoadCase:
    id: int
    title: str
    category: str = "DL"  # Must be one of LOAD_CATEGORIES


@dataclass
class ComboFactor:
    category: str  # e.g. "DL", "LL", "WL", "EQ"
    factor: float


@dataclass
class ComboRule:
    name_template: str  # e.g. "1.5(DL + LL)" or "1.2DL + 1.5LL + 1.2WL"
    combo_type: str  # "ULS" or "SLS"
    factors: dict[str, float]  # e.g. {"DL": 1.5, "LL": 1.5, "WL": 1.2}


@dataclass
class CombinationPreset:
    name: str
    description: str
    is_builtin: bool = False
    rules: list[ComboRule] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "description": self.description,
            "is_builtin": self.is_builtin,
            "rules": [
                {
                    "name_template": r.name_template,
                    "combo_type": r.combo_type,
                    "factors": r.factors,
                }
                for r in self.rules
            ],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> CombinationPreset:
        rules = [
            ComboRule(
                name_template=r.get("name_template", ""),
                combo_type=r.get("combo_type", "ULS"),
                factors=r.get("factors", {}),
            )
            for r in data.get("rules", [])
        ]
        return cls(
            name=data.get("name", "Custom"),
            description=data.get("description", ""),
            is_builtin=data.get("is_builtin", False),
            rules=rules,
        )


@dataclass
class GeneratedCombo:
    number: int
    title: str
    combo_type: str  # "ULS" or "SLS"
    factors: list[tuple[int, float]]  # list of (primary_load_case_id, factor)


# Built-in Presets
BUILTIN_PRESETS: list[CombinationPreset] = [
    CombinationPreset(
        name="IS 800:2007 (Limit State)",
        description="Indian Standard IS 800:2007 Table 4 Limit State Design combinations.",
        is_builtin=True,
        rules=[
            # ULS Rules
            ComboRule("1.5(DL + LL)", "ULS", {"DL": 1.5, "CL": 1.5, "LL": 1.5, "RLL": 1.5}),
            ComboRule("1.5(DL + WL)", "ULS", {"DL": 1.5, "CL": 1.5, "WL": 1.5}),
            ComboRule("1.5(DL + EQ)", "ULS", {"DL": 1.5, "CL": 1.5, "EQ": 1.5}),
            ComboRule("1.2(DL + LL + WL)", "ULS", {"DL": 1.2, "CL": 1.2, "LL": 1.2, "RLL": 1.2, "WL": 1.2}),
            ComboRule("1.2(DL + LL + EQ)", "ULS", {"DL": 1.2, "CL": 1.2, "LL": 1.2, "RLL": 1.2, "EQ": 1.2}),
            ComboRule("0.9DL + 1.5WL", "ULS", {"DL": 0.9, "CL": 0.9, "WL": 1.5}),
            ComboRule("0.9DL + 1.5EQ", "ULS", {"DL": 0.9, "CL": 0.9, "EQ": 1.5}),
            ComboRule("1.5DL + 1.5CRANE", "ULS", {"DL": 1.5, "CL": 1.5, "CRANE": 1.5}),
            # SLS Rules
            ComboRule("1.0(DL + LL)", "SLS", {"DL": 1.0, "CL": 1.0, "LL": 1.0, "RLL": 1.0}),
            ComboRule("1.0(DL + WL)", "SLS", {"DL": 1.0, "CL": 1.0, "WL": 1.0}),
            ComboRule("1.0(DL + EQ)", "SLS", {"DL": 1.0, "CL": 1.0, "EQ": 1.0}),
            ComboRule("1.0DL + 0.8LL + 0.8WL", "SLS", {"DL": 1.0, "CL": 1.0, "LL": 0.8, "RLL": 0.8, "WL": 0.8}),
            ComboRule("1.0DL + 0.8LL + 0.8EQ", "SLS", {"DL": 1.0, "CL": 1.0, "LL": 0.8, "RLL": 0.8, "EQ": 0.8}),
        ],
    ),
    CombinationPreset(
        name="ASCE 7-16 / AISC 360-16 (LRFD)",
        description="ASCE 7-16 Strength Design (LRFD) combinations.",
        is_builtin=True,
        rules=[
            ComboRule("1.4DL", "ULS", {"DL": 1.4, "CL": 1.4}),
            ComboRule("1.2DL + 1.6LL + 0.5RLL", "ULS", {"DL": 1.2, "CL": 1.2, "LL": 1.6, "RLL": 0.5}),
            ComboRule("1.2DL + 1.6RLL + 1.0LL", "ULS", {"DL": 1.2, "CL": 1.2, "RLL": 1.6, "LL": 1.0}),
            ComboRule("1.2DL + 1.0WL + 1.0LL + 0.5RLL", "ULS", {"DL": 1.2, "CL": 1.2, "WL": 1.0, "LL": 1.0, "RLL": 0.5}),
            ComboRule("1.2DL + 1.0EQ + 1.0LL", "ULS", {"DL": 1.2, "CL": 1.2, "EQ": 1.0, "LL": 1.0}),
            ComboRule("0.9DL + 1.0WL", "ULS", {"DL": 0.9, "CL": 0.9, "WL": 1.0}),
            ComboRule("0.9DL + 1.0EQ", "ULS", {"DL": 0.9, "CL": 0.9, "EQ": 1.0}),
        ],
    ),
    CombinationPreset(
        name="ASCE 7-16 / AISC 360-16 (ASD)",
        description="ASCE 7-16 Allowable Stress Design (ASD) combinations.",
        is_builtin=True,
        rules=[
            ComboRule("1.0DL", "ULS", {"DL": 1.0, "CL": 1.0}),
            ComboRule("1.0DL + 1.0LL", "ULS", {"DL": 1.0, "CL": 1.0, "LL": 1.0}),
            ComboRule("1.0DL + 1.0RLL", "ULS", {"DL": 1.0, "CL": 1.0, "RLL": 1.0}),
            ComboRule("1.0DL + 0.75LL + 0.75RLL", "ULS", {"DL": 1.0, "CL": 1.0, "LL": 0.75, "RLL": 0.75}),
            ComboRule("1.0DL + 0.6WL", "ULS", {"DL": 1.0, "CL": 1.0, "WL": 0.6}),
            ComboRule("1.0DL + 0.75LL + 0.45WL", "ULS", {"DL": 1.0, "CL": 1.0, "LL": 0.75, "WL": 0.45}),
            ComboRule("1.0DL + 0.7EQ", "ULS", {"DL": 1.0, "CL": 1.0, "EQ": 0.7}),
            ComboRule("0.6DL + 0.6WL", "ULS", {"DL": 0.6, "CL": 0.6, "WL": 0.6}),
            ComboRule("0.6DL + 0.7EQ", "ULS", {"DL": 0.6, "CL": 0.6, "EQ": 0.7}),
        ],
    ),
]


def presets_dir() -> Path:
    p = Path.home() / ".staad_ext" / "presets"
    p.mkdir(parents=True, exist_ok=True)
    return p


def load_all_presets() -> list[CombinationPreset]:
    """Return built-in presets plus any custom user presets saved on disk."""
    presets = list(BUILTIN_PRESETS)
    pdir = presets_dir()
    for f in pdir.glob("*.json"):
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
            preset = CombinationPreset.from_dict(data)
            preset.is_builtin = False
            presets.append(preset)
        except Exception:
            pass
    return presets


def save_custom_preset(preset: CombinationPreset) -> None:
    """Save user custom preset to disk."""
    preset.is_builtin = False
    filename = "".join(c for c in preset.name if c.isalnum() or c in (" ", "_", "-")).strip() or "custom"
    filepath = presets_dir() / f"{filename}.json"
    filepath.write_text(json.dumps(preset.to_dict(), indent=2), encoding="utf-8")


def delete_custom_preset(preset_name: str) -> bool:
    """Delete custom user preset from disk by name."""
    pdir = presets_dir()
    for f in pdir.glob("*.json"):
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
            if data.get("name") == preset_name and not data.get("is_builtin"):
                f.unlink()
                return True
        except Exception:
            pass
    return False


def auto_detect_category(title: str, staad_type: int | None = None) -> str:
    """Infer category from STAAD load type or title keywords."""
    if staad_type is not None:
        if staad_type > 100:
            staad_type = staad_type // 101
        mapping = {0: "DL", 1: "LL", 2: "RLL", 3: "WL", 4: "EQ", 19: "CRANE"}
        if staad_type in mapping:
            return mapping[staad_type]

    t = title.upper()
    if "DEAD" in t or "SELF" in t or "SW" in t or "COLLATERAL" in t:
        return "DL"
    if "ROOF" in t and "LIVE" in t:
        return "RLL"
    if "LIVE" in t or "IMPOSED" in t:
        return "LL"
    if "WIND" in t or "WL" in t or "W+" in t or "W-" in t:
        return "WL"
    if "EQ" in t or "SEISMIC" in t or "EL" in t or "EARTHQUAKE" in t:
        return "EQ"
    if "CRANE" in t:
        return "CRANE"
    if "TEMP" in t or "THERMAL" in t:
        return "TEMP"
    return "DL"


def generate_combinations(
    primary_cases: list[PrimaryLoadCase],
    preset: CombinationPreset,
    aggregate_same_type: bool = False,
    start_uls: int = 101,
    start_sls: int = 201,
) -> list[GeneratedCombo]:
    """Generate list of combination objects based on assigned primary cases and rules."""
    # Filter active non-skip cases
    active_cases = [c for c in primary_cases if c.category != "SKIP"]

    # Group primary cases by category
    by_cat: dict[str, list[PrimaryLoadCase]] = {}
    for c in active_cases:
        by_cat.setdefault(c.category, []).append(c)

    results: list[GeneratedCombo] = []
    curr_uls = start_uls
    curr_sls = start_sls

    for rule in preset.rules:
        # Check which required categories exist in active_cases
        req_cats = [cat for cat, factor in rule.factors.items() if factor != 0]

        # Determine lateral / variable categories in this rule (WL, EQ, CRANE)
        lateral_cats = [cat for cat in req_cats if cat in ("WL", "EQ", "CRANE")]

        if aggregate_same_type or not lateral_cats:
            # Aggregate mode: combine all primary cases of each category together
            factors_list: list[tuple[int, float]] = []
            title_parts: list[str] = []

            has_all_required = True
            for cat, rule_factor in rule.factors.items():
                if cat in by_cat:
                    for plc in by_cat[cat]:
                        factors_list.append((plc.id, rule_factor))
                    title_parts.append(f"{rule_factor:g}{cat}")
                elif cat in ("DL", "LL", "WL", "EQ") and cat in req_cats:
                    # Mandatory category missing, skip rule if core load missing
                    pass

            if factors_list:
                c_num = curr_uls if rule.combo_type == "ULS" else curr_sls
                if rule.combo_type == "ULS":
                    curr_uls += 1
                else:
                    curr_sls += 1

                combo_name = f"{rule.name_template}"
                results.append(
                    GeneratedCombo(
                        number=c_num,
                        title=combo_name,
                        combo_type=rule.combo_type,
                        factors=factors_list,
                    )
                )
        else:
            # Separate mode: Generate distinct combinations for each lateral load case
            # e.g. for rule 1.5DL + 1.5WL, if WL has WL+X (10) and WL-X (11), generate 2 combos!
            # Find primary cases for lateral categories
            lat_combinations: list[list[PrimaryLoadCase]] = []
            lat_cat_names: list[str] = []

            for cat in lateral_cats:
                if cat in by_cat:
                    lat_combinations.append(by_cat[cat])
                    lat_cat_names.append(cat)

            if not lat_combinations:
                # No lateral cases found for this rule, treat standard non-lateral part
                factors_list = []
                for cat, rule_factor in rule.factors.items():
                    if cat not in ("WL", "EQ", "CRANE") and cat in by_cat:
                        for plc in by_cat[cat]:
                            factors_list.append((plc.id, rule_factor))

                if factors_list:
                    c_num = curr_uls if rule.combo_type == "ULS" else curr_sls
                    if rule.combo_type == "ULS":
                        curr_uls += 1
                    else:
                        curr_sls += 1

                    results.append(
                        GeneratedCombo(
                            number=c_num,
                            title=rule.name_template,
                            combo_type=rule.combo_type,
                            factors=factors_list,
                        )
                    )
            else:
                # Cartesian product over lateral combinations
                import itertools

                for lat_tuple in itertools.product(*lat_combinations):
                    factors_list = []
                    name_suffix = []

                    # Non-lateral loads
                    for cat, rule_factor in rule.factors.items():
                        if cat not in ("WL", "EQ", "CRANE") and cat in by_cat:
                            for plc in by_cat[cat]:
                                factors_list.append((plc.id, rule_factor))

                    # Selected lateral loads
                    for plc in lat_tuple:
                        rule_factor = rule.factors.get(plc.category, 1.0)
                        factors_list.append((plc.id, rule_factor))
                        name_suffix.append(f"LC{plc.id}")

                    c_num = curr_uls if rule.combo_type == "ULS" else curr_sls
                    if rule.combo_type == "ULS":
                        curr_uls += 1
                    else:
                        curr_sls += 1

                    combo_name = f"{rule.name_template} ({', '.join(name_suffix)})"
                    results.append(
                        GeneratedCombo(
                            number=c_num,
                            title=combo_name,
                            combo_type=rule.combo_type,
                            factors=factors_list,
                        )
                    )

    return results


def format_staad_combo_text(combos: list[GeneratedCombo]) -> str:
    """Format generated combinations into STAAD.Pro text commands."""
    lines: list[str] = []
    lines.append("*** ------------------------------------------------------------------")
    lines.append("*** LOAD COMBINATIONS GENERATED BY STAAD_EXT")
    lines.append("*** ------------------------------------------------------------------")
    for c in combos:
        lines.append(f"LOAD COMB {c.number} {c.title}")
        factor_items = " ".join(f"{lc_id} {factor:g}" for lc_id, factor in c.factors)
        lines.append(f" {factor_items}")
    return "\n".join(lines)


def fetch_primary_cases_from_openstaad(staad: Any) -> list[PrimaryLoadCase]:
    """Fetch primary load case IDs and titles from active STAAD model."""
    from comtypes.safearray import _midlSAFEARRAY
    from ctypes import c_long

    try:
        count = int(staad.load.GetPrimaryLoadCaseCount())
    except Exception:
        count = 0

    if count <= 0:
        return []

    numbers = _midlSAFEARRAY(c_long).create([0] * count)
    staad.load.GetPrimaryLoadCaseNumbers(numbers)

    cases: list[PrimaryLoadCase] = []
    for nid in numbers.unpack():
        lc_id = int(nid)
        title = ""
        stype = None
        try:
            title = str(staad.load.GetLoadCaseTitle(lc_id) or f"LOAD CASE {lc_id}")
        except Exception:
            title = f"LOAD CASE {lc_id}"

        try:
            stype = int(staad.load.GetLoadType(lc_id))
        except Exception:
            stype = None

        cat = auto_detect_category(title, stype)
        cases.append(PrimaryLoadCase(id=lc_id, title=title, category=cat))

    return cases


def push_combos_to_openstaad(staad: Any, combos: list[GeneratedCombo]) -> int:
    """Push generated combinations into active STAAD.Pro session via OpenSTAAD COM interface."""
    pushed = 0
    for c in combos:
        try:
            staad.load.CreateNewLoadCombination(c.title, c.number)
            for lc_id, factor in c.factors:
                staad.load.AddLoadAndFactorToCombination(c.number, lc_id, float(factor))
            pushed += 1
        except Exception:
            pass
    return pushed
