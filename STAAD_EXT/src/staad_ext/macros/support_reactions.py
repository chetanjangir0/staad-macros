from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Iterable

from staad_ext.openstaad import OpenStaad, OpenStaadError

COMPONENTS = ("FX", "FY", "FZ", "MX", "MY", "MZ")


@dataclass(frozen=True, slots=True)
class SupportReaction:
    node: int
    load_case: int
    fx: float
    fy: float
    fz: float
    mx: float
    my: float
    mz: float

    @property
    def values(self) -> tuple[float, float, float, float, float, float]:
        return self.fx, self.fy, self.fz, self.mx, self.my, self.mz


@dataclass(frozen=True, slots=True)
class ReactionEnvelope:
    component: str
    minimum: float
    minimum_node: int
    minimum_load_case: int
    maximum: float
    maximum_node: int
    maximum_load_case: int


def parse_load_cases(text: str) -> list[int]:
    """Parse comma/space-separated case IDs and inclusive ranges such as 101-105."""
    tokens = [token for token in re.split(r"[\s,;]+", text.strip()) if token]
    if not tokens:
        raise ValueError("Enter at least one load combination number.")
    cases: list[int] = []
    for token in tokens:
        if re.fullmatch(r"\d+", token):
            cases.append(int(token))
            continue
        match = re.fullmatch(r"(\d+)\s*-\s*(\d+)", token)
        if not match:
            raise ValueError(
                f"Invalid load combination entry '{token}'. Use numbers "
                "separated by commas, or ranges such as 101-105."
            )
        start, end = (int(value) for value in match.groups())
        if end < start:
            raise ValueError(f"Load combination range '{token}' is reversed.")
        cases.extend(range(start, end + 1))
    if any(case <= 0 for case in cases):
        raise ValueError("Load combination numbers must be greater than zero.")
    return list(dict.fromkeys(cases))


def reaction_envelopes(
    reactions: Iterable[SupportReaction],
) -> list[ReactionEnvelope]:
    rows = list(reactions)
    if not rows:
        return []
    envelopes = []
    for index, component in enumerate(COMPONENTS):
        minimum = min(rows, key=lambda row: row.values[index])
        maximum = max(rows, key=lambda row: row.values[index])
        envelopes.append(ReactionEnvelope(
            component,
            minimum.values[index], minimum.node, minimum.load_case,
            maximum.values[index], maximum.node, maximum.load_case,
        ))
    return envelopes


def get_support_reactions(
    staad: OpenStaad, load_cases: Iterable[int]
) -> list[SupportReaction]:
    """Read global reactions at every support for the requested result cases."""
    requested = list(load_cases)
    available = set(staad.load_combination_cases())
    missing = [case for case in requested if case not in available]
    if missing:
        missing_text = ", ".join(str(case) for case in missing)
        raise OpenStaadError(
            f"Load combination(s) not found in the active model: {missing_text}."
        )
    if not staad.results_available():
        raise OpenStaadError(
            "Analysis results are not available. Run the STAAD.Pro analysis first."
        )
    nodes = staad.support_nodes()
    if not nodes:
        raise OpenStaadError("The active STAAD.Pro model has no support nodes.")
    rows = []
    for load_case in requested:
        for node in nodes:
            rows.append(SupportReaction(
                node, load_case, *staad.support_reactions(node, load_case)
            ))
    return rows
