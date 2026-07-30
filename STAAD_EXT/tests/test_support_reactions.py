import pytest

from staad_ext.macros.support_reactions import (
    get_support_reactions,
    parse_load_cases,
    reaction_envelopes,
)


class FakeStaad:
    def load_combination_cases(self) -> list[int]:
        return [2, 4]

    def results_available(self) -> bool:
        return True

    def support_nodes(self) -> list[int]:
        return [1, 3]

    def support_reactions(self, node: int, load_case: int) -> tuple[float, ...]:
        value = node * load_case
        return value, -value, value + 1, 0.0, value / 2, -value / 2


def test_load_case_input_accepts_lists_ranges_and_removes_duplicates() -> None:
    assert parse_load_cases("101, 103-105 101; 110") == [
        101, 103, 104, 105, 110
    ]


@pytest.mark.parametrize("text", ["", "1-A", "5-3", "0"])
def test_invalid_load_case_input_is_rejected(text: str) -> None:
    with pytest.raises(ValueError):
        parse_load_cases(text)


def test_reactions_and_min_max_envelopes_include_governing_node_and_case() -> None:
    rows = get_support_reactions(FakeStaad(), [2, 4])  # type: ignore[arg-type]
    assert len(rows) == 4
    envelopes = {row.component: row for row in reaction_envelopes(rows)}
    assert envelopes["FX"].minimum == 2
    assert envelopes["FX"].minimum_node == 1
    assert envelopes["FX"].minimum_load_case == 2
    assert envelopes["FX"].maximum == 12
    assert envelopes["FX"].maximum_node == 3
    assert envelopes["FX"].maximum_load_case == 4
    assert envelopes["FY"].minimum == -12

def test_missing_combination_is_rejected_before_reaction_calls() -> None:
    class TrackingStaad(FakeStaad):
        reaction_called = False

        def support_reactions(self, node: int, load_case: int) -> tuple[float, ...]:
            self.reaction_called = True
            return super().support_reactions(node, load_case)

    staad = TrackingStaad()
    with pytest.raises(RuntimeError, match="3"):
        get_support_reactions(staad, [2, 3])  # type: ignore[arg-type]
    assert not staad.reaction_called
