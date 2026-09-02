"""A refused analysis has to be told apart from a rejected model.

Driven back to back, STAAD.Pro declines to start a run at all once it has done
30-odd of them: AnalyzeEx hands back "terminated" having burned no CPU and
written no output, and the very same model then analyses cleanly on the next
call. Reporting that as a fault in the structure sends the engineer hunting
for a problem that is not there, and throws away a search that was minutes
from finishing. These tests pin the two apart.

The facade is exercised without building one, so no STAAD.Pro and no COM
registration are needed: ``analyze`` touches only ``_application`` and
``output``, which the stubs below stand in for.
"""

from __future__ import annotations

import pytest

from staad_ext import openstaad
from staad_ext.openstaad import OpenStaad, OpenStaadError

TERMINATED = -1
STILL_RUNNING = 1
CLEAN = 2
WARNINGS = 3
MODEL_ERRORS = 4


class Application:
    """The handful of OpenSTAAD root methods ``analyze`` calls."""

    def __init__(self, statuses: list[int], analyzing: list[bool] | None = None) -> None:
        self.statuses = list(statuses)
        self.analyzing = list(analyzing or [])
        self.calls = 0
        self.silent: int | None = None

    def AnalyzeEx(self, silent: int, hidden: int, wait: int) -> int:  # noqa: N802
        self.calls += 1
        # The last status repeats, so a test only has to list what changes.
        return self.statuses[min(self.calls - 1, len(self.statuses) - 1)]

    def IsAnalyzing(self) -> bool:  # noqa: N802
        return self.analyzing.pop(0) if self.analyzing else False

    def SetSilentMode(self, flag: int) -> None:  # noqa: N802
        self.silent = flag


class Output:
    def __init__(self, available: bool = True) -> None:
        self.available = available

    def AreResultsAvailable(self) -> bool:  # noqa: N802
        return self.available


def facade(statuses: list[int], analyzing: list[bool] | None = None,
           results: bool = True) -> tuple[OpenStaad, Application]:
    staad = object.__new__(OpenStaad)
    application = Application(statuses, analyzing)
    staad._application = application
    staad.output = Output(results)
    return staad, application


@pytest.fixture(autouse=True)
def slept(monkeypatch: pytest.MonkeyPatch) -> list[float]:
    """Record the waits instead of sitting through them."""
    waits: list[float] = []
    monkeypatch.setattr(openstaad, "sleep", waits.append)
    return waits


def test_a_refused_analysis_is_tried_again(slept: list[float]) -> None:
    # The refusal is STAAD.Pro declining to start, not an objection to the
    # model, so the run that follows it is the answer -- and it goes straight
    # in, because a refusal clears at once and each analysis costs seconds.
    staad, application = facade([TERMINATED, WARNINGS])
    staad.analyze()
    assert application.calls == 2
    assert slept == []


def test_a_run_that_keeps_being_refused_is_reported_as_a_refusal(
        slept: list[float]) -> None:
    staad, application = facade([TERMINATED])
    with pytest.raises(OpenStaadError) as raised:
        staad.analyze()

    assert application.calls == openstaad._ANALYSIS_ATTEMPTS
    # Something slower than the usual refusal is wrong by now, so the later
    # attempts do back off.
    assert slept == [openstaad._ANALYSIS_SETTLE_SECONDS]
    message = str(raised.value)
    # The engine never read the model, so the message must not send the
    # engineer off to hunt for a fault in the structure.
    assert "refused to start" in message
    assert "not what it is objecting to" in message
    assert "run the analysis once by hand" not in message


def test_a_model_the_engine_rejects_is_not_tried_again() -> None:
    # That engine did run and did read the model. Repeating it would only take
    # another few minutes to object again.
    staad, application = facade([MODEL_ERRORS])
    with pytest.raises(OpenStaadError) as raised:
        staad.analyze()

    assert application.calls == 1
    assert "completed with errors" in str(raised.value)
    assert "run the analysis once by hand" in str(raised.value)


def test_a_clean_run_is_not_repeated() -> None:
    for status in (CLEAN, WARNINGS):
        staad, application = facade([status])
        staad.analyze()
        assert application.calls == 1


def test_a_retry_waits_for_the_running_analysis_to_finish() -> None:
    # Launching a second analysis on top of a running one is what gets
    # refused, so the retry has to wait the first one out rather than race it.
    staad, application = facade([STILL_RUNNING, WARNINGS],
                                analyzing=[True, True, False])
    staad.analyze()
    assert application.calls == 2
    assert application.analyzing == []      # the wait ran until STAAD went idle


def test_silent_mode_is_on_before_the_analysis() -> None:
    staad, application = facade([WARNINGS])
    staad.analyze()
    assert application.silent == 1


def test_a_run_that_leaves_no_results_is_reported_even_when_the_status_is_clean() -> None:
    staad, _application = facade([WARNINGS], results=False)
    with pytest.raises(OpenStaadError, match="no results"):
        staad.analyze()
