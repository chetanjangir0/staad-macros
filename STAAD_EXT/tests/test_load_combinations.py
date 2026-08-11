import pytest
from staad_ext.macros.load_combinations import (
    CombinationPreset,
    ComboRule,
    PrimaryLoadCase,
    auto_detect_category,
    delete_custom_preset,
    format_staad_combo_text,
    generate_combinations,
    load_all_presets,
    save_custom_preset,
)


def test_auto_detect_category():
    assert auto_detect_category("SELFWEIGHT AND DEAD LOAD") == "DL"
    assert auto_detect_category("ROOF LIVE LOAD") == "RLL"
    assert auto_detect_category("LIVE LOAD 1") == "LL"
    assert auto_detect_category("WIND LOAD +X") == "WL"
    assert auto_detect_category("EARTHQUAKE -Z") == "EQ"
    assert auto_detect_category("CRANE HOOK LOAD") == "CRANE"


def test_generate_combinations_aggregate():
    primary_cases = [
        PrimaryLoadCase(id=1, title="DEAD LOAD", category="DL"),
        PrimaryLoadCase(id=2, title="COLLATERAL LOAD", category="DL"),
        PrimaryLoadCase(id=3, title="LIVE LOAD", category="LL"),
        PrimaryLoadCase(id=4, title="WIND +X", category="WL"),
        PrimaryLoadCase(id=5, title="WIND -X", category="WL"),
    ]

    preset = load_all_presets()[0]  # IS 800:2007 (Limit State)
    combos = generate_combinations(
        primary_cases=primary_cases,
        preset=preset,
        aggregate_same_type=True,
        start_uls=101,
        start_sls=201,
    )

    assert len(combos) > 0

    # In aggregate mode, rule 1.5(DL + LL) should combine DL1 and DL2, plus LL3
    c101 = combos[0]
    assert c101.number == 101
    # Check that both DL cases (1 and 2) are present with factor 1.5
    factors_dict = dict(c101.factors)
    assert factors_dict[1] == 1.5
    assert factors_dict[2] == 1.5
    assert factors_dict[3] == 1.5


def test_generate_combinations_separate():
    primary_cases = [
        PrimaryLoadCase(id=1, title="DEAD LOAD", category="DL"),
        PrimaryLoadCase(id=2, title="LIVE LOAD", category="LL"),
        PrimaryLoadCase(id=4, title="WIND +X", category="WL"),
        PrimaryLoadCase(id=5, title="WIND -X", category="WL"),
    ]

    preset = load_all_presets()[0]  # IS 800:2007 (Limit State)
    combos = generate_combinations(
        primary_cases=primary_cases,
        preset=preset,
        aggregate_same_type=False,
        start_uls=101,
        start_sls=201,
    )

    # For 1.5(DL + WL), separate mode should produce 2 combos (one for WL4, one for WL5)
    wl_combos = [c for c in combos if "1.5(DL + WL)" in c.title]
    assert len(wl_combos) == 2
    assert "LC4" in wl_combos[0].title or "LC5" in wl_combos[0].title


def test_format_staad_combo_text():
    primary_cases = [
        PrimaryLoadCase(id=1, title="DL", category="DL"),
        PrimaryLoadCase(id=2, title="LL", category="LL"),
    ]
    preset = CombinationPreset(
        name="Test",
        description="Test",
        rules=[ComboRule("1.5(DL + LL)", "ULS", {"DL": 1.5, "LL": 1.5})],
    )
    combos = generate_combinations(primary_cases, preset)
    text = format_staad_combo_text(combos)
    assert "LOAD COMB 101" in text
    assert "1 1.5 2 1.5" in text


def test_preset_save_and_delete():
    custom = CombinationPreset(
        name="Unit Test Preset 123",
        description="Custom Test Preset",
        rules=[ComboRule("1.4DL", "ULS", {"DL": 1.4})],
    )
    save_custom_preset(custom)

    all_presets = load_all_presets()
    names = [p.name for p in all_presets]
    assert "Unit Test Preset 123" in names

    success = delete_custom_preset("Unit Test Preset 123")
    assert success is True

    all_presets_after = load_all_presets()
    names_after = [p.name for p in all_presets_after]
    assert "Unit Test Preset 123" not in names_after
