import pytest
from staad_ext.macros.frame_generator import (
    FrameParameters,
    compute_frame_geometry,
    generate_std_file_content,
    rafter_y,
    parse_interior_columns,
)


def test_rafter_y_symmetric():
    # Width 20m, Eave 7m, Ridge 10m, Slope 5 (1:5)
    # Ridge height = 7 + 10/5 = 9.0m
    assert pytest.approx(rafter_y(0.0, 20.0, 7.0, 10.0, 5.0)) == 7.0
    assert pytest.approx(rafter_y(10.0, 20.0, 7.0, 10.0, 5.0)) == 9.0
    assert pytest.approx(rafter_y(20.0, 20.0, 7.0, 10.0, 5.0)) == 7.0
    assert pytest.approx(rafter_y(5.0, 20.0, 7.0, 10.0, 5.0)) == 8.0


def test_parse_interior_columns_count():
    params = FrameParameters(width=20.0, col_mode="count", col_input="3")
    cols = parse_interior_columns(params)
    assert cols == [5.0, 10.0, 15.0]


def test_parse_interior_columns_spacing():
    params = FrameParameters(width=20.0, col_mode="spacing", col_input="4, 6, 5")
    cols = parse_interior_columns(params)
    assert cols == [4.0, 10.0, 15.0]


def test_compute_frame_geometry_simple():
    params = FrameParameters(
        width=20.0,
        eave_height=7.0,
        ridge_distance=10.0,
        slope=5.0,
        col_mode="count",
        col_input="1",  # Column at X=10 (ridge)
        brick_wall_height=0.0,
        mezzanine_enabled=False,
    )
    geom = compute_frame_geometry(params)
    # Check node count: base(0,0), base(20,0), base(10,0), eave(0,7), eave(20,7), ridge(10,9)
    assert len(geom.nodes) == 6
    assert geom.ridge_height == 9.0
    assert geom.interior_x_positions == [10.0]
    # Check member count: left col (1), right col (1), int col (1), left rafter (1), right rafter (1)
    assert len(geom.members) == 5


def test_compute_frame_geometry_with_brick_wall_and_mezzanine():
    params = FrameParameters(
        width=20.0,
        eave_height=7.0,
        ridge_distance=10.0,
        slope=5.0,
        col_mode="spacing",
        col_input="5, 10",  # Cols at X=5, 15
        brick_wall_height=3.0,
        mezzanine_enabled=True,
        mezzanine_height=4.0,
        mezzanine_start_x=0.0,
        mezzanine_end_x=10.0,
    )
    geom = compute_frame_geometry(params)

    # Check brick wall split nodes on outer columns
    left_col_nodes = [n for n in geom.nodes if abs(n.x) < 1e-4]
    y_coords_left = sorted([n.y for n in left_col_nodes])
    assert 3.0 in y_coords_left  # Brick wall node
    assert 4.0 in y_coords_left  # Mezzanine node
    assert 7.0 in y_coords_left  # Eave node

    # Check mezzanine beams exist
    assert len(geom.mezzanine_beams) > 0


def test_generate_std_file_content():
    params = FrameParameters(
        width=20.0,
        eave_height=7.0,
        ridge_distance=10.0,
        slope=5.0,
        dead_load=0.2,
        roof_live_load=0.75,
        bay_spacing=6.0,
        design_code="IS 800:2007",
    )
    std_text = generate_std_file_content(params)
    assert "STAAD PLANE" in std_text
    assert "*** Roof Dead Load Calculation:" in std_text
    assert "JOINT COORDINATES" in std_text
    assert "MEMBER INCIDENCES" in std_text
    assert "LOAD 1 LOADTYPE Seismic-H  TITLE EL1" in std_text
    assert "1893 LOAD X 1" in std_text
    assert "LOAD 5 LOADTYPE Dead  TITLE DL" in std_text
    assert "LOAD 6 LOADTYPE Roof Live  TITLE RL" in std_text
    assert "LOAD 7 LOADTYPE Dead  TITLE CL" in std_text
    assert "CODE IS800 LSD" in std_text
    assert "FYLD 345000 ALL" in std_text
    assert "FU 490000 ALL" in std_text
    assert "RATIO 0.99 ALL" in std_text
    assert "STP 2 ALL" in std_text
    assert "BEAM 1 ALL" in std_text
    # Check converted line load: 0.2 * 6.0 = 1.2 kN/m
    assert "1.200" in std_text


def test_invalid_parameters():
    with pytest.raises(ValueError, match="Brick wall height"):
        FrameParameters(width=20.0, eave_height=7.0, brick_wall_height=8.0).validate()

    with pytest.raises(ValueError, match="Ridge distance"):
        FrameParameters(width=20.0, ridge_distance=25.0).validate()
