from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
import math
from typing import Any


@dataclass
class FrameParameters:
    width: float = 20.0  # m
    eave_height: float = 7.0  # m
    ridge_distance: float = 10.0  # m (from left column)
    slope: float = 5.0  # 1:x slope (e.g. 5 means 1 in 5 slope)
    col_mode: str = "count"  # "count" or "spacing"
    col_input: str = "0"  # count or comma-separated spacing
    brick_wall_height: float = 0.0  # m
    mezzanine_enabled: bool = False
    mezzanine_height: float = 3.5  # m
    mezzanine_start_x: float = 0.0  # m
    mezzanine_end_x: float = 10.0  # m
    bay_spacing: float = 6.0  # m
    left_support: str = "Fixed"  # "Fixed" or "Pinned"
    right_support: str = "Fixed"  # "Fixed" or "Pinned"
    int_support: str = "Fixed"  # "Fixed" or "Pinned"
    basic_wind_speed: float = 39.0  # m/s
    seismic_zone: str = "Zone III (0.16)"
    dead_load: float = 0.15  # kN/m2
    roof_live_load: float = 0.75  # kN/m2
    collateral_load: float = 0.10  # kN/m2
    mezzanine_live_load: float = 3.0  # kN/m2
    mezzanine_dead_load: float = 1.5  # kN/m2
    design_code: str = "IS 800:2007"  # "IS 800:2007", "AISC 360-16 (LRFD)", "AISC 360-16 (ASD)"

    def validate(self) -> None:
        if self.width <= 0:
            raise ValueError("Frame width must be greater than 0.")
        if self.eave_height <= 0:
            raise ValueError("Eave height must be greater than 0.")
        if self.ridge_distance <= 0 or self.ridge_distance >= self.width:
            raise ValueError(
                f"Ridge distance ({self.ridge_distance} m) must be between 0 and width ({self.width} m)."
            )
        if self.slope <= 0:
            raise ValueError("Roof slope (1:x) must be greater than 0.")
        if self.brick_wall_height < 0:
            raise ValueError("Brick wall height cannot be negative.")
        if self.brick_wall_height >= self.eave_height:
            raise ValueError(
                f"Brick wall height ({self.brick_wall_height} m) must be less than eave height ({self.eave_height} m)."
            )
        if self.bay_spacing <= 0:
            raise ValueError("Bay spacing must be greater than 0.")
        if self.mezzanine_enabled:
            if self.mezzanine_height <= 0 or self.mezzanine_height >= self.eave_height:
                raise ValueError(
                    f"Mezzanine height ({self.mezzanine_height} m) must be between 0 and eave height ({self.eave_height} m)."
                )
            if self.mezzanine_start_x < 0 or self.mezzanine_end_x > self.width or self.mezzanine_start_x >= self.mezzanine_end_x:
                raise ValueError(
                    f"Invalid mezzanine span (Start: {self.mezzanine_start_x} m, End: {self.mezzanine_end_x} m) for frame width {self.width} m."
                )


@dataclass
class Node:
    id: int
    x: float
    y: float
    z: float = 0.0
    tag: str = "standard"  # "base", "eave", "ridge", "brick_wall", "mezzanine"


@dataclass
class Member:
    id: int
    start_node: int
    end_node: int
    group_type: str  # "outer_column", "rafter", "interior_column", "mezzanine"


@dataclass
class FrameGeometryData:
    nodes: list[Node]
    members: list[Member]
    interior_x_positions: list[float]
    ridge_height: float
    rafter_beams: list[int]
    mezzanine_beams: list[int]
    left_support_node: int
    right_support_node: int
    interior_support_nodes: list[int]


def rafter_y(x: float, width: float, eave_height: float, ridge_distance: float, slope: float) -> float:
    """Calculate Y coordinate along rafter profile at distance x from left column."""
    ridge_h = eave_height + (ridge_distance / slope)
    if x <= ridge_distance:
        return eave_height + (x / slope)
    else:
        return ridge_h - ((x - ridge_distance) / slope)


def parse_interior_columns(params: FrameParameters) -> list[float]:
    """Return list of X coordinates for interior columns."""
    if params.col_mode == "count":
        try:
            count = int(abs(float(params.col_input)))
        except ValueError:
            count = 0
        if count <= 0:
            return []
        step = params.width / (count + 1)
        return [round((i + 1) * step, 4) for i in range(count)]
    else:  # spacing
        raw = params.col_input.strip()
        if not raw or raw == "0":
            return []
        parts = [p.strip() for p in raw.split(",") if p.strip()]
        x_list: list[float] = []
        cum_x = 0.0
        for p in parts:
            try:
                val = abs(float(p))
                cum_x += val
                if cum_x >= params.width - 1e-4:
                    raise ValueError(
                        f"Cumulative spacing ({cum_x:.2f} m) reaches or exceeds frame width ({params.width} m)."
                    )
                x_list.append(round(cum_x, 4))
            except ValueError as exc:
                if "reaches or exceeds" in str(exc):
                    raise
                pass
        return x_list


def compute_frame_geometry(params: FrameParameters) -> FrameGeometryData:
    """Calculate exact node coordinates, member connectivity, and element classifications."""
    params.validate()

    int_x = parse_interior_columns(params)
    ridge_h = params.eave_height + (params.ridge_distance / params.slope)

    # Collect key X coordinates
    key_x_coords = sorted(list(set([0.0, params.width, params.ridge_distance] + int_x)))

    nodes: list[Node] = []
    node_lookup: dict[tuple[float, float], int] = {}
    next_node_id = 1

    def get_or_add_node(x: float, y: float, tag: str = "standard") -> int:
        nonlocal next_node_id
        key = (round(x, 4), round(y, 4))
        if key in node_lookup:
            return node_lookup[key]
        n = Node(id=next_node_id, x=round(x, 4), y=round(y, 4), z=0.0, tag=tag)
        nodes.append(n)
        node_lookup[key] = next_node_id
        next_node_id += 1
        return n.id

    # 1. Base Nodes
    left_base = get_or_add_node(0.0, 0.0, "base")
    right_base = get_or_add_node(params.width, 0.0, "base")
    int_base_nodes: list[int] = []
    for x in int_x:
        int_base_nodes.append(get_or_add_node(x, 0.0, "base"))

    # 2. Key Height Nodes at each X location
    has_bw = params.brick_wall_height > 0.001
    has_mezz = params.mezzanine_enabled

    if has_mezz:
        if params.mezzanine_start_x not in key_x_coords:
            key_x_coords.append(params.mezzanine_start_x)
        if params.mezzanine_end_x not in key_x_coords:
            key_x_coords.append(params.mezzanine_end_x)
        key_x_coords.sort()

    for x in key_x_coords:
        y_top = rafter_y(x, params.width, params.eave_height, params.ridge_distance, params.slope)
        get_or_add_node(x, y_top, "ridge" if abs(x - params.ridge_distance) < 1e-4 else "rafter_top")

        if has_bw and (abs(x) < 1e-4 or abs(x - params.width) < 1e-4):
            get_or_add_node(x, params.brick_wall_height, "brick_wall")

        if has_mezz and (params.mezzanine_start_x - 1e-4 <= x <= params.mezzanine_end_x + 1e-4):
            if params.mezzanine_height < y_top:
                get_or_add_node(x, params.mezzanine_height, "mezzanine")

    members: list[Member] = []
    next_member_id = 1

    def add_beam(n1: int, n2: int, group: str) -> Member:
        nonlocal next_member_id
        m = Member(id=next_member_id, start_node=n1, end_node=n2, group_type=group)
        members.append(m)
        next_member_id += 1
        return m

    # 3. Outer Left Column Beams (X = 0)
    col_left_nodes = sorted(
        [n for n in nodes if abs(n.x) < 1e-4 and n.y <= params.eave_height + 1e-4],
        key=lambda n: n.y,
    )
    for i in range(len(col_left_nodes) - 1):
        add_beam(col_left_nodes[i].id, col_left_nodes[i + 1].id, "outer_column")

    # 4. Outer Right Column Beams (X = width)
    y_right_eave = rafter_y(params.width, params.width, params.eave_height, params.ridge_distance, params.slope)
    col_right_nodes = sorted(
        [n for n in nodes if abs(n.x - params.width) < 1e-4 and n.y <= y_right_eave + 1e-4],
        key=lambda n: n.y,
    )
    for i in range(len(col_right_nodes) - 1):
        add_beam(col_right_nodes[i].id, col_right_nodes[i + 1].id, "outer_column")

    # 5. Interior Column Beams (at each X in int_x)
    for x in int_x:
        y_top = rafter_y(x, params.width, params.eave_height, params.ridge_distance, params.slope)
        int_nodes = sorted(
            [n for n in nodes if abs(n.x - x) < 1e-4 and n.y <= y_top + 1e-4],
            key=lambda n: n.y,
        )
        for i in range(len(int_nodes) - 1):
            add_beam(int_nodes[i].id, int_nodes[i + 1].id, "interior_column")

    # 6. Additional vertical posts at mezzanine bounds if start/end X is not an outer/interior column
    if has_mezz:
        for x_m in (params.mezzanine_start_x, params.mezzanine_end_x):
            if abs(x_m) > 1e-4 and abs(x_m - params.width) > 1e-4 and x_m not in int_x:
                y_top = rafter_y(x_m, params.width, params.eave_height, params.ridge_distance, params.slope)
                get_or_add_node(x_m, 0.0, "base")
                posts = sorted(
                    [n for n in nodes if abs(n.x - x_m) < 1e-4 and n.y <= y_top + 1e-4],
                    key=lambda n: n.y,
                )
                for i in range(len(posts) - 1):
                    add_beam(posts[i].id, posts[i + 1].id, "interior_column")

    # 7. Rafter Beams (Along top nodes at each key_x_coords)
    rafter_nodes: list[Node] = []
    for x in key_x_coords:
        y_top = rafter_y(x, params.width, params.eave_height, params.ridge_distance, params.slope)
        matching = [n for n in nodes if abs(n.x - x) < 1e-4 and abs(n.y - y_top) < 1e-4]
        if matching:
            rafter_nodes.append(matching[0])
    rafter_nodes.sort(key=lambda n: n.x)

    rafter_beams: list[int] = []
    for i in range(len(rafter_nodes) - 1):
        m = add_beam(rafter_nodes[i].id, rafter_nodes[i + 1].id, "rafter")
        rafter_beams.append(m.id)

    # 8. Mezzanine Beams (Along Y = mezzanine_height from start_x to end_x)
    mezzanine_beams: list[int] = []
    if has_mezz:
        mezz_nodes = sorted(
            [
                n
                for n in nodes
                if abs(n.y - params.mezzanine_height) < 1e-4
                and params.mezzanine_start_x - 1e-4 <= n.x <= params.mezzanine_end_x + 1e-4
            ],
            key=lambda n: n.x,
        )
        for i in range(len(mezz_nodes) - 1):
            m = add_beam(mezz_nodes[i].id, mezz_nodes[i + 1].id, "mezzanine")
            mezzanine_beams.append(m.id)

    return FrameGeometryData(
        nodes=nodes,
        members=members,
        interior_x_positions=int_x,
        ridge_height=ridge_h,
        rafter_beams=rafter_beams,
        mezzanine_beams=mezzanine_beams,
        left_support_node=left_base,
        right_support_node=right_base,
        interior_support_nodes=int_base_nodes,
    )


def generate_std_file_content(params: FrameParameters) -> str:
    """Generate complete .STD file text content for STAAD.Pro."""
    geom = compute_frame_geometry(params)

    w_dl = params.dead_load * params.bay_spacing
    w_rll = params.roof_live_load * params.bay_spacing
    w_cl = params.collateral_load * params.bay_spacing
    w_mll = params.mezzanine_live_load * params.bay_spacing if params.mezzanine_enabled else 0.0
    w_mdl = params.mezzanine_dead_load * params.bay_spacing if params.mezzanine_enabled else 0.0

    lines: list[str] = []
    lines.append("STAAD PLANE")
    lines.append("*** ------------------------------------------------------------------")
    lines.append(f"*** 2D PORTAL / GABLED FRAME MODEL GENERATED BY STAAD_EXT")
    lines.append(f"*** Frame Width: {params.width:.2f} m | Eave Height: {params.eave_height:.2f} m | Roof Slope: 1:{params.slope}")
    lines.append(f"*** Bay Spacing: {params.bay_spacing:.2f} m | Design Code: {params.design_code}")
    lines.append("*** ------------------------------------------------------------------")
    lines.append("START JOB INFORMATION")
    lines.append(f"ENGINEER DATE {date.today().strftime('%d-%b-%y')}")
    lines.append("END JOB INFORMATION")
    lines.append("INPUT WIDTH 79")
    lines.append("UNIT METER KN")
    lines.append("JOINT COORDINATES")

    for n in geom.nodes:
        lines.append(f"{n.id} {n.x:.4f} {n.y:.4f} {n.z:.4f}")

    lines.append("MEMBER INCIDENCES")
    for m in geom.members:
        lines.append(f"{m.id} {m.start_node} {m.end_node}")

    lines.append("DEFINE MATERIAL START")
    lines.append("ISOTROPIC STEEL")
    lines.append("E 2.05e+08")
    lines.append("POISSON 0.3")
    lines.append("DENSITY 76.8195")
    lines.append("ALPHA 1.2e-05")
    lines.append("DAMP 0.03")
    lines.append("G 7.88462e+07")
    lines.append("TYPE STEEL")
    lines.append("STRENGTH FY 250000 FU 400000 RY 1.5 RT 1.2")
    lines.append("END DEFINE MATERIAL")

    lines.append("MEMBER PROPERTY INDIAN")
    lines.append(f"1 TO {len(geom.members)} TABLE ST ISMB350")
    lines.append("CONSTANTS")
    lines.append(" MATERIAL STEEL ALL")

    lines.append("SUPPORTS")
    supp_left = "FIXED" if params.left_support == "Fixed" else "PINNED"
    supp_right = "FIXED" if params.right_support == "Fixed" else "PINNED"
    supp_int = "FIXED" if params.int_support == "Fixed" else "PINNED"

    lines.append(f"{geom.left_support_node} {supp_left}")
    lines.append(f"{geom.right_support_node} {supp_right}")
    if geom.interior_support_nodes:
        int_nodes_str = " ".join(str(nid) for nid in geom.interior_support_nodes)
        lines.append(f"{int_nodes_str} {supp_int}")

    # Load Cases
    rafter_str = " ".join(str(bid) for bid in geom.rafter_beams)
    mezz_str = " ".join(str(bid) for bid in geom.mezzanine_beams) if geom.mezzanine_beams else ""

    lines.append("LOAD 1 TITLE DEAD LOAD")
    lines.append("*** Structure Selfweight")
    lines.append("SELFWEIGHT Y -1")
    lines.append("MEMBER LOAD")
    if rafter_str:
        lines.append("*** Roof Dead Load Calculation:")
        lines.append(f"*** {params.dead_load:.2f} kN/m^2 x {params.bay_spacing:.2f} m (bay spacing) = {w_dl:.3f} kN/m")
        lines.append(f"{rafter_str} UNI GY -{w_dl:.3f}")
    if mezz_str and w_mdl > 0:
        lines.append("*** Mezzanine Dead Load Calculation:")
        lines.append(f"*** {params.mezzanine_dead_load:.2f} kN/m^2 x {params.bay_spacing:.2f} m (bay spacing) = {w_mdl:.3f} kN/m")
        lines.append(f"{mezz_str} UNI GY -{w_mdl:.3f}")

    lines.append("LOAD 2 TITLE ROOF LIVE LOAD")
    if rafter_str:
        lines.append("MEMBER LOAD")
        lines.append("*** Roof Live Load Calculation:")
        lines.append(f"*** {params.roof_live_load:.2f} kN/m^2 x {params.bay_spacing:.2f} m (bay spacing) = {w_rll:.3f} kN/m")
        lines.append(f"{rafter_str} UNI GY -{w_rll:.3f}")

    lines.append("LOAD 3 TITLE COLLATERAL LOAD")
    if rafter_str:
        lines.append("MEMBER LOAD")
        lines.append("*** Collateral Load Calculation:")
        lines.append(f"*** {params.collateral_load:.2f} kN/m^2 x {params.bay_spacing:.2f} m (bay spacing) = {w_cl:.3f} kN/m")
        lines.append(f"{rafter_str} UNI GY -{w_cl:.3f}")

    if params.mezzanine_enabled and mezz_str and w_mll > 0:
        lines.append("LOAD 4 TITLE MEZZANINE LIVE LOAD")
        lines.append("MEMBER LOAD")
        lines.append("*** Mezzanine Live Load Calculation:")
        lines.append(f"*** {params.mezzanine_live_load:.2f} kN/m^2 x {params.bay_spacing:.2f} m (bay spacing) = {w_mll:.3f} kN/m")
        lines.append(f"{mezz_str} UNI GY -{w_mll:.3f}")

    lines.append(f"LOAD 5 TITLE WIND LOAD (BASIC SPEED = {params.basic_wind_speed:.1f} M/S)")
    lines.append(f"LOAD 6 TITLE SEISMIC LOAD (ZONE = {params.seismic_zone})")

    # Combinations
    lines.append("LOAD COMB 101 1.5(DL + RLL + CL)")
    if params.mezzanine_enabled and mezz_str:
        lines.append(" 1 1.5 2 1.5 3 1.5 4 1.5")
    else:
        lines.append(" 1 1.5 2 1.5 3 1.5")

    lines.append("LOAD COMB 102 1.0(DL + RLL + CL) SERVICE")
    if params.mezzanine_enabled and mezz_str:
        lines.append(" 1 1.0 2 1.0 3 1.0 4 1.0")
    else:
        lines.append(" 1 1.0 2 1.0 3 1.0")

    lines.append("PERFORM ANALYSIS")

    # Design Parameters
    lines.append("PARAMETER 1")
    if "AISC" in params.design_code:
        lines.append("CODE AISC UNIFIED")
        if "LRFD" in params.design_code:
            lines.append("METHOD LRFD")
        else:
            lines.append("METHOD ASD")
    else:
        lines.append("CODE INDIAN")

    lines.append("CHECK CODE ALL")
    lines.append("SELECT ALL")
    lines.append("PERFORM ANALYSIS")
    lines.append("FINISH")

    return "\n".join(lines)


def build_model_in_openstaad(staad: Any, params: FrameParameters) -> None:
    """Construct nodes, beams, and supports in active STAAD.Pro model via OpenSTAAD COM interface."""
    geom = compute_frame_geometry(params)

    # Units to Meter, KN (4=Meter, 5=KN in OpenSTAAD)
    try:
        staad._application.SetInputUnits(4, 5)
    except Exception:
        pass

    geometry = staad.geometry
    support = staad.support

    # Add Nodes
    for n in geom.nodes:
        try:
            geometry.AddNode(n.x, n.y, n.z)
        except Exception:
            pass

    # Add Beams
    for m in geom.members:
        try:
            geometry.AddBeam(m.start_node, m.end_node)
        except Exception:
            pass

    # Supports
    try:
        s_fixed = support.CreateSupportFixed()
        s_pinned = support.CreateSupportPinned()

        s_left = s_fixed if params.left_support == "Fixed" else s_pinned
        s_right = s_fixed if params.right_support == "Fixed" else s_pinned
        s_int = s_fixed if params.int_support == "Fixed" else s_pinned

        support.AssignSupportToNode(geom.left_support_node, s_left)
        support.AssignSupportToNode(geom.right_support_node, s_right)

        for nid in geom.interior_support_nodes:
            support.AssignSupportToNode(nid, s_int)
    except Exception:
        pass
