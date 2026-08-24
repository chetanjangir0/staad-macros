from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
import math
import re
from typing import Any

_SOIL_TYPE_CODES = {"Hard": 1, "Medium": 2, "Soft": 3}


def seismic_zone_factor(seismic_zone: str) -> float:
    """Extract the numeric zone factor Z from a label like 'Zone III (0.16)'."""
    match = re.search(r"\(([\d.]+)\)", seismic_zone)
    return float(match.group(1)) if match else 0.16


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
    # "IS 875 Part 3" or "MBMA 12". Defaults to "MBMA 12" (no computed wind
    # load, matching the pre-wind-load placeholder behavior) so constructing
    # FrameParameters() never has the side effect of launching Excel COM.
    wind_standard: str = "MBMA 12"
    wind_building_length: float = 30.0  # m (overall building length, not bay spacing)
    wind_design_life: int = 50  # years
    wind_terrain_category: int = 2
    wind_opening: str = "<5%"  # "<5%", "5-20%", ">20%"
    seismic_zone: str = "Zone III (0.16)"
    response_reduction_factor: float = 5.0
    importance_factor: float = 1.0
    soil_type: str = "Medium"  # "Hard", "Medium", "Soft"
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
        if self.wind_standard == "IS 875 Part 3" and self.wind_building_length <= 0:
            raise ValueError("Wind load building length must be greater than 0.")
        if self.response_reduction_factor <= 0:
            raise ValueError("Response reduction factor (RF) must be greater than 0.")
        if self.importance_factor <= 0:
            raise ValueError("Importance factor (I) must be greater than 0.")
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


def wind_load_member_groups(
    geom: FrameGeometryData, params: FrameParameters
) -> tuple[list[int], list[int], list[int], list[int]]:
    """Return (left_column, left_rafter, right_rafter, right_column) member ids
    for mapping onto the IS 875 Part 3 wind load workbook's 'For STAAD' sheet.

    Assumes the standard single-span gable case the workbook is built for:
    one column line at x=0, one at x=width, and rafter segments split at the
    ridge. Any interior columns/mezzanine beams are not part of the wind path.

    A brick wall segment (base up to brick_wall_height) is excluded from the
    column wind path -- the wall itself carries that portion of the wind
    pressure, not the sheeted column above it.
    """
    node_map = {n.id: n for n in geom.nodes}

    def column_ids(x_target: float) -> list[int]:
        ids = [
            m.id
            for m in geom.members
            if m.group_type == "outer_column"
            and abs(node_map[m.start_node].x - x_target) < 1e-4
            and abs(node_map[m.end_node].x - x_target) < 1e-4
            and max(node_map[m.start_node].y, node_map[m.end_node].y) > params.brick_wall_height + 1e-4
        ]
        return sorted(ids, key=lambda mid: node_map[next(m for m in geom.members if m.id == mid).start_node].y)

    def rafter_ids(before_ridge: bool) -> list[int]:
        ids = []
        for mid in geom.rafter_beams:
            m = next(mm for mm in geom.members if mm.id == mid)
            mid_x = (node_map[m.start_node].x + node_map[m.end_node].x) / 2
            if (mid_x < params.ridge_distance) == before_ridge:
                ids.append(mid)
        return ids

    left_column = column_ids(0.0)
    right_column = column_ids(params.width)
    left_rafter = rafter_ids(True)
    right_rafter = rafter_ids(False)
    return left_column, left_rafter, right_rafter, right_column


def _format_member_ranges(member_ids: list[int]) -> str:
    """Format member ids as STAAD list syntax, collapsing runs of 2+
    consecutive ids into "start TO end" (e.g. [1, 3, 5, 7, 8, 9] -> "1 3 5 7 TO 9")."""
    ids = sorted(set(member_ids))
    parts: list[str] = []
    i = 0
    while i < len(ids):
        j = i
        while j + 1 < len(ids) and ids[j + 1] == ids[j] + 1:
            j += 1
        if j > i:
            parts.append(f"{ids[i]} TO {ids[j]}")
        else:
            parts.append(str(ids[i]))
        i = j + 1
    return " ".join(parts)


def _emit_merged_param_lines(name: str, value_to_ids: dict[float, list[int]]) -> list[str]:
    """One line per distinct value, merging every member id assigned that
    value (within this single component block) into one ranged list."""
    lines = []
    for value in sorted(value_to_ids):
        lines.append(f"{name} {value:g} MEMB {_format_member_ranges(value_to_ids[value])}")
    return lines


def unbraced_length_parameter_lines(geom: FrameGeometryData, params: FrameParameters) -> list[str]:
    """Build the KZ/LX/LY/LZ design-parameter lines for columns and rafters.

    Column KZ follows its own base support (Pinned -> 2, Fixed -> 1.2). LX/LY
    are the sheeting-braced height (max of brick wall height and 1.5 m) for
    outer columns, or the full column length for unbraced interior columns.
    LZ is always the full physical column length (base to top), computed per
    physical column rather than per STAAD member, since a column may be
    split into several members at a brick wall or mezzanine node. Columns
    that land on the same value are reported on a single merged line.

    Rafter LX/LY are fixed at 1.5 m (purlin bracing). Rafter LZ is the
    actual (slope-corrected) length of the rafter between unbraced points:
    for a clear span (no interior columns) that's column-to-ridge on each
    side; otherwise it's column-to-column, following the rafter profile
    rather than the horizontal projection.
    """
    node_map = {n.id: n for n in geom.nodes}

    def member_length(m: Member) -> float:
        n1, n2 = node_map[m.start_node], node_map[m.end_node]
        return math.hypot(n2.x - n1.x, n2.y - n1.y)

    lines: list[str] = []

    # ---- Columns ----
    col_groups: dict[float, list[Member]] = {}
    for m in geom.members:
        if m.group_type in ("outer_column", "interior_column"):
            x = round(node_map[m.start_node].x, 4)
            col_groups.setdefault(x, []).append(m)

    if col_groups:
        kz_map: dict[float, list[int]] = {}
        lx_map: dict[float, list[int]] = {}
        ly_map: dict[float, list[int]] = {}
        lz_map: dict[float, list[int]] = {}
        braced_len = max(params.brick_wall_height, 1.5)
        for x in sorted(col_groups):
            members_at_x = col_groups[x]
            ids = [m.id for m in members_at_x]
            ys = [node_map[m.start_node].y for m in members_at_x] + [
                node_map[m.end_node].y for m in members_at_x
            ]
            full_len = round(max(ys) - min(ys), 4)
            is_outer = abs(x) < 1e-4 or abs(x - params.width) < 1e-4
            if abs(x) < 1e-4:
                support = params.left_support
            elif abs(x - params.width) < 1e-4:
                support = params.right_support
            else:
                support = params.int_support
            kz = 2 if support == "Pinned" else 1.2
            lx_ly = braced_len if is_outer else full_len
            kz_map.setdefault(kz, []).extend(ids)
            lx_map.setdefault(lx_ly, []).extend(ids)
            ly_map.setdefault(lx_ly, []).extend(ids)
            lz_map.setdefault(full_len, []).extend(ids)

        lines.append("******** COLUMNS *******")
        lines.extend(_emit_merged_param_lines("KZ", kz_map))
        lines.extend(_emit_merged_param_lines("LX", lx_map))
        lines.extend(_emit_merged_param_lines("LY", ly_map))
        lines.extend(_emit_merged_param_lines("LZ", lz_map))

    # ---- Rafters ----
    if geom.rafter_beams:
        lines.append("******** RAFTERS *******")
        lines.append(f"LX 1.5 MEMB {_format_member_ranges(geom.rafter_beams)}")
        lines.append(f"LY 1.5 MEMB {_format_member_ranges(geom.rafter_beams)}")

        int_x = geom.interior_x_positions
        if int_x:
            breakpoints = sorted(set([0.0, params.width] + int_x))
        else:
            breakpoints = sorted(set([0.0, params.ridge_distance, params.width]))

        rafter_members = [m for m in geom.members if m.id in geom.rafter_beams]
        seg_ids: dict[int, list[int]] = {}
        for m in rafter_members:
            mid_x = (node_map[m.start_node].x + node_map[m.end_node].x) / 2
            for i in range(len(breakpoints) - 1):
                lo, hi = breakpoints[i], breakpoints[i + 1]
                if lo - 1e-4 <= mid_x <= hi + 1e-4:
                    seg_ids.setdefault(i, []).append(m.id)
                    break

        lz_map: dict[float, list[int]] = {}
        for seg_i, ids in seg_ids.items():
            seg_members = [m for m in rafter_members if m.id in ids]
            seg_len = round(sum(member_length(m) for m in seg_members), 4)
            lz_map.setdefault(seg_len, []).extend(ids)

        lines.extend(_emit_merged_param_lines("LZ", lz_map))

    return lines


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

    # IS 1893 (Part 1) seismic definition and the four static-equivalent
    # seismic load cases (EL1-EL4). Each case is run through its own
    # PERFORM ANALYSIS / CHANGE before the remaining static loads are
    # defined, per the required IS 1893 STAAD input sequence.
    zone_factor = seismic_zone_factor(params.seismic_zone)
    soil_code = _SOIL_TYPE_CODES.get(params.soil_type, 2)
    lines.append("DEFINE 1893 LOAD")
    lines.append(
        f"ZONE {zone_factor:g} RF {params.response_reduction_factor:g} "
        f"I {params.importance_factor:g} SS {soil_code} ST 3"
    )

    # Seismic weight: structure selfweight plus the same dead/collateral
    # loads used in the DL/CL static cases, applied to the same members.
    # Roof live load makes no contribution. Mezzanine live load contributes
    # only a fraction per IS 1893:2016 (25% up to and including 3 kN/m^2,
    # 50% above 3 kN/m^2).
    lines.append("SELFWEIGHT 1.1")
    lines.append("MEMBER WEIGHT")
    if rafter_str:
        lines.append(f"{rafter_str} UNI {w_dl:.3f} 0 0")
        if w_cl > 0:
            lines.append(f"{rafter_str} UNI {w_cl:.3f} 0 0")
    if mezz_str:
        if w_mdl > 0:
            lines.append(f"{mezz_str} UNI {w_mdl:.3f} 0 0")
        if w_mll > 0:
            mll_fraction = 0.25 if params.mezzanine_live_load <= 3.0 else 0.5
            lines.append(f"{mezz_str} UNI {w_mll * mll_fraction:.3f} 0 0")

    seismic_cases = [("X", 1, "EL1"), ("X", -1, "EL2"), ("Z", 1, "EL3"), ("Z", -1, "EL4")]
    for case_num, (direction, sign, title) in enumerate(seismic_cases, start=1):
        lines.append(f"LOAD {case_num} LOADTYPE Seismic-H  TITLE {title}")
        lines.append(f"1893 LOAD {direction} {sign}")
        lines.append("PERFORM ANALYSIS")
        lines.append("CHANGE")

    dl_num = 5
    rl_num = 6
    cl_num = 7
    next_load_num = 8

    lines.append(f"LOAD {dl_num} LOADTYPE Dead  TITLE DL")
    lines.append("*** Structure Selfweight")
    lines.append("SELFWEIGHT Y -1.1")
    lines.append("MEMBER LOAD")
    if rafter_str:
        lines.append("*** Roof Dead Load Calculation:")
        lines.append(f"*** {params.dead_load:.2f} kN/m^2 x {params.bay_spacing:.2f} m (bay spacing) = {w_dl:.3f} kN/m")
        lines.append(f"{rafter_str} UNI GY -{w_dl:.3f}")
    if mezz_str and w_mdl > 0:
        lines.append("*** Mezzanine Dead Load Calculation:")
        lines.append(f"*** {params.mezzanine_dead_load:.2f} kN/m^2 x {params.bay_spacing:.2f} m (bay spacing) = {w_mdl:.3f} kN/m")
        lines.append(f"{mezz_str} UNI GY -{w_mdl:.3f}")

    lines.append(f"LOAD {rl_num} LOADTYPE Roof Live  TITLE RL")
    if rafter_str:
        lines.append("MEMBER LOAD")
        lines.append("*** Roof Live Load Calculation:")
        lines.append(f"*** {params.roof_live_load:.2f} kN/m^2 x {params.bay_spacing:.2f} m (bay spacing) = {w_rll:.3f} kN/m")
        lines.append(f"{rafter_str} UNI GY -{w_rll:.3f}")

    lines.append(f"LOAD {cl_num} LOADTYPE Dead  TITLE CL")
    if rafter_str:
        lines.append("MEMBER LOAD")
        lines.append("*** Collateral Load Calculation:")
        lines.append(f"*** {params.collateral_load:.2f} kN/m^2 x {params.bay_spacing:.2f} m (bay spacing) = {w_cl:.3f} kN/m")
        lines.append(f"{rafter_str} UNI GY -{w_cl:.3f}")

    mezz_num = None
    if params.mezzanine_enabled and mezz_str and w_mll > 0:
        mezz_num = next_load_num
        lines.append(f"LOAD {mezz_num} TITLE MEZZANINE LIVE LOAD")
        lines.append("MEMBER LOAD")
        lines.append("*** Mezzanine Live Load Calculation:")
        lines.append(f"*** {params.mezzanine_live_load:.2f} kN/m^2 x {params.bay_spacing:.2f} m (bay spacing) = {w_mll:.3f} kN/m")
        lines.append(f"{mezz_str} UNI GY -{w_mll:.3f}")
        next_load_num += 1

    if params.wind_standard == "IS 875 Part 3":
        from staad_ext.macros.wind_load import Is875WindParameters, generate_is875_wind_load_lines

        left_col, left_raf, right_raf, right_col = wind_load_member_groups(geom, params)
        wind_params = Is875WindParameters(
            building_length=params.wind_building_length,
            width=params.width,
            height=params.eave_height,
            roof_slope_x=params.slope,
            basic_wind_speed=params.basic_wind_speed,
            design_life=params.wind_design_life,
            terrain_category=params.wind_terrain_category,
            bay_spacing=params.bay_spacing,
            opening=params.wind_opening,
        )
        wind_lines = generate_is875_wind_load_lines(
            wind_params, next_load_num, left_col, left_raf, right_raf, right_col
        )
        lines.extend(wind_lines)
        # The workbook emits one "LOAD n ... TITLE WLx" header per wind load
        # case (pressure/suction x wind direction x gable parallel cases) —
        # count them rather than assuming a fixed number of cases.
        next_load_num += sum(1 for wl in wind_lines if wl.strip().upper().startswith("LOAD "))
    else:
        lines.append(f"LOAD {next_load_num} TITLE WIND LOAD (BASIC SPEED = {params.basic_wind_speed:.1f} M/S)")
        next_load_num += 1

    # Combinations
    lines.append("LOAD COMB 101 1.5(DL + RLL + CL)")
    if mezz_num is not None:
        lines.append(f" {dl_num} 1.5 {rl_num} 1.5 {cl_num} 1.5 {mezz_num} 1.5")
    else:
        lines.append(f" {dl_num} 1.5 {rl_num} 1.5 {cl_num} 1.5")

    lines.append("LOAD COMB 102 1.0(DL + RLL + CL) SERVICE")
    if mezz_num is not None:
        lines.append(f" {dl_num} 1.0 {rl_num} 1.0 {cl_num} 1.0 {mezz_num} 1.0")
    else:
        lines.append(f" {dl_num} 1.0 {rl_num} 1.0 {cl_num} 1.0")

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
        lines.append("CODE IS800 LSD")
        lines.append("FYLD 345000 ALL")
        lines.append("FU 490000 ALL")
        lines.append("RATIO 0.99 ALL")
        lines.append("STP 2 ALL")
        lines.append("BEAM 1 ALL")

    lines.extend(unbraced_length_parameter_lines(geom, params))

    lines.append("CHECK CODE ALL")
    lines.append("SELECT ALL")
    lines.append("PERFORM ANALYSIS")
    lines.append("FINISH")

    return "\n".join(lines)


def build_model_in_openstaad(staad: Any, params: FrameParameters) -> None:
    """Construct nodes, beams, and supports in active STAAD.Pro model via OpenSTAAD COM interface."""
    from staad_ext.openstaad import OpenStaadError

    geom = compute_frame_geometry(params)

    # Units to Meter, KN (4=Meter, 5=KN in OpenSTAAD)
    try:
        staad._application.SetInputUnits(4, 5)
    except Exception:
        pass

    geometry = staad.geometry
    support = staad.support

    try:
        # AddNode/AddBeam return the node/beam number STAAD.Pro actually
        # assigned, which will not match this module's local 1..N node ids
        # unless the active model is empty. Map local ids to the real
        # STAAD node numbers before wiring up beams and supports.
        node_id_map: dict[int, int] = {}
        for n in geom.nodes:
            node_id_map[n.id] = int(geometry.AddNode(n.x, n.y, n.z))

        for m in geom.members:
            geometry.AddBeam(node_id_map[m.start_node], node_id_map[m.end_node])

        s_fixed = support.CreateSupportFixed()
        s_pinned = support.CreateSupportPinned()

        s_left = s_fixed if params.left_support == "Fixed" else s_pinned
        s_right = s_fixed if params.right_support == "Fixed" else s_pinned
        s_int = s_fixed if params.int_support == "Fixed" else s_pinned

        support.AssignSupportToNode(node_id_map[geom.left_support_node], s_left)
        support.AssignSupportToNode(node_id_map[geom.right_support_node], s_right)

        for nid in geom.interior_support_nodes:
            support.AssignSupportToNode(node_id_map[nid], s_int)
    except OpenStaadError:
        raise
    except Exception as exc:
        raise OpenStaadError(
            f"STAAD.Pro rejected the 2D frame model: {exc}"
        ) from exc
