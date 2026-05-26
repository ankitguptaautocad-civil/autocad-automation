from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

import dxf_col_rectangles_to_excel as base


DEFAULT_DXF_PATH = base.DEFAULT_DXF_PATH
DEFAULT_COLUMN_LAYER = base.DEFAULT_COLUMN_LAYER
DEFAULT_COLUMN_HATCH_LAYER = base.DEFAULT_COLUMN_HATCH_LAYER
DEFAULT_TARGET_RECT_COUNT = base.DEFAULT_TARGET_RECT_COUNT
DEFAULT_PREFERRED_BLOCK_NAME = base.DEFAULT_PREFERRED_BLOCK_NAME
DEFAULT_MM_PER_UNIT = base.DEFAULT_MM_PER_UNIT
DEFAULT_ROW_TOLERANCE_MM = base.DEFAULT_ROW_TOLERANCE_MM
DEFAULT_INCH_MULTIPLE_TOLERANCE = base.DEFAULT_INCH_MULTIPLE_TOLERANCE
DEFAULT_HATCH_OVERLAP_TOLERANCE = base.DEFAULT_HATCH_OVERLAP_TOLERANCE
DEFAULT_X_GRID_TOLERANCE_M = 0.12
DEFAULT_Y_GRID_TOLERANCE_M = 0.10
DEFAULT_X_CONSENSUS_TOLERANCE_M = 0.36
DEFAULT_Y_CONSENSUS_TOLERANCE_M = 0.30


@dataclass(frozen=True)
class OrderedRect:
    idx: int
    rect: base.Rect
    xmin_m: float
    xmax_m: float
    ymin_m: float
    ymax_m: float

    @property
    def width_m(self) -> float:
        return self.xmax_m - self.xmin_m

    @property
    def height_m(self) -> float:
        return self.ymax_m - self.ymin_m

    @property
    def cx_m(self) -> float:
        return (self.xmin_m + self.xmax_m) / 2.0

    @property
    def cy_m(self) -> float:
        return (self.ymin_m + self.ymax_m) / 2.0


@dataclass(frozen=True)
class Face:
    rect_idx: int
    side: str
    value_m: float


@dataclass(frozen=True)
class Family:
    mean_m: float
    spread_m: float
    members: tuple[Face, ...]

    @property
    def unique_rect_count(self) -> int:
        return len({member.rect_idx for member in self.members})


def build_local_rects(ordered: list[base.Rect], anchor: base.Rect, mm_per_unit: float) -> list[OrderedRect]:
    scale = mm_per_unit / 1000.0
    local_rects: list[OrderedRect] = []
    for idx, rect in enumerate(ordered, start=1):
        local_rects.append(
            OrderedRect(
                idx=idx,
                rect=rect,
                xmin_m=round((rect.xmin - anchor.xmin) * scale, 3),
                xmax_m=round((rect.xmax - anchor.xmin) * scale, 3),
                ymin_m=round((rect.ymin - anchor.ymin) * scale, 3),
                ymax_m=round((rect.ymax - anchor.ymin) * scale, 3),
            )
        )
    return local_rects


def cluster_numeric(values: list[float], tolerance_m: float) -> list[list[float]]:
    sorted_values = sorted(values)
    groups: list[list[float]] = [[sorted_values[0]]]
    for value in sorted_values[1:]:
        if value - groups[-1][-1] <= tolerance_m:
            groups[-1].append(value)
        else:
            groups.append([value])
    return groups


def build_families(faces: list[Face], tolerance_m: float) -> tuple[list[Family], dict[tuple[int, str], int]]:
    sorted_faces = sorted(faces, key=lambda face: face.value_m)
    grouped_faces: list[list[Face]] = [[sorted_faces[0]]]
    for face in sorted_faces[1:]:
        if face.value_m - grouped_faces[-1][-1].value_m <= tolerance_m:
            grouped_faces[-1].append(face)
        else:
            grouped_faces.append([face])

    families: list[Family] = []
    lookup: dict[tuple[int, str], int] = {}
    for family_idx, grouped in enumerate(grouped_faces):
        values = [face.value_m for face in grouped]
        family = Family(
            mean_m=sum(values) / len(values),
            spread_m=max(values) - min(values),
            members=tuple(grouped),
        )
        families.append(family)
        for member in grouped:
            lookup[(member.rect_idx, member.side)] = family_idx
    return families, lookup


def family_score(family: Family, face_value_m: float) -> tuple[int, float, float]:
    return (
        family.unique_rect_count,
        -family.spread_m,
        -abs(face_value_m - family.mean_m),
    )


def consensus_score(face_value_m: float, other_anchor_values: list[float], tolerance_m: float) -> tuple[int, float, float]:
    deltas = [abs(face_value_m - value) for value in other_anchor_values if abs(face_value_m - value) <= tolerance_m]
    if not deltas:
        return (0, float("-inf"), float("-inf"))
    return (
        len(deltas),
        -(sum(deltas) / len(deltas)),
        -min(deltas),
    )


def choose_x_tag(
    rect: OrderedRect,
    x_families: list[Family],
    x_lookup: dict[tuple[int, str], int],
    global_min_x: float,
    global_max_x: float,
) -> str | None:
    left_family_idx = x_lookup[(rect.idx, "Left")]
    right_family_idx = x_lookup[(rect.idx, "Right")]
    left_family = x_families[left_family_idx]
    right_family = x_families[right_family_idx]

    if rect.xmin_m == global_min_x:
        return "Left"
    if rect.xmax_m == global_max_x:
        return "Right"

    left_score = family_score(left_family, rect.xmin_m)
    right_score = family_score(right_family, rect.xmax_m)
    if max(left_score[0], right_score[0]) < 2:
        return None
    if left_score > right_score:
        return "Left"
    if right_score > left_score:
        return "Right"
    return None


def choose_y_tag(
    rect: OrderedRect,
    y_families: list[Family],
    y_lookup: dict[tuple[int, str], int],
    global_min_y: float,
    global_max_y: float,
) -> str | None:
    if rect.ymin_m == global_min_y:
        return "Front"
    if rect.ymax_m == global_max_y:
        return "Back"

    front_family = y_families[y_lookup[(rect.idx, "Front")]]
    back_family = y_families[y_lookup[(rect.idx, "Back")]]
    front_score = family_score(front_family, rect.ymin_m)
    back_score = family_score(back_family, rect.ymax_m)

    if max(front_score[0], back_score[0]) < 2:
        return None
    if front_score > back_score:
        return "Front"
    if back_score > front_score:
        return "Back"
    return "Front"


def chosen_anchor_value(rect: OrderedRect, tag: str | None, axis: str) -> float | None:
    if axis == "X":
        if tag == "Left":
            return rect.xmin_m
        if tag == "Right":
            return rect.xmax_m
    else:
        if tag == "Front":
            return rect.ymin_m
        if tag == "Back":
            return rect.ymax_m
    return None


def refine_x_tag_with_consensus(
    rect: OrderedRect,
    provisional_tag: str | None,
    provisional_x_anchors: dict[int, float],
    consensus_tolerance_m: float,
    global_min_x: float,
    global_max_x: float,
) -> str | None:
    if rect.xmin_m == global_min_x:
        return "Left"
    if rect.xmax_m == global_max_x:
        return "Right"

    other_values = [value for rect_idx, value in provisional_x_anchors.items() if rect_idx != rect.idx]
    if not other_values:
        return provisional_tag

    left_score = consensus_score(rect.xmin_m, other_values, consensus_tolerance_m)
    right_score = consensus_score(rect.xmax_m, other_values, consensus_tolerance_m)
    if max(left_score[0], right_score[0]) == 0:
        return provisional_tag
    if left_score > right_score:
        return "Left"
    if right_score > left_score:
        return "Right"
    return provisional_tag


def refine_y_tag_with_consensus(
    rect: OrderedRect,
    provisional_tag: str | None,
    provisional_y_anchors: dict[int, float],
    consensus_tolerance_m: float,
    global_min_y: float,
    global_max_y: float,
) -> str | None:
    if rect.ymin_m == global_min_y:
        return "Front"
    if rect.ymax_m == global_max_y:
        return "Back"

    other_values = [value for rect_idx, value in provisional_y_anchors.items() if rect_idx != rect.idx]
    if not other_values:
        return provisional_tag

    front_score = consensus_score(rect.ymin_m, other_values, consensus_tolerance_m)
    back_score = consensus_score(rect.ymax_m, other_values, consensus_tolerance_m)
    if max(front_score[0], back_score[0]) == 0:
        return provisional_tag
    if front_score > back_score:
        return "Front"
    if back_score > front_score:
        return "Back"
    return provisional_tag


def auto_tag_rects(
    local_rects: list[OrderedRect],
    x_grid_tolerance_m: float,
    y_grid_tolerance_m: float,
) -> dict[int, tuple[str | None, str | None, str | None]]:
    # Build face lists SEPARATED BY SIDE. Previously all four side values
    # (Left/Right Xs and Front/Back Ys) were lumped into two single lists and
    # clustered together — which incorrectly let a Left face cluster with a
    # Right face, and a Front face cluster with a Back face. That mismatch
    # caused columns like C7 (whose Front face was geometrically close to
    # OTHER columns' Back faces) to be mis-tagged as "Front" when their beam
    # actually attaches on the Back face. With side-separated clustering,
    # only same-side faces influence each side's family size.
    left_faces = [Face(rect_idx=r.idx, side="Left",  value_m=r.xmin_m) for r in local_rects]
    right_faces = [Face(rect_idx=r.idx, side="Right", value_m=r.xmax_m) for r in local_rects]
    front_faces = [Face(rect_idx=r.idx, side="Front", value_m=r.ymin_m) for r in local_rects]
    back_faces  = [Face(rect_idx=r.idx, side="Back",  value_m=r.ymax_m) for r in local_rects]

    # Keep combined x_faces / y_faces around for Pass 2 below.
    x_faces = left_faces + right_faces
    y_faces = front_faces + back_faces

    def _combine_side_families(
        low_fs: list[Face], high_fs: list[Face], tol: float
    ) -> tuple[list[Family], dict[tuple[int, str], int]]:
        """Build families for the two sides independently, then offset the
        high-side indices and concatenate. The downstream `choose_x_tag`,
        `choose_y_tag`, and consensus refinement code reads from a single
        (families, lookup) pair via `lookup[(rect_idx, side)]`, so this keeps
        their signatures unchanged while guaranteeing low-side and high-side
        faces never share a family.
        """
        low_families, low_lookup = build_families(low_fs, tol)
        high_families, high_lookup = build_families(high_fs, tol)
        offset = len(low_families)
        combined_families = list(low_families) + list(high_families)
        combined_lookup = dict(low_lookup)
        for key, value in high_lookup.items():
            combined_lookup[key] = value + offset
        return combined_families, combined_lookup

    x_families, x_lookup = _combine_side_families(left_faces, right_faces, x_grid_tolerance_m)
    y_families, y_lookup = _combine_side_families(front_faces, back_faces, y_grid_tolerance_m)
    global_min_x = min(rect.xmin_m for rect in local_rects)
    global_max_x = max(rect.xmax_m for rect in local_rects)
    global_min_y = min(rect.ymin_m for rect in local_rects)
    global_max_y = max(rect.ymax_m for rect in local_rects)

    provisional: dict[int, tuple[str | None, str | None, str | None]] = {}
    for rect in local_rects:
        x_tag = choose_x_tag(rect, x_families, x_lookup, global_min_x, global_max_x)
        y_tag = choose_y_tag(rect, y_families, y_lookup, global_min_y, global_max_y)
        provisional[rect.idx] = (x_tag, y_tag, None)

    x_consensus_tolerance_m = max(x_grid_tolerance_m * 3.0, DEFAULT_X_CONSENSUS_TOLERANCE_M)
    y_consensus_tolerance_m = max(y_grid_tolerance_m * 3.0, DEFAULT_Y_CONSENSUS_TOLERANCE_M)

    provisional_x_anchors = {
        rect.idx: anchor_value
        for rect in local_rects
        if (anchor_value := chosen_anchor_value(rect, provisional[rect.idx][0], "X")) is not None
    }
    provisional_y_anchors = {
        rect.idx: anchor_value
        for rect in local_rects
        if (anchor_value := chosen_anchor_value(rect, provisional[rect.idx][1], "Y")) is not None
    }

    tags: dict[int, tuple[str | None, str | None, str | None]] = {}
    for rect in local_rects:
        provisional_x_tag, provisional_y_tag, location = provisional[rect.idx]
        x_tag = refine_x_tag_with_consensus(
            rect,
            provisional_x_tag,
            provisional_x_anchors,
            x_consensus_tolerance_m,
            global_min_x,
            global_max_x,
        )
        y_tag = refine_y_tag_with_consensus(
            rect,
            provisional_y_tag,
            provisional_y_anchors,
            y_consensus_tolerance_m,
            global_min_y,
            global_max_y,
        )
        tags[rect.idx] = (x_tag, y_tag, location)

    # ── Pass 2: Relaxed family merge for unresolved columns ──
    # Rebuild families with 2× tolerance, then re-tag only unresolved columns
    relaxed_x_tol = x_grid_tolerance_m * 2.0
    relaxed_y_tol = y_grid_tolerance_m * 2.0

    rect_by_idx = {rect.idx: rect for rect in local_rects}
    unresolved = {idx for idx, (xt, yt, _) in tags.items() if xt is None or yt is None}

    if unresolved:
        # Rebuild families with relaxed tolerance — merges nearby size-1 families.
        # Still SIDE-SEPARATED so Left/Right and Front/Back never cross-cluster
        # even at the 2× tolerance.
        relaxed_x_families, relaxed_x_lookup = _combine_side_families(left_faces, right_faces, relaxed_x_tol)
        relaxed_y_families, relaxed_y_lookup = _combine_side_families(front_faces, back_faces, relaxed_y_tol)

        for idx in unresolved:
            rect = rect_by_idx[idx]
            x_tag, y_tag, loc = tags[idx]

            # Try to resolve X (Left/Right) if None
            if x_tag is None:
                left_fam = relaxed_x_families[relaxed_x_lookup[(idx, "Left")]]
                right_fam = relaxed_x_families[relaxed_x_lookup[(idx, "Right")]]
                if max(left_fam.unique_rect_count, right_fam.unique_rect_count) >= 2:
                    ls = family_score(left_fam, rect.xmin_m)
                    rs = family_score(right_fam, rect.xmax_m)
                    if ls >= rs:
                        x_tag = "Left"  # Left wins ties (convention)
                    else:
                        x_tag = "Right"

            # Try to resolve Y (Front/Back) if None
            if y_tag is None:
                front_fam = relaxed_y_families[relaxed_y_lookup[(idx, "Front")]]
                back_fam = relaxed_y_families[relaxed_y_lookup[(idx, "Back")]]
                if max(front_fam.unique_rect_count, back_fam.unique_rect_count) >= 2:
                    fs = family_score(front_fam, rect.ymin_m)
                    bs = family_score(back_fam, rect.ymax_m)
                    if fs >= bs:
                        y_tag = "Front"  # Front wins ties (convention)
                    else:
                        y_tag = "Back"

            tags[idx] = (x_tag, y_tag, loc)

    return tags


def autosize(ws) -> None:
    for col in ws.columns:
        values = [str(cell.value) if cell.value is not None else "" for cell in col]
        width = min(max(len(value) for value in values) + 2, 40)
        ws.column_dimensions[col[0].column_letter].width = width


def write_excel(out_path: Path, local_rects: list[OrderedRect], tags: dict[int, tuple[str | None, str | None, str | None]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Rectangles_m"
    ws.append(
        [
            "Column No",
            "Xmin (m)",
            "Xmax (m)",
            "Ymin (m)",
            "Ymax (m)",
            "Width (m)",
            "Height (m)",
            "Left/Right",
            "Front/Back",
            "Location",
        ]
    )

    for rect in local_rects:
        x_tag, y_tag, location = tags[rect.idx]
        ws.append(
            [
                f"C{rect.idx}",
                rect.xmin_m,
                rect.xmax_m,
                rect.ymin_m,
                rect.ymax_m,
                rect.width_m,
                rect.height_m,
                x_tag,
                y_tag,
                location,
            ]
        )

    left_right_validation = DataValidation(type="list", formula1='"Left,Right,Centre"', allow_blank=True)
    front_back_validation = DataValidation(type="list", formula1='"Front,Back,Centre"', allow_blank=True)
    location_validation = DataValidation(type="list", formula1='"Lift,Staircase,Interior"', allow_blank=True)
    ws.add_data_validation(left_right_validation)
    ws.add_data_validation(front_back_validation)
    ws.add_data_validation(location_validation)

    last_row = len(local_rects) + 1
    if last_row >= 2:
        left_right_validation.add(f"H2:H{last_row}")
        front_back_validation.add(f"I2:I{last_row}")
        location_validation.add(f"J2:J{last_row}")

    autosize(ws)
    wb.save(out_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract COL rectangles from a DXF or DWG block and auto-tag clear grid-line-driven faces.")
    parser.add_argument("--dxf", type=Path, default=DEFAULT_DXF_PATH, help="Optional DXF or DWG path. If omitted, the script uses the only drawing currently open in AutoCAD.")
    parser.add_argument("--column-layer", default=DEFAULT_COLUMN_LAYER, help="Layer name for column rectangles.")
    parser.add_argument("--column-hatch-layer", default=DEFAULT_COLUMN_HATCH_LAYER, help="Layer name for column hatches. If the layer exists in the source, only hatch-overlapping rectangles are kept.")
    parser.add_argument("--target-rect-count", type=int, default=DEFAULT_TARGET_RECT_COUNT, help="Optional expected rectangle count for auto-detect. Leave unset to avoid count-based selection.")
    parser.add_argument("--block-name", default=DEFAULT_PREFERRED_BLOCK_NAME, help="Optional explicit block name override.")
    parser.add_argument("--mm-per-unit", type=float, default=DEFAULT_MM_PER_UNIT, help="Unit conversion to millimeters.")
    parser.add_argument("--inch-multiple-tolerance", type=float, default=DEFAULT_INCH_MULTIPLE_TOLERANCE, help="Tolerance in drawing units used to keep only rectangles whose width and height are near whole-inch multiples. Use 0 to disable.")
    parser.add_argument("--hatch-overlap-tolerance", type=float, default=DEFAULT_HATCH_OVERLAP_TOLERANCE, help="Tolerance in drawing units used for rectangle-to-hatch overlap checks.")
    parser.add_argument("--row-tolerance-mm", type=float, default=DEFAULT_ROW_TOLERANCE_MM, help="Tolerance in millimeters used to group rectangles into the same row.")
    parser.add_argument("--x-grid-tolerance-m", type=float, default=DEFAULT_X_GRID_TOLERANCE_M, help="Tolerance in meters used to form vertical control-line families.")
    parser.add_argument("--y-grid-tolerance-m", type=float, default=DEFAULT_Y_GRID_TOLERANCE_M, help="Tolerance in meters used to form horizontal control-line families.")
    parser.add_argument("--output", type=Path, default=None, help="Output Excel path. Defaults to <drawing_stem>_col_rectangles_m_v2.xlsx beside the source drawing.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    dxf_path = base.resolve_open_autocad_dxf(args.dxf)

    doc = base.read_cad_document(dxf_path)
    candidates = base.find_candidates(
        doc,
        args.column_layer,
        args.column_hatch_layer,
        args.inch_multiple_tolerance,
        args.hatch_overlap_tolerance,
    )
    candidate = base.select_candidate(candidates, args.block_name, args.target_rect_count)
    row_tolerance_units = args.row_tolerance_mm / args.mm_per_unit
    anchor, ordered = base.order_rects(candidate.rects, row_tolerance_units)

    local_rects = build_local_rects(ordered, anchor, args.mm_per_unit)
    tags = auto_tag_rects(local_rects, args.x_grid_tolerance_m, args.y_grid_tolerance_m)

    out_path = args.output or dxf_path.with_name(f"{dxf_path.stem}_col_rectangles_m_v2.xlsx")
    write_excel(out_path, local_rects, tags)

    print(f"Selected block   : {candidate.name}")
    print(f"Inserted count   : {candidate.insert_count}")
    print(f"Rectangles found : {len(local_rects)}")
    print(f"Excel saved      : {out_path}")


if __name__ == "__main__":
    main()
