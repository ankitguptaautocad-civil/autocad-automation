from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path

import ezdxf
from openpyxl import Workbook

import dxf_col_rectangles_to_excel as base
import dxf_col_rectangles_to_excel_v2 as col_v2
import dxf_walls_to_excel_v2 as wall_v2


DEFAULT_DXF_PATH = base.DEFAULT_DXF_PATH
DEFAULT_COLUMN_LAYER = base.DEFAULT_COLUMN_LAYER
DEFAULT_COLUMN_HATCH_LAYER = base.DEFAULT_COLUMN_HATCH_LAYER
DEFAULT_TARGET_RECT_COUNT = base.DEFAULT_TARGET_RECT_COUNT
DEFAULT_PREFERRED_BLOCK_NAME = base.DEFAULT_PREFERRED_BLOCK_NAME
DEFAULT_MM_PER_UNIT = base.DEFAULT_MM_PER_UNIT
DEFAULT_ROW_TOLERANCE_MM = base.DEFAULT_ROW_TOLERANCE_MM
DEFAULT_INCH_MULTIPLE_TOLERANCE = base.DEFAULT_INCH_MULTIPLE_TOLERANCE
DEFAULT_HATCH_OVERLAP_TOLERANCE = base.DEFAULT_HATCH_OVERLAP_TOLERANCE
DEFAULT_X_GRID_TOLERANCE_M = col_v2.DEFAULT_X_GRID_TOLERANCE_M
DEFAULT_Y_GRID_TOLERANCE_M = col_v2.DEFAULT_Y_GRID_TOLERANCE_M
DEFAULT_WALL_LINE_LAYER = wall_v2.DEFAULT_WALL_LINE_LAYER
DEFAULT_HATCH_LAYER = wall_v2.DEFAULT_HATCH_LAYER
DEFAULT_WALL_THICKNESSES_MM = wall_v2.DEFAULT_WALL_THICKNESSES_MM
DEFAULT_WALL_THICKNESS_TOLERANCE_MM = wall_v2.DEFAULT_WALL_THICKNESS_TOLERANCE_MM
DEFAULT_MIN_WALL_LENGTH_MM = wall_v2.DEFAULT_MIN_WALL_LENGTH_MM
DEFAULT_COLUMN_OVERLAP_RATIO = wall_v2.DEFAULT_COLUMN_OVERLAP_RATIO
DEFAULT_DUPLICATE_CENTER_TOL_MM = wall_v2.DEFAULT_DUPLICATE_CENTER_TOL_RAW * wall_v2.DEFAULT_MM_PER_UNIT
DEFAULT_DUPLICATE_OVERLAP_RATIO = wall_v2.DEFAULT_DUPLICATE_OVERLAP_RATIO
DEFAULT_WALL_FACE_TOLERANCE_M = 0.05
DEFAULT_AXIS_ANGLE_TOLERANCE_DEG = wall_v2.DEFAULT_AXIS_ANGLE_TOLERANCE_DEG
DEFAULT_LIFT_WALL_SEARCH_RADIUS_M = 3.0
DEFAULT_LIFT_TEXT_SEARCH_RADIUS_M = 5.0
DEFAULT_LIFT_BOX_MARGIN_M = 0.3
DEFAULT_DXF_DIR = Path("STD ANL model")
DEFAULT_MATCH_CLUSTER_TOLERANCE_M = 0.15
DEFAULT_MATCH_DISTANCE_TOLERANCE_M = 1.5
FLOOR_KEYS = ("typical", "plinth", "terrace")
FLOOR_KEYWORDS = {
    "typical": "typical",
    "plinth": "plinth",
    "terrace": "terrace",
}


@dataclass(frozen=True)
class LocalWall:
    xmin_m: float
    xmax_m: float
    ymin_m: float
    ymax_m: float
    orientation: str

    @property
    def center_x(self) -> float:
        return (self.xmin_m + self.xmax_m) / 2.0

    @property
    def center_y(self) -> float:
        return (self.ymin_m + self.ymax_m) / 2.0


@dataclass(frozen=True)
class FloorArtifacts:
    floor_key: str
    dxf_path: Path
    candidate_name: str
    anchor: base.Rect
    ordered_rects: list[base.Rect]
    local_rects: list[col_v2.OrderedRect]
    ordered_walls: list[wall_v2.WallRect]
    primary_wall_count: int
    fallback_wall_count: int
    match_count: int | None = None
    match_shift_x: float | None = None
    match_shift_y: float | None = None
    max_local_coord_diff_m: float | None = None
    column_match_note: str | None = None


def autosize(ws) -> None:
    for col in ws.columns:
        values = [str(cell.value) if cell.value is not None else "" for cell in col]
        width = min(max(len(value) for value in values) + 2, 40)
        ws.column_dimensions[col[0].column_letter].width = width


def localize_walls(walls: list[wall_v2.WallRect], anchor: base.Rect, mm_per_unit: float) -> list[LocalWall]:
    scale = mm_per_unit / 1000.0
    local_walls: list[LocalWall] = []
    for wall in walls:
        local_walls.append(
            LocalWall(
                xmin_m=round((wall.xmin - anchor.xmin) * scale, 3),
                xmax_m=round((wall.xmax - anchor.xmin) * scale, 3),
                ymin_m=round((wall.ymin - anchor.ymin) * scale, 3),
                ymax_m=round((wall.ymax - anchor.ymin) * scale, 3),
                orientation=wall.orientation,
            )
        )
    return local_walls


def wall_face_values(local_walls: list[LocalWall]) -> tuple[list[float], list[float]]:
    x_faces: list[float] = []
    y_faces: list[float] = []
    for wall in local_walls:
        if wall.orientation == "Vertical":
            x_faces.extend([wall.xmin_m, wall.xmax_m])
        else:
            y_faces.extend([wall.ymin_m, wall.ymax_m])
    return x_faces, y_faces


def iter_text_entities(doc: ezdxf.EzDxf, candidate_name: str):
    if candidate_name == base.MODELSPACE_CANDIDATE_NAME:
        return [entity for entity in doc.modelspace() if entity.dxftype() in {"TEXT", "MTEXT"}]
    try:
        block = doc.blocks.get(candidate_name)
    except Exception:
        return []
    return [entity for entity in block if entity.dxftype() in {"TEXT", "MTEXT"}]


def extract_exact_lift_points(doc: ezdxf.EzDxf, candidate_name: str, anchor: base.Rect, mm_per_unit: float) -> list[tuple[float, float]]:
    scale = mm_per_unit / 1000.0
    points: list[tuple[float, float]] = []
    for entity in iter_text_entities(doc, candidate_name):
        if entity.dxftype() == "MTEXT":
            raw_text = entity.plain_text()
            insert = entity.dxf.insert
        else:
            raw_text = entity.dxf.text
            insert = entity.dxf.insert
        words = [word for word in re.findall(r"[A-Za-z]+", raw_text.upper()) if len(word) > 1]
        if words != ["LIFT"]:
            continue
        points.append(
            (
                round((float(insert.x) - anchor.xmin) * scale, 3),
                round((float(insert.y) - anchor.ymin) * scale, 3),
            )
        )
    return points


def choose_unique_nearest(columns_by_corner: list[list[tuple[float, int]]]) -> set[int]:
    chosen: set[int] = set()
    for options in columns_by_corner:
        for _, rect_idx in options:
            if rect_idx not in chosen:
                chosen.add(rect_idx)
                break
    return chosen


def _build_lift_bounding_box(
    text_x: float,
    text_y: float,
    left_wall: LocalWall | None,
    right_wall: LocalWall | None,
    front_wall: LocalWall | None,
    back_wall: LocalWall | None,
) -> tuple[float, float, float, float] | None:
    """Build a bounding box from shaft walls. Returns (xmin, xmax, ymin, ymax) or None."""
    walls_found = sum(w is not None for w in (left_wall, right_wall, front_wall, back_wall))
    if walls_found < 2:
        return None

    # Use wall positions where available; mirror from opposite wall for missing side
    if left_wall is not None:
        box_xmin = left_wall.center_x
    elif right_wall is not None:
        box_xmin = text_x - (right_wall.center_x - text_x)
    else:
        return None

    if right_wall is not None:
        box_xmax = right_wall.center_x
    elif left_wall is not None:
        box_xmax = text_x + (text_x - left_wall.center_x)
    else:
        return None

    if front_wall is not None:
        box_ymin = front_wall.center_y
    elif back_wall is not None:
        box_ymin = text_y - (back_wall.center_y - text_y)
    else:
        return None

    if back_wall is not None:
        box_ymax = back_wall.center_y
    elif front_wall is not None:
        box_ymax = text_y + (text_y - front_wall.center_y)
    else:
        return None

    # Sanity: box must be reasonable (0.5m to 5m each side)
    width = box_xmax - box_xmin
    height = box_ymax - box_ymin
    if width < 0.5 or width > 5.0 or height < 0.5 or height > 5.0:
        return None

    return (box_xmin, box_xmax, box_ymin, box_ymax)


def detect_lift_columns(
    local_rects: list[col_v2.OrderedRect],
    local_walls: list[LocalWall],
    lift_points: list[tuple[float, float]],
    wall_search_radius_m: float,
    text_search_radius_m: float,
    box_margin_m: float = DEFAULT_LIFT_BOX_MARGIN_M,
) -> set[int]:
    detected: set[int] = set()
    if not lift_points:
        return detected

    for text_x, text_y in lift_points:
        verticals = [
            wall for wall in local_walls
            if wall.orientation == "Vertical"
            and wall.ymin_m - 0.5 <= text_y <= wall.ymax_m + 0.5
            and abs(wall.center_x - text_x) <= wall_search_radius_m
        ]
        horizontals = [
            wall for wall in local_walls
            if wall.orientation == "Horizontal"
            and wall.xmin_m - 0.5 <= text_x <= wall.xmax_m + 0.5
            and abs(wall.center_y - text_y) <= wall_search_radius_m
        ]

        left_wall = min((wall for wall in verticals if wall.center_x < text_x), key=lambda wall: text_x - wall.center_x, default=None)
        right_wall = min((wall for wall in verticals if wall.center_x > text_x), key=lambda wall: wall.center_x - text_x, default=None)
        front_wall = min((wall for wall in horizontals if wall.center_y < text_y), key=lambda wall: text_y - wall.center_y, default=None)
        back_wall = min((wall for wall in horizontals if wall.center_y > text_y), key=lambda wall: wall.center_y - text_y, default=None)

        walls_found = sum(w is not None for w in (left_wall, right_wall, front_wall, back_wall))

        # ── Priority 1: Wall Bounding Box (>= 2 walls) ──
        # Rectangle overlap check — handles wide shear wall columns whose center
        # lies outside the shaft box but whose body forms the shaft boundary
        box = _build_lift_bounding_box(text_x, text_y, left_wall, right_wall, front_wall, back_wall)
        if box is not None:
            bx_min, bx_max, by_min, by_max = box
            inside = [
                rect for rect in local_rects
                if rect.xmax_m >= bx_min - box_margin_m
                and rect.xmin_m <= bx_max + box_margin_m
                and rect.ymax_m >= by_min - box_margin_m
                and rect.ymin_m <= by_max + box_margin_m
            ]
            if len(inside) >= 2:
                if walls_found < 3:
                    names = ", ".join(f"C{r.idx}" for r in inside)
                    print(f"  VERIFY: Lift near ({text_x:.3f}, {text_y:.3f}) — only {walls_found} walls, "
                          f"detected {len(inside)} columns by bounding box: {names}")
                else:
                    names = ", ".join(f"C{r.idx}" for r in inside)
                    print(f"  Lift near ({text_x:.3f}, {text_y:.3f}) — {walls_found} walls, "
                          f"{len(inside)} columns by bounding box: {names}")
                detected.update(r.idx for r in inside)
                continue

        # ── Priority 2: Constrained Quadrant (fallback — reduced radius, no forced-4) ──
        constrained_radius = min(text_search_radius_m, 2.0)
        nearby_rects = [
            rect for rect in local_rects
            if ((rect.cx_m - text_x) ** 2 + (rect.cy_m - text_y) ** 2) ** 0.5 <= constrained_radius
        ]

        quadrant_columns: list[list[tuple[float, int]]] = []
        quadrant_filters = (
            lambda rect: rect.cx_m <= text_x and rect.cy_m <= text_y,
            lambda rect: rect.cx_m >= text_x and rect.cy_m <= text_y,
            lambda rect: rect.cx_m <= text_x and rect.cy_m >= text_y,
            lambda rect: rect.cx_m >= text_x and rect.cy_m >= text_y,
        )
        for quadrant_filter in quadrant_filters:
            ranked = sorted(
                (
                    (
                        ((rect.cx_m - text_x) ** 2 + (rect.cy_m - text_y) ** 2) ** 0.5,
                        rect.idx,
                    )
                    for rect in nearby_rects
                    if quadrant_filter(rect)
                ),
                key=lambda item: item[0],
            )
            quadrant_columns.append(ranked)

        selected = choose_unique_nearest(quadrant_columns)
        if len(selected) >= 2:
            names = ", ".join(f"C{idx}" for idx in sorted(selected))
            print(f"  VERIFY: Lift near ({text_x:.3f}, {text_y:.3f}) — no bounding box, "
                  f"detected {len(selected)} columns by quadrant (radius {constrained_radius}m): {names}")
            detected.update(selected)
        else:
            print(f"  WARNING: Lift near ({text_x:.3f}, {text_y:.3f}) — could not detect lift columns "
                  f"(walls={walls_found}, nearby columns={len(nearby_rects)})")

    return detected


def face_match_score(target: float, face_values: list[float], tolerance_m: float) -> tuple[int, float]:
    deltas = [abs(target - value) for value in face_values if abs(target - value) <= tolerance_m]
    if not deltas:
        return 0, float("inf")
    return len(deltas), min(deltas)


def choose_wall_face_tag(
    low_value: float,
    high_value: float,
    face_values: list[float],
    low_label: str,
    high_label: str,
    tolerance_m: float,
) -> str | None:
    low_score = face_match_score(low_value, face_values, tolerance_m)
    high_score = face_match_score(high_value, face_values, tolerance_m)
    if low_score[0] == 0 and high_score[0] == 0:
        return None
    if low_score[0] > high_score[0]:
        return low_label
    if high_score[0] > low_score[0]:
        return high_label
    if low_score[1] < high_score[1]:
        return low_label
    if high_score[1] < low_score[1]:
        return high_label
    return None


def apply_wall_assisted_tags(
    local_rects: list[col_v2.OrderedRect],
    initial_tags: dict[int, tuple[str | None, str | None, str | None]],
    local_walls: list[LocalWall],
    wall_face_tolerance_m: float,
) -> tuple[dict[int, tuple[str | None, str | None, str | None]], list[str]]:
    x_faces, y_faces = wall_face_values(local_walls)
    updated_tags = dict(initial_tags)
    changes: list[str] = []
    for rect in local_rects:
        x_tag, y_tag, location = updated_tags[rect.idx]
        new_x = x_tag
        new_y = y_tag
        if new_x is None:
            new_x = choose_wall_face_tag(rect.xmin_m, rect.xmax_m, x_faces, "Left", "Right", wall_face_tolerance_m)
            if new_x is not None:
                changes.append(f"C{rect.idx}: Left/Right filled from walls -> {new_x}")
        if new_y is None:
            new_y = choose_wall_face_tag(rect.ymin_m, rect.ymax_m, y_faces, "Front", "Back", wall_face_tolerance_m)
            if new_y is not None:
                changes.append(f"C{rect.idx}: Front/Back filled from walls -> {new_y}")
        updated_tags[rect.idx] = (new_x, new_y, location)
    return updated_tags, changes


def apply_lift_locations(
    tags: dict[int, tuple[str | None, str | None, str | None]],
    lift_rect_ids: set[int],
    local_rects: list[col_v2.OrderedRect] | None = None,
) -> dict[int, tuple[str | None, str | None, str | None]]:
    updated: dict[int, tuple[str | None, str | None, str | None]] = {}

    # ── Lift center rule: fix L/R and F/B for lift columns ──
    # Outer face (away from lift center) = anchor face
    lift_lr_fb: dict[int, tuple[str, str]] = {}
    if local_rects and lift_rect_ids:
        rect_by_idx = {r.idx: r for r in local_rects}
        lift_rects = [rect_by_idx[idx] for idx in lift_rect_ids if idx in rect_by_idx]
        if lift_rects:
            lift_cx = sum(r.cx_m for r in lift_rects) / len(lift_rects)
            lift_cy = sum(r.cy_m for r in lift_rects) / len(lift_rects)
            for r in lift_rects:
                lr = "Left" if r.cx_m < lift_cx else "Right"
                fb = "Front" if r.cy_m < lift_cy else "Back"
                lift_lr_fb[r.idx] = (lr, fb)

    for rect_idx, (x_tag, y_tag, location) in tags.items():
        if rect_idx in lift_rect_ids:
            if rect_idx in lift_lr_fb:
                lr, fb = lift_lr_fb[rect_idx]
                updated[rect_idx] = (lr, fb, "Lift")
            else:
                updated[rect_idx] = (x_tag, y_tag, "Lift")
        else:
            updated[rect_idx] = (x_tag, y_tag, location)
    return updated


def default_dxf_dir() -> Path:
    candidate = (Path.cwd() / DEFAULT_DXF_DIR).resolve()
    if candidate.exists():
        return candidate
    return Path.cwd().resolve()


def resolve_optional_path(path: Path | None) -> Path | None:
    if path is None:
        return None
    return base.validate_drawing_path(path)


def discover_single_floor_dxf(search_dir: Path, floor_key: str) -> Path:
    keyword = FLOOR_KEYWORDS[floor_key]
    dxf_matches = sorted(path.resolve() for path in search_dir.glob("*.dxf") if keyword in path.name.lower())
    if len(dxf_matches) == 1:
        return dxf_matches[0]
    if len(dxf_matches) > 1:
        names = ", ".join(path.name for path in dxf_matches)
        raise SystemExit(f"Multiple DXFs matched '{keyword}' in {search_dir}: {names}. Pass the {floor_key} path explicitly.")

    dwg_matches = sorted(path.resolve() for path in search_dir.glob("*.dwg") if keyword in path.name.lower())
    if len(dwg_matches) == 1:
        return dwg_matches[0]
    if len(dwg_matches) > 1:
        names = ", ".join(path.name for path in dwg_matches)
        raise SystemExit(f"Multiple DWGs matched '{keyword}' in {search_dir}: {names}. Pass the {floor_key} path explicitly.")

    raise SystemExit(f"Could not find any DXF or DWG containing '{keyword}' in {search_dir}")


def resolve_floor_dxfs(args: argparse.Namespace) -> dict[str, Path]:
    search_dir = args.dxf_dir.resolve()
    if not search_dir.exists():
        raise SystemExit(f"Drawing search directory not found: {search_dir}")

    explicit = {
        "typical": resolve_optional_path(args.typical_dxf),
        "plinth": resolve_optional_path(args.plinth_dxf),
        "terrace": resolve_optional_path(args.terrace_dxf),
    }
    if args.dxf is not None and explicit["typical"] is None:
        explicit["typical"] = base.validate_drawing_path(args.dxf)

    resolved: dict[str, Path] = {}
    for floor_key in FLOOR_KEYS:
        resolved[floor_key] = explicit[floor_key] or discover_single_floor_dxf(search_dir, floor_key)
    return resolved


def rect_center(rect: base.Rect) -> tuple[float, float]:
    return ((rect.xmin + rect.xmax) / 2.0, (rect.ymin + rect.ymax) / 2.0)


def cluster_mode(values: list[float], tolerance: float) -> tuple[float, int]:
    ordered = sorted(values)
    best: list[float] = []
    current: list[float] = [ordered[0]]
    for value in ordered[1:]:
        if abs(value - (sum(current) / len(current))) <= tolerance:
            current.append(value)
        else:
            if len(current) > len(best):
                best = current
            current = [value]
    if len(current) > len(best):
        best = current
    return round(sum(best) / len(best), 6), len(best)


def match_rects_to_reference(
    reference_rects: list[base.Rect],
    candidate_rects: list[base.Rect],
    cluster_tolerance_m: float,
    match_tolerance_m: float,
    mm_per_unit: float,
) -> tuple[list[base.Rect], float, float]:
    scale = 1000.0 / mm_per_unit
    ref_centers = [rect_center(rect) for rect in reference_rects]
    candidate_centers = [rect_center(rect) for rect in candidate_rects]

    dx_values = [candidate_x - ref_x for ref_x, _ in ref_centers for candidate_x, _ in candidate_centers]
    dy_values = [candidate_y - ref_y for _, ref_y in ref_centers for _, candidate_y in candidate_centers]
    dx, _ = cluster_mode(dx_values, cluster_tolerance_m * scale)
    dy, _ = cluster_mode(dy_values, cluster_tolerance_m * scale)

    matched: list[base.Rect] = []
    used_indexes: set[int] = set()
    threshold = match_tolerance_m * scale

    for ref_x, ref_y in ref_centers:
        target_x = ref_x + dx
        target_y = ref_y + dy
        best: tuple[float, int] | None = None
        for idx, (cand_x, cand_y) in enumerate(candidate_centers):
            if idx in used_indexes:
                continue
            dist = ((cand_x - target_x) ** 2 + (cand_y - target_y) ** 2) ** 0.5
            if best is None or dist < best[0]:
                best = (dist, idx)
        if best is None or best[0] > threshold:
            raise SystemExit(
                "Could not align the floor-specific columns to the typical-floor reference. "
                "Check that the plinth/terrace drawings contain the same structural column layout."
            )
        used_indexes.add(best[1])
        matched.append(candidate_rects[best[1]])

    return matched, dx, dy


def local_coord_diffs_m(reference_rects: list[base.Rect], candidate_rects: list[base.Rect], mm_per_unit: float) -> float:
    scale = mm_per_unit / 1000.0
    ref_anchor = reference_rects[0]
    cand_anchor = candidate_rects[0]
    max_diff = 0.0
    for ref_rect, cand_rect in zip(reference_rects, candidate_rects):
        diffs = (
            abs((ref_rect.xmin - ref_anchor.xmin) - (cand_rect.xmin - cand_anchor.xmin)),
            abs((ref_rect.xmax - ref_anchor.xmin) - (cand_rect.xmax - cand_anchor.xmin)),
            abs((ref_rect.ymin - ref_anchor.ymin) - (cand_rect.ymin - cand_anchor.ymin)),
            abs((ref_rect.ymax - ref_anchor.ymin) - (cand_rect.ymax - cand_anchor.ymin)),
        )
        max_diff = max(max_diff, *(value * scale for value in diffs))
    return round(max_diff, 3)


def layout_for_candidate(doc: ezdxf.EzDxf, candidate_name: str):
    if candidate_name == base.MODELSPACE_CANDIDATE_NAME:
        return doc.modelspace()
    try:
        return doc.blocks.get(candidate_name)
    except Exception as exc:
        raise SystemExit(f"Block layout not found for {candidate_name}") from exc


def relaxed_candidate_rects(
    doc: ezdxf.EzDxf,
    candidate_name: str,
    column_layer: str,
    inch_multiple_tolerance: float,
) -> list[base.Rect]:
    layout = layout_for_candidate(doc, candidate_name)
    rects = base.extract_rects_from_layout(layout, column_layer)
    return base.filter_rects(
        rects,
        hatch_boxes=[],
        inch_multiple_tolerance=inch_multiple_tolerance,
        hatch_overlap_tolerance=0.0,
    )


def match_rects_with_filtered_then_relaxed_fallback(
    reference_rects: list[base.Rect],
    candidate_rects: list[base.Rect],
    doc: ezdxf.EzDxf,
    candidate_name: str,
    args: argparse.Namespace,
) -> tuple[list[base.Rect], float, float, str | None]:
    try:
        ordered_rects, shift_x, shift_y = match_rects_to_reference(
            reference_rects,
            candidate_rects,
            args.match_cluster_tolerance_m,
            args.match_distance_tolerance_m,
            args.mm_per_unit,
        )
        return ordered_rects, shift_x, shift_y, None
    except SystemExit as primary_exc:
        relaxed_rects = relaxed_candidate_rects(
            doc,
            candidate_name,
            args.column_layer,
            args.inch_multiple_tolerance,
        )
        if len(relaxed_rects) <= len(candidate_rects):
            raise primary_exc

        try:
            ordered_rects, shift_x, shift_y = match_rects_to_reference(
                reference_rects,
                relaxed_rects,
                args.match_cluster_tolerance_m,
                args.match_distance_tolerance_m,
                args.mm_per_unit,
            )
        except SystemExit:
            raise primary_exc

        note = (
            "Used relaxed column fallback: "
            f"{len(candidate_rects)} filtered rects did not align, "
            f"but {len(relaxed_rects)} same-source rects without the column-hatch gate did."
        )
        return ordered_rects, shift_x, shift_y, note


def extract_floor_artifacts(
    floor_key: str,
    dxf_path: Path,
    args: argparse.Namespace,
    reference_rects: list[base.Rect] | None = None,
) -> FloorArtifacts:
    doc = base.read_cad_document(dxf_path)
    pairs = base.parse_code_pairs(dxf_path)
    candidates = base.find_candidates(
        doc,
        args.column_layer,
        args.column_hatch_layer,
        args.inch_multiple_tolerance,
        args.hatch_overlap_tolerance,
    )
    candidate = base.select_candidate(candidates, args.block_name, args.target_rect_count if floor_key == "typical" else None)
    row_tolerance_units = args.row_tolerance_mm / args.mm_per_unit

    if reference_rects is None:
        anchor, ordered_rects = base.order_rects(candidate.rects, row_tolerance_units)
        match_count = None
        shift_x = None
        shift_y = None
        max_local_diff = None
        column_match_note = None
    else:
        ordered_rects, shift_x, shift_y, column_match_note = match_rects_with_filtered_then_relaxed_fallback(
            reference_rects,
            candidate.rects,
            doc,
            candidate.name,
            args,
        )
        anchor = ordered_rects[0]
        match_count = len(ordered_rects)
        max_local_diff = local_coord_diffs_m(reference_rects, ordered_rects, args.mm_per_unit)

    local_rects = col_v2.build_local_rects(ordered_rects, anchor, args.mm_per_unit)
    block_body = wall_v2.get_block_body(pairs, candidate.name)
    allowed_thicknesses = tuple(thickness_mm / args.mm_per_unit for thickness_mm in args.wall_thicknesses_mm)
    wall_thickness_tolerance = args.wall_thickness_tolerance_mm / args.mm_per_unit
    min_wall_length = args.min_wall_length_mm / args.mm_per_unit
    duplicate_center_tol = args.duplicate_center_tol_mm / args.mm_per_unit

    horizontal_lines, vertical_lines = wall_v2.extract_wall_lines(block_body, args.wall_line_layer, args.axis_angle_tolerance_deg)
    primary_walls = wall_v2.extract_from_wall_line_pairs(
        horizontal_lines=horizontal_lines,
        vertical_lines=vertical_lines,
        allowed_thicknesses=allowed_thicknesses,
        wall_thickness_tolerance=wall_thickness_tolerance,
        min_wall_length=min_wall_length,
        column_rects=ordered_rects,
        column_overlap_ratio=args.column_overlap_ratio,
    )
    fallback_walls = wall_v2.extract_from_hatch_fallback(
        block_body=block_body,
        hatch_layer=args.hatch_layer,
        allowed_thicknesses=allowed_thicknesses,
        wall_thickness_tolerance=wall_thickness_tolerance,
        min_wall_length=min_wall_length,
        column_rects=ordered_rects,
        column_overlap_ratio=args.column_overlap_ratio,
    )
    ordered_walls = wall_v2.order_walls(
        wall_v2.combine_primary_and_fallback(
            primary_walls=primary_walls,
            fallback_walls=fallback_walls,
            duplicate_center_tol=duplicate_center_tol,
            duplicate_overlap_ratio=args.duplicate_overlap_ratio,
        )
    )

    return FloorArtifacts(
        floor_key=floor_key,
        dxf_path=dxf_path,
        candidate_name=candidate.name,
        anchor=anchor,
        ordered_rects=ordered_rects,
        local_rects=local_rects,
        ordered_walls=ordered_walls,
        primary_wall_count=len(primary_walls),
        fallback_wall_count=len(fallback_walls),
        match_count=match_count,
        match_shift_x=shift_x,
        match_shift_y=shift_y,
        max_local_coord_diff_m=max_local_diff,
        column_match_note=column_match_note,
    )


def localized_wall_row(wall: wall_v2.WallRect, anchor: base.Rect, mm_per_unit: float) -> tuple[float, float, float, float, float, str]:
    scale = mm_per_unit / 1000.0
    if wall.orientation == "Vertical":
        start_x = end_x = round((wall.center_x - anchor.xmin) * scale, 3)
        start_y = round((wall.ymin - anchor.ymin) * scale, 3)
        end_y = round((wall.ymax - anchor.ymin) * scale, 3)
        return start_x, start_y, end_x, end_y, round(wall.thickness * scale, 3), wall.orientation
    start_x = round((wall.xmin - anchor.xmin) * scale, 3)
    end_x = round((wall.xmax - anchor.xmin) * scale, 3)
    start_y = end_y = round((wall.center_y - anchor.ymin) * scale, 3)
    return start_x, start_y, end_x, end_y, round(wall.thickness * scale, 3), wall.orientation


def write_combined_wall_excel(out_path: Path, floors: list[FloorArtifacts], mm_per_unit: float) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Walls_m"
    ws.append(
        [
            "Wall No",
            "Start X (m)",
            "Start Y (m)",
            "End X (m)",
            "End Y (m)",
            "Thickness (m)",
            "Orientation",
            "Source",
            "DXF Source",
        ]
    )

    wall_no = 1
    for floor in floors:
        for wall in floor.ordered_walls:
            start_x, start_y, end_x, end_y, thickness, orientation = localized_wall_row(wall, floor.anchor, mm_per_unit)
            ws.append(
                [
                    f"W{wall_no}",
                    start_x,
                    start_y,
                    end_x,
                    end_y,
                    thickness,
                    orientation,
                    wall.source,
                    floor.floor_key,
                ]
            )
            wall_no += 1

    autosize(ws)
    wb.save(out_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="One-click 3-drawing pipeline: tag columns from the typical-floor drawing, extract walls from typical/plinth/terrace, and write a combined wall workbook."
    )
    parser.add_argument("--dxf-dir", type=Path, default=default_dxf_dir(), help="Directory used to discover the typical/plinth/terrace DXF or DWG files by filename phrase.")
    parser.add_argument("--dxf", type=Path, default=DEFAULT_DXF_PATH, help="Deprecated alias for --typical-dxf. Accepts DXF or DWG.")
    parser.add_argument("--typical-dxf", type=Path, default=None, help="Optional typical-floor DXF or DWG path.")
    parser.add_argument("--plinth-dxf", type=Path, default=None, help="Optional plinth DXF or DWG path.")
    parser.add_argument("--terrace-dxf", type=Path, default=None, help="Optional terrace DXF or DWG path.")
    parser.add_argument("--column-layer", default=DEFAULT_COLUMN_LAYER, help="Layer name for column rectangles.")
    parser.add_argument("--column-hatch-layer", default=DEFAULT_COLUMN_HATCH_LAYER, help="Layer name for column hatches.")
    parser.add_argument("--target-rect-count", type=int, default=DEFAULT_TARGET_RECT_COUNT, help="Optional expected rectangle count for auto-detecting the typical-floor columns.")
    parser.add_argument("--block-name", default=DEFAULT_PREFERRED_BLOCK_NAME, help="Optional explicit block name override.")
    parser.add_argument("--mm-per-unit", type=float, default=DEFAULT_MM_PER_UNIT, help="Unit conversion to millimeters.")
    parser.add_argument("--inch-multiple-tolerance", type=float, default=DEFAULT_INCH_MULTIPLE_TOLERANCE, help="Tolerance in drawing units used while filtering candidate column rectangles.")
    parser.add_argument("--hatch-overlap-tolerance", type=float, default=DEFAULT_HATCH_OVERLAP_TOLERANCE, help="Tolerance in drawing units used while matching candidate column rectangles to hatches.")
    parser.add_argument("--row-tolerance-mm", type=float, default=DEFAULT_ROW_TOLERANCE_MM, help="Tolerance in millimeters used to group columns into the same row.")
    parser.add_argument("--x-grid-tolerance-m", type=float, default=DEFAULT_X_GRID_TOLERANCE_M, help="Tolerance in meters used to form vertical control-line families.")
    parser.add_argument("--y-grid-tolerance-m", type=float, default=DEFAULT_Y_GRID_TOLERANCE_M, help="Tolerance in meters used to form horizontal control-line families.")
    parser.add_argument("--wall-line-layer", default=DEFAULT_WALL_LINE_LAYER, help="Primary line layer used for wall edge detection.")
    parser.add_argument("--hatch-layer", default=DEFAULT_HATCH_LAYER, help="Fallback hatch layer used for missing wall bands.")
    parser.add_argument("--wall-thicknesses-mm", type=float, nargs="+", default=list(DEFAULT_WALL_THICKNESSES_MM), help="Allowed wall thicknesses in millimeters.")
    parser.add_argument("--wall-thickness-tolerance-mm", type=float, default=DEFAULT_WALL_THICKNESS_TOLERANCE_MM, help="Tolerance in millimeters around the allowed wall thicknesses.")
    parser.add_argument("--min-wall-length-mm", type=float, default=DEFAULT_MIN_WALL_LENGTH_MM, help="Minimum wall length in millimeters.")
    parser.add_argument("--column-overlap-ratio", type=float, default=DEFAULT_COLUMN_OVERLAP_RATIO, help="Maximum allowed overlap ratio with known columns before a wall candidate is discarded.")
    parser.add_argument("--duplicate-center-tol-mm", type=float, default=DEFAULT_DUPLICATE_CENTER_TOL_MM, help="Tolerance in millimeters used to treat a HACH fallback wall as already covered by a WALL-line wall.")
    parser.add_argument("--duplicate-overlap-ratio", type=float, default=DEFAULT_DUPLICATE_OVERLAP_RATIO, help="Minimum overlap ratio used to deduplicate WALL and HACH wall candidates.")
    parser.add_argument("--axis-angle-tolerance-deg", type=float, default=DEFAULT_AXIS_ANGLE_TOLERANCE_DEG, help="Maximum angular deviation from horizontal/vertical used while classifying WALL lines.")
    parser.add_argument("--wall-face-tolerance-m", type=float, default=DEFAULT_WALL_FACE_TOLERANCE_M, help="Tolerance in meters used to match unresolved column faces to nearby wall faces.")
    parser.add_argument("--lift-wall-search-radius-m", type=float, default=DEFAULT_LIFT_WALL_SEARCH_RADIUS_M, help="Search radius in meters used while finding shaft walls around exact LIFT text.")
    parser.add_argument("--lift-text-search-radius-m", type=float, default=DEFAULT_LIFT_TEXT_SEARCH_RADIUS_M, help="Search radius in meters used while finding the nearest supporting columns around LIFT text.")
    parser.add_argument("--match-cluster-tolerance-m", type=float, default=DEFAULT_MATCH_CLUSTER_TOLERANCE_M, help="Tolerance in meters used while finding the translation between the typical columns and the plinth/terrace columns.")
    parser.add_argument("--match-distance-tolerance-m", type=float, default=DEFAULT_MATCH_DISTANCE_TOLERANCE_M, help="Maximum center-point mismatch in meters allowed while mapping plinth/terrace columns back to the typical-floor column order.")
    parser.add_argument("--column-output", type=Path, default=None, help="Output path for the wall-assisted column workbook.")
    parser.add_argument("--wall-output", type=Path, default=None, help="Output path for the combined wall workbook.")
    parser.add_argument("--output-dir", type=Path, default=None, help="Directory for generated Excel files. Overrides the default (next to the typical DXF). Ignored for outputs that set --column-output/--wall-output explicitly.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    floor_dxfs = resolve_floor_dxfs(args)

    typical = extract_floor_artifacts("typical", floor_dxfs["typical"], args)
    initial_tags = col_v2.auto_tag_rects(typical.local_rects, args.x_grid_tolerance_m, args.y_grid_tolerance_m)
    typical_local_walls = localize_walls(typical.ordered_walls, typical.anchor, args.mm_per_unit)
    final_tags, wall_changes = apply_wall_assisted_tags(typical.local_rects, initial_tags, typical_local_walls, args.wall_face_tolerance_m)

    typical_doc = base.read_cad_document(floor_dxfs["typical"])
    lift_points = extract_exact_lift_points(typical_doc, typical.candidate_name, typical.anchor, args.mm_per_unit)
    lift_rect_ids = detect_lift_columns(
        typical.local_rects,
        typical_local_walls,
        lift_points,
        args.lift_wall_search_radius_m,
        args.lift_text_search_radius_m,
        DEFAULT_LIFT_BOX_MARGIN_M,
    )
    final_tags = apply_lift_locations(final_tags, lift_rect_ids, local_rects=typical.local_rects)

    # ── Post-lift consensus: re-check non-lift columns with corrected lift anchors ──
    post_lift_changes = []
    if lift_rect_ids:
        rect_by_idx_pl = {r.idx: r for r in typical.local_rects}
        # Build corrected anchor values from all tags (including lift-fixed)
        corrected_x_anchors: dict[int, float] = {}
        corrected_y_anchors: dict[int, float] = {}
        for idx, (xt, yt, _) in final_tags.items():
            r = rect_by_idx_pl[idx]
            if xt == "Left":
                corrected_x_anchors[idx] = r.xmin_m
            elif xt == "Right":
                corrected_x_anchors[idx] = r.xmax_m
            if yt == "Front":
                corrected_y_anchors[idx] = r.ymin_m
            elif yt == "Back":
                corrected_y_anchors[idx] = r.ymax_m

        x_consensus_tol = args.x_grid_tolerance_m * 3.0
        y_consensus_tol = args.y_grid_tolerance_m * 3.0

        for idx, (xt, yt, loc) in list(final_tags.items()):
            if idx in lift_rect_ids:
                continue
            r = rect_by_idx_pl[idx]
            new_xt, new_yt = xt, yt

            # Re-check X
            if xt is not None:
                other_x = [v for k, v in corrected_x_anchors.items() if k != idx]
                if other_x:
                    l_d = [abs(r.xmin_m - v) for v in other_x if abs(r.xmin_m - v) <= x_consensus_tol]
                    r_d = [abs(r.xmax_m - v) for v in other_x if abs(r.xmax_m - v) <= x_consensus_tol]
                    l_s = (len(l_d), -(sum(l_d)/len(l_d)) if l_d else -999, -(min(l_d)) if l_d else -999)
                    r_s = (len(r_d), -(sum(r_d)/len(r_d)) if r_d else -999, -(min(r_d)) if r_d else -999)
                    if l_s > r_s and l_s[0] > 0:
                        new_xt = "Left"
                    elif r_s > l_s and r_s[0] > 0:
                        new_xt = "Right"

            # Re-check Y
            if yt is not None:
                other_y = [v for k, v in corrected_y_anchors.items() if k != idx]
                if other_y:
                    f_d = [abs(r.ymin_m - v) for v in other_y if abs(r.ymin_m - v) <= y_consensus_tol]
                    b_d = [abs(r.ymax_m - v) for v in other_y if abs(r.ymax_m - v) <= y_consensus_tol]
                    f_s = (len(f_d), -(sum(f_d)/len(f_d)) if f_d else -999, -(min(f_d)) if f_d else -999)
                    b_s = (len(b_d), -(sum(b_d)/len(b_d)) if b_d else -999, -(min(b_d)) if b_d else -999)
                    if f_s > b_s and f_s[0] > 0:
                        new_yt = "Front"
                    elif b_s > f_s and b_s[0] > 0:
                        new_yt = "Back"

            if new_xt != xt or new_yt != yt:
                if new_xt != xt:
                    post_lift_changes.append(f"C{idx}: L/R {xt} -> {new_xt} (post-lift consensus)")
                if new_yt != yt:
                    post_lift_changes.append(f"C{idx}: F/B {yt} -> {new_yt} (post-lift consensus)")
                final_tags[idx] = (new_xt, new_yt, loc)

    plinth = extract_floor_artifacts("plinth", floor_dxfs["plinth"], args, reference_rects=typical.ordered_rects)
    terrace = extract_floor_artifacts("terrace", floor_dxfs["terrace"], args, reference_rects=typical.ordered_rects)

    typical_dxf = floor_dxfs["typical"]
    output_dir = args.output_dir.resolve() if args.output_dir else typical_dxf.parent
    output_dir.mkdir(parents=True, exist_ok=True)
    from datetime import datetime as _dt
    timestamp = _dt.now().strftime("%Y%m%d_%H%M%S")
    column_output = args.column_output.resolve() if args.column_output else output_dir / f"{typical_dxf.stem}_{timestamp}_col_rectangles_m_v2_wall_assisted.xlsx"
    wall_output = args.wall_output.resolve() if args.wall_output else output_dir / f"{typical_dxf.stem}_{timestamp}_walls_m_v2.xlsx"
    col_v2.write_excel(column_output, typical.local_rects, final_tags)
    write_combined_wall_excel(wall_output, [typical, plinth, terrace], args.mm_per_unit)

    unresolved_lr = sum(1 for x_tag, _, _ in final_tags.values() if x_tag is None)
    unresolved_fb = sum(1 for _, y_tag, _ in final_tags.values() if y_tag is None)

    print(f"Typical drawing        : {floor_dxfs['typical']}")
    print(f"Plinth drawing         : {floor_dxfs['plinth']}")
    print(f"Terrace drawing        : {floor_dxfs['terrace']}")
    print(f"Typical candidate      : {typical.candidate_name}")
    print(f"Typical columns found  : {len(typical.local_rects)}")
    for floor in (plinth, terrace):
        print(f"{floor.floor_key.title()} candidate       : {floor.candidate_name}")
        print(f"{floor.floor_key.title()} matched cols    : {floor.match_count}")
        print(f"{floor.floor_key.title()} shift (x, y)    : ({floor.match_shift_x:.3f}, {floor.match_shift_y:.3f})")
        print(f"{floor.floor_key.title()} max local diff  : {floor.max_local_coord_diff_m:.3f} m")
        if floor.column_match_note:
            print(f"{floor.floor_key.title()} column fallback : {floor.column_match_note}")
    print(f"Typical WALL walls     : {typical.primary_wall_count}")
    print(f"Typical HACH walls     : {typical.fallback_wall_count}")
    print(f"Plinth WALL walls      : {plinth.primary_wall_count}")
    print(f"Plinth HACH walls      : {plinth.fallback_wall_count}")
    print(f"Terrace WALL walls     : {terrace.primary_wall_count}")
    print(f"Terrace HACH walls     : {terrace.fallback_wall_count}")
    print(f"Combined wall rows     : {len(typical.ordered_walls) + len(plinth.ordered_walls) + len(terrace.ordered_walls)}")
    print(f"Wall-assisted fills    : {len(wall_changes)}")
    for change in wall_changes:
        print(f"  - {change}")
    print(f"Lift-tagged columns    : {len(lift_rect_ids)}")
    if lift_rect_ids:
        print("  - " + ", ".join(f"C{rect_idx}" for rect_idx in sorted(lift_rect_ids)))
    print(f"Post-lift consensus    : {len(post_lift_changes)}")
    for change in post_lift_changes:
        print(f"  - {change}")
    print(f"Unresolved Left/Right  : {unresolved_lr}")
    print(f"Unresolved Front/Back  : {unresolved_fb}")
    print(f"Column workbook saved  : {column_output}")
    print(f"Wall workbook saved    : {wall_output}")


if __name__ == "__main__":
    main()
