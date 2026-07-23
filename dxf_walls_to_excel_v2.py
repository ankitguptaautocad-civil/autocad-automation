from __future__ import annotations

import argparse
import math
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook

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
DEFAULT_WALL_LINE_LAYER = "WALL"
DEFAULT_HATCH_LAYER = "HACH"
# Allowed wall thicknesses (mm). 100 and 200 added for consultant drawings that
# use them (e.g. MAX) alongside the 115/230 brick sizes.
DEFAULT_WALL_THICKNESSES_MM = (100.0, 115.0, 200.0, 230.0)
# Tolerance MUST stay below half the smallest gap between two allowed thicknesses
# (100 vs 115 are only 15 mm apart), otherwise their bands overlap and a wall can
# be classified as the wrong thickness - which would change its mass and stiffness.
DEFAULT_WALL_THICKNESS_TOLERANCE_MM = 7.0
DEFAULT_MIN_WALL_LENGTH_MM = 300.0
DEFAULT_COLUMN_OVERLAP_RATIO = 0.5
DEFAULT_DUPLICATE_CENTER_TOL_RAW = 1.0
DEFAULT_DUPLICATE_OVERLAP_RATIO = 0.8
DEFAULT_AXIS_ANGLE_TOLERANCE_DEG = 1.0


@dataclass(frozen=True)
class WallRect:
    xmin: float
    xmax: float
    ymin: float
    ymax: float
    source: str

    @property
    def width(self) -> float:
        return self.xmax - self.xmin

    @property
    def height(self) -> float:
        return self.ymax - self.ymin

    @property
    def thickness(self) -> float:
        return min(self.width, self.height)

    @property
    def length(self) -> float:
        return max(self.width, self.height)

    @property
    def orientation(self) -> str:
        return "Vertical" if self.width <= self.height else "Horizontal"

    @property
    def center_x(self) -> float:
        return (self.xmin + self.xmax) / 2.0

    @property
    def center_y(self) -> float:
        return (self.ymin + self.ymax) / 2.0


@dataclass(frozen=True)
class AxisLine:
    start: float
    end: float
    fixed: float
    handle: str


def autosize(ws) -> None:
    for col in ws.columns:
        values = [str(cell.value) if cell.value is not None else "" for cell in col]
        width = min(max(len(value) for value in values) + 2, 40)
        ws.column_dimensions[col[0].column_letter].width = width


def overlap_len(a_start: float, a_end: float, b_start: float, b_end: float) -> float:
    return max(0.0, min(a_end, b_end) - max(a_start, b_start))


def rect_overlap_area(a: WallRect | base.Rect, b: WallRect | base.Rect) -> float:
    xmin = max(a.xmin, b.xmin)
    xmax = min(a.xmax, b.xmax)
    ymin = max(a.ymin, b.ymin)
    ymax = min(a.ymax, b.ymax)
    return max(0.0, xmax - xmin) * max(0.0, ymax - ymin)


def gap_matches(gap: float, allowed_thicknesses: tuple[float, ...], tolerance: float) -> bool:
    return any(abs(gap - thickness) <= tolerance for thickness in allowed_thicknesses)


def get_block_body(pairs: list[tuple[str, str]], block_name: str) -> list[tuple[str, str]]:
    if block_name == base.MODELSPACE_CANDIDATE_NAME:
        return base.section_pairs(pairs, "ENTITIES")
    block_pairs = base.section_pairs(pairs, "BLOCKS")
    for name, body in base.iter_block_defs(block_pairs):
        if name == block_name:
            return body
    raise SystemExit(f"Block body not found for {block_name}")


def parse_hatch_polyline_loops(entity_pairs: list[tuple[str, str]]) -> list[list[tuple[float, float]]]:
    loops: list[list[tuple[float, float]]] = []
    i = 0
    while i < len(entity_pairs):
        code, value = entity_pairs[i]
        if code != "92":
            i += 1
            continue

        i += 1
        if i + 2 >= len(entity_pairs):
            break

        if entity_pairs[i][0] == "72" and entity_pairs[i + 1][0] == "73" and entity_pairs[i + 2][0] == "93":
            i += 2
            vertex_count = int(entity_pairs[i][1])
            i += 1
            loop: list[tuple[float, float]] = []
            for _ in range(vertex_count):
                if i + 1 >= len(entity_pairs):
                    break
                if entity_pairs[i][0] != "10" or entity_pairs[i + 1][0] != "20":
                    break
                loop.append((float(entity_pairs[i][1]), float(entity_pairs[i + 1][1])))
                i += 2
            if loop:
                loops.append(loop)
            while i < len(entity_pairs) and entity_pairs[i][0] not in {"92", "0"}:
                i += 1
            continue

        while i < len(entity_pairs) and entity_pairs[i][0] not in {"92", "0"}:
            i += 1

    return loops


def is_axis_aligned(loop: list[tuple[float, float]], tol: float = 1e-3) -> bool:
    for (x1, y1), (x2, y2) in zip(loop, loop[1:] + loop[:1]):
        if abs(x1 - x2) > tol and abs(y1 - y2) > tol:
            return False
    return True


def iter_entities(block_body: list[tuple[str, str]]) -> list[list[tuple[str, str]]]:
    entities: list[list[tuple[str, str]]] = []
    i = 0
    while i < len(block_body):
        code, value = block_body[i]
        if code == "0":
            entity = [(code, value)]
            i += 1
            while i < len(block_body) and block_body[i][0] != "0":
                entity.append(block_body[i])
                i += 1
            entities.append(entity)
        else:
            i += 1
    return entities


def line_angle_deg(dx: float, dy: float) -> float:
    return abs(math.degrees(math.atan2(dy, dx))) % 180.0


def is_near_horizontal(angle_deg: float, tolerance_deg: float) -> bool:
    return min(angle_deg, abs(angle_deg - 180.0)) <= tolerance_deg


def is_near_vertical(angle_deg: float, tolerance_deg: float) -> bool:
    return abs(angle_deg - 90.0) <= tolerance_deg


def extract_wall_lines(
    block_body: list[tuple[str, str]],
    wall_line_layer: str,
    axis_angle_tolerance_deg: float = DEFAULT_AXIS_ANGLE_TOLERANCE_DEG,
) -> tuple[list[AxisLine], list[AxisLine]]:
    horizontal: list[AxisLine] = []
    vertical: list[AxisLine] = []
    i = 0
    while i < len(block_body):
        code, value = block_body[i]
        if code == "0" and value == "LINE":
            layer = None
            handle = ""
            x1 = y1 = x2 = y2 = None
            i += 1
            while i < len(block_body) and block_body[i][0] != "0":
                code2, value2 = block_body[i]
                if code2 == "8":
                    layer = value2
                elif code2 == "5":
                    handle = value2
                elif code2 == "10":
                    x1 = float(value2)
                elif code2 == "20":
                    y1 = float(value2)
                elif code2 == "11":
                    x2 = float(value2)
                elif code2 == "21":
                    y2 = float(value2)
                i += 1
            if layer != wall_line_layer or x1 is None or y1 is None or x2 is None or y2 is None:
                continue
            dx = x2 - x1
            dy = y2 - y1
            angle_deg = line_angle_deg(dx, dy)
            if is_near_horizontal(angle_deg, axis_angle_tolerance_deg):
                horizontal.append(AxisLine(min(x1, x2), max(x1, x2), (y1 + y2) / 2.0, handle))
            elif is_near_vertical(angle_deg, axis_angle_tolerance_deg):
                vertical.append(AxisLine(min(y1, y2), max(y1, y2), (x1 + x2) / 2.0, handle))
        else:
            i += 1
    return horizontal, vertical


def overlaps_columns(rect: WallRect, column_rects: list[base.Rect], column_overlap_ratio: float) -> bool:
    rect_area = rect.width * rect.height
    for col_rect in column_rects:
        col_area = col_rect.width * col_rect.height
        overlap = rect_overlap_area(rect, col_rect)
        if overlap > column_overlap_ratio * min(rect_area, col_area):
            return True
    return False


def overlap_interval_with_column(rect: WallRect, col_rect: base.Rect) -> tuple[float, float] | None:
    overlap_x = overlap_len(rect.xmin, rect.xmax, col_rect.xmin, col_rect.xmax)
    overlap_y = overlap_len(rect.ymin, rect.ymax, col_rect.ymin, col_rect.ymax)
    if overlap_x <= 0 or overlap_y <= 0:
        return None
    if rect.orientation == "Horizontal":
        return (max(rect.xmin, col_rect.xmin), min(rect.xmax, col_rect.xmax))
    return (max(rect.ymin, col_rect.ymin), min(rect.ymax, col_rect.ymax))


def subtract_intervals(base_interval: tuple[float, float], blocked: list[tuple[float, float]]) -> list[tuple[float, float]]:
    remaining = [base_interval]
    for block_lo, block_hi in sorted(blocked):
        next_remaining: list[tuple[float, float]] = []
        for curr_lo, curr_hi in remaining:
            if block_hi <= curr_lo or block_lo >= curr_hi:
                next_remaining.append((curr_lo, curr_hi))
                continue
            if block_lo > curr_lo:
                next_remaining.append((curr_lo, block_lo))
            if block_hi < curr_hi:
                next_remaining.append((block_hi, curr_hi))
        remaining = next_remaining
    return remaining


def trim_wall_rect_by_columns(
    rect: WallRect,
    column_rects: list[base.Rect],
    column_overlap_ratio: float,
    min_wall_length: float,
) -> list[WallRect]:
    blockers: list[tuple[float, float]] = []
    rect_area = rect.width * rect.height
    for col_rect in column_rects:
        col_area = col_rect.width * col_rect.height
        overlap = rect_overlap_area(rect, col_rect)
        if overlap <= column_overlap_ratio * min(rect_area, col_area):
            continue
        interval = overlap_interval_with_column(rect, col_rect)
        if interval is not None:
            blockers.append(interval)

    if not blockers:
        return [rect]

    if rect.orientation == "Horizontal":
        remaining = subtract_intervals((rect.xmin, rect.xmax), blockers)
        return [
            WallRect(xmin=lo, xmax=hi, ymin=rect.ymin, ymax=rect.ymax, source=rect.source)
            for lo, hi in remaining
            if (hi - lo) >= min_wall_length
        ]

    remaining = subtract_intervals((rect.ymin, rect.ymax), blockers)
    return [
        WallRect(xmin=rect.xmin, xmax=rect.xmax, ymin=lo, ymax=hi, source=rect.source)
        for lo, hi in remaining
        if (hi - lo) >= min_wall_length
    ]


def dedupe_exact_rects(rects: list[WallRect]) -> list[WallRect]:
    out: list[WallRect] = []
    seen: set[tuple[float, float, float, float]] = set()
    for rect in sorted(rects, key=lambda r: (r.ymin, r.xmin, r.ymax, r.xmax, r.source)):
        key = (round(rect.xmin, 3), round(rect.xmax, 3), round(rect.ymin, 3), round(rect.ymax, 3))
        if key in seen:
            continue
        seen.add(key)
        out.append(rect)
    return out


def extract_from_wall_line_pairs(
    horizontal_lines: list[AxisLine],
    vertical_lines: list[AxisLine],
    allowed_thicknesses: tuple[float, ...],
    wall_thickness_tolerance: float,
    min_wall_length: float,
    column_rects: list[base.Rect],
    column_overlap_ratio: float,
) -> list[WallRect]:
    walls: list[WallRect] = []

    for i, line_a in enumerate(horizontal_lines):
        for line_b in horizontal_lines[i + 1 :]:
            gap = abs(line_b.fixed - line_a.fixed)
            if not gap_matches(gap, allowed_thicknesses, wall_thickness_tolerance):
                continue
            overlap = overlap_len(line_a.start, line_a.end, line_b.start, line_b.end)
            if overlap < min_wall_length:
                continue
            rect = WallRect(
                xmin=max(line_a.start, line_b.start),
                xmax=min(line_a.end, line_b.end),
                ymin=min(line_a.fixed, line_b.fixed),
                ymax=max(line_a.fixed, line_b.fixed),
                source="WALL",
            )
            walls.extend(trim_wall_rect_by_columns(rect, column_rects, column_overlap_ratio, min_wall_length))

    for i, line_a in enumerate(vertical_lines):
        for line_b in vertical_lines[i + 1 :]:
            gap = abs(line_b.fixed - line_a.fixed)
            if not gap_matches(gap, allowed_thicknesses, wall_thickness_tolerance):
                continue
            overlap = overlap_len(line_a.start, line_a.end, line_b.start, line_b.end)
            if overlap < min_wall_length:
                continue
            rect = WallRect(
                xmin=min(line_a.fixed, line_b.fixed),
                xmax=max(line_a.fixed, line_b.fixed),
                ymin=max(line_a.start, line_b.start),
                ymax=min(line_a.end, line_b.end),
                source="WALL",
            )
            walls.extend(trim_wall_rect_by_columns(rect, column_rects, column_overlap_ratio, min_wall_length))

    return dedupe_exact_rects(walls)


def extract_wall_rectangles(
    block_body: list[tuple[str, str]],
    hatch_layer: str,
    allowed_thicknesses: tuple[float, ...],
    wall_thickness_tolerance: float,
    min_wall_length: float,
    column_rects: list[base.Rect],
    column_overlap_ratio: float,
) -> list[WallRect]:
    walls: list[WallRect] = []
    for entity in iter_entities(block_body):
        if entity[0] != ("0", "HATCH"):
            continue
        layer = next((value for code, value in entity if code == "8"), None)
        if layer != hatch_layer:
            continue

        for loop in parse_hatch_polyline_loops(entity[1:]):
            if not is_axis_aligned(loop):
                continue

            xs = [point[0] for point in loop]
            ys = [point[1] for point in loop]
            rect = WallRect(min(xs), max(xs), min(ys), max(ys), "HACH")
            if rect.length < min_wall_length:
                continue
            if not any(abs(rect.thickness - thickness) <= wall_thickness_tolerance for thickness in allowed_thicknesses):
                continue
            walls.extend(trim_wall_rect_by_columns(rect, column_rects, column_overlap_ratio, min_wall_length))

    return dedupe_exact_rects(walls)


def extract_from_hatch_fallback(
    block_body: list[tuple[str, str]],
    hatch_layer: str,
    allowed_thicknesses: tuple[float, ...],
    wall_thickness_tolerance: float,
    min_wall_length: float,
    column_rects: list[base.Rect],
    column_overlap_ratio: float,
) -> list[WallRect]:
    return extract_wall_rectangles(
        block_body=block_body,
        hatch_layer=hatch_layer,
        allowed_thicknesses=allowed_thicknesses,
        wall_thickness_tolerance=wall_thickness_tolerance,
        min_wall_length=min_wall_length,
        column_rects=column_rects,
        column_overlap_ratio=column_overlap_ratio,
    )


def is_duplicate_wall(
    existing: WallRect,
    candidate: WallRect,
    duplicate_center_tol: float,
    duplicate_overlap_ratio: float,
) -> bool:
    if existing.orientation != candidate.orientation:
        return False

    if existing.orientation == "Horizontal":
        if abs(existing.center_y - candidate.center_y) > duplicate_center_tol:
            return False
        overlap = overlap_len(existing.xmin, existing.xmax, candidate.xmin, candidate.xmax)
        return overlap >= duplicate_overlap_ratio * min(existing.length, candidate.length)

    if abs(existing.center_x - candidate.center_x) > duplicate_center_tol:
        return False
    overlap = overlap_len(existing.ymin, existing.ymax, candidate.ymin, candidate.ymax)
    return overlap >= duplicate_overlap_ratio * min(existing.length, candidate.length)


def combine_primary_and_fallback(
    primary_walls: list[WallRect],
    fallback_walls: list[WallRect],
    duplicate_center_tol: float,
    duplicate_overlap_ratio: float,
) -> list[WallRect]:
    combined = list(primary_walls)
    for fallback_wall in fallback_walls:
        if any(is_duplicate_wall(existing, fallback_wall, duplicate_center_tol, duplicate_overlap_ratio) for existing in combined):
            continue
        combined.append(fallback_wall)
    return dedupe_exact_rects(combined)


def order_walls(walls: list[WallRect]) -> list[WallRect]:
    return sorted(
        walls,
        key=lambda wall: (
            round(min(wall.ymin, wall.ymax), 3),
            round(min(wall.xmin, wall.xmax), 3),
            round(max(wall.ymin, wall.ymax), 3),
            round(max(wall.xmin, wall.xmax), 3),
            wall.source,
        ),
    )


def write_excel(out_path: Path, walls: list[WallRect], anchor: base.Rect, mm_per_unit: float) -> None:
    scale = mm_per_unit / 1000.0
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
        ]
    )

    for idx, wall in enumerate(walls, start=1):
        if wall.orientation == "Vertical":
            start_x = end_x = round((wall.center_x - anchor.xmin) * scale, 3)
            start_y = round((wall.ymin - anchor.ymin) * scale, 3)
            end_y = round((wall.ymax - anchor.ymin) * scale, 3)
            start_end = (start_x, start_y, end_x, end_y)
        else:
            start_x = round((wall.xmin - anchor.xmin) * scale, 3)
            end_x = round((wall.xmax - anchor.xmin) * scale, 3)
            start_y = end_y = round((wall.center_y - anchor.ymin) * scale, 3)
            start_end = (start_x, start_y, end_x, end_y)

        ws.append(
            [
                f"W{idx}",
                start_end[0],
                start_end[1],
                start_end[2],
                start_end[3],
                round(wall.thickness * scale, 3),
                wall.orientation,
                wall.source,
            ]
        )

    autosize(ws)
    wb.save(out_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract wall centerlines and thicknesses from the selected DXF or DWG layout using WALL lines first and HACH as fallback."
    )
    parser.add_argument("--dxf", type=Path, default=DEFAULT_DXF_PATH, help="Optional DXF or DWG path. If omitted, the script uses the only drawing currently open in AutoCAD.")
    parser.add_argument("--column-layer", default=DEFAULT_COLUMN_LAYER, help="Layer name for column rectangles.")
    parser.add_argument("--column-hatch-layer", default=DEFAULT_COLUMN_HATCH_LAYER, help="Layer name for column hatches used while locating the same layout as the column extractor.")
    parser.add_argument("--target-rect-count", type=int, default=DEFAULT_TARGET_RECT_COUNT, help="Optional expected column rectangle count for auto-detecting the layout block.")
    parser.add_argument("--block-name", default=DEFAULT_PREFERRED_BLOCK_NAME, help="Optional explicit block name override.")
    parser.add_argument("--mm-per-unit", type=float, default=DEFAULT_MM_PER_UNIT, help="Unit conversion to millimeters.")
    parser.add_argument("--inch-multiple-tolerance", type=float, default=DEFAULT_INCH_MULTIPLE_TOLERANCE, help="Tolerance in drawing units used while filtering candidate column rectangles.")
    parser.add_argument("--hatch-overlap-tolerance", type=float, default=DEFAULT_HATCH_OVERLAP_TOLERANCE, help="Tolerance in drawing units used while matching candidate column rectangles to hatches.")
    parser.add_argument("--row-tolerance-mm", type=float, default=DEFAULT_ROW_TOLERANCE_MM, help="Tolerance in millimeters used to order the column anchor rows.")
    parser.add_argument("--wall-line-layer", default=DEFAULT_WALL_LINE_LAYER, help="Primary line layer used for wall edge detection.")
    parser.add_argument("--hatch-layer", default=DEFAULT_HATCH_LAYER, help="Fallback hatch layer used for missing wall bands.")
    parser.add_argument("--wall-thicknesses-mm", type=float, nargs="+", default=list(DEFAULT_WALL_THICKNESSES_MM), help="Allowed wall thicknesses in millimeters.")
    parser.add_argument("--wall-thickness-tolerance-mm", type=float, default=DEFAULT_WALL_THICKNESS_TOLERANCE_MM, help="Tolerance in millimeters around the allowed wall thicknesses.")
    parser.add_argument("--min-wall-length-mm", type=float, default=DEFAULT_MIN_WALL_LENGTH_MM, help="Minimum wall length in millimeters.")
    parser.add_argument("--column-overlap-ratio", type=float, default=DEFAULT_COLUMN_OVERLAP_RATIO, help="Maximum allowed overlap ratio with known columns before a wall candidate is discarded.")
    parser.add_argument("--duplicate-center-tol-mm", type=float, default=DEFAULT_DUPLICATE_CENTER_TOL_RAW * DEFAULT_MM_PER_UNIT, help="Tolerance in millimeters used to treat a HACH fallback wall as already covered by a WALL-line wall.")
    parser.add_argument("--duplicate-overlap-ratio", type=float, default=DEFAULT_DUPLICATE_OVERLAP_RATIO, help="Minimum overlap ratio used to deduplicate WALL and HACH wall candidates.")
    parser.add_argument("--axis-angle-tolerance-deg", type=float, default=DEFAULT_AXIS_ANGLE_TOLERANCE_DEG, help="Maximum angular deviation from horizontal/vertical used while classifying WALL lines.")
    parser.add_argument("--output", type=Path, default=None, help="Output Excel path. Defaults to <drawing_stem>_walls_m_v2.xlsx beside the source drawing.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    dxf_path = base.resolve_open_autocad_dxf(args.dxf)

    pairs = base.parse_code_pairs(dxf_path)
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
    anchor, ordered_columns = base.order_rects(candidate.rects, row_tolerance_units)

    block_body = get_block_body(pairs, candidate.name)
    allowed_thicknesses = tuple(thickness_mm / args.mm_per_unit for thickness_mm in args.wall_thicknesses_mm)
    wall_thickness_tolerance = args.wall_thickness_tolerance_mm / args.mm_per_unit
    min_wall_length = args.min_wall_length_mm / args.mm_per_unit
    duplicate_center_tol = args.duplicate_center_tol_mm / args.mm_per_unit

    horizontal_lines, vertical_lines = extract_wall_lines(block_body, args.wall_line_layer, args.axis_angle_tolerance_deg)
    primary_walls = extract_from_wall_line_pairs(
        horizontal_lines=horizontal_lines,
        vertical_lines=vertical_lines,
        allowed_thicknesses=allowed_thicknesses,
        wall_thickness_tolerance=wall_thickness_tolerance,
        min_wall_length=min_wall_length,
        column_rects=ordered_columns,
        column_overlap_ratio=args.column_overlap_ratio,
    )
    fallback_walls = extract_from_hatch_fallback(
        block_body=block_body,
        hatch_layer=args.hatch_layer,
        allowed_thicknesses=allowed_thicknesses,
        wall_thickness_tolerance=wall_thickness_tolerance,
        min_wall_length=min_wall_length,
        column_rects=ordered_columns,
        column_overlap_ratio=args.column_overlap_ratio,
    )
    ordered_walls = order_walls(
        combine_primary_and_fallback(
            primary_walls=primary_walls,
            fallback_walls=fallback_walls,
            duplicate_center_tol=duplicate_center_tol,
            duplicate_overlap_ratio=args.duplicate_overlap_ratio,
        )
    )

    out_path = args.output or dxf_path.with_name(f"{dxf_path.stem}_walls_m_v2.xlsx")
    write_excel(out_path, ordered_walls, anchor, args.mm_per_unit)

    print(f"Selected block     : {candidate.name}")
    print(f"WALL-line walls    : {len(primary_walls)}")
    print(f"HACH fallback walls: {len(fallback_walls)}")
    print(f"Combined walls     : {len(ordered_walls)}")
    print(f"Excel saved        : {out_path}")


if __name__ == "__main__":
    main()
