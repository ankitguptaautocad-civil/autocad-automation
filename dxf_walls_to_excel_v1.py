from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook

import dxf_col_rectangles_to_excel as base


DEFAULT_DXF_PATH = base.DEFAULT_DXF_PATH
DEFAULT_COLUMN_LAYER = base.DEFAULT_COLUMN_LAYER
DEFAULT_TARGET_RECT_COUNT = base.DEFAULT_TARGET_RECT_COUNT
DEFAULT_PREFERRED_BLOCK_NAME = base.DEFAULT_PREFERRED_BLOCK_NAME
DEFAULT_MM_PER_UNIT = base.DEFAULT_MM_PER_UNIT
DEFAULT_ROW_TOLERANCE_MM = base.DEFAULT_ROW_TOLERANCE_MM
DEFAULT_HATCH_LAYER = "HACH"
DEFAULT_WALL_THICKNESSES_MM = (115.0, 230.0)
DEFAULT_WALL_THICKNESS_TOLERANCE_MM = 20.0
DEFAULT_MIN_WALL_LENGTH_MM = 300.0
DEFAULT_COLUMN_OVERLAP_RATIO = 0.5


@dataclass(frozen=True)
class WallRect:
    xmin: float
    xmax: float
    ymin: float
    ymax: float

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
    def cx(self) -> float:
        return (self.xmin + self.xmax) / 2.0

    @property
    def cy(self) -> float:
        return (self.ymin + self.ymax) / 2.0


def autosize(ws) -> None:
    for col in ws.columns:
        values = [str(cell.value) if cell.value is not None else "" for cell in col]
        width = min(max(len(value) for value in values) + 2, 40)
        ws.column_dimensions[col[0].column_letter].width = width


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


def get_block_body(pairs: list[tuple[str, str]], block_name: str) -> list[tuple[str, str]]:
    block_pairs = base.section_pairs(pairs, "BLOCKS")
    for name, body in base.iter_block_defs(block_pairs):
        if name == block_name:
            return body
    raise SystemExit(f"Block body not found for {block_name}")


def rect_overlap_area(a: WallRect | base.Rect, b: WallRect | base.Rect) -> float:
    xmin = max(a.xmin, b.xmin)
    xmax = min(a.xmax, b.xmax)
    ymin = max(a.ymin, b.ymin)
    ymax = min(a.ymax, b.ymax)
    return max(0.0, xmax - xmin) * max(0.0, ymax - ymin)


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
            rect = WallRect(min(xs), max(xs), min(ys), max(ys))
            if rect.length < min_wall_length:
                continue
            if not any(abs(rect.thickness - thickness) <= wall_thickness_tolerance for thickness in allowed_thicknesses):
                continue

            rect_area = rect.width * rect.height
            overlaps_column = False
            for col_rect in column_rects:
                col_area = col_rect.width * col_rect.height
                overlap = rect_overlap_area(rect, col_rect)
                if overlap > column_overlap_ratio * min(rect_area, col_area):
                    overlaps_column = True
                    break
            if overlaps_column:
                continue

            walls.append(rect)

    deduped: list[WallRect] = []
    seen: set[tuple[float, float, float, float]] = set()
    for rect in sorted(walls, key=lambda r: (r.ymin, r.xmin, r.ymax, r.xmax)):
        key = (round(rect.xmin, 3), round(rect.xmax, 3), round(rect.ymin, 3), round(rect.ymax, 3))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(rect)
    return deduped


def order_walls(walls: list[WallRect]) -> list[WallRect]:
    return sorted(
        walls,
        key=lambda wall: (
            round(min(wall.ymin, wall.ymax), 3),
            round(min(wall.xmin, wall.xmax), 3),
            round(max(wall.ymin, wall.ymax), 3),
            round(max(wall.xmin, wall.xmax), 3),
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
        ]
    )

    for idx, wall in enumerate(walls, start=1):
        if wall.orientation == "Vertical":
            start_x = end_x = round(((wall.xmin + wall.xmax) / 2.0 - anchor.xmin) * scale, 3)
            start_y = round((wall.ymin - anchor.ymin) * scale, 3)
            end_y = round((wall.ymax - anchor.ymin) * scale, 3)
        else:
            start_x = round((wall.xmin - anchor.xmin) * scale, 3)
            end_x = round((wall.xmax - anchor.xmin) * scale, 3)
            start_y = end_y = round(((wall.ymin + wall.ymax) / 2.0 - anchor.ymin) * scale, 3)

        ws.append(
            [
                f"W{idx}",
                start_x,
                start_y,
                end_x,
                end_y,
                round(wall.thickness * scale, 3),
                wall.orientation,
            ]
        )

    autosize(ws)
    wb.save(out_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract wall centerlines and thicknesses from the selected DXF or DWG layout using HACH wall bands."
    )
    parser.add_argument("--dxf", type=Path, default=DEFAULT_DXF_PATH, help="Path to the DXF or DWG file.")
    parser.add_argument("--column-layer", default=DEFAULT_COLUMN_LAYER, help="Layer name for column rectangles.")
    parser.add_argument("--target-rect-count", type=int, default=DEFAULT_TARGET_RECT_COUNT, help="Expected column rectangle count for auto-detecting the layout block.")
    parser.add_argument("--block-name", default=DEFAULT_PREFERRED_BLOCK_NAME, help="Optional explicit block name override.")
    parser.add_argument("--mm-per-unit", type=float, default=DEFAULT_MM_PER_UNIT, help="Unit conversion to millimeters.")
    parser.add_argument("--row-tolerance-mm", type=float, default=DEFAULT_ROW_TOLERANCE_MM, help="Tolerance in millimeters used to order the column anchor rows.")
    parser.add_argument("--hatch-layer", default=DEFAULT_HATCH_LAYER, help="Hatch layer used for wall fills.")
    parser.add_argument("--wall-thicknesses-mm", type=float, nargs="+", default=list(DEFAULT_WALL_THICKNESSES_MM), help="Allowed wall thicknesses in millimeters.")
    parser.add_argument("--wall-thickness-tolerance-mm", type=float, default=DEFAULT_WALL_THICKNESS_TOLERANCE_MM, help="Tolerance in millimeters around the allowed wall thicknesses.")
    parser.add_argument("--min-wall-length-mm", type=float, default=DEFAULT_MIN_WALL_LENGTH_MM, help="Minimum wall length in millimeters.")
    parser.add_argument("--column-overlap-ratio", type=float, default=DEFAULT_COLUMN_OVERLAP_RATIO, help="Maximum allowed overlap ratio with known columns before a hatch rectangle is discarded.")
    parser.add_argument("--output", type=Path, default=None, help="Output Excel path. Defaults to <drawing_stem>_walls_m.xlsx beside the source drawing.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    dxf_path = base.validate_drawing_path(args.dxf)

    pairs = base.parse_code_pairs(dxf_path)
    candidates = base.find_candidates(pairs, args.column_layer)
    candidate = base.select_candidate(candidates, args.block_name, args.target_rect_count)
    row_tolerance_units = args.row_tolerance_mm / args.mm_per_unit
    anchor, ordered_columns = base.order_rects(candidate.rects, row_tolerance_units)

    block_body = get_block_body(pairs, candidate.name)
    allowed_thicknesses = tuple(thickness_mm / args.mm_per_unit for thickness_mm in args.wall_thicknesses_mm)
    wall_thickness_tolerance = args.wall_thickness_tolerance_mm / args.mm_per_unit
    min_wall_length = args.min_wall_length_mm / args.mm_per_unit
    walls = extract_wall_rectangles(
        block_body=block_body,
        hatch_layer=args.hatch_layer,
        allowed_thicknesses=allowed_thicknesses,
        wall_thickness_tolerance=wall_thickness_tolerance,
        min_wall_length=min_wall_length,
        column_rects=ordered_columns,
        column_overlap_ratio=args.column_overlap_ratio,
    )
    ordered_walls = order_walls(walls)

    out_path = args.output or dxf_path.with_name(f"{dxf_path.stem}_walls_m.xlsx")
    write_excel(out_path, ordered_walls, anchor, args.mm_per_unit)

    print(f"Selected block : {candidate.name}")
    print(f"Walls found    : {len(ordered_walls)}")
    print(f"Excel saved    : {out_path}")


if __name__ == "__main__":
    main()
