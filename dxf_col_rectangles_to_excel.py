from __future__ import annotations

import argparse
import ctypes
import shutil
import tempfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import ezdxf
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

try:
    import win32com.client as win32
except ImportError:  # pragma: no cover - pywin32 is expected on the user's AutoCAD machine.
    win32 = None


# Edit these defaults for future drawings if needed.
DEFAULT_DXF_PATH = None
DEFAULT_COLUMN_LAYER = "COL"
DEFAULT_COLUMN_HATCH_LAYER = "column hatch"
DEFAULT_TARGET_RECT_COUNT = None
DEFAULT_PREFERRED_BLOCK_NAME = None  # Example: "A$Cffac18ee"
DEFAULT_MM_PER_UNIT = 25.4
DEFAULT_ROW_TOLERANCE_MM = 300.0
DEFAULT_INCH_MULTIPLE_TOLERANCE = 0.15
DEFAULT_HATCH_OVERLAP_TOLERANCE = 0.20
MODELSPACE_CANDIDATE_NAME = "__MODELSPACE__"
POPUP_TITLE = "Column Coordinates"
SUPPORTED_DRAWING_SUFFIXES = {".dxf", ".dwg"}
_TEMP_DXF_CACHE: dict[Path, Path] = {}
_TEMP_DXF_DIRS: list[tempfile.TemporaryDirectory[str]] = []
# AutoCAD type library enum: AcSaveAsType.ac2018_dxf
AUTOCAD_SAVEAS_DXF = 65


@dataclass(frozen=True)
class Rect:
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
    def cx(self) -> float:
        return (self.xmin + self.xmax) / 2.0

    @property
    def cy(self) -> float:
        return (self.ymin + self.ymax) / 2.0


@dataclass
class BlockCandidate:
    name: str
    insert_count: int
    rects: list[Rect]


def show_popup(message: str, title: str = POPUP_TITLE) -> None:
    ctypes.windll.user32.MessageBoxW(0, message, title, 0x10)


def validate_drawing_path(path: Path) -> Path:
    resolved = path.resolve()
    if not resolved.is_file():
        raise SystemExit(f"Drawing not found: {resolved}")
    if resolved.suffix.lower() not in SUPPORTED_DRAWING_SUFFIXES:
        raise SystemExit(f"Unsupported drawing format '{resolved.suffix}' for: {resolved}")
    return resolved


def _get_autocad_application(start_if_needed: bool):
    if win32 is None:
        raise SystemExit("pywin32 is required to read or convert AutoCAD drawings. Install with: pip install pywin32")

    try:
        return win32.GetActiveObject("AutoCAD.Application")
    except Exception as exc:
        if not start_if_needed:
            raise exc
        try:
            return win32.Dispatch("AutoCAD.Application")
        except Exception as dispatch_exc:
            raise SystemExit("AutoCAD is required to convert DWG drawings to temporary DXF files.") from dispatch_exc


def coerce_drawing_to_dxf_path(path: Path) -> Path:
    source_path = validate_drawing_path(path)
    if source_path.suffix.lower() == ".dxf":
        return source_path

    cached = _TEMP_DXF_CACHE.get(source_path)
    if cached is not None and cached.exists():
        return cached

    tmp_dir = tempfile.TemporaryDirectory(prefix=f"{source_path.stem}_dwg_")
    temp_root = Path(tmp_dir.name)
    temp_source = temp_root / source_path.name
    temp_dxf = temp_root / f"{source_path.stem}.dxf"
    shutil.copy2(source_path, temp_source)

    acad = _get_autocad_application(start_if_needed=True)
    temp_doc = None
    try:
        temp_doc = acad.Documents.Open(str(temp_source))
        temp_doc.SaveAs(str(temp_dxf), AUTOCAD_SAVEAS_DXF)
    except Exception as exc:
        raise SystemExit(f"Failed to convert DWG to temporary DXF: {source_path}\n{exc}") from exc
    finally:
        if temp_doc is not None:
            try:
                temp_doc.Close(False)
            except Exception:
                pass

    if not temp_dxf.exists():
        raise SystemExit(f"AutoCAD did not create the temporary DXF copy for: {source_path}")

    resolved_dxf = temp_dxf.resolve()
    _TEMP_DXF_DIRS.append(tmp_dir)
    _TEMP_DXF_CACHE[source_path] = resolved_dxf
    return resolved_dxf


def read_cad_document(path: Path):
    return ezdxf.readfile(coerce_drawing_to_dxf_path(path))


def resolve_open_autocad_dxf(explicit_path: Path | None) -> Path:
    if explicit_path is not None:
        return validate_drawing_path(explicit_path)

    if win32 is None:
        raise SystemExit("pywin32 is required to read the currently open AutoCAD drawing. Install with: pip install pywin32")

    try:
        acad = _get_autocad_application(start_if_needed=False)
    except Exception as exc:  # pragma: no cover - depends on a live AutoCAD session.
        message = "No AutoCAD session found. Open the relevant DXF or DWG in AutoCAD first."
        show_popup(message)
        raise SystemExit(message) from exc

    docs = acad.Documents
    open_paths: list[Path] = []
    for idx in range(docs.Count):
        doc = docs.Item(idx)
        full_name = str(getattr(doc, "FullName", "") or "").strip()
        if full_name:
            open_paths.append(Path(full_name))

    if len(open_paths) > 1:
        message = "Multiple drawings open. Keep only relevant drawing open"
        show_popup(message)
        raise SystemExit(message)
    if not open_paths:
        message = "No drawing open in AutoCAD. Open the relevant drawing first."
        show_popup(message)
        raise SystemExit(message)

    path = open_paths[0].resolve()
    if path.suffix.lower() not in SUPPORTED_DRAWING_SUFFIXES:
        message = "The open AutoCAD drawing is not a DXF or DWG. Open the relevant drawing."
        show_popup(message)
        raise SystemExit(message)
    return validate_drawing_path(path)


def parse_code_pairs(path: Path) -> list[tuple[str, str]]:
    lines = coerce_drawing_to_dxf_path(path).read_text(errors="ignore").splitlines()
    return [(lines[i].strip(), lines[i + 1].rstrip("\r\n")) for i in range(0, len(lines) - 1, 2)]


def section_pairs(pairs: list[tuple[str, str]], section_name: str) -> list[tuple[str, str]]:
    in_section = False
    out: list[tuple[str, str]] = []
    for code, value in pairs:
        if code == "2" and value == section_name:
            in_section = True
            continue
        if in_section and code == "0" and value == "ENDSEC":
            break
        if in_section:
            out.append((code, value))
    return out


def extract_insert_names(entity_pairs: list[tuple[str, str]]) -> Counter:
    counts: Counter[str] = Counter()
    i = 0
    while i < len(entity_pairs):
        code, value = entity_pairs[i]
        if code == "0" and value == "INSERT":
            name = None
            i += 1
            while i < len(entity_pairs):
                code2, value2 = entity_pairs[i]
                if code2 == "0":
                    break
                if code2 == "2":
                    name = value2
                i += 1
            if name:
                counts[name] += 1
        else:
            i += 1
    return counts


def iter_block_defs(block_pairs: list[tuple[str, str]]) -> Iterable[tuple[str, list[tuple[str, str]]]]:
    i = 0
    while i < len(block_pairs):
        code, value = block_pairs[i]
        if code == "0" and value == "BLOCK":
            name = None
            body: list[tuple[str, str]] = []
            i += 1
            while i < len(block_pairs):
                code2, value2 = block_pairs[i]
                if code2 == "2" and name is None:
                    name = value2
                if code2 == "0" and value2 == "ENDBLK":
                    if name:
                        yield name, body
                    break
                body.append((code2, value2))
                i += 1
        i += 1


def extract_rects_from_entity_stream(entity_pairs: list[tuple[str, str]], column_layer: str) -> list[Rect]:
    rects: list[Rect] = []
    i = 0
    while i < len(entity_pairs):
        code, value = entity_pairs[i]
        if code == "0" and value == "LWPOLYLINE":
            layer = None
            flags = 0
            vertices: list[list[float | None]] = []
            i += 1
            while i < len(entity_pairs):
                code2, value2 = entity_pairs[i]
                if code2 == "0":
                    break
                if code2 == "8":
                    layer = value2
                elif code2 == "70":
                    try:
                        flags = int(value2)
                    except ValueError:
                        flags = 0
                elif code2 == "10":
                    try:
                        vertices.append([float(value2), None])
                    except ValueError:
                        pass
                elif code2 == "20":
                    try:
                        if vertices:
                            vertices[-1][1] = float(value2)
                    except ValueError:
                        pass
                i += 1

            cleaned = [tuple(v) for v in vertices if v[0] is not None and v[1] is not None]
            if layer == column_layer and (flags & 1) and len(cleaned) == 4:
                xs = [p[0] for p in cleaned]
                ys = [p[1] for p in cleaned]
                rects.append(Rect(min(xs), max(xs), min(ys), max(ys)))
        else:
            i += 1
    return dedupe_rects(rects)


def dedupe_rects(rects: list[Rect], tol: float = 1e-3) -> list[Rect]:
    out: list[Rect] = []
    for rect in sorted(rects, key=lambda r: (r.xmin, r.ymin, r.xmax, r.ymax)):
        if any(
            abs(rect.xmin - seen.xmin) < tol
            and abs(rect.xmax - seen.xmax) < tol
            and abs(rect.ymin - seen.ymin) < tol
            and abs(rect.ymax - seen.ymax) < tol
            for seen in out
        ):
            continue
        out.append(rect)
    return out


def extract_insert_names_from_doc(doc: ezdxf.EzDxfDocument) -> Counter:
    return Counter(entity.dxf.name for entity in doc.modelspace().query("INSERT"))


def extract_rects_from_layout(layout, column_layer: str) -> list[Rect]:
    rects: list[Rect] = []
    for entity in layout:
        if entity.dxftype() != "LWPOLYLINE" or entity.dxf.layer != column_layer or not entity.closed:
            continue
        points = [tuple(map(float, point[:2])) for point in entity.get_points("xy")]
        if len(points) != 4:
            continue
        xs = [point[0] for point in points]
        ys = [point[1] for point in points]
        rects.append(Rect(min(xs), max(xs), min(ys), max(ys)))
    return dedupe_rects(rects)


def hatch_bbox(entity) -> Rect | None:
    xs: list[float] = []
    ys: list[float] = []
    for path in entity.paths:
        if hasattr(path, "vertices"):
            for vertex in path.vertices:
                xs.append(float(vertex[0]))
                ys.append(float(vertex[1]))
            continue
        if hasattr(path, "edges"):
            for edge in path.edges:
                if getattr(edge, "start", None) is not None:
                    xs.append(float(edge.start[0]))
                    ys.append(float(edge.start[1]))
                if getattr(edge, "end", None) is not None:
                    xs.append(float(edge.end[0]))
                    ys.append(float(edge.end[1]))
                if getattr(edge, "center", None) is not None:
                    xs.append(float(edge.center[0]))
                    ys.append(float(edge.center[1]))
    if not xs or not ys:
        return None
    return Rect(min(xs), max(xs), min(ys), max(ys))


def extract_hatch_boxes_from_layout(layout, hatch_layer: str | None) -> list[Rect]:
    if not hatch_layer:
        return []
    boxes: list[Rect] = []
    for entity in layout:
        if entity.dxftype() != "HATCH" or entity.dxf.layer != hatch_layer:
            continue
        bbox = hatch_bbox(entity)
        if bbox is not None:
            boxes.append(bbox)
    return dedupe_rects(boxes)


def rects_overlap(a: Rect, b: Rect, tol: float) -> bool:
    return not (
        a.xmax < b.xmin - tol
        or b.xmax < a.xmin - tol
        or a.ymax < b.ymin - tol
        or b.ymax < a.ymin - tol
    )


def is_near_integer(value: float, tolerance: float) -> bool:
    return abs(value - round(value)) <= tolerance


def _is_structural_size(rect: Rect, mm_per_unit: float) -> bool:
    """Check if rectangle dimensions are within reasonable column/shear-wall range (100-2000mm)."""
    MIN_STRUCTURAL_MM = 100.0
    MAX_STRUCTURAL_MM = 2000.0
    w_mm = rect.width * mm_per_unit
    h_mm = rect.height * mm_per_unit
    return MIN_STRUCTURAL_MM <= w_mm <= MAX_STRUCTURAL_MM and MIN_STRUCTURAL_MM <= h_mm <= MAX_STRUCTURAL_MM


def filter_rects(
    rects: list[Rect],
    hatch_boxes: list[Rect],
    inch_multiple_tolerance: float,
    hatch_overlap_tolerance: float,
    mm_per_unit: float = DEFAULT_MM_PER_UNIT,
) -> list[Rect]:
    require_hatch = bool(hatch_boxes)
    filtered: list[Rect] = []
    dropped: list[Rect] = []
    for rect in rects:
        if inch_multiple_tolerance > 0:
            if not is_near_integer(rect.width, inch_multiple_tolerance):
                dropped.append(rect)
                continue
            if not is_near_integer(rect.height, inch_multiple_tolerance):
                dropped.append(rect)
                continue
        if require_hatch and not any(rects_overlap(rect, hatch_box, hatch_overlap_tolerance) for hatch_box in hatch_boxes):
            continue
        filtered.append(rect)

    # Auto-rescue: dropped by inch filter but structurally valid size → include
    if dropped:
        for rect in dropped:
            if _is_structural_size(rect, mm_per_unit):
                if require_hatch and not any(rects_overlap(rect, hatch_box, hatch_overlap_tolerance) for hatch_box in hatch_boxes):
                    continue  # hatch filter still applies
                filtered.append(rect)

    return dedupe_rects(filtered)


def extract_filtered_rects_from_layout(
    layout,
    column_layer: str,
    column_hatch_layer: str | None,
    inch_multiple_tolerance: float,
    hatch_overlap_tolerance: float,
    mm_per_unit: float = DEFAULT_MM_PER_UNIT,
) -> list[Rect]:
    rects = extract_rects_from_layout(layout, column_layer)
    hatch_boxes = extract_hatch_boxes_from_layout(layout, column_hatch_layer)
    return filter_rects(rects, hatch_boxes, inch_multiple_tolerance, hatch_overlap_tolerance, mm_per_unit)


def find_candidates(
    doc: ezdxf.EzDxfDocument,
    column_layer: str,
    column_hatch_layer: str | None,
    inch_multiple_tolerance: float,
    hatch_overlap_tolerance: float,
    mm_per_unit: float = DEFAULT_MM_PER_UNIT,
) -> list[BlockCandidate]:
    insert_counts = extract_insert_names_from_doc(doc)
    candidates: list[BlockCandidate] = []
    for block in doc.blocks:
        block_name = block.name
        if block_name.startswith("*"):
            continue
        rects = extract_filtered_rects_from_layout(
            block,
            column_layer,
            column_hatch_layer,
            inch_multiple_tolerance,
            hatch_overlap_tolerance,
            mm_per_unit,
        )
        if rects:
            candidates.append(BlockCandidate(block_name, insert_counts.get(block_name, 0), rects))
    modelspace_rects = extract_filtered_rects_from_layout(
        doc.modelspace(),
        column_layer,
        column_hatch_layer,
        inch_multiple_tolerance,
        hatch_overlap_tolerance,
        mm_per_unit,
    )
    if modelspace_rects:
        candidates.append(BlockCandidate(MODELSPACE_CANDIDATE_NAME, 0, modelspace_rects))
    return candidates


def select_candidate(
    candidates: list[BlockCandidate],
    preferred_block_name: str | None,
    target_rect_count: int | None,
) -> BlockCandidate:
    name_map = {c.name: c for c in candidates}
    if preferred_block_name:
        if preferred_block_name not in name_map:
            raise SystemExit(f"Preferred block '{preferred_block_name}' not found.\nAvailable: {sorted(name_map)}")
        return name_map[preferred_block_name]

    modelspace_candidate = name_map.get(MODELSPACE_CANDIDATE_NAME)
    inserted = [c for c in candidates if c.insert_count > 0]
    if not inserted:
        if modelspace_candidate:
            return modelspace_candidate
        raise SystemExit("No inserted block definitions or modelspace candidates with matching column rectangles were found.")

    if target_rect_count is not None:
        exact = [c for c in inserted if len(c.rects) == target_rect_count]
        if len(exact) == 1:
            return exact[0]
        if len(exact) > 1:
            details = "\n".join(f"  - {c.name}: {len(c.rects)} rects, {c.insert_count} insert(s)" for c in exact)
            raise SystemExit(
                "Multiple inserted blocks match the target rectangle count.\n"
                "Set DEFAULT_PREFERRED_BLOCK_NAME or pass --block-name.\n"
                f"{details}"
            )
    if len(inserted) == 1:
        return inserted[0]

    details = "\n".join(f"  - {c.name}: {len(c.rects)} rects, {c.insert_count} insert(s)" for c in inserted)
    raise SystemExit(
        "Could not auto-select a unique block.\n"
        "Set DEFAULT_PREFERRED_BLOCK_NAME or pass --block-name.\n"
        "Optionally pass --target-rect-count if you know the expected count.\n"
        f"{details}"
    )


def order_rects(rects: list[Rect], row_tolerance_units: float) -> tuple[Rect, list[Rect]]:
    anchor = min(rects, key=lambda r: (r.ymin, r.xmin))
    by_center = sorted(rects, key=lambda r: (r.cy, r.xmin, r.ymin))
    rows: list[list[Rect]] = []
    row_centerlines: list[float] = []
    for rect in by_center:
        if not rows or abs(rect.cy - row_centerlines[-1]) > row_tolerance_units:
            rows.append([rect])
            row_centerlines.append(rect.cy)
        else:
            rows[-1].append(rect)
            row_centerlines[-1] = sum(r.cy for r in rows[-1]) / len(rows[-1])
    ordered: list[Rect] = []
    for row in sorted(rows, key=lambda current_row: sum(r.cy for r in current_row) / len(current_row)):
        ordered.extend(sorted(row, key=lambda r: (r.xmin, r.xmax)))
    return anchor, ordered


def autosize(ws) -> None:
    for col in ws.columns:
        values = [str(cell.value) if cell.value is not None else "" for cell in col]
        width = min(max(len(v) for v in values) + 2, 40)
        ws.column_dimensions[col[0].column_letter].width = width


def write_excel(
    out_path: Path,
    dxf_path: Path,
    candidate: BlockCandidate,
    anchor: Rect,
    ordered: list[Rect],
    mm_per_unit: float,
    row_tolerance_mm: float,
    row_tolerance_units: float,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Rectangles_m"
    headers = [
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
    ws.append(headers)
    for idx, rect in enumerate(ordered, start=1):
        xmin_m = round((rect.xmin - anchor.xmin) * mm_per_unit / 1000.0, 3)
        xmax_m = round((rect.xmax - anchor.xmin) * mm_per_unit / 1000.0, 3)
        ymin_m = round((rect.ymin - anchor.ymin) * mm_per_unit / 1000.0, 3)
        ymax_m = round((rect.ymax - anchor.ymin) * mm_per_unit / 1000.0, 3)
        ws.append(
            [
                f"C{idx}",
                xmin_m,
                xmax_m,
                ymin_m,
                ymax_m,
                round((rect.width) * mm_per_unit / 1000.0, 3),
                round((rect.height) * mm_per_unit / 1000.0, 3),
                None,
                None,
                None,
            ]
        )

    left_right_validation = DataValidation(type="list", formula1='"Left,Right,Centre"', allow_blank=True)
    front_back_validation = DataValidation(type="list", formula1='"Front,Back,Centre"', allow_blank=True)
    location_validation = DataValidation(type="list", formula1='"Lift,Staircase,Interior"', allow_blank=True)
    ws.add_data_validation(left_right_validation)
    ws.add_data_validation(front_back_validation)
    ws.add_data_validation(location_validation)

    last_row = len(ordered) + 1
    if last_row >= 2:
        left_right_validation.add(f"H2:H{last_row}")
        front_back_validation.add(f"I2:I{last_row}")
        location_validation.add(f"J2:J{last_row}")

    autosize(ws)
    wb.save(out_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract COL rectangles from a DXF or DWG block and export local meter coordinates to Excel.")
    parser.add_argument("--dxf", type=Path, default=DEFAULT_DXF_PATH, help="Optional DXF or DWG path. If omitted, the script uses the only drawing currently open in AutoCAD.")
    parser.add_argument("--column-layer", default=DEFAULT_COLUMN_LAYER, help="Layer name for column rectangles.")
    parser.add_argument("--column-hatch-layer", default=DEFAULT_COLUMN_HATCH_LAYER, help="Layer name for column hatches. If the layer exists in the source, only hatch-overlapping rectangles are kept.")
    parser.add_argument("--target-rect-count", type=int, default=DEFAULT_TARGET_RECT_COUNT, help="Optional expected rectangle count for auto-detect. Leave unset to avoid count-based selection.")
    parser.add_argument("--block-name", default=DEFAULT_PREFERRED_BLOCK_NAME, help="Optional explicit block name override.")
    parser.add_argument("--mm-per-unit", type=float, default=DEFAULT_MM_PER_UNIT, help="Unit conversion to millimeters.")
    parser.add_argument("--inch-multiple-tolerance", type=float, default=DEFAULT_INCH_MULTIPLE_TOLERANCE, help="Tolerance in drawing units used to keep only rectangles whose width and height are near whole-inch multiples. Use 0 to disable.")
    parser.add_argument("--hatch-overlap-tolerance", type=float, default=DEFAULT_HATCH_OVERLAP_TOLERANCE, help="Tolerance in drawing units used for rectangle-to-hatch overlap checks.")
    parser.add_argument(
        "--row-tolerance-mm",
        type=float,
        default=DEFAULT_ROW_TOLERANCE_MM,
        help="Tolerance in millimeters used to group rectangles into the same row.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output Excel path. Defaults to <drawing_stem>_col_rectangles_m.xlsx beside the source drawing.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    dxf_path = resolve_open_autocad_dxf(args.dxf)

    doc = read_cad_document(dxf_path)
    candidates = find_candidates(
        doc,
        args.column_layer,
        args.column_hatch_layer,
        args.inch_multiple_tolerance,
        args.hatch_overlap_tolerance,
        args.mm_per_unit,
    )
    candidate = select_candidate(candidates, args.block_name, args.target_rect_count)
    row_tolerance_units = args.row_tolerance_mm / args.mm_per_unit
    anchor, ordered = order_rects(candidate.rects, row_tolerance_units)

    out_path = args.output or dxf_path.with_name(f"{dxf_path.stem}_col_rectangles_m.xlsx")
    write_excel(
        out_path,
        dxf_path,
        candidate,
        anchor,
        ordered,
        args.mm_per_unit,
        args.row_tolerance_mm,
        row_tolerance_units,
    )

    print(f"Selected block   : {candidate.name}")
    print(f"Inserted count   : {candidate.insert_count}")
    print(f"Rectangles found : {len(ordered)}")
    print(f"Anchor (raw)     : xmin={anchor.xmin:.3f}, ymin={anchor.ymin:.3f}")
    print(f"Excel saved      : {out_path}")


if __name__ == "__main__":
    main()
