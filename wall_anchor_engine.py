"""Wall-based hierarchical anchor-tagging engine (axis-generic).

Decides which face of a column is its structural anchor (the beam/grid line the
column sits on) from the *control lines* reconstructed from wall faces — rather
than from inter-column face alignment alone (the family method), which ties and
mis-defaults on symmetric interior columns.

AXIS-GENERIC BY DESIGN. Front/Back (Y axis) is switched on now. Left/Right
(X axis) is the SAME code path — activate it by adding "X" to ENABLED_AXES.
Y-anchor evidence comes from HORIZONTAL walls (their Y faces); X-anchor evidence
comes from VERTICAL walls (their X faces). Everything below is written once and
parameterised by `axis`.

Hierarchy (design doc s2.3), in priority order:
  T1 Boundary   - a face on the plan's global extreme IS the perimeter grid.
  T2 Single     - exactly one face lies on a wall control line -> that face.
  T3 Conflict   - both faces on lines -> decisive support margin wins, else defer.
  T4 No evidence- neither face on a line -> defer.
Where the engine "defers" it returns None and the existing pipeline tag stands
(family + wall-assist + lift + consensus), so no column is ever left blank and
nothing regresses. Lift columns are never overridden (they keep their dedicated
lift-centre rule).
"""
from __future__ import annotations

from dataclasses import dataclass


# ---- parameters (design doc s2.5; calibrated starting values) ----
TAU_FACE = 0.06       # m, column face -> control-line incidence tolerance
TAU_CLUSTER = 0.06    # m, 1-D gap for clustering wall faces into a control line
THETA_MARGIN = 0.25   # fractional support gap needed to settle a both-faces conflict
BOUNDARY_EPS = 0.01   # m, equality to the plan's global extreme
STUB_LENGTH_M = 0.60  # m, walls shorter than this are discounted (fragments, not beams)
STUB_FACTOR = 0.25    # multiplier applied to sub-stub-length walls
KAPPA = {230: 1.0, 115: 0.5}  # thickness weight: 230 = beam-bearing, 115 = partition


@dataclass(frozen=True)
class ControlLine:
    coord: float      # support-weighted mean coordinate of the line
    support: float    # total wall weight backing the line
    ext_lo: float     # physical extent of the line along the perpendicular axis
    ext_hi: float


def _thickness_factor(t_m: float) -> float:
    t_mm = t_m * 1000.0
    return KAPPA[min(KAPPA, key=lambda k: abs(k - t_mm))]


def _length_factor(length_m: float) -> float:
    f = min(length_m, 3.0)
    if length_m < STUB_LENGTH_M:
        f *= STUB_FACTOR
    return f


def _aligned_walls(local_walls, axis):
    # Y-anchor is set by horizontal beams/walls; X-anchor by vertical ones.
    want = "Horizontal" if axis == "Y" else "Vertical"
    return [w for w in local_walls if w.orientation == want]


def _wall_faces(w, axis):
    """Return (face_lo, face_hi, length, ext_lo, ext_hi, thickness) for a wall.

    For axis Y: faces are the wall's Ymin/Ymax, length/extent are along X.
    For axis X: faces are the wall's Xmin/Xmax, length/extent are along Y.
    """
    if axis == "Y":
        return w.ymin_m, w.ymax_m, (w.xmax_m - w.xmin_m), w.xmin_m, w.xmax_m, (w.ymax_m - w.ymin_m)
    return w.xmin_m, w.xmax_m, (w.ymax_m - w.ymin_m), w.ymin_m, w.ymax_m, (w.xmax_m - w.xmin_m)


def build_control_lines(local_walls, axis):
    """Reconstruct control lines: cluster all aligned-wall faces (1-D) into
    support-weighted lines. Support = sum of length*thickness weights."""
    faces = []  # (coord, weight, ext_lo, ext_hi)
    for w in _aligned_walls(local_walls, axis):
        lo, hi, length, elo, ehi, th = _wall_faces(w, axis)
        wt = _length_factor(length) * _thickness_factor(th)
        faces.append((lo, wt, elo, ehi))
        faces.append((hi, wt, elo, ehi))
    faces.sort(key=lambda f: f[0])

    lines = []
    cluster = []
    for f in faces:
        if cluster and f[0] - cluster[-1][0] > TAU_CLUSTER:
            lines.append(_finalize_cluster(cluster))
            cluster = []
        cluster.append(f)
    if cluster:
        lines.append(_finalize_cluster(cluster))
    return lines


def _finalize_cluster(cluster):
    sw = sum(w for _, w, _, _ in cluster)
    if sw > 0:
        coord = sum(c * w for c, w, _, _ in cluster) / sw
    else:
        coord = sum(c for c, _, _, _ in cluster) / len(cluster)
    return ControlLine(
        coord=round(coord, 4),
        support=round(sw, 4),
        ext_lo=min(e for _, _, e, _ in cluster),
        ext_hi=max(e for _, _, _, e in cluster),
    )


def _evidence(face_value, lines):
    """Strongest control line within TAU_FACE of the face; evidence weighted by
    the line's support and the closeness of the match."""
    best = None
    for L in lines:
        d = abs(face_value - L.coord)
        if d <= TAU_FACE:
            ev = L.support * (1.0 - d / TAU_FACE)
            if best is None or ev > best[0]:
                best = (ev, L, d)
    return best


def resolve_axis(axis, local_rects, local_walls, g_lo, g_hi):
    """Return {idx: (tag_or_None, tier, ev_lo, ev_hi)} for one axis.
    tag None means 'defer to the existing pipeline tag'."""
    lines = build_control_lines(local_walls, axis)
    lo_label, hi_label = ("Front", "Back") if axis == "Y" else ("Left", "Right")
    out = {}
    for r in local_rects:
        f_lo, f_hi = (r.ymin_m, r.ymax_m) if axis == "Y" else (r.xmin_m, r.xmax_m)

        # T1 - boundary: perimeter face is the grid, deterministic.
        if abs(f_lo - g_lo) <= BOUNDARY_EPS:
            out[r.idx] = (lo_label, "T1-boundary", None, None)
            continue
        if abs(f_hi - g_hi) <= BOUNDARY_EPS:
            out[r.idx] = (hi_label, "T1-boundary", None, None)
            continue

        elo = _evidence(f_lo, lines)
        ehi = _evidence(f_hi, lines)
        vlo = elo[0] if elo else 0.0
        vhi = ehi[0] if ehi else 0.0

        # T2 - exactly one face on a control line.
        if vlo > 0 and vhi == 0:
            out[r.idx] = (lo_label, "T2-single", vlo, vhi)
            continue
        if vhi > 0 and vlo == 0:
            out[r.idx] = (hi_label, "T2-single", vlo, vhi)
            continue

        # T3 - both faces on lines: settle only on a decisive support margin.
        if vlo > 0 and vhi > 0:
            m = THETA_MARGIN * max(vlo, vhi)
            if vhi - vlo >= m:
                out[r.idx] = (hi_label, "T3-margin", vlo, vhi)
            elif vlo - vhi >= m:
                out[r.idx] = (lo_label, "T3-margin", vlo, vhi)
            else:
                out[r.idx] = (None, "T3-defer", vlo, vhi)
            continue

        # T4 - no wall evidence either face: defer to existing tag.
        out[r.idx] = (None, "T4-defer", vlo, vhi)
    return out


ENABLED_AXES = ("Y", "X")  # Front/Back + Left/Right both active (promoted 2026-07-01)


def apply_wall_anchor_engine(tags, local_rects, local_walls, lift_rect_ids, axes=ENABLED_AXES):
    """Override the anchor tag on each enabled axis where the engine is
    confident; defer otherwise. Never touches lift columns. Returns
    (new_tags, report) where report[idx][axis] = (tag, tier, ev_lo, ev_hi)."""
    g_minx = min(r.xmin_m for r in local_rects)
    g_maxx = max(r.xmax_m for r in local_rects)
    g_miny = min(r.ymin_m for r in local_rects)
    g_maxy = max(r.ymax_m for r in local_rects)

    result = dict(tags)
    report = {}
    for axis in axes:
        g_lo, g_hi = (g_miny, g_maxy) if axis == "Y" else (g_minx, g_maxx)
        decisions = resolve_axis(axis, local_rects, local_walls, g_lo, g_hi)
        for r in local_rects:
            tag, tier, vlo, vhi = decisions[r.idx]
            report.setdefault(r.idx, {})[axis] = (tag, tier, vlo, vhi)
            if r.idx in lift_rect_ids:
                continue          # lift columns keep their dedicated rule
            if tag is None:
                continue          # defer: existing pipeline tag stands
            x, y, loc = result[r.idx]
            if axis == "Y":
                result[r.idx] = (x, tag, loc)
            else:
                result[r.idx] = (tag, y, loc)
    return result, report
