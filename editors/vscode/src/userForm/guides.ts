// Smart guides for moving controls in the Designer, as in Figma: a control being dragged
// snaps to the edges and centres of its siblings and of its container, to the gap that is
// already between siblings in the same row or column (equal spacing), and to 8 / 16 / 24
// from the nearest sibling or the container's edge. Points, in the container's coordinates.

import type { Bounds, Geometry } from "./geometry";

/** The distances from a neighbour or the container's edge that a control snaps to. */
export const guideDistances = [8, 16, 24];

/**
 * What the canvas draws while a control is being moved. `axis` is the axis the guide
 * measures along: an "x" line is vertical (at x = `at`, from y = `from` to `to`); an "x" gap
 * is horizontal (from x = `from` to `to`, at y = `at`). Y is the same, turned.
 */
export type Guide =
  | { kind: "line"; axis: "x" | "y"; at: number; from: number; to: number }
  | { kind: "gap"; axis: "x" | "y"; at: number; from: number; to: number; label: string };

export interface GuideResult {
  left: number;
  top: number;
  snappedX: boolean;
  snappedY: boolean;
}

type Axis = "x" | "y";
const lo = (r: Geometry, a: Axis) => (a === "x" ? r.left : r.top);
const size = (r: Geometry, a: Axis) => (a === "x" ? r.width : r.height);
const hi = (r: Geometry, a: Axis) => lo(r, a) + size(r, a);
const extent = (b: Bounds, a: Axis) => (a === "x" ? b.width : b.height);
const cross = (a: Axis): Axis => (a === "x" ? "y" : "x");
const eps = 0.01;
const near = (a: number, b: number) => Math.abs(a - b) < eps;
/** Whether two rectangles share some of the span of `a`. */
const overlaps = (r: Geometry, s: Geometry, a: Axis) => lo(r, a) < hi(s, a) - eps && lo(s, a) < hi(r, a) - eps;
const round = (value: number) => Math.round(value * 1e6) / 1e6;
const label = (value: number) => String(Math.round(value * 100) / 100);

/** Nothing of `row` is in the open span between `from` and `to` (along `a`), but `except`. */
const clear = (row: Geometry[], a: Axis, from: number, to: number, except: Geometry[]) =>
  !row.some((q) => !except.includes(q) && lo(q, a) < to - eps && hi(q, a) > from + eps);

/** The gaps between neighbours of `row` along `a`: pairs side by side with nothing between. */
function gaps(row: Geometry[], a: Axis): { before: Geometry; after: Geometry; gap: number }[] {
  const found: { before: Geometry; after: Geometry; gap: number }[] = [];
  for (const p of row)
    for (const q of row) {
      if (p === q || hi(p, a) > lo(q, a) + eps || !overlaps(p, q, cross(a))) continue;
      if (!clear(row, a, hi(p, a), lo(q, a), [p, q])) continue;
      const gap = round(lo(q, a) - hi(p, a));
      if (gap > eps) found.push({ before: p, after: q, gap });
    }
  return found;
}

/** Where a control could go along `a`, closest first; undefined when nothing is near. */
function snapAxis(moving: Geometry, others: Geometry[], bounds: Bounds, a: Axis, threshold: number) {
  const s = size(moving, a);
  const limit = extent(bounds, a) - s;
  const raw = lo(moving, a);
  const row = others.filter((o) => overlaps(o, moving, cross(a)));
  const candidates: { value: number; priority: number }[] = [];
  const add = (value: number, priority: number) => candidates.push({ value: round(value), priority });
  // Edges and centres.
  for (const edge of [0, extent(bounds, a) / 2, extent(bounds, a)])
    for (const offset of [0, s / 2, s]) add(edge - offset, 0);
  for (const o of others)
    for (const edge of [lo(o, a), lo(o, a) + size(o, a) / 2, hi(o, a)])
      for (const offset of [0, s / 2, s]) add(edge - offset, 0);
  // The gaps there already are, and the fixed distances, from a neighbour in the row.
  const spacing = [...new Set(gaps(row, a).map((g) => g.gap))];
  for (const o of row) {
    for (const [gap, priority] of [...spacing.map((g) => [g, 1]), ...guideDistances.map((d) => [d, 2])]) {
      const after = hi(o, a) + gap;
      if (clear(row, a, hi(o, a), after + s, [o])) add(after, priority);
      const before = lo(o, a) - gap - s;
      if (clear(row, a, before, lo(o, a), [o])) add(before, priority);
    }
    // Midway between two neighbours: the same gap on each side.
    for (const p of row)
      if (p !== o && hi(o, a) + s <= lo(p, a) + eps && clear(row, a, hi(o, a), lo(p, a), [o, p]))
        add((hi(o, a) + lo(p, a) - s) / 2, 1);
  }
  for (const d of guideDistances) {
    if (clear(row, a, 0, d + s, [])) add(d, 2);
    if (clear(row, a, extent(bounds, a) - d - s, extent(bounds, a), [])) add(extent(bounds, a) - d - s, 2);
  }
  let best: { value: number; priority: number } | undefined;
  for (const c of candidates) {
    if (c.value < -eps || c.value > limit + eps) continue;
    const distance = Math.abs(c.value - raw);
    if (distance > threshold) continue;
    const bestDistance = best ? Math.abs(best.value - raw) : Infinity;
    if (!best || distance < bestDistance - eps || (near(distance, bestDistance) && c.priority < best.priority))
      best = c;
  }
  return best && Math.max(0, Math.min(limit, best.value));
}

/**
 * Where a control dragged to `moving` snaps to, within `threshold` of where it is: each axis
 * on its own. `others` are its siblings (not the controls being moved).
 */
export function snapToGuides(moving: Geometry, others: Geometry[], bounds: Bounds, threshold: number): GuideResult {
  const x = snapAxis(moving, others, bounds, "x", threshold);
  const y = snapAxis(moving, others, bounds, "y", threshold);
  return { left: x ?? moving.left, top: y ?? moving.top, snappedX: x !== undefined, snappedY: y !== undefined };
}

/** The guides that hold for a control at `moving`: what it lines up with and the gaps it matches. */
export function describeGuides(moving: Geometry, others: Geometry[], bounds: Bounds): Guide[] {
  const guides: Guide[] = [];
  const keys = new Set<string>();
  const push = (guide: Guide) => {
    const key = `${guide.kind}:${guide.axis}:${label(guide.at)}:${label(guide.from)}:${label(guide.to)}`;
    if (!keys.has(key)) {
      keys.add(key);
      guides.push(guide);
    }
  };
  const lines = new Map<string, { axis: Axis; at: number; from: number; to: number }>();
  const line = (axis: Axis, at: number, from: number, to: number) => {
    const key = `${axis}:${label(at)}`;
    const old = lines.get(key);
    lines.set(key, old ? { ...old, from: Math.min(old.from, from), to: Math.max(old.to, to) } : { axis, at, from, to });
  };
  for (const a of ["x", "y"] as const) {
    const c = cross(a);
    const edges = (r: Geometry) => [lo(r, a), lo(r, a) + size(r, a) / 2, hi(r, a)];
    const mine = edges(moving);
    for (const o of others) {
      for (const edge of edges(o))
        if (mine.some((m) => near(m, edge)))
          line(a, edge, Math.min(lo(o, c), lo(moving, c)), Math.max(hi(o, c), hi(moving, c)));
    }
    const whole = extent(bounds, a);
    for (const edge of [whole / 2]) if (mine.some((m) => near(m, edge))) line(a, edge, 0, extent(bounds, c));

    // Gaps along `a`, drawn across the middle of what the two have in common.
    const row = others.filter((o) => overlaps(o, moving, c));
    const middle = (r: Geometry, s?: Geometry) => {
      const from = s ? Math.max(lo(r, c), lo(s, c)) : lo(r, c);
      const to = s ? Math.min(hi(r, c), hi(s, c)) : hi(r, c);
      return (from + to) / 2;
    };
    const gap = (from: number, to: number, at: number) =>
      push({ kind: "gap", axis: a, at, from, to, label: label(to - from) });
    const before = row
      .filter((o) => hi(o, a) <= lo(moving, a) + eps && clear(row, a, hi(o, a), lo(moving, a), [o]))
      .sort((p, q) => hi(q, a) - hi(p, a))[0];
    const after = row
      .filter((o) => lo(o, a) >= hi(moving, a) - eps && clear(row, a, hi(moving, a), lo(o, a), [o]))
      .sort((p, q) => lo(p, a) - lo(q, a))[0];
    const existing = gaps(row, a);
    const sides: { from: number; to: number; at: number }[] = [];
    if (before) sides.push({ from: hi(before, a), to: lo(moving, a), at: middle(moving, before) });
    if (after) sides.push({ from: hi(moving, a), to: lo(after, a), at: middle(moving, after) });
    const edgeSides: { from: number; to: number; at: number }[] = [];
    if (!before && clear(row, a, 0, lo(moving, a), []))
      edgeSides.push({ from: 0, to: lo(moving, a), at: middle(moving) });
    if (!after && clear(row, a, hi(moving, a), whole, []))
      edgeSides.push({ from: hi(moving, a), to: whole, at: middle(moving) });
    for (const side of sides) {
      const length = round(side.to - side.from);
      if (length <= eps) continue;
      const same = existing.filter((g) => near(g.gap, length));
      const other = sides.find((s) => s !== side && near(round(s.to - s.from), length));
      if (same.length || other || guideDistances.some((d) => near(d, length))) gap(side.from, side.to, side.at);
      for (const g of same) gap(hi(g.before, a), lo(g.after, a), middle(g.before, g.after));
    }
    for (const side of edgeSides) {
      const length = round(side.to - side.from);
      if (guideDistances.some((d) => near(d, length))) gap(side.from, side.to, side.at);
    }
  }
  for (const l of lines.values()) push({ kind: "line", ...l });
  return guides;
}
