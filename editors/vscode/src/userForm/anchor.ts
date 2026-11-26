// The `anchor` of a control (RAS language specification §14.11.16): the edges of its
// container it keeps its distance to when the container is resized, as the Anchor
// property of Windows Forms. The same rules are in ras-form (anchor.rs, the VBA the form
// runs on UserForm_Resize) and in the debug run (ras-compiler/src/js/forms.js).
//
// * left and right: the control is stretched;
// * right only: it moves with the right edge;
// * neither: it stays centred (moves by half the change);
// * left only (the default): it stays where it is.
//
// The same for top and bottom. A Frame or MultiPage that is stretched passes the change on
// to the controls in it.

export interface Anchor {
  top: boolean;
  bottom: boolean;
  left: boolean;
  right: boolean;
}

export const anchorSides = ["top", "bottom", "left", "right"] as const;
export type AnchorSide = (typeof anchorSides)[number];
export const defaultAnchor: Anchor = { top: true, bottom: false, left: true, right: false };
/** Control types whose edges cannot be anchored (a Page fills its MultiPage). */
export const unanchoredTypes = ["page", "progressbar"];

/**
 * Parses an `anchor` value: `"top,left"`, `"Top, Bottom, Left, Right"`, `"none"`. Absent
 * (`undefined` / `null`) is the default; anything else that is not such a list is undefined.
 */
export function parseAnchor(value: unknown): Anchor | undefined {
  if (value === undefined || value === null) return { ...defaultAnchor };
  if (typeof value !== "string") return undefined;
  const words = value
    .split(/[\s,|]+/)
    .map((word) => word.trim().toLowerCase())
    .filter(Boolean);
  if (words.length === 1 && words[0] === "none") return { top: false, bottom: false, left: false, right: false };
  if (!words.length) return undefined;
  const anchor: Anchor = { top: false, bottom: false, left: false, right: false };
  for (const word of words) {
    if (!(anchorSides as readonly string[]).includes(word)) return undefined;
    if (anchor[word as AnchorSide]) return undefined;
    anchor[word as AnchorSide] = true;
  }
  return anchor;
}

/** The canonical text of an anchor (`"top,left"`, `"none"`). */
export function formatAnchor(anchor: Anchor): string {
  const sides = anchorSides.filter((side) => anchor[side]);
  return sides.length ? sides.join(",") : "none";
}

export const isDefaultAnchor = (anchor: Anchor) =>
  anchorSides.every((side) => anchor[side] === defaultAnchor[side]);

/** The VBA-style label of an anchor (`Top, Left`, `None`). */
export function anchorLabel(anchor: Anchor): string {
  const sides = anchorSides.filter((side) => anchor[side]);
  return sides.length ? sides.map((side) => side[0].toUpperCase() + side.slice(1)).join(", ") : "None";
}

/** Start and size along one axis after the container grew by `delta`. */
export function anchorAxis(
  near: boolean,
  far: boolean,
  start: number,
  size: number,
  delta: number,
  minimum = 0,
): [number, number] {
  if (delta === 0) return [start, size];
  if (near && far) return [start, Math.max(minimum, size + delta)];
  if (far) return [start + delta, size];
  if (!near) return [start + delta / 2, size];
  return [start, size];
}

export interface AnchorItem {
  id: string;
  parentId?: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  anchor?: unknown;
}
export interface AnchoredGeometry {
  left: number;
  top: number;
  width: number;
  height: number;
}

const round = (value: number) => Math.round(value * 1e6) / 1e6;

/**
 * Where the controls in the container `containerId` (`""` for the form) go when its inside
 * grows by `dw` × `dh` points (negative shrinks it). Only the controls that change are in the
 * result. Sizes stay at `minimum` or more; positions stay at 0 or more.
 */
export function anchorLayout(
  items: readonly AnchorItem[],
  containerId: string,
  dw: number,
  dh: number,
  minimum = 1,
): Map<string, AnchoredGeometry> {
  const result = new Map<string, AnchoredGeometry>();
  const visit = (parentId: string, dx: number, dy: number, seen: Set<string>) => {
    if ((dx === 0 && dy === 0) || seen.has(parentId)) return;
    seen.add(parentId);
    for (const item of items) {
      if ((item.parentId ?? "") !== parentId) continue;
      if (item.type.toLowerCase() === "page") {
        // A Page is as large as its MultiPage's inside.
        visit(item.id, dx, dy, seen);
        continue;
      }
      if (unanchoredTypes.includes(item.type.toLowerCase())) continue;
      const anchor = parseAnchor(item.anchor) ?? defaultAnchor;
      let [left, width] = anchorAxis(anchor.left, anchor.right, item.left, item.width, dx, minimum);
      let [top, height] = anchorAxis(anchor.top, anchor.bottom, item.top, item.height, dy, minimum);
      left = round(Math.max(0, left));
      top = round(Math.max(0, top));
      width = round(width);
      height = round(height);
      if (left !== item.left || top !== item.top || width !== item.width || height !== item.height)
        result.set(item.id, { left, top, width, height });
      visit(item.id, round(width - item.width), round(height - item.height), seen);
    }
  };
  visit(containerId, dw, dh, new Set());
  return result;
}

/**
 * How far the container `containerId` can shrink before an anchored control in it would
 * get smaller than `minimum` or go past an edge it is not anchored to: the least inside
 * size, given the current one.
 */
export function anchorMinimumSize(
  items: readonly AnchorItem[],
  containerId: string,
  width: number,
  height: number,
  minimum = 1,
): { width: number; height: number } {
  const least = { width: 1, height: 1 };
  const children = items.filter((item) => (item.parentId ?? "") === containerId);
  for (const item of children) {
    const type = item.type.toLowerCase();
    if (type === "page") {
      // The controls on a Page limit their MultiPage (its inside is the Page's size).
      const inner = anchorMinimumSize(items, item.id, width, height, minimum);
      least.width = Math.max(least.width, inner.width);
      least.height = Math.max(least.height, inner.height);
      continue;
    }
    const anchor = unanchoredTypes.includes(type) ? defaultAnchor : (parseAnchor(item.anchor) ?? defaultAnchor);
    // A stretched Frame or MultiPage cannot get smaller than what is anchored in it.
    const own = { width: minimum, height: minimum };
    if (type === "frame" || type === "multipage") {
      const [x, y] = type === "frame" ? [1.5, 1.5] : [3, 19.5];
      const inner = anchorMinimumSize(items, item.id, Math.max(0, item.width - x), Math.max(0, item.height - y), minimum);
      own.width = Math.max(minimum, inner.width + x);
      own.height = Math.max(minimum, inner.height + y);
    }
    const axis = (near: boolean, far: boolean, start: number, size: number, least: number, outer: number) => {
      // Left only: the right edge has to stay inside. Stretched: the size has to stay. Right
      // only: the left edge has to stay at 0 or more. Centred: both.
      if (near && far) return outer - (size - Math.min(size, least));
      if (far) return outer - start;
      if (!near) return outer - 2 * start;
      return start + size;
    };
    least.width = Math.max(least.width, axis(anchor.left, anchor.right, item.left, item.width, own.width, width));
    least.height = Math.max(least.height, axis(anchor.top, anchor.bottom, item.top, item.height, own.height, height));
  }
  return { width: round(least.width), height: round(least.height) };
}
