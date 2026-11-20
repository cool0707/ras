/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
import type { DesignerDocument, DesignerProblem, PreviewDocument } from "./protocol";
import { controlDefaults } from "./toolbox";

export function pointsToPixels(points: number, zoom = 1): number {
  return (points * 96 * zoom) / 72;
}

export function pixelsToPoints(pixels: number, zoom = 1): number {
  return (pixels * 72) / (96 * zoom);
}

export function designerDocument(
  spec: PreviewDocument,
  warnings: DesignerProblem[] = [],
  eventEditable = false,
): DesignerDocument {
  const form = spec.form;
  const width = form.width;
  const height = form.height;
  return {
    font: form.font,
    foreColor: form.foreColor,
    backColor: form.backColor,
    borderStyle: form.borderStyle,
    name: form.name,
    caption: form.caption ?? form.name,
    width: width ?? 240,
    height: height ?? 180,
    approximate: width == null || height == null,
    eventEditable,
    warnings: warnings.map((warning) => ({ ...warning })),
    controls: spec.controls.map((control) => {
      const effective = control;
      const defaults = controlDefaults(control.type);
      return {
        ...effective,
        left: effective.left ?? 0,
        top: effective.top ?? 0,
        width: effective.width ?? defaults[0],
        height: effective.height ?? defaults[1],
        // RAS: Page bounds always derive from the MultiPage; that is not an approximation.
        approximate:
          control.type.toLowerCase() !== "page" &&
          (effective.width === undefined || effective.height === undefined),
      };
    }),
  };
}

/**
 * The state the Value of a CheckBox, OptionButton or ToggleButton stands for: true, 1, -1 and
 * "True" are on, "Null" is neither (TripleState), anything else is off.
 */
export function toggleState(value: unknown): "on" | "off" | "null" {
  if (value === true || value === 1 || value === -1) return "on";
  if (typeof value === "string") {
    const text = value.trim().toLowerCase();
    if (text === "true" || text === "1" || text === "-1") return "on";
    if (text === "null") return "null";
  }
  return "off";
}
