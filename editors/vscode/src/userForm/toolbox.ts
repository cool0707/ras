/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
import type { ColorValue, DesignerDocument, FontValue, SpecControl } from "./protocol";
import { gridSize, isDesignerFrame, isDesignerPage, parentBounds, stableDelta } from "./geometry";

// Shared by rendering fallback and Toolbox creation. Dimensions are points,
// matching the pure-Go compiler's built-in generation defaults.
export const controlMetadata = [
  { type: "Label", width: 72, height: 18, caption: true },
  { type: "TextBox", width: 120, height: 18 },
  { type: "ComboBox", width: 120, height: 18 },
  { type: "ListBox", width: 120, height: 72 },
  { type: "CommandButton", width: 72, height: 24, caption: true },
  { type: "CheckBox", width: 72, height: 18, caption: true },
  { type: "OptionButton", width: 72, height: 18, caption: true },
  { type: "ToggleButton", width: 72, height: 18, caption: true },
  { type: "SpinButton", width: 18, height: 36 },
  { type: "ScrollBar", width: 120, height: 18 },
  { type: "Image", width: 72, height: 72 },
  { type: "Frame", width: 144, height: 108, caption: true },
  { type: "MultiPage", width: 240, height: 180 },
  { type: "TabStrip", width: 240, height: 48 },
  // RAS's own: written as four labels (spec §14.11.13).
  { type: "ProgressBar", width: 200, height: 24, caption: true },
  { type: "FilePicker", width: 200, height: 18 },
] as const;

export type ToolboxType = (typeof controlMetadata)[number]["type"];
export type NewControl = Pick<
  SpecControl,
  "id" | "parentId" | "name" | "caption" | "selectedIndex" | "tabs" | "backStyle" | "backColor" | "font"
> & {
  type: ToolboxType;
  left: number;
  top: number;
  width: number;
  height: number;
};

// Controls whose background is the container's surface. Created transparent
// (fmBackStyleTransparent) so they follow the form or frame BackColor.
export const transparentOnCreate: ReadonlySet<ToolboxType> = new Set([
  "Label",
  "CheckBox",
  "OptionButton",
]);

// Containers with no transparent BackStyle in MSForms. As in the VBA editor, a new
// one takes the BackColor of the form, frame or page it is placed on.
export const containerColorOnCreate: ReadonlySet<ToolboxType> = new Set([
  "Frame",
  "MultiPage",
  "TabStrip",
]);

/** The BackColor shown behind a control placed in `parentId` (a Page shows its MultiPage's). */
export function containerBackColor(
  document: DesignerDocument,
  parentId?: string,
): ColorValue | undefined {
  let parent = document.controls.find((c) => c.id === parentId);
  if (parent?.type.toLowerCase() === "page")
    parent = document.controls.find((c) => c.id === parent!.parentId);
  return (parentId ? parent?.backColor : document.backColor) ?? undefined;
}

/** The size new forms and controls are made with, in points. */
export const newFontSize = 10;
/** The UI font of Japanese Windows 10 and later, when the system's cannot be looked up. */
export const fallbackNewFont: Readonly<FontValue> = { name: "Yu Gothic UI", size: newFontSize };
/** Controls without a Font property. */
const fontless: ReadonlySet<ToolboxType> = new Set(["SpinButton", "ScrollBar", "Image"]);

export function controlDefaults(type: string): readonly [number, number] {
  const metadata = controlMetadata.find((m) => m.type.toLowerCase() === type.toLowerCase());
  return metadata
    ? [metadata.width, metadata.height]
    : type.toLowerCase() === "page"
      ? [0, 0]
      : [72, 24];
}

export function createControl(
  document: DesignerDocument,
  type: ToolboxType,
  left: number,
  top: number,
  snapping: boolean,
  uuid: () => string = () => crypto.randomUUID(),
  parentId?: string,
  pageEditable = false,
  size?: { width: number; height: number },
): NewControl | undefined {
  const found = controlMetadata.find((m) => m.type === type);
  // A control drawn by dragging on the form takes the dragged size (as in the VBA editor);
  // a click makes it with the default size.
  const metadata =
    found && size && size.width > 0 && size.height > 0
      ? { ...found, width: size.width, height: size.height }
      : found;
  const parent = document.controls.find((c) => c.id === parentId);
  if (
    parentId &&
    (!parent || (!isDesignerFrame(parent) && !(pageEditable && isDesignerPage(document, parent))))
  )
    return;
  const bounds = parent ? parentBounds(document, { ...parent, parentId: parent.id }) : document;
  if (
    !metadata ||
    !Number.isFinite(left) ||
    !Number.isFinite(top) ||
    metadata.width > bounds.width ||
    metadata.height > bounds.height
  )
    return;
  const names = new Set(
    [document.name, ...document.controls.map((c) => c.name)].map((n) => n.toLowerCase()),
  );
  let suffix = 1;
  while (names.has(`${type}${suffix}`.toLowerCase())) suffix++;
  const ids = new Set(document.controls.map((c) => c.id));
  let id: string;
  // Fail safely if the random source is broken rather than spin forever.
  for (let attempt = 0; attempt < 100; attempt++) {
    id = `control-${uuid()}`;
    if (ids.has(id)) continue;
    const position = (value: number, maximum: number) =>
      Math.max(
        0,
        Math.min(maximum, snapping ? Math.round(value / gridSize) * gridSize : stableDelta(value)),
      );
    const control: NewControl = {
      id,
      name: `${type}${suffix}`,
      type,
      left: position(left, bounds.width - metadata.width),
      top: position(top, bounds.height - metadata.height),
      width: metadata.width,
      height: metadata.height,
    };
    if (parentId) control.parentId = parentId;
    if ("caption" in metadata && metadata.caption) control.caption = control.name;
    if (transparentOnCreate.has(type)) control.backStyle = 0;
    if (containerColorOnCreate.has(type)) {
      const backColor = containerBackColor(document, parentId);
      if (backColor) control.backColor = { kind: backColor.kind, value: backColor.value };
    }
    // As in the VBA editor, a new control takes the form's font; a form without one
    // gives the system UI font instead of MSForms' MS UI Gothic 9pt.
    if (!fontless.has(type)) {
      const font = document.font ?? document.newControlFont ?? fallbackNewFont;
      control.font = Object.fromEntries(
        Object.entries(font).filter(([, v]) => v !== null && v !== undefined),
      ) as FontValue;
    }
    if (type === "MultiPage" || type === "TabStrip") control.selectedIndex = -1;
    if (type === "TabStrip") {
      control.tabs = [
        { name: "Tab1", caption: "Tab1" },
        { name: "Tab2", caption: "Tab2" },
      ];
      control.selectedIndex = 0;
    }
    return control;
  }
  return undefined;
}

// Preview controls are canonical and flattened, including legacy nested input.
export function subtreeIds(
  controls: readonly Pick<SpecControl, "id" | "parentId">[],
  id: string,
): Set<string> {
  const result = new Set([id]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const control of controls) {
      if (control.parentId && result.has(control.parentId) && !result.has(control.id)) {
        result.add(control.id);
        changed = true;
      }
    }
  }
  return result;
}
