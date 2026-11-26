/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
import type { DesignerControl, DesignerDocument } from "../../src/userForm/protocol";
import { pointsToPixels } from "../../src/userForm/model";
import { pixelsToPoints, toggleState } from "../../src/userForm/model";
import { columnLayout, itemColumns, textColumnIndex } from "../../src/userForm/listItems";
import { backgroundPicture, picturePositionLayout } from "../../src/userForm/pictureLayout";
import { useEffect, useMemo, useRef, useState } from "preact/hooks";
import { displayPages } from "../../src/userForm/collections";
import type { ActivePages } from "../../src/userForm/collections";
import type { JSX } from "preact";
import type {
  GeometryOperation,
  ParentOperation,
  PropertyOperation,
  ProductivityOperation,
} from "../../src/userForm/protocol";
import {
  fullyEnclosedChildren,
  moveSiblingGroup,
  promoteSelection,
  selectedWorldBounds,
  siblingSelection,
  updateSelection,
} from "../../src/userForm/productivity";
import type { SelectionRect } from "../../src/userForm/productivity";
import { subtreeIds } from "../../src/userForm/toolbox";
import { progressBarLayout } from "../../src/userForm/progressBar";
import {
  canEditGeometry,
  isControlVisible,
  isDesignerFrame,
  isDesignerPage,
  contentInsets,
  geometryOperations,
  gridSize,
  moveGeometry,
  parentBounds,
  requiredSize,
  resizeGeometry,
  resizeHandles,
  dropTarget,
  placeGeometry,
  placementRect,
  worldGeometry,
  orderedChildren,
  selectionPath,
  stableDelta,
} from "../../src/userForm/geometry";
import type { Geometry, ResizeHandle, DropTarget } from "../../src/userForm/geometry";
import { appearanceStyle, borderClass, colorStyle } from "./propertyEditors/appearance";
const px = (points: number) => `${pointsToPixels(points)}px`;

/** The picture of an Image, or a placeholder when there is none that can be drawn. */
function ControlPicture({ control }: { control: DesignerControl }) {
  // A picture the browser cannot draw must not leave a broken-image icon.
  const [broken, setBroken] = useState<string>();
  const source = control.picture?.remove === true ? undefined : control.picturePreview;
  if (source && source !== broken)
    return (
      <img
        class="control-picture"
        src={source}
        alt=""
        draggable={false}
        onError={() => setBroken(source)}
        style={{
          // Transparent pixels are blended onto BackColor (buttonFace by default).
          backgroundColor: colorStyle(control.backColor) ?? "#f0f0f0",
          objectFit:
            control.pictureSizeMode === 1
              ? "fill"
              : control.pictureSizeMode === 3
                ? "contain"
                : "none",
          objectPosition:
            ["left top", "right top", "center", "left bottom", "right bottom"][
              control.pictureAlignment ?? 2
            ] ?? "center",
        }}
      />
    );
  // A picture that is set but not shown is named, so that it is not mistaken for none.
  const path = control.picture?.remove === true ? undefined : control.picture?.path;
  const name = typeof path === "string" ? path.split("/").pop() : undefined;
  return <span class="placeholder">▧ {name || "Image"}</span>;
}

/** A caption with the control's picture beside it, as PicturePosition places it. */
function PictureCaption({ control, caption }: { control: DesignerControl; caption: string }) {
  const source = control.picture?.remove === true ? undefined : control.picturePreview;
  if (!source) return <span>{caption}</span>;
  const layout = picturePositionLayout(control.picturePosition);
  return (
    <span
      class={`picture-caption ${layout.direction}`}
      style={layout.direction === "stack" ? undefined : { flexDirection: layout.direction, alignItems: layout.align }}
    >
      <img class="caption-picture" src={source} alt="" draggable={false} />
      {caption && <span class="caption-text">{caption}</span>}
    </span>
  );
}

/** Orientation Horizontal (1) or Vertical (0), else Auto: horizontal when wider than high. */
const horizontal = (control: DesignerControl) =>
  control.orientation === 1 || (control.orientation !== 0 && control.width >= control.height);

/** Classes for the plain properties that change how a control is drawn. */
export function scalarClasses(control: DesignerControl): string {
  return [
    typeof control.specialEffect === "number" ? `effect-${control.specialEffect}` : "",
    control.backStyle === 0 ? "transparent" : "",
    control.alignment === 0 ? "caption-left" : "",
    control.wordWrap === false ? "no-wrap" : control.wordWrap === true ? "wrap" : "",
    control.multiLine ? "multi-line" : "",
    typeof control.style === "number" && ["multipage", "tabstrip"].includes(control.type.toLowerCase())
      ? `tab-style-${control.style}`
      : "",
    // TabOrientation Bottom, Left or Right (Top is the default drawing).
    typeof control.tabOrientation === "number" &&
    control.tabOrientation > 0 &&
    ["multipage", "tabstrip"].includes(control.type.toLowerCase())
      ? `tab-orient-${control.tabOrientation}`
      : "",
  ]
    .filter(Boolean)
    .join(" ");
}

export function ControlGlyph({ control }: { control: DesignerControl }) {
  const type = control.type.toLowerCase();
  const caption = control.caption ?? control.name;
  switch (type) {
    case "label":
    case "commandbutton":
    case "togglebutton":
      return <PictureCaption control={control} caption={caption} />;
    case "textbox": {
      const text = control.text ?? String(control.value ?? "");
      // PasswordChar shows in place of every character.
      return <span>{control.passwordChar ? control.passwordChar.repeat([...text].length) : text}</span>;
    }
    case "checkbox":
    case "optionbutton": {
      // MSForms draws the box at the left, vertically centred, with the caption beside it.
      const state = { on: "checked", null: "mixed", off: "" }[toggleState(control.value)];
      return (
        <>
          <span class={`toggle-box ${type === "checkbox" ? "check" : "option"} ${state}`} aria-hidden="true" />
          <span class="toggle-caption">
            <PictureCaption control={control} caption={caption} />
          </span>
        </>
      );
    }
    case "combobox": {
      // Without an authored text or value, the selected list item is what the form shows.
      const authored = control.text || (control.value == null ? "" : String(control.value));
      // The selected item shows its TextColumn.
      const rows = (control.list ?? []).map(itemColumns);
      const columns = columnLayout(
        control.columnCount,
        control.columnWidths,
        control.width,
        Math.max(1, ...rows.map((row) => row.length)),
      );
      const shownColumn = textColumnIndex(control.textColumn, columns);
      const selectedRow = rows[control.selectedIndex ?? -1];
      const selectedText = selectedRow
        ? shownColumn < 0
          ? String(control.selectedIndex)
          : (selectedRow[shownColumn] ?? "")
        : "";
      // The drop button shows always (what Excel gives a new ComboBox), never, or only
      // with the focus, which a control on the design surface never has.
      const button = (control.showDropButtonWhen ?? 2) === 2;
      return (
        <>
          <span class="combo-text">
            {authored || selectedText}
          </span>
          {button && (
            <span class="combo-arrow" aria-hidden="true">
              {/* DropButtonStyle: Plain, Arrow, Ellipsis, Reduce. */}
              {[null, <span class="combo-chevron" />, "…", <span class="combo-reduce" />][
                control.dropButtonStyle ?? 1
              ] ?? null}
            </span>
          )}
        </>
      );
    }
    case "listbox": {
      // ListStyle Option puts a box before each item: a check box when several
      // items can be selected, an option button when only one.
      const box = control.listStyle === 1 ? (control.multiSelect ? "check" : "option") : undefined;
      // The items of a RowSource are in the workbook; the canvas names the cells.
      if (control.rowSource && !control.list?.length)
        return (
          <span class="list-items">
            <div class="list-source">{control.rowSource}</div>
          </span>
        );
      // Several columns (or headings) are laid out by ColumnCount and ColumnWidths.
      const rows = (control.list ?? []).map(itemColumns);
      const columns = columnLayout(
        control.columnCount,
        control.columnWidths,
        Math.max(0, control.width - 4),
        Math.max(1, ...rows.map((row) => row.length)),
      );
      const cells = columns.length > 1 || control.columnHeads === true;
      const line = (row: string[]) =>
        cells
          ? columns.map((width, column) => (
              <span class="list-cell" style={{ width: px(width) }}>
                {row[column] ?? ""}
              </span>
            ))
          : row[0];
      return (
        <span class="list-items">
          {control.columnHeads === true && (
            <div class="list-item list-heads">{line([])}</div>
          )}
          {rows.map((row, index) => {
            const selected = index === control.selectedIndex;
            return (
              <div
                key={index}
                class={`list-item ${selected ? "selected" : ""} ${box ? "with-box" : ""}`}
              >
                {box && (
                  <span
                    class={`toggle-box ${box} ${selected ? "checked" : ""}`}
                    aria-hidden="true"
                  />
                )}
                {line(row)}
              </div>
            );
          })}
        </span>
      );
    }
    case "spinbutton":
      return (
        <span
          class={`spinbutton-chrome ${horizontal(control) ? "horizontal" : "vertical"}`}
          aria-hidden="true"
        >
          <span class="spinbutton-arrow decrement">
            <span />
          </span>
          <span class="spinbutton-arrow increment">
            <span />
          </span>
        </span>
      );
    case "scrollbar":
      return (
        <span
          class={`scrollbar-chrome ${horizontal(control) ? "horizontal" : "vertical"}`}
          aria-hidden="true"
        >
          <span class="scrollbar-arrow decrement">
            <span />
          </span>
          <span class="scrollbar-track">
            <span class="scrollbar-thumb" />
          </span>
          <span class="scrollbar-arrow increment">
            <span />
          </span>
        </span>
      );
    case "image":
      return <ControlPicture control={control} />;
    case "frame":
      return (
        <>
          <span class="frame-edge" aria-hidden="true" />
          {control.caption ? <span class="frame-caption">{control.caption}</span> : null}
        </>
      );
    case "multipage":
      return null;
    case "page":
      return null;
    case "progressbar": {
      // As written into the workbook: caption and percentage over a bar (ras-form progress.rs).
      const layout = progressBarLayout(control);
      const colors = control as unknown as Record<string, unknown>;
      const box = ([left, top, width, height]: [number, number, number, number]) => ({
        left: px(left),
        top: px(top),
        width: px(width),
        height: px(height),
      });
      return (
        <>
          <span
            class="progress-track"
            style={{ ...box(layout.track), backgroundColor: colorStyle(colors.trackColor) }}
          />
          {layout.bar[2] > 0 && (
            <span
              class="progress-bar"
              style={{ ...box(layout.bar), backgroundColor: colorStyle(colors.barColor) }}
            />
          )}
          <span class="progress-caption" style={box(layout.caption)}>
            {control.caption ?? ""}
          </span>
          {layout.showValue && (
            <span
              class="progress-value"
              style={{ ...box(layout.value), color: colorStyle(colors.valueColor) }}
            >
              {layout.text}
            </span>
          )}
        </>
      );
    }
    case "filepicker": {
      // As written into the workbook: a text box with a button at its right (ras-form picker.rs).
      const raw = control as unknown as Record<string, unknown>;
      const buttonWidth = Math.max(
        Math.min(typeof raw.buttonWidth === "number" ? raw.buttonWidth : 24, control.width),
        0.75,
      );
      const caption = typeof raw.buttonCaption === "string" ? raw.buttonCaption : "...";
      return (
        <>
          <span
            class="picker-text"
            style={{ width: px(Math.max(control.width - buttonWidth, 0.75)), height: px(control.height) }}
          >
            {control.text ?? ""}
          </span>
          <span
            class="picker-button"
            style={{ left: px(control.width - buttonWidth), width: px(buttonWidth), height: px(control.height) }}
          >
            {caption}
          </span>
        </>
      );
    }
    case "tabstrip":
      return (
        <div class="tabs">
          {(control.tabs ?? []).map(
            (tab, index) =>
              tab.visible !== false && (
                <span key={tab.name} class={index === (control.selectedIndex ?? 0) ? "active" : ""}>
                  <span class="tab-text">{tab.caption ?? tab.name}</span>
                </span>
              ),
          )}
        </div>
      );
    default:
      return (
        <span class="placeholder">
          {control.type}: {control.name}
        </span>
      );
  }
}

export function DesignerCanvas({
  document: sourceDocument,
  approximateBounds = "approximate bounds",
  editable = false,
  parentEditable = false,
  pageEditable = false,
  formResizable = false,
  activePages,
  onActivePageChange,
  busy = false,
  eventEditable = false,
  eventPending = false,
  zoom = 1,
  showGrid = true,
  snapping = false,
  productivityEditable = false,
  cannotMove = "This control is larger than its parent. Resize it to fit before moving.",
  selectedId: externalSelected,
  selectedIds: externalSelectedIds,
  onSelectionChange,
  onSelectionChangeMany,
  onCommit = () => {},
  onFormResize,
  onContextMenu,
  onOpenEvent,
  onShortcut,
  placing = false,
  onPlace,
  onCancelPlacement,
  onDelete,
  hierarchy = "Selection path",
  invalidDrop = "This control cannot be placed in this container.",
  dropTargetLabel = "Drop target",
  placementSize,
  resizeWidthLabel = "Resize form width",
  resizeHeightLabel = "Resize form height",
  resizeBothLabel = "Resize form width and height",
}: {
  document: DesignerDocument;
  approximateBounds?: string;
  editable?: boolean;
  parentEditable?: boolean;
  pageEditable?: boolean;
  formResizable?: boolean;
  activePages?: ActivePages;
  onActivePageChange?: (ownerId: string, pageId: string) => void;
  busy?: boolean;
  eventEditable?: boolean;
  eventPending?: boolean;
  zoom?: number;
  showGrid?: boolean;
  snapping?: boolean;
  productivityEditable?: boolean;
  cannotMove?: string;
  selectedId?: string;
  selectedIds?: string[];
  onSelectionChange?: (id: string | undefined) => void;
  onSelectionChangeMany?: (ids: string[]) => void;
  onCommit?: (operations: (GeometryOperation | ParentOperation | ProductivityOperation)[]) => void;
  onFormResize?: (operations: PropertyOperation[]) => void;
  onContextMenu?: (controlId: string, x: number, y: number) => void;
  onOpenEvent?: (targetId?: string, choose?: boolean) => void;
  onShortcut?: (action: "copy" | "cut" | "paste" | "duplicate" | "delete") => boolean;
  placing?: boolean;
  onPlace?: (
    left: number,
    top: number,
    parentId?: string,
    size?: { width: number; height: number },
  ) => void;
  onCancelPlacement?: () => void;
  onDelete?: () => void;
  hierarchy?: string;
  invalidDrop?: string;
  dropTargetLabel?: string;
  placementSize?: { width: number; height: number };
  resizeWidthLabel?: string;
  resizeHeightLabel?: string;
  resizeBothLabel?: string;
}) {
  const [localActivePages, setLocalActivePages] = useState<ActivePages>({});
  const document = useMemo(
    () => displayPages(sourceDocument, activePages ?? localActivePages),
    [sourceDocument, activePages, localActivePages],
  );
  const [localSelection, setLocalSelection] = useState<string[]>([]);
  const controlledSelection = !!onSelectionChangeMany || !!onSelectionChange;
  const selectionIds = controlledSelection
    ? (externalSelectedIds ?? (externalSelected ? [externalSelected] : []))
    : localSelection;
  const selected = selectionIds.at(-1);
  const setSelection = (ids: string[]) => {
    if (onSelectionChangeMany) onSelectionChangeMany(ids);
    else if (onSelectionChange) onSelectionChange(ids.at(-1));
    else setLocalSelection(ids);
  };
  const [preview, setPreview] = useState<{ id: string; geometry: Geometry; floating?: boolean }>();
  const [formResizePreview, setFormResizePreview] = useState<{ width: number; height: number }>();
  const [groupPreview, setGroupPreview] = useState<{ id: string; left: number; top: number }[]>();
  const [marqueeBox, setMarqueeBox] = useState<SelectionRect>();
  const [destination, setDestination] = useState<{ target?: DropTarget; valid: boolean }>();
  const [notice, setNotice] = useState<string>();
  // A control being drawn with the Toolbox: a click makes it with the default size, a drag
  // with the dragged rectangle (as in the VBA editor). Points, in the container's coordinates.
  const placement = useRef<{
    pointerId: number;
    target: DropTarget;
    x: number;
    y: number;
    start: { left: number; top: number };
    rect?: Geometry;
  }>();
  const [placementBox, setPlacementBox] = useState<Geometry>();
  const gesture = useRef<{
    control: DesignerControl;
    start: Geometry;
    end: Geometry;
    x: number;
    y: number;
    handle?: ResizeHandle;
    pointerId?: number;
    keys?: Set<string>;
    dx: number;
    dy: number;
    target?: DropTarget;
    valid?: boolean;
    parenting?: boolean;
    group?: DesignerControl[];
    groupMoves?: Extract<ProductivityOperation, { type: "moveControls" }>;
  }>();
  const formResizeGesture = useRef<{
    handle: "e" | "s" | "se";
    pointerId: number;
    x: number;
    y: number;
    start: { width: number; height: number };
    end: { width: number; height: number };
  }>();
  const marquee = useRef<{
    pointerId: number;
    parentId: string;
    startX: number;
    startY: number;
    endX: number;
    endY: number;
  }>();
  const canvas = useRef<HTMLDivElement>(null);
  const cancel = () => {
    placement.current = undefined;
    setPlacementBox(undefined);
    gesture.current = undefined;
    formResizeGesture.current = undefined;
    marquee.current = undefined;
    setPreview(undefined);
    setFormResizePreview(undefined);
    setGroupPreview(undefined);
    setMarqueeBox(undefined);
    setDestination(undefined);
  };
  const cancelPointer = (event: JSX.TargetedPointerEvent<HTMLDivElement>) => {
    if (
      gesture.current?.pointerId === event.pointerId ||
      formResizeGesture.current?.pointerId === event.pointerId ||
      marquee.current?.pointerId === event.pointerId ||
      placement.current?.pointerId === event.pointerId
    )
      cancel();
  };
  const finish = () => {
    const current = gesture.current;
    gesture.current = undefined;
    if (!current) return;
    setDestination(undefined);
    if (current.group) {
      const operation = current.groupMoves;
      const changed = operation?.moves.some((move) => {
        const start = current.group!.find((control) => control.id === move.controlId)!;
        return move.left !== start.left || move.top !== start.top;
      });
      if (changed && editable) onCommit([operation!]);
      else setGroupPreview(undefined);
      return;
    }
    if (current.parenting && !current.valid) {
      setPreview(undefined);
      return;
    }
    const operations: (GeometryOperation | ParentOperation)[] =
      current.parenting &&
      current.target &&
      current.target.parentId !== (current.control.parentId ?? "")
        ? [
            {
              type: "setParent",
              controlId: current.control.id,
              parentId: current.target.parentId,
              left: current.end.left,
              top: current.end.top,
            },
          ]
        : geometryOperations(current.control.id, current.start, current.end);
    if (operations.length && editable) onCommit(operations);
    else setPreview(undefined);
  };
  useEffect(() => {
    cancel();
    if (selectionIds.length > 1) {
      const visible = siblingSelection(document, selectionIds)
        .filter((control) => isControlVisible(document, control.id))
        .map((control) => control.id);
      if (visible.length !== selectionIds.length) setSelection(visible);
    }
    if (
      selected &&
      !isControlVisible(document, selected) &&
      !document.controls.some(
        (c) =>
          c.id === selected &&
          c.type.toLowerCase() === "page" &&
          !!c.parentId &&
          document.controls.some((owner) => owner.id === c.parentId),
      )
    )
      setSelection([]);
  }, [document]);
  useEffect(() => {
    if (!formResizable) {
      formResizeGesture.current = undefined;
      setFormResizePreview(undefined);
    } else if (busy || placing) {
      if (formResizeGesture.current) {
        formResizeGesture.current = undefined;
        setFormResizePreview(undefined);
      }
    }
  }, [formResizable, busy, placing]);
  useEffect(() => {
    if (!editable) {
      gesture.current = undefined;
      if (!busy) setPreview(undefined);
    }
  }, [editable, busy]);
  useEffect(() => {
    cancel();
  }, [parentEditable, pageEditable, placing]);
  useEffect(() => {
    if (!busy && !gesture.current) cancel();
  }, [busy]);
  useEffect(() => {
    cancel();
  }, [zoom, snapping]);
  useEffect(() => {
    const blur = () => {
      if (gesture.current?.keys) finish();
      else cancel();
    };
    window.addEventListener("blur", blur);
    return () => window.removeEventListener("blur", blur);
  }, [editable, onCommit]);
  const startGesture = (
    control: DesignerControl,
    x: number,
    y: number,
    handle?: ResizeHandle,
    pointerId?: number,
  ) => {
    if (!editable || busy || !canEditGeometry(control)) return;
    const start = {
      left: control.left,
      top: control.top,
      width: control.width,
      height: control.height,
    };
    setGroupPreview(undefined);
    setMarqueeBox(undefined);
    gesture.current = { control, start, end: start, x, y, handle, pointerId, dx: 0, dy: 0 };
    setNotice(undefined);
    return gesture.current;
  };
  const startFormResize = (handle: "e" | "s" | "se", x: number, y: number, pointerId: number) => {
    if (!formResizable || busy || placing || selectionIds.length !== 0) return;
    const start = { width: document.width, height: document.height };
    formResizeGesture.current = { handle, x, y, pointerId, start, end: start };
    setFormResizePreview(start);
  };
  const updateFormResize = (dx: number, dy: number) => {
    const current = formResizeGesture.current;
    if (!current) return;
    const rootChildren = document.controls.filter((control) => !control.parentId);
    const minimumWidth = Math.max(
      1,
      ...rootChildren.map((control) => control.left + control.width),
    );
    const minimumHeight = Math.max(
      1,
      ...rootChildren.map((control) => control.top + control.height),
    );
    const resizeAxis = (start: number, delta: number, minimum: number) => {
      if (stableDelta(delta) === 0) return start;
      const target = start + stableDelta(delta);
      return stableDelta(
        Math.max(minimum, snapping ? Math.round(target / gridSize) * gridSize : target),
      );
    };
    const end = {
      width:
        current.handle === "e" || current.handle === "se"
          ? resizeAxis(current.start.width, dx, minimumWidth)
          : current.start.width,
      height:
        current.handle === "s" || current.handle === "se"
          ? resizeAxis(current.start.height, dy, minimumHeight)
          : current.start.height,
    };
    current.end = end;
    setFormResizePreview(end);
  };
  const finishFormResize = () => {
    const current = formResizeGesture.current;
    formResizeGesture.current = undefined;
    if (!current) return;
    const operations: PropertyOperation[] = [];
    if (current.end.width !== current.start.width)
      operations.push({ type: "setFormProperty", field: "width", value: current.end.width });
    if (current.end.height !== current.start.height)
      operations.push({ type: "setFormProperty", field: "height", value: current.end.height });
    if (!operations.length || !onFormResize || !formResizable || busy || placing) {
      setFormResizePreview(undefined);
      return;
    }
    onFormResize?.(operations);
  };
  const startGroupGesture = (
    controls: DesignerControl[],
    x: number,
    y: number,
    pointerId?: number,
  ) => {
    if (
      !productivityEditable ||
      !editable ||
      busy ||
      controls.length < 2 ||
      controls.some((control) => !canEditGeometry(control))
    )
      return;
    const primary = controls[controls.length - 1];
    const start = {
      left: primary.left,
      top: primary.top,
      width: primary.width,
      height: primary.height,
    };
    setPreview(undefined);
    setGroupPreview(undefined);
    setMarqueeBox(undefined);
    gesture.current = {
      control: primary,
      start,
      end: start,
      x,
      y,
      pointerId,
      dx: 0,
      dy: 0,
      group: controls,
    };
    setNotice(undefined);
    return gesture.current;
  };
  const updateGroupGesture = (dx: number, dy: number) => {
    const current = gesture.current;
    if (!current?.group) return;
    const operation = moveSiblingGroup(
      document,
      current.group.map((control) => control.id),
      dx,
      dy,
      snapping,
    );
    if (!operation || operation.type !== "moveControls") {
      current.groupMoves = undefined;
      setGroupPreview(undefined);
      setNotice(cannotMove);
      return;
    }
    current.groupMoves = operation;
    const primaryMove = operation.moves.find((move) => move.controlId === current.control.id);
    if (primaryMove) current.end = { ...current.end, left: primaryMove.left, top: primaryMove.top };
    setGroupPreview(
      operation.moves.map(({ controlId, left, top }) => ({ id: controlId, left, top })),
    );
    setNotice(undefined);
  };
  const updateGesture = (dx: number, dy: number) => {
    const current = gesture.current;
    if (!current) return;
    if (current.group) {
      updateGroupGesture(dx, dy);
      return;
    }
    const bounds = parentBounds(document, current.control);
    const oldParent = document.controls.find((c) => c.id === current.control.parentId);
    if (
      parentEditable &&
      !current.handle &&
      !current.keys &&
      (!oldParent ||
        isDesignerFrame(oldParent) ||
        (pageEditable && isDesignerPage(document, oldParent))) &&
      (dx !== 0 || dy !== 0)
    ) {
      const form = canvas.current?.querySelector<HTMLElement>(".form");
      if (!form) return;
      const rect = form.getBoundingClientRect();
      const point = {
        left: pixelsToPoints(current.x - rect.left, zoom) + dx,
        top: pixelsToPoints(current.y - rect.top, zoom) + dy,
      };
      const target = dropTarget(
        document,
        point,
        subtreeIds(document.controls, current.control.id),
        pageEditable,
      );
      const startWorld = worldGeometry(document, current.control);
      const world = { ...startWorld, left: startWorld.left + dx, top: startWorld.top + dy };
      const end =
        target &&
        (target.parentId === (current.control.parentId ?? "")
          ? moveGeometry(current.start, dx, dy, target.bounds, snapping)
          : placeGeometry(
              {
                ...world,
                left: world.left - target.origin.left,
                top: world.top - target.origin.top,
              },
              target.bounds,
              snapping,
            ));
      current.parenting = true;
      current.target = target;
      current.valid = !!end;
      if (end && target) current.end = end;
      setDestination({ target, valid: !!end });
      setNotice(end ? undefined : invalidDrop);
      setPreview({
        id: current.control.id,
        geometry:
          end && target
            ? { ...end, left: end.left + target.origin.left, top: end.top + target.origin.top }
            : world,
        floating: true,
      });
      return;
    }
    current.parenting = false;
    current.target = undefined;
    current.valid = undefined;
    setDestination(undefined);
    const end =
      dx === 0 && dy === 0
        ? current.start
        : current.handle
          ? resizeGeometry(
              current.start,
              current.handle,
              dx,
              dy,
              bounds,
              requiredSize(document, current.control),
              snapping,
            )
          : moveGeometry(current.start, dx, dy, bounds, current.keys ? false : snapping);
    if (!end) {
      setNotice(cannotMove);
      return;
    }
    current.end = end;
    setPreview({ id: current.control.id, geometry: end });
  };
  /** Follows the pointer while a control is drawn; a move of a few pixels is still a click. */
  const updatePlacement = (clientX: number, clientY: number) => {
    const current = placement.current;
    const form = canvas.current?.querySelector<HTMLElement>(".form");
    if (!current || !form) return;
    const moved = Math.max(Math.abs(clientX - current.x), Math.abs(clientY - current.y));
    if (!current.rect && moved < 4) return;
    const rect = form.getBoundingClientRect();
    const end = {
      left: pixelsToPoints(clientX - rect.left, zoom) - current.target.origin.left,
      top: pixelsToPoints(clientY - rect.top, zoom) - current.target.origin.top,
    };
    current.rect = placementRect(current.start, end, current.target.bounds, snapping);
    setPlacementBox(
      current.rect && {
        ...current.rect,
        left: current.rect.left + current.target.origin.left,
        top: current.rect.top + current.target.origin.top,
      },
    );
  };
  const pointerDown = (event: JSX.TargetedPointerEvent<HTMLDivElement>) => {
    if (
      event.button !== 0 ||
      gesture.current?.pointerId !== undefined ||
      formResizeGesture.current?.pointerId !== undefined ||
      marquee.current?.pointerId !== undefined ||
      placement.current !== undefined
    )
      return;
    if (gesture.current?.keys) finish();
    const target = event.target as Element;
    if (placing) {
      const form = canvas.current?.querySelector<HTMLElement>(".form");
      if (form && form.contains(target) && !busy) {
        event.preventDefault();
        cancel();
        canvas.current?.focus({ preventScroll: true });
        const bounds = form.getBoundingClientRect();
        const point = {
          left: pixelsToPoints(event.clientX - bounds.left, zoom),
          top: pixelsToPoints(event.clientY - bounds.top, zoom),
        };
        const target = parentEditable
          ? dropTarget(document, point, new Set(), pageEditable)
          : { parentId: "", origin: { left: 0, top: 0 }, bounds: document };
        if (!target) {
          setNotice(invalidDrop);
          return;
        }
        // The control is made when the button is released: at the size dragged, or the
        // default size for a click.
        placement.current = {
          pointerId: event.pointerId,
          target,
          x: event.clientX,
          y: event.clientY,
          start: { left: point.left - target.origin.left, top: point.top - target.origin.top },
        };
        try {
          event.currentTarget.setPointerCapture(event.pointerId);
        } catch {
          /* no active pointer */
        }
      }
      return;
    }
    const formHandle = target.closest<HTMLElement>("[data-form-handle]")?.dataset.formHandle as
      | "e"
      | "s"
      | "se"
      | undefined;
    if (formHandle) {
      event.preventDefault();
      canvas.current?.focus({ preventScroll: true });
      startFormResize(formHandle, event.clientX, event.clientY, event.pointerId);
      if (formResizeGesture.current) {
        try {
          event.currentTarget.setPointerCapture(event.pointerId);
        } catch {
          /* no active pointer */
        }
      }
      return;
    }
    if (target.closest("[data-page-tab]")) return;
    const blankArea = target.closest<HTMLElement>(".container-content, .page-content, .form");
    const candidate = target.closest<HTMLElement>(
      ".control[data-control-id], .selection-outline[data-control-id]",
    );
    const id =
      candidate && (!blankArea || blankArea.contains(candidate))
        ? candidate.dataset.controlId
        : undefined;
    const control = document.controls.find((c) => c.id === id);
    setNotice(undefined);
    canvas.current?.focus({ preventScroll: true });
    if (!control) {
      setSelection([]);
      const form = canvas.current?.querySelector<HTMLElement>(".form");
      if (productivityEditable && form?.contains(target) && !busy) {
        const container = target.closest<HTMLElement>("[data-container-id]");
        const current = {
          pointerId: event.pointerId,
          parentId: container?.dataset.containerId ?? "",
          startX: event.clientX,
          startY: event.clientY,
          endX: event.clientX,
          endY: event.clientY,
        };
        marquee.current = current;
        setMarqueeBox({
          left: current.startX,
          top: current.startY,
          right: current.startX,
          bottom: current.startY,
        });
        try {
          event.currentTarget.setPointerCapture(event.pointerId);
        } catch {
          /* no active pointer */
        }
      }
      return;
    }
    event.preventDefault();
    const modified = event.ctrlKey || event.metaKey || event.shiftKey;
    if (productivityEditable && modified) {
      setSelection(updateSelection(document, selectionIds, control.id, true));
      setMarqueeBox(undefined);
      return;
    }
    const remainsSelected = selectionIds.includes(control.id);
    const nextSelection =
      productivityEditable && remainsSelected && siblingSelection(document, selectionIds).length > 1
        ? promoteSelection(document, selectionIds, control.id)
        : [control.id];
    if (nextSelection.some((id, index) => id !== selectionIds[index])) setSelection(nextSelection);
    const selectedControls =
      productivityEditable && nextSelection.length > 1
        ? siblingSelection(document, nextSelection)
        : [];
    const handle = target.closest<HTMLElement>("[data-handle]")?.dataset.handle as
      | ResizeHandle
      | undefined;
    if (selectedControls.length > 1 && !handle)
      startGroupGesture(selectedControls, event.clientX, event.clientY, event.pointerId);
    else startGesture(control, event.clientX, event.clientY, handle, event.pointerId);
    if (gesture.current) {
      // Synthetic Webview tests have no active native pointer to capture.
      try {
        event.currentTarget.setPointerCapture(event.pointerId);
      } catch {
        /* no active pointer */
      }
    }
  };
  const contextMenu = (event: JSX.TargetedMouseEvent<HTMLDivElement>) => {
    if (!productivityEditable && !eventEditable) return;
    const target = event.target as Element;
    const blankArea = target.closest<HTMLElement>(".container-content, .page-content, .form");
    const candidate = target.closest<HTMLElement>(
      ".control[data-control-id], .selection-outline[data-control-id]",
    );
    const tab = target.closest<HTMLElement>("[data-page-tab]");
    const page = target.closest<HTMLElement>(".page-content[data-container-id]");
    const id =
      tab?.dataset.pageTab ??
      (page && (!candidate || !page.contains(candidate))
        ? page.dataset.containerId
        : candidate && (!blankArea || blankArea.contains(candidate))
          ? candidate.dataset.controlId
          : undefined);
    if (!id) return;
    event.preventDefault();
    if (document.controls.find((control) => control.id === id)?.type.toLowerCase() === "page")
      setSelection([id]);
    else if (!selectionIds.includes(id)) setSelection([id]);
    else setSelection(promoteSelection(document, selectionIds, id));
    canvas.current?.focus({ preventScroll: true });
    onContextMenu?.(id, event.clientX, event.clientY);
  };
  const openEventOnDoubleClick = (event: JSX.TargetedMouseEvent<HTMLDivElement>) => {
    event.stopPropagation();
    if (
      event.detail < 2 ||
      !eventEditable ||
      !onOpenEvent ||
      busy ||
      eventPending ||
      placing ||
      selectionIds.length > 1 ||
      gesture.current ||
      formResizeGesture.current ||
      marquee.current ||
      preview ||
      formResizePreview ||
      groupPreview ||
      marqueeBox ||
      destination?.valid === false
    )
      return;

    // Chromium targets click/dblclick at the pointer-capture element even
    // when capture was released during pointerup. Recover the visual target
    // when that element is our viewport, keeping ordinary bubbling intact.
    const target =
      event.target === event.currentTarget
        ? event.currentTarget.ownerDocument.elementFromPoint?.(event.clientX, event.clientY)
        : (event.target as Element);
    if (!target) return;
    if (
      target.closest(
        "[data-handle], [data-form-handle], .marquee-selection, .drop-target, .group-selection-outline",
      )
    )
      return;
    const form = canvas.current?.querySelector<HTMLElement>(".form");
    if (!form?.contains(target)) return;

    const page = target.closest<HTMLElement>(".page-content[data-container-id]");
    const candidate = target.closest<HTMLElement>(
      ".control[data-control-id], .selection-outline[data-control-id]",
    );
    const tab = target.closest<HTMLElement>("[data-page-tab]");
    const targetId =
      tab?.dataset.pageTab ??
      (page && (!candidate || !page.contains(candidate))
        ? page.dataset.containerId
        : candidate?.dataset.controlId);
    onOpenEvent(targetId, false);
  };
  const keyDown = (event: JSX.TargetedKeyboardEvent<HTMLDivElement>) => {
    if (
      event.isComposing ||
      (event.target as Element).closest(
        "input, textarea, select, [contenteditable]:not([contenteditable='false'])",
      )
    )
      return;
    if (
      productivityEditable &&
      (event.ctrlKey || event.metaKey) &&
      !event.altKey &&
      !event.repeat &&
      ["c", "x", "v", "d"].includes(event.key.toLowerCase())
    ) {
      const action =
        event.key.toLowerCase() === "c"
          ? "copy"
          : event.key.toLowerCase() === "x"
            ? "cut"
            : event.key.toLowerCase() === "v"
              ? "paste"
              : "duplicate";
      if (onShortcut?.(action)) event.preventDefault();
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      cancel();
      onCancelPlacement?.();
      return;
    }
    if (
      (event.key === "Delete" || event.key === "Backspace") &&
      event.target === event.currentTarget &&
      !busy &&
      !placing &&
      !event.repeat &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.shiftKey &&
      !event.altKey
    ) {
      if (productivityEditable && onShortcut?.("delete")) {
        event.preventDefault();
        cancel();
        return;
      }
      if (onDelete) {
        event.preventDefault();
        cancel();
        onDelete();
        return;
      }
    }
    if (
      placing ||
      !editable ||
      busy ||
      event.ctrlKey ||
      event.metaKey ||
      event.altKey ||
      !["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"].includes(event.key)
    )
      return;
    const control = document.controls.find((c) => c.id === selected);
    if (!control || !canEditGeometry(control) || gesture.current?.pointerId !== undefined) return;
    event.preventDefault();
    if (!gesture.current) {
      const selectedControls =
        productivityEditable && selectionIds.length > 1
          ? siblingSelection(document, selectionIds)
          : [];
      const created =
        selectedControls.length > 1
          ? startGroupGesture(selectedControls, 0, 0)
          : startGesture(control, 0, 0);
      if (created) created.keys = new Set();
    }
    const current = gesture.current;
    if (!current) return;
    current.keys!.add(event.key);
    const step = event.shiftKey ? 10 : 1;
    current.dx =
      current.end.left -
      current.start.left +
      (event.key === "ArrowLeft" ? -step : event.key === "ArrowRight" ? step : 0);
    current.dy =
      current.end.top -
      current.start.top +
      (event.key === "ArrowUp" ? -step : event.key === "ArrowDown" ? step : 0);
    updateGesture(current.dx, current.dy);
  };
  const children = (parentId?: string) => orderedChildren(document, parentId);
  const draw = (control: DesignerControl | undefined, floating = false): JSX.Element | null => {
    if (!control || control.visible === false) return null;
    if (preview?.floating && preview.id === control.id && !floating) return null;
    const type = control.type.toLowerCase();
    const pages = type === "multipage" ? children(control.id) : [];
    const selected = control.selectedIndex ?? (pages.length ? 0 : -1);
    const page = pages[selected];
    const groupPosition = groupPreview?.find((item) => item.id === control.id);
    const shown =
      preview?.id === control.id
        ? { ...control, ...preview.geometry }
        : groupPosition
          ? { ...control, left: groupPosition.left, top: groupPosition.top }
          : control;
    const insets = contentInsets(type, control.tabOrientation);
    const selectedControl = selectionIds.includes(control.id);
    return (
      <>
        <div
          key={control.id}
          class={`control ${type} ${borderClass(shown)} ${scalarClasses(shown)} ${control.enabled === false ? "disabled" : ""} ${type === "frame" && control.caption ? "has-caption" : ""} ${type === "togglebutton" ? { on: "pressed", null: "mixed", off: "" }[toggleState(control.value)] : ""}`}
          data-control-id={control.id}
          title={`${control.name} (${control.type})${control.approximate ? ` — ${approximateBounds}` : ""}`}
          style={{
            ...appearanceStyle(shown),
            left: px(shown.left),
            top: px(shown.top),
            width: px(shown.width),
            height: px(shown.height),
          }}
        >
          <ControlGlyph control={shown} />
          {type === "multipage" ? (
            <>
              <div class="tabs">
                {pages.map(
                  (p, index) =>
                    p.visible !== false && (
                      <button
                        type="button"
                        key={p.id}
                        data-page-tab={p.id}
                        class={`${index === selected ? "active" : ""} ${selectedId === p.id ? "selected-page" : ""}`}
                        aria-pressed={selectedId === p.id}
                        onClick={(event) => {
                          event.stopPropagation();
                          cancel();
                          if (onActivePageChange) onActivePageChange(control.id, p.id);
                          else
                            setLocalActivePages((current) => ({ ...current, [control.id]: p.id }));
                          setSelection([p.id]);
                        }}
                      >
                        <span class="tab-text">{p.caption ?? p.name}</span>
                      </button>
                    ),
                )}
              </div>
              {page && page.visible !== false && (
                <div
                  class="control page page-content"
                  data-container-id={page.id}
                  title={page.name}
                  style={{
                    ...appearanceStyle(page),
                    ...backgroundPicture(
                      page.picture?.remove === true ? undefined : page.picturePreview,
                      page.pictureAlignment,
                      page.pictureSizeMode,
                      page.pictureTiling,
                    ),
                    left: px(insets.left),
                    top: px(insets.top),
                    right: px(insets.right),
                    bottom: px(insets.bottom),
                  }}
                >
                  {children(page.id).map((c) => draw(c))}
                </div>
              )}
            </>
          ) : type === "frame" ? (
            <div
              class="container-content"
              data-container-id={control.id}
              style={{
                ...backgroundPicture(
                  control.picture?.remove === true ? undefined : control.picturePreview,
                  control.pictureAlignment,
                  control.pictureSizeMode,
                  control.pictureTiling,
                ),
                left: px(insets.left),
                top: px(insets.top),
                right: px(insets.right),
                bottom: px(insets.bottom),
              }}
            >
              {children(control.id).map((c) => draw(c))}
            </div>
          ) : null}
        </div>
        {selectedControl && (
          <div
            class="selection-outline"
            data-control-id={control.id}
            style={{
              left: px(shown.left),
              top: px(shown.top),
              width: px(shown.width),
              height: px(shown.height),
            }}
          >
            {editable &&
              selectionIds.length === 1 &&
              canEditGeometry(control) &&
              resizeHandles.map((handle) => (
                <span
                  class={`resize-handle handle-${handle}`}
                  data-handle={handle}
                  style={{ transform: `translate(-50%, -50%) scale(${1 / zoom})` }}
                />
              ))}
          </div>
        )}
      </>
    );
  };
  const selectedId = selected;
  const groupBounds =
    productivityEditable && selectionIds.length > 1
      ? selectedWorldBounds(
          groupPreview
            ? {
                ...document,
                controls: document.controls.map((control) => {
                  const position = groupPreview.find((item) => item.id === control.id);
                  return position
                    ? { ...control, left: position.left, top: position.top }
                    : control;
                }),
              }
            : document,
          selectionIds,
        )
      : undefined;
  return (
    <div class="designer-canvas">
      <nav class="selection-path" aria-label={hierarchy}>
        <button
          type="button"
          disabled={busy}
          onClick={() => {
            cancel();
            setSelection([]);
            canvas.current?.focus({ preventScroll: true });
          }}
        >
          {document.name}
        </button>
        {selectionPath(document, selected).map((control) => (
          <span key={control.id}>
            {" "}
            ›{" "}
            <button
              type="button"
              disabled={busy}
              aria-current={control.id === selected ? "true" : undefined}
              onClick={() => {
                cancel();
                setSelection([control.id]);
                canvas.current?.focus({ preventScroll: true });
              }}
            >
              {control.name}
            </button>
          </span>
        ))}
      </nav>
      {notice && (
        <p class="notice" role="status">
          {notice}
        </p>
      )}
      <div
        ref={canvas}
        class={`designer-viewport ${placing ? "placing-control" : ""}`}
        tabIndex={0}
        aria-label={document.name}
        onPointerDown={pointerDown}
        onDblClick={openEventOnDoubleClick}
        onContextMenu={contextMenu}
        onPointerMove={(event) => {
          const current = gesture.current;
          const formResize = formResizeGesture.current;
          const drawing = placement.current;
          if (drawing?.pointerId === event.pointerId) {
            updatePlacement(event.clientX, event.clientY);
            return;
          }
          if (placing && parentEditable && !busy) {
            const form = canvas.current?.querySelector<HTMLElement>(".form");
            if (form) {
              const rect = form.getBoundingClientRect();
              const target = dropTarget(
                document,
                {
                  left: pixelsToPoints(event.clientX - rect.left, zoom),
                  top: pixelsToPoints(event.clientY - rect.top, zoom),
                },
                new Set(),
                pageEditable,
              );
              const valid =
                !!target &&
                (!placementSize ||
                  (placementSize.width <= target.bounds.width &&
                    placementSize.height <= target.bounds.height));
              setDestination({ target, valid });
              setNotice(valid ? undefined : invalidDrop);
            }
          }
          const activeMarquee = marquee.current;
          if (activeMarquee?.pointerId === event.pointerId) {
            activeMarquee.endX = event.clientX;
            activeMarquee.endY = event.clientY;
            setMarqueeBox({
              left: Math.min(activeMarquee.startX, activeMarquee.endX),
              top: Math.min(activeMarquee.startY, activeMarquee.endY),
              right: Math.max(activeMarquee.startX, activeMarquee.endX),
              bottom: Math.max(activeMarquee.startY, activeMarquee.endY),
            });
          }
          if (current?.pointerId === event.pointerId)
            updateGesture(
              pixelsToPoints(event.clientX - current.x, zoom),
              pixelsToPoints(event.clientY - current.y, zoom),
            );
          if (formResize?.pointerId === event.pointerId)
            updateFormResize(
              pixelsToPoints(event.clientX - formResize.x, zoom),
              pixelsToPoints(event.clientY - formResize.y, zoom),
            );
        }}
        onPointerUp={(event) => {
          const drawing = placement.current;
          if (drawing?.pointerId === event.pointerId) {
            updatePlacement(event.clientX, event.clientY);
            placement.current = undefined;
            setPlacementBox(undefined);
            if (event.currentTarget.hasPointerCapture?.(event.pointerId))
              event.currentTarget.releasePointerCapture(event.pointerId);
            if (!placing || busy) return;
            const parentId = drawing.target.parentId || undefined;
            if (drawing.rect)
              onPlace?.(drawing.rect.left, drawing.rect.top, parentId, {
                width: drawing.rect.width,
                height: drawing.rect.height,
              });
            else onPlace?.(drawing.start.left, drawing.start.top, parentId);
            return;
          }
          const formResize = formResizeGesture.current;
          if (formResize?.pointerId === event.pointerId) {
            updateFormResize(
              pixelsToPoints(event.clientX - formResize.x, zoom),
              pixelsToPoints(event.clientY - formResize.y, zoom),
            );
            finishFormResize();
            if (event.currentTarget.hasPointerCapture?.(event.pointerId))
              event.currentTarget.releasePointerCapture(event.pointerId);
            return;
          }
          const current = gesture.current;
          if (current?.pointerId === event.pointerId) {
            updateGesture(
              pixelsToPoints(event.clientX - current.x, zoom),
              pixelsToPoints(event.clientY - current.y, zoom),
            );
            finish();
            if (event.currentTarget.hasPointerCapture?.(event.pointerId))
              event.currentTarget.releasePointerCapture(event.pointerId);
            return;
          }
          const activeMarquee = marquee.current;
          if (activeMarquee?.pointerId === event.pointerId) {
            activeMarquee.endX = event.clientX;
            activeMarquee.endY = event.clientY;
            const selection = {
              left: Math.min(activeMarquee.startX, activeMarquee.endX),
              top: Math.min(activeMarquee.startY, activeMarquee.endY),
              right: Math.max(activeMarquee.startX, activeMarquee.endX),
              bottom: Math.max(activeMarquee.startY, activeMarquee.endY),
            };
            const form = canvas.current?.querySelector<HTMLElement>(".form");
            const rectFor = (id: string) => {
              const element = Array.from(
                form?.querySelectorAll<HTMLElement>(".control[data-control-id]") ?? [],
              ).find((candidate) => candidate.dataset.controlId === id);
              if (!element) return undefined;
              const rect = element.getBoundingClientRect();
              return { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom };
            };
            setSelection(
              fullyEnclosedChildren(document, activeMarquee.parentId, selection, rectFor),
            );
            marquee.current = undefined;
            setMarqueeBox(undefined);
            if (event.currentTarget.hasPointerCapture?.(event.pointerId))
              event.currentTarget.releasePointerCapture(event.pointerId);
          }
        }}
        onPointerCancel={cancelPointer}
        onLostPointerCapture={cancelPointer}
        onKeyDown={keyDown}
        onKeyUp={(event) => {
          const current = gesture.current;
          current?.keys?.delete(event.key);
          if (current?.keys && !current.keys.size) finish();
        }}
        onBlur={() => {
          if (gesture.current?.keys) finish();
          else cancel();
        }}
        style={{
          width: `${pointsToPixels((formResizePreview ?? document).width + 2, zoom)}px`,
          height: `${pointsToPixels((formResizePreview ?? document).height + 24, zoom)}px`,
        }}
      >
        <div class="form-shell" style={{ transform: `scale(${zoom})` }}>
          <div class="form-title">
            <span class="form-title-text">{document.caption}</span>
            <span class="form-title-close" aria-hidden="true">
              ✕
            </span>
          </div>
          <div class="form-client-area">
            <div
              class={`form ${borderClass(document)} ${typeof document.specialEffect === "number" ? `effect-${document.specialEffect}` : ""} ${showGrid ? "show-grid" : ""} ${selectedId === undefined ? "form-selected" : ""}`}
              data-container-id=""
              style={{
                ...appearanceStyle(document),
                ...backgroundPicture(
                  document.picture?.remove === true ? undefined : document.picturePreview,
                  document.pictureAlignment,
                  document.pictureSizeMode,
                  document.pictureTiling,
                ),
                width: px((formResizePreview ?? document).width),
                height: px((formResizePreview ?? document).height),
                backgroundSize: `${px(gridSize)} ${px(gridSize)}`,
              }}
            >
              {children().map((c) => draw(c))}
              {groupBounds && (
                <div
                  class="group-selection-outline"
                  aria-hidden="true"
                  style={{
                    left: px(groupBounds.left),
                    top: px(groupBounds.top),
                    width: px(groupBounds.right - groupBounds.left),
                    height: px(groupBounds.bottom - groupBounds.top),
                  }}
                />
              )}
              {marqueeBox &&
                (() => {
                  const form = canvas.current?.querySelector<HTMLElement>(".form");
                  const rect = form?.getBoundingClientRect();
                  if (!rect) return null;
                  return (
                    <div
                      class="marquee-selection"
                      aria-hidden="true"
                      style={{
                        left: px(pixelsToPoints(marqueeBox.left - rect.left, zoom)),
                        top: px(pixelsToPoints(marqueeBox.top - rect.top, zoom)),
                        width: px(pixelsToPoints(marqueeBox.right - marqueeBox.left, zoom)),
                        height: px(pixelsToPoints(marqueeBox.bottom - marqueeBox.top, zoom)),
                      }}
                    />
                  );
                })()}
              {placementBox && (
                <div
                  class="placement-box"
                  aria-hidden="true"
                  style={{
                    left: px(placementBox.left),
                    top: px(placementBox.top),
                    width: px(placementBox.width),
                    height: px(placementBox.height),
                  }}
                />
              )}
              {destination?.target && (
                <div
                  class={`drop-target ${destination.valid ? "valid" : "invalid"}`}
                  data-drop-parent-id={destination.target.parentId}
                  aria-label={dropTargetLabel}
                  style={{
                    left: px(destination.target.origin.left),
                    top: px(destination.target.origin.top),
                    width: px(destination.target.bounds.width),
                    height: px(destination.target.bounds.height),
                  }}
                />
              )}
              {preview?.floating && (
                <div class="drag-preview" aria-hidden="true">
                  {draw(
                    document.controls.find((c) => c.id === preview.id),
                    true,
                  )}
                </div>
              )}
            </div>
            {formResizable && !busy && !placing && selectionIds.length === 0 && (
              <>
                {(
                  [
                    ["e", resizeWidthLabel],
                    ["s", resizeHeightLabel],
                    ["se", resizeBothLabel],
                  ] as const
                ).map(([handle, label]) => (
                  <span
                    key={handle}
                    class={`form-resize-handle form-resize-${handle}`}
                    data-form-handle={handle}
                    role="button"
                    tabIndex={-1}
                    aria-label={label}
                    title={label}
                    style={{ transform: `translate(-50%, -50%) scale(${1 / zoom})` }}
                  />
                ))}
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
