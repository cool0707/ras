/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
import type { JSX } from "preact";
import type {
  DesignerDocument,
  PropertyGridData,
  PropertyTarget,
} from "../../../src/userForm/protocol";
import type { EditorPreview } from "./types";
import { objectValue } from "./types";

// Display approximations only: neither CSS resolution nor fallback fonts are
// ever written back to an authored color/font. System colors are drawn in the
// Windows 10/11 defaults (what Excel shows at run time with the default
// theme), not CSS system colors, which follow the editor's dark/light scheme.
export const systemColorTable: Record<string, { hex: string; label: string }> = {
  scrollBar: { hex: "#C8C8C8", label: "スクロール バー" },
  desktop: { hex: "#000000", label: "デスクトップ" },
  activeCaption: { hex: "#99B4D1", label: "アクティブ タイトル バー" },
  inactiveCaption: { hex: "#BFCDDB", label: "非アクティブ タイトル バー" },
  menu: { hex: "#F0F0F0", label: "メニュー バー" },
  window: { hex: "#FFFFFF", label: "ウィンドウの背景" },
  windowFrame: { hex: "#646464", label: "ウィンドウの枠" },
  menuText: { hex: "#000000", label: "メニューのテキスト" },
  windowText: { hex: "#000000", label: "ウィンドウのテキスト" },
  captionText: { hex: "#000000", label: "アクティブ タイトル バーの文字" },
  activeBorder: { hex: "#B4B4B4", label: "アクティブ境界" },
  inactiveBorder: { hex: "#F4F7FC", label: "非アクティブ境界" },
  appWorkspace: { hex: "#ABABAB", label: "アプリケーション作業領域" },
  highlight: { hex: "#0078D7", label: "強調表示" },
  highlightText: { hex: "#FFFFFF", label: "強調表示されたテキスト" },
  buttonFace: { hex: "#F0F0F0", label: "ボタンの表面" },
  buttonShadow: { hex: "#A0A0A0", label: "ボタンの影" },
  grayText: { hex: "#6D6D6D", label: "淡色表示のテキスト" },
  buttonText: { hex: "#000000", label: "ボタンのテキスト" },
  inactiveCaptionText: { hex: "#000000", label: "非アクティブ タイトル バーの文字" },
  buttonHighlight: { hex: "#FFFFFF", label: "ボタンの強調表示" },
  buttonDarkShadow: { hex: "#696969", label: "ボタンの濃い影" },
  buttonLight: { hex: "#E3E3E3", label: "ボタンの明るい部分" },
  infoText: { hex: "#000000", label: "ヒントのテキスト" },
  infoBackground: { hex: "#FFFFE1", label: "ヒントの背景" },
  hotLight: { hex: "#0066CC", label: "ホット トラック" },
  gradientActiveCaption: { hex: "#B9D1EA", label: "アクティブ タイトル バーのグラデーション" },
  gradientInactiveCaption: { hex: "#D7E4F2", label: "非アクティブ タイトル バーのグラデーション" },
  menuHighlight: { hex: "#3399FF", label: "メニューの強調表示" },
  menuBar: { hex: "#F0F0F0", label: "メニュー バーの背景" },
};
export function colorStyle(value: unknown): string | undefined {
  const color = objectValue(value);
  if (color.kind === "rgb" && typeof color.value === "string" && /^#[\da-f]{6}$/i.test(color.value))
    return color.value;
  if (color.kind === "system" && typeof color.value === "string")
    return Object.hasOwn(systemColorTable, color.value)
      ? systemColorTable[color.value].hex
      : undefined;
  if (
    color.kind === "ole" &&
    typeof color.value === "string" &&
    /^0x[\da-f]{8}$/i.test(color.value)
  ) {
    const raw = Number.parseInt(color.value.slice(2), 16);
    if (raw <= 0xffffff)
      return `#${[raw & 255, (raw >>> 8) & 255, (raw >>> 16) & 255].map((n) => n.toString(16).padStart(2, "0")).join("")}`;
    // Unknown OLE values retain their representation; there is no RGB guess.
  }
  return undefined;
}
export function fontStyle(value: unknown): JSX.CSSProperties {
  const font = objectValue(value);
  const decoration = [
    font.underline === true ? "underline" : "",
    font.strikeThrough === true ? "line-through" : "",
  ]
    .filter(Boolean)
    .join(" ");
  return {
    fontFamily:
      typeof font.name === "string" ? `${JSON.stringify(font.name)}, sans-serif` : undefined,
    fontSize:
      typeof font.size === "number" && Number.isFinite(font.size) ? `${font.size}pt` : undefined,
    fontWeight: typeof font.bold === "boolean" ? (font.bold ? "bold" : "normal") : undefined,
    fontStyle: typeof font.italic === "boolean" ? (font.italic ? "italic" : "normal") : undefined,
    textDecoration:
      decoration || (font.underline === false && font.strikeThrough === false ? "none" : undefined),
  };
}
export function appearanceStyle(value: unknown): JSX.CSSProperties {
  const appearance = objectValue(value);
  const borderColor = colorStyle(appearance.borderColor);
  return {
    ...fontStyle(appearance.font),
    // BorderColor is the color of a Single border (app.css .border-single).
    ...(borderColor ? ({ "--border-color": borderColor } as JSX.CSSProperties) : {}),
    color: colorStyle(appearance.foreColor),
    backgroundColor: colorStyle(appearance.backColor),
    textAlign:
      appearance.textAlign === 1
        ? "left"
        : appearance.textAlign === 2
          ? "center"
          : appearance.textAlign === 3
            ? "right"
            : undefined,
  };
}
/**
 * The class that draws BorderStyle. MSForms draws `fmBorderStyleSingle` as a
 * 1px line in place of the SpecialEffect edge; `fmBorderStyleNone` leaves the
 * control's own edge (a TextBox stays sunken) and only removes the line of
 * controls whose edge is that line (an Image).
 */
export function borderClass(value: unknown): string {
  const style = objectValue(value).borderStyle;
  return style === 1 ? "border-single" : style === 0 ? "border-none" : "";
}
export function previewDocument(
  document: DesignerDocument,
  preview?: EditorPreview,
): DesignerDocument {
  if (!preview) return document;
  if (preview.controlId === undefined) return { ...document, ...preview.values };
  return {
    ...document,
    controls: document.controls.map((control) =>
      control.id === preview.controlId ? { ...control, ...preview.values } : control,
    ),
  };
}

export function metadataAppearanceDocument(
  document: DesignerDocument,
  grid?: PropertyGridData,
): DesignerDocument {
  if (!grid) return document;
  const appearance = (target?: PropertyTarget) =>
    Object.fromEntries(
      (target?.descriptors ?? [])
        .filter(
          (descriptor) =>
            [
              "font",
              "foreColor",
              "backColor",
              "textAlign",
              "borderStyle",
              "pictureAlignment",
              "pictureSizeMode",
            ].includes(descriptor.field) && target?.values[descriptor.field]?.present,
        )
        .map((descriptor) => [descriptor.field, target!.values[descriptor.field].value]),
    );
  // The server supplies canonical values for legacy properties. Reuse those
  // values for display instead of decoding a second legacy schema in the UI.
  return {
    ...document,
    ...appearance(grid.form),
    controls: document.controls.map((control) => ({
      ...control,
      ...appearance(grid.controls[control.id]),
    })),
  };
}
