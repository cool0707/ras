/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
import type { PropertyDescriptor, PropertyValue } from "../../../src/userForm/protocol";

export type EditorValue = PropertyValue;
export type EditorKind = "font" | "color" | "picture" | "enum" | "boolean";
export type EditorDescriptor = PropertyDescriptor;
export interface EditorPreview {
  controlId?: string;
  values: Record<string, unknown>;
}
export interface EditorSession {
  descriptor: EditorDescriptor;
  controlId?: string;
  version: number;
  epoch: number;
  value: EditorValue | undefined;
  title: string;
  returnFocus?: HTMLElement;
  picturePreview?: string;
}
export interface PictureSelection {
  token: string;
  name: string;
  dataURI: string;
  width?: number;
  height?: number;
  converted?: boolean;
}

export const objectValue = (value: unknown): Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value) ? { ...value } : {};
