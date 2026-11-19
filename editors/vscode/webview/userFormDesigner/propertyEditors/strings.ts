/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
import type { DesignerStrings } from "../../../src/userForm/protocol";
import { propertyEditorStrings } from "../../../src/userForm/propertyEditorStrings";

export const editorStrings = propertyEditorStrings;
export function localizedEditorStrings(strings: DesignerStrings) {
  return { ...editorStrings, ...strings };
}
