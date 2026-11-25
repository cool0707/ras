// The events of the shapes and controls of a worksheet, for the sheet panel (sheetView.ts): the
// events an ActiveX control raises on a sheet (spec §12.2.1), and the macros of the RAS files a
// shape or a form control can run (`OnAction`).

import * as path from "path";
import { eventsForType, jsName } from "./userForm/events";
import type { EventDescriptor } from "./userForm/events";

/** What a form raises and a sheet does not, and the other way round (spec §12.2.1). */
const FORM_ONLY_EVENTS = ["Enter", "Exit", "BeforeUpdate", "AfterUpdate"];
const SHEET_ONLY_EVENTS = ["GotFocus", "LostFocus"];

/** The events of an ActiveX control on a sheet. */
export function sheetControlEvents(type: string, objectName: string): EventDescriptor[] {
	const events = eventsForType(type, objectName).filter((e) => !FORM_ONLY_EVENTS.some((f) => f.toLowerCase() === e.name.toLowerCase()));
	if (events.length === 0) {
		return [];
	}
	return [...events, ...SHEET_ONLY_EVENTS.map((name) => ({ name, signature: `Private Sub ${objectName}_${name}()`, isDefault: false }))];
}

/** The VBA module of a RAS file of `src/` (spec §12.1.1): `mod_util.ras` → `ModUtil`. */
export function moduleName(file: string): string {
	const base = path.basename(file).replace(/\.ras$/i, "");
	let name = base.replace(/(^|[_\-. ]+)(.)/gu, (_, _sep: string, c: string) => c.toUpperCase());
	if (/^[0-9]/.test(name)) {
		name = `Mod${name}`;
	}
	return name.slice(0, 31);
}

/** The functions of a RAS file that can be a macro: `export function f()` and `@macro`. */
export function macrosOf(text: string): { name: string; line: number }[] {
	const lines = text.split(/\r\n|\r|\n/);
	const out: { name: string; line: number }[] = [];
	for (let i = 0; i < lines.length; i++) {
		const m = /^\s*(export\s+)?function\s+([\p{L}_$][\p{L}\p{N}_$]*)\s*\(\s*\)/u.exec(lines[i]);
		if (m === null) {
			continue;
		}
		// `@macro` on a line of its own above it.
		let marked = false;
		for (let k = i - 1; k >= 0 && /^\s*@/.test(lines[k]); k--) {
			marked ||= /^\s*@macro\b/.test(lines[k]);
		}
		if (m[1] !== undefined || marked) {
			out.push({ name: m[2], line: i });
		}
	}
	return out;
}

/** The name of a new macro for an object: `chart1Click`, `ボタン1Click`, unused in `text`. */
export function newMacroName(text: string, object: string, shape: number | undefined): string {
	let base = object.replace(/[^\p{L}\p{N}_]/gu, "");
	if (base === "" || /^[0-9]/.test(base)) {
		base = `shape${shape ?? ""}${base}`;
	}
	base = `${jsName(base)}Click`;
	const taken = (name: string) => new RegExp(`(^|[^\\p{L}\\p{N}_$])(function|let|const)\\s+${name}(?![\\p{L}\\p{N}_$])`, "u").test(text);
	if (!taken(base)) {
		return base;
	}
	for (let i = 2; ; i++) {
		if (!taken(`${base}${i}`)) {
			return `${base}${i}`;
		}
	}
}

/**
 * The object part of the macro the compiler generates for the `@on("<name>.Click")` handlers of a
 * shape or a form control (spec §12.2.1), as Excel names a new macro: `ボタン 1` → `ボタン1`.
 */
export function shapeMacroObject(name: string): string {
	const kept = name.replace(/[^\p{L}\p{N}_]/gu, "");
	return /^\p{L}/u.test(kept) ? kept : `Shape${kept}`;
}

/** The macro a shape's `@on("<name>.Click")` handlers are called from: `Sheet1.ボタン1_Click`. */
export function shapeMacro(codeName: string, name: string): string {
	return `${codeName}.${shapeMacroObject(name)}_Click`;
}

/** Whether `macro` (as `OnAction` has it: `'Book.xlsm'!Sheet1.ボタン1_Click` too) is that macro. */
export function isShapeMacro(macro: string, codeName: string, name: string): boolean {
	return macro.replace(/^.*!/, "").toLowerCase() === shapeMacro(codeName, name).toLowerCase();
}

/** Text to append to a sheet's file for a new `@on("<name>.Click")` handler of a shape. */
export function shapeHandlerText(existing: string, name: string, indent: string, newline: string): string {
	const separator = !existing.trim() ? "" : /(\r\n|\r|\n)$/.test(existing) ? newline : newline + newline;
	return `${separator}@on(${JSON.stringify(`${name}.Click`)})${newline}function ${newMacroName(existing, name, undefined)}() {${newline}${indent}${newline}}${newline}`;
}
