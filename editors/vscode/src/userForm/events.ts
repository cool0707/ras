// UserForm event handlers in the code-behind: `@on("<object>.<event>")` functions in
// forms/<Name>.ras.
//
// The catalogue (eventCatalogue.json) is xlflow's VBE-observed list of event
// declarations (internal/vba/userforms/events/catalogue.json).
// Ported from xlflow (https://github.com/harumiWeb/xlflow),
// MIT License, Copyright (c) 2026 harumiWeb.

import catalogue from "./eventCatalogue.json";

export interface EventDescriptor {
	name: string;
	/** VBA declaration of the event procedure. */
	signature: string;
	isDefault: boolean;
}

/** camelCase of a host name, as the RAS compiler spells it (`KeyCode` → `keyCode`, `RGB` → `rgb`). */
export function jsName(host: string): string {
	const run = /^[A-Z]*/.exec(host)?.[0].length ?? 0;
	if (run === 0) return host;
	const lower = run > 1 && /[a-z]/.test(host.charAt(run)) ? run - 1 : run;
	return host.slice(0, lower).toLowerCase() + host.slice(lower);
}

const RAS_TYPES: Record<string, string> = { Integer: "short", Long: "int", Single: "float", Double: "double", String: "string", Boolean: "boolean", Variant: "any", Object: "object" };

/** RAS parameters of an event: `ByVal Cancel As MSForms.ReturnBoolean` → `cancel: MSForms.ReturnBoolean`. */
export function rasParams(signature: string): string {
	const inner = /\((.*)\)\s*$/.exec(signature)?.[1] ?? "";
	return inner
		.split(",")
		.map((p) => p.trim())
		.filter((p) => p)
		.map((p) => {
			const byval = /^ByVal\s+/i.test(p);
			const rest = p.replace(/^(ByVal|ByRef)\s+/i, "");
			const [name, type] = rest.split(/\s+As\s+/i);
			return `${byval ? "" : "byref "}${jsName(name.trim())}: ${RAS_TYPES[type?.trim()] ?? type?.trim() ?? "any"}`;
		})
		.join(", ");
}

interface CatalogueEntry extends EventDescriptor {
	type: string;
}

/** Events of a host type (`UserForm` or a built-in control type), with `{object}` filled in. */
export function eventsForType(type: string, objectName: string): EventDescriptor[] {
	return (catalogue as CatalogueEntry[])
		.filter((entry) => entry.type.toLowerCase() === type.toLowerCase())
		.map((entry) => ({ name: entry.name, signature: entry.signature.replace("{object}", objectName), isDefault: entry.isDefault }));
}

export interface EventTarget {
	/** Catalogue host type. */
	type: string;
	/** Name used in the procedure name (`UserForm` for the form itself). */
	objectName: string;
}

/**
 * Page events belong to the owning MultiPage; the form is `UserForm`.
 * Returns undefined for custom controls.
 */
export function eventTarget(
	controls: readonly { id: string; parentId?: string; name: string; type: string; progId?: string }[],
	targetId: string | undefined,
): EventTarget | undefined {
	if (!targetId) return { type: "UserForm", objectName: "UserForm" };
	let control = controls.find((c) => c.id === targetId);
	if (control?.type.toLowerCase() === "page") control = controls.find((c) => c.id === control!.parentId);
	if (!control) return undefined;
	const type = control.type;
	if (control.progId && control.progId.toLowerCase() !== `forms.${type.toLowerCase()}.1`) return undefined;
	return eventsForType(type, control.name).length ? { type, objectName: control.name } : undefined;
}

export interface HandlerLocation {
	/** Line of the `@on` decorator (or of the `Sub` declaration in VBA). */
	line: number;
	/** First body line (where the cursor goes). */
	bodyLine: number;
}

/** Finds the function decorated with `@on("<object>.<event>"…)` (case-insensitive). */
export function findHandler(text: string, objectName: string, eventName: string): HandlerLocation | undefined {
	const lines = text.split(/\r\n|\r|\n/);
	const wanted = `${objectName}.${eventName}`.toLowerCase();
	for (let line = 0; line < lines.length; line++) {
		const match = /^\s*@on\(\s*["']([^"']*)["']/.exec(lines[line]);
		if (!match || match[1].toLowerCase() !== wanted) continue;
		// The body starts after the line that opens the function.
		for (let open = line; open < lines.length; open++) {
			if (/\{\s*(\/\/.*)?$/.test(lines[open]) && /\bfunction\b/.test(lines.slice(line, open + 1).join("\n"))) {
				return { line, bodyLine: Math.min(open + 1, lines.length - 1) };
			}
		}
		return { line, bodyLine: line };
	}
	return undefined;
}

/** Indentation of existing function bodies, or `fallback`. */
export function bodyIndent(text: string, fallback: string): string {
	const lines = text.split(/\r\n|\r|\n/);
	for (let index = 0; index < lines.length; index++) {
		if (!/^\s*(export\s+)?function\b.*\{\s*$/.test(lines[index])) continue;
		for (const body of lines.slice(index + 1)) {
			if (!body.trim() || /^\s*\}/.test(body)) continue;
			const indentation = /^[\t ]*/.exec(body)?.[0] ?? "";
			if (indentation) return indentation;
		}
	}
	return fallback;
}

/** Name of a new handler: `btnOKClick`, `userFormInitialize`, unused in `existing`. */
export function handlerName(existing: string, objectName: string, eventName: string): string {
	const base = `${jsName(objectName)}${eventName}`;
	const taken = (name: string) => new RegExp(`\\b(function|let|const)\\s+${name}\\b`).test(existing);
	if (!taken(base)) return base;
	for (let i = 2; ; i++) if (!taken(`${base}${i}`)) return `${base}${i}`;
}

/** Text to append to the code-behind for a new, empty handler. */
export function handlerText(existing: string, objectName: string, event: EventDescriptor, indent: string, newline: string): string {
	const separator = !existing.trim() ? "" : /(\r\n|\r|\n)$/.test(existing) ? newline : newline + newline;
	const name = handlerName(existing, objectName, event.name);
	return `${separator}@on("${objectName}.${event.name}")${newline}function ${name}(${rasParams(event.signature)}) {${newline}${indent}${newline}}${newline}`;
}
