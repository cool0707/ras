// UserForm event handlers in the code-behind (forms/<Name>.bas).
//
// The catalogue (eventCatalogue.json) is xlflow's VBE-observed list of event
// declarations (internal/vba/userforms/events/catalogue.json).
// Ported from xlflow (https://github.com/harumiWeb/xlflow),
// MIT License, Copyright (c) 2026 harumiWeb.

import catalogue from "./eventCatalogue.json";

export interface EventDescriptor {
	name: string;
	signature: string;
	isDefault: boolean;
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
	/** Line of the `Sub` declaration. */
	line: number;
	/** First body line (where the cursor goes). */
	bodyLine: number;
}

/** Finds `[Private|Public] Sub <object>_<event>(` (case-insensitive). */
export function findHandler(text: string, objectName: string, eventName: string): HandlerLocation | undefined {
	const lines = text.split(/\r\n|\r|\n/);
	const name = `${objectName}_${eventName}`.toLowerCase();
	for (let line = 0; line < lines.length; line++) {
		const match = /^\s*(?:(?:Private|Public|Friend)\s+)?(?:Static\s+)?Sub\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(/i.exec(lines[line]);
		if (match && match[1].toLowerCase() === name) return { line, bodyLine: Math.min(line + 1, lines.length - 1) };
	}
	return undefined;
}

/** Indentation of existing procedure bodies, or `fallback`. */
export function bodyIndent(text: string, fallback: string): string {
	const lines = text.split(/\r\n|\r|\n/);
	for (let index = 0; index < lines.length; index++) {
		if (!/^\s*(?:(?:Private|Public|Friend)\s+)?(?:Sub|Function)\b/i.test(lines[index])) continue;
		for (const body of lines.slice(index + 1)) {
			if (!body.trim() || /^\s*End\s+(?:Sub|Function)\b/i.test(body)) continue;
			const indentation = /^[\t ]*/.exec(body)?.[0] ?? "";
			if (indentation) return indentation;
		}
	}
	return fallback;
}

/** Text to append to the code-behind for a new, empty handler. */
export function handlerText(existing: string, signature: string, indent: string, newline: string): string {
	const separator = !existing.trim() ? "" : /(\r\n|\r|\n)$/.test(existing) ? newline : newline + newline;
	return `${separator}${signature}${newline}${indent}${newline}End Sub${newline}`;
}
