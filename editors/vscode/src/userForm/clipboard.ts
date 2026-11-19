// Name and ID allocation for pasted controls (xlflow's clipboard.ts pasteOperation).
// Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License, Copyright (c) 2026 harumiWeb.

import { randomUUID } from "node:crypto";
import { EditError } from "./engine";
import type { DesignerDocument } from "./protocol";
import { controlMetadata } from "./toolbox";

/** Allocates fresh `control-<UUID>` IDs and conflict-free names for a paste. */
export function pasteAllocations(content: { controls: readonly { id: string; type: string }[] }, target: DesignerDocument, uuid: () => string = randomUUID) {
	const names = new Set([target.name, ...target.controls.map((c) => c.name)].map((n) => n.toLowerCase()));
	const ids = new Set(target.controls.map((c) => c.id));
	return content.controls.map((control) => {
		const prefix =
			controlMetadata.find((m) => m.type.toLowerCase() === control.type.toLowerCase())?.type ??
			(control.type.toLowerCase() === "page" ? "Page" : undefined);
		if (!prefix) throw new EditError("unsupportedCopy", "カスタム コントロールはコピーできません。");
		let suffix = 1;
		while (names.has(`${prefix}${suffix}`.toLowerCase())) suffix++;
		const name = `${prefix}${suffix}`;
		let id = `control-${uuid()}`;
		while (ids.has(id)) id = `control-${uuid()}`;
		names.add(name.toLowerCase());
		ids.add(id);
		return { sourceId: control.id, id, name };
	});
}

