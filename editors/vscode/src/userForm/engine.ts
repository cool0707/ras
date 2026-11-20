// Applies the Designer's semantic operations to a FormSpec JSON source.
//
// xlflow applies these operations in its Go language server
// (internal/vba/userforms/spec/edit). RAS has no such server support, so the
// same semantics are implemented here on the parsed JSON value. The result is
// re-serialised with the document's indentation and replaces only the changed
// span, so VS Code's own undo/redo and dirty state apply to every edit.
//
// Ported from xlflow (https://github.com/harumiWeb/xlflow),
// MIT License, Copyright (c) 2026 harumiWeb.

import {
	controlContracts,
	editableControlProperties,
	editableFormProperties,
	editableTabProperties,
	isEditorProperty,
	lookupControlContractByProgId,
} from "./contract";
import type { PropertyContract } from "./contract";
import type { JsonValue } from "./jsonSource";
import { formatJson } from "./jsonSource";
import { minimumEditHeight } from "./editHeight";
import type { ClipboardSnapshot, DesignerError, SemanticOperation, SpecTab } from "./protocol";
import type { ControlRecord, JsonObject, ParsedSpec } from "./spec";
import { errorKeys, flattenControls, isObject, parseSpec, validateSpec } from "./spec";

export class EditError extends Error {
	constructor(
		readonly code: string,
		message: string,
	) {
		super(message);
	}
	toDesignerError(): DesignerError {
		return { code: this.code, message: this.message };
	}
}

const fail = (code: string, message: string): never => {
	throw new EditError(code, message);
};
const tolerance = 1e-6;
const stable = (value: number) => Math.round(value * 1e6) / 1e6;

/** Working state: the mutable JSON root plus the flat control index. */
class Model {
	records: ControlRecord[] = [];
	constructor(readonly root: JsonObject) {
		this.reindex();
	}
	get form(): JsonObject {
		return this.root.form as JsonObject;
	}
	get controls(): JsonValue[] {
		return this.root.controls as JsonValue[];
	}
	reindex() {
		this.records = flattenControls(this.controls);
	}
	record(id: string): ControlRecord {
		const record = this.records.find((r) => r.control.id === id);
		if (!record) fail("UFE004", "コントロールが見つかりません。");
		if (typeof record!.raw.id !== "string") fail("UFE005", "id が明示されていないコントロールは編集できません。");
		return record!;
	}
	find(id: string | undefined): ControlRecord | undefined {
		return id ? this.records.find((r) => r.control.id === id) : undefined;
	}
	children(parentId: string, type?: string): ControlRecord[] {
		return this.records
			.filter((r) => (r.control.parentId ?? "") === parentId && (!type || r.control.type.toLowerCase() === type.toLowerCase()))
			.map((record, index) => ({ record, index }))
			.sort((a, b) => a.record.control.zIndex - b.record.control.zIndex || a.index - b.index)
			.map(({ record }) => record);
	}
	/** Removes a raw control object from whichever array (flat or nested) holds it. */
	removeRaw(raw: JsonObject) {
		const remove = (items: JsonValue[]): boolean => {
			const index = items.indexOf(raw);
			if (index >= 0) {
				items.splice(index, 1);
				return true;
			}
			return items.some((item) => isObject(item) && Array.isArray(item.controls) && remove(item.controls));
		};
		remove(this.controls);
	}
	append(raw: JsonObject) {
		this.controls.push(raw);
	}
}

const defaults: Record<string, [number, number]> = {
	label: [72, 18],
	textbox: [120, 18],
	combobox: [120, 18],
	listbox: [120, 72],
	commandbutton: [72, 24],
	checkbox: [72, 18],
	optionbutton: [72, 18],
	togglebutton: [72, 18],
	spinbutton: [18, 36],
	scrollbar: [120, 18],
	image: [72, 72],
	frame: [144, 108],
	multipage: [240, 180],
	tabstrip: [240, 48],
	page: [0, 0],
};

interface Geometry {
	record: ControlRecord;
	left: number;
	top: number;
	width: number;
	height: number;
	parentWidth: number;
	parentHeight: number;
}

const num = (value: unknown, fallback: number) => (typeof value === "number" && Number.isFinite(value) ? value : fallback);

function observed(record: ControlRecord, key: string): number | undefined {
	const value = isObject(record.raw.observed) ? record.raw.observed[key] : undefined;
	return typeof value === "number" ? value : undefined;
}

function geometry(model: Model, record: ControlRecord): Geometry {
	const [w, h] = defaults[record.control.type.toLowerCase()] ?? [72, 24];
	const [parentWidth, parentHeight] = parentBounds(model, record.control.parentId ?? "");
	return {
		record,
		left: num(record.control.left, observed(record, "left") ?? 0),
		top: num(record.control.top, observed(record, "top") ?? 0),
		width: num(record.control.width, observed(record, "width") ?? w),
		height: num(record.control.height, observed(record, "height") ?? h),
		parentWidth,
		parentHeight,
	};
}

function contentSize(type: string, width: number, height: number): [number, number] {
	const t = type.toLowerCase();
	const [l, top, r, b] = t === "frame" ? [0.75, 0.75, 0.75, 0.75] : t === "multipage" ? [1.5, 18, 1.5, 1.5] : [0, 0, 0, 0];
	return [Math.max(0, width - l - r), Math.max(0, height - top - b)];
}

function parentBounds(model: Model, parentId: string): [number, number] {
	if (!parentId) return [Math.max(0, num(model.form.width, 240)), Math.max(0, num(model.form.height, 180))];
	const parent = model.find(parentId);
	if (!parent) return [0, 0];
	if (parent.control.type.toLowerCase() === "page") {
		const owner = model.find(parent.control.parentId);
		if (!owner) return [0, 0];
		const g = geometry(model, owner);
		return contentSize(owner.control.type, g.width, g.height);
	}
	const g = geometry(model, parent);
	return contentSize(parent.control.type, g.width, g.height);
}

function requiredSize(model: Model, record: ControlRecord): [number, number] {
	const type = record.control.type.toLowerCase();
	if (type !== "frame" && type !== "multipage") return [1, 1];
	const owners = new Set([record.control.id]);
	const [l, t, r, b] = type === "frame" ? [0.75, 0.75, 0.75, 0.75] : [1.5, 18, 1.5, 1.5];
	if (type === "multipage")
		for (const page of model.records) if (page.control.parentId === record.control.id && page.control.type.toLowerCase() === "page") owners.add(page.control.id);
	let width = 1;
	let height = 1;
	for (const child of model.records) {
		if (!owners.has(child.control.parentId ?? "") || child.control.type.toLowerCase() === "page") continue;
		const g = geometry(model, child);
		width = Math.max(width, g.left + g.width + l + r);
		height = Math.max(height, g.top + g.height + t + b);
	}
	return [width, height];
}

function checkGeometry(model: Model, g: Geometry): string | undefined {
	const [rw, rh] = requiredSize(model, g.record);
	if (!(g.width >= 1 && g.height >= 1) || g.width + tolerance < rw || g.height + tolerance < rh) return "コントロールのサイズがデザイナーの最小値を下回ります。";
	if (
		!Number.isFinite(g.left) ||
		!Number.isFinite(g.top) ||
		g.left < 0 ||
		g.top < 0 ||
		g.width > g.parentWidth ||
		g.height > g.parentHeight ||
		g.left + g.width > g.parentWidth + tolerance ||
		g.top + g.height > g.parentHeight + tolerance
	)
		return "親コンテナーの内側に収まりません。";
	return undefined;
}

const isBuiltIn = (record: ControlRecord) => {
	const type = record.control.type.toLowerCase();
	if (!controlContracts[type]) return false;
	const progId = record.control.progId;
	return !progId || lookupControlContractByProgId(progId)?.type.toLowerCase() === type;
};

function setGeometryFields(model: Model, g: Geometry) {
	const current = geometry(model, g.record);
	for (const key of ["left", "top", "width", "height"] as const) if (current[key] !== g[key]) g.record.raw[key] = g[key];
}

// ---------------------------------------------------------------------------

function checkValue(field: string, contract: PropertyContract, value: unknown) {
	if (value === undefined) fail("UFE003", "プロパティの値がありません。");
	if (value === null) {
		if (contract.required) fail("UFE003", "必須のプロパティに null は設定できません。");
		return;
	}
	if (contract.editor === "font" || contract.editor === "color" || contract.editor === "picture") {
		if (!isObject(value)) fail("UFE003", "このプロパティにはオブジェクトを指定します。");
		for (const [name, member] of Object.entries(value as JsonObject)) {
			const child = contract.children?.[name];
			if (!child) fail("UFE003", `${field} に ${name} という項目はありません。`);
			if (member === null) continue;
			const ok =
				child!.valueType === "string"
					? typeof member === "string"
					: child!.valueType === "boolean"
						? typeof member === "boolean"
						: typeof member === "number" && Number.isFinite(member) && (child!.valueType !== "integer" || Number.isInteger(member));
			if (!ok) fail("UFE003", `${field}.${name} の値の型が正しくありません。`);
		}
		for (const [name, child] of Object.entries(contract.children ?? {}))
			if (child.required && !Object.hasOwn(value as JsonObject, name)) fail("UFE003", `${field}.${name} は必須です。`);
		return;
	}
	if (contract.valueType === "listItems") {
		const item = (v: unknown) => typeof v === "string" || (Array.isArray(v) && v.every((column) => typeof column === "string"));
		if (!Array.isArray(value) || !value.every(item)) fail("UFE003", `${field} には文字列か、文字列の配列（各列）の配列を指定します。`);
		return;
	}
	if (contract.valueType === "stringArray") {
		if (!Array.isArray(value) || !value.every((item) => typeof item === "string")) fail("UFE003", `${field} には文字列の配列を指定します。`);
		return;
	}
	const ok =
		contract.valueType === "any"
			? ["string", "number", "boolean"].includes(typeof value)
			: contract.valueType === "string"
				? typeof value === "string"
				: contract.valueType === "boolean"
					? typeof value === "boolean"
					: typeof value === "number" && Number.isFinite(value) && (contract.valueType !== "integer" || Number.isInteger(value));
	if (!ok) fail("UFE003", `${field} の値の型が正しくありません。`);
	if (contract.allowedValues?.length && !contract.allowedValues.includes(String(value))) fail("UFE003", `${field} は ${contract.allowedValues.join(", ")} のいずれかです。`);
}

const legacyNames: Record<string, string> = {
	font: "font",
	foreColor: "foreColor",
	backColor: "backColor",
	borderColor: "borderColor",
	textAlign: "textAlign",
	borderStyle: "borderStyle",
	pictureAlignment: "pictureAlignment",
	pictureSizeMode: "pictureSizeMode",
};

function setProperty(owner: JsonObject, field: string, contract: PropertyContract, value: JsonValue) {
	checkValue(field, contract, value);
	if (contract.editor === "picture" && (value === null || (isObject(value) && value.remove === true))) {
		// No picture is written as an absent field.
		delete owner[field];
	} else if (contract.editor === "list") {
		if (value === null) delete owner[field];
		else owner[field] = value;
		// A selection that the new list no longer contains becomes "none".
		const count = Array.isArray(value) ? value.length : 0;
		if (typeof owner.selectedIndex === "number" && owner.selectedIndex >= count) owner.selectedIndex = -1;
	} else if (contract.editor === "font" && isObject(value)) {
		// Font is partial: members not given keep their authored value.
		const current = isObject(owner.font) ? owner.font : isObject(owner.properties) && isObject(owner.properties.font) ? owner.properties.font : {};
		owner.font = { ...current, ...value };
	} else owner[field] = value;
	// Confirming a typed value migrates only the matching legacy property-bag key.
	if (isEditorProperty(field) && isObject(owner.properties)) {
		const legacy = Object.keys(owner.properties).find((key) => key.toLowerCase() === legacyNames[field]?.toLowerCase());
		if (legacy !== undefined) {
			delete owner.properties[legacy];
			if (Object.keys(owner.properties).length === 0) delete owner.properties;
		}
	}
}

function newControlObject(control: Record<string, unknown>, order: string[]): JsonObject {
	const result: JsonObject = {};
	for (const key of order) if (control[key] !== undefined) result[key] = control[key] as JsonValue;
	for (const key of Object.keys(control)) if (!order.includes(key) && control[key] !== undefined) result[key] = control[key] as JsonValue;
	return result;
}
const controlKeyOrder = ["id", "parentId", "name", "type", "caption", "left", "top", "width", "height", "zIndex", "selectedIndex", "tabs"];

function selectedAfterAdd(selected: unknown, count: number): number | undefined {
	return selected === -1 && count === 0 ? 0 : undefined;
}
function selectedAfterRemove(selected: unknown, removed: number, count: number): number | undefined {
	if (typeof selected !== "number") return;
	if (selected === removed) return count === 1 ? -1 : 0;
	if (removed < selected) return selected - 1;
	return undefined;
}
function selectedAfterReorder(selected: unknown, from: number, to: number): number | undefined {
	if (typeof selected !== "number" || from === to) return;
	if (selected === from) return to;
	if (from < selected && to >= selected) return selected - 1;
	if (from > selected && to <= selected) return selected + 1;
	return undefined;
}

/** Moves one sibling to `target` by renumbering the siblings' zIndex contiguously. */
function reorder(model: Model, id: string, target: number) {
	const record = model.record(id);
	const siblings = model.children(record.control.parentId ?? "", record.control.type.toLowerCase() === "page" ? "Page" : undefined);
	const from = siblings.indexOf(record);
	if (target < 0 || target >= siblings.length) fail("UFE003", "並べ替え先のインデックスが範囲外です。");
	if (from === target) return;
	siblings.splice(from, 1);
	siblings.splice(target, 0, record);
	const base = Math.min(0, ...siblings.map((s) => s.control.zIndex));
	siblings.forEach((sibling, index) => {
		if (sibling.control.zIndex !== base + index) sibling.raw.zIndex = base + index;
	});
	model.reindex();
}

function removeSubtree(model: Model, id: string, cascade: boolean) {
	const remove = new Set([id]);
	for (let changed = true; changed; ) {
		changed = false;
		for (const r of model.records)
			if (r.control.parentId && remove.has(r.control.parentId) && !remove.has(r.control.id)) {
				remove.add(r.control.id);
				changed = true;
			}
	}
	if (remove.size > 1 && !cascade) fail("UFE007", "このコントロールには子コントロールがあります。");
	for (const record of [...model.records].reverse()) if (remove.has(record.control.id)) model.removeRaw(record.raw);
	model.reindex();
}

function pageTarget(model: Model, id: string) {
	const page = model.record(id);
	if (page.control.type.toLowerCase() !== "page") fail("UFE004", "Page が見つかりません。");
	const owner = model.find(page.control.parentId);
	if (!owner || owner.control.type.toLowerCase() !== "multipage") fail("UFE005", "Page の親が MultiPage ではありません。");
	const pages = model.children(owner!.control.id, "Page");
	return { page, owner: owner!, pages, index: pages.indexOf(page) };
}

function applyOne(model: Model, op: SemanticOperation) {
	switch (op.type) {
		case "setGeometry": {
			if (op.controlId === undefined) {
				if (op.width !== undefined) model.form.width = op.width;
				if (op.height !== undefined) model.form.height = op.height;
				return;
			}
			const record = model.record(op.controlId);
			if (record.control.type.toLowerCase() === "page") fail("UFE005", "Page の位置とサイズは MultiPage から決まります。");
			for (const key of ["left", "top", "width", "height"] as const) if (op[key] !== undefined) record.raw[key] = op[key]!;
			return;
		}
		case "moveControl":
		case "resizeControl": {
			const record = model.record(op.controlId);
			if (op.type === "moveControl") Object.assign(record.raw, { left: op.left, top: op.top });
			else Object.assign(record.raw, { width: op.width, height: op.height });
			return;
		}
		case "setFormProperty": {
			const contract = editableFormProperties()[op.field];
			if (!contract) fail("UFE005", `${op.field} はデザイナーで編集できるプロパティではありません。`);
			setProperty(model.form, op.field, contract!, op.value as JsonValue);
			return;
		}
		case "setControlProperty": {
			const record = model.record(op.controlId);
			const contract = editableControlProperties(record.control.type, record.control.progId)[op.field];
			if (!contract) fail("UFE005", `${op.field} はこのコントロールで編集できるプロパティではありません。`);
			setProperty(record.raw, op.field, contract!, op.value as JsonValue);
			return;
		}
		case "addControl": {
			const control = op.control;
			if (!control.id || !control.name || !control.type) fail("UFE003", "追加するコントロールには id・name・type が必要です。");
			if (model.find(control.id)) fail("UFE008", "同じ ID のコントロールがすでにあります。");
			model.append(newControlObject(control as unknown as Record<string, unknown>, controlKeyOrder));
			model.reindex();
			return;
		}
		case "removeControl": {
			const record = model.record(op.controlId);
			if (record.control.type.toLowerCase() === "page") fail("UFE005", "Page はページの一覧から削除してください。");
			removeSubtree(model, op.controlId, op.cascade);
			return;
		}
		case "setParent": {
			const record = model.record(op.controlId);
			if (op.parentId) {
				const parent = model.record(op.parentId);
				const type = parent.control.type.toLowerCase();
				if (type !== "frame" && type !== "page") fail("UFE005", "このコンテナーには配置できません。");
				let ancestor: ControlRecord | undefined = parent;
				while (ancestor) {
					if (ancestor === record) fail("UFE005", "コントロールを自分自身の中には移動できません。");
					ancestor = model.find(ancestor.control.parentId);
				}
			}
			if (record.nested) {
				// Move a legacy nested item to the flat collection.
				model.removeRaw(record.raw);
				if (record.raw.zIndex === undefined) record.raw.zIndex = record.control.zIndex;
				model.append(record.raw);
			}
			if (op.parentId) record.raw.parentId = op.parentId;
			else delete record.raw.parentId;
			record.raw.left = op.left;
			record.raw.top = op.top;
			model.reindex();
			return;
		}
		case "addPage": {
			const control = op.control;
			const owner = model.record(control.parentId!);
			if (owner.control.type.toLowerCase() !== "multipage") fail("UFE005", "Page は MultiPage にだけ追加できます。");
			if (model.find(control.id)) fail("UFE008", "同じ ID のコントロールがすでにあります。");
			const pages = model.children(owner.control.id, "Page");
			const zIndex = pages.length ? pages[pages.length - 1].control.zIndex + 1 : 0;
			model.append(newControlObject({ ...control, zIndex }, controlKeyOrder));
			const selected = selectedAfterAdd(owner.raw.selectedIndex, pages.length);
			if (selected !== undefined) owner.raw.selectedIndex = selected;
			model.reindex();
			return;
		}
		case "removePage": {
			const { owner, pages, index } = pageTarget(model, op.controlId);
			const selected = selectedAfterRemove(owner.raw.selectedIndex, index, pages.length);
			if (selected !== undefined) owner.raw.selectedIndex = selected;
			removeSubtree(model, op.controlId, op.cascade);
			return;
		}
		case "reorderPage": {
			const { owner, pages, index } = pageTarget(model, op.controlId);
			if (op.index < 0 || op.index >= pages.length) fail("UFE003", "ページのインデックスが範囲外です。");
			const selected = selectedAfterReorder(owner.raw.selectedIndex, index, op.index);
			if (selected !== undefined) owner.raw.selectedIndex = selected;
			reorder(model, op.controlId, op.index);
			return;
		}
		case "addTab":
		case "removeTab":
		case "reorderTab":
		case "setTabProperty":
			applyTab(model, op);
			return;
		default:
			applyProductivity(model, op);
	}
}

function applyTab(model: Model, op: Extract<SemanticOperation, { type: "addTab" | "removeTab" | "reorderTab" | "setTabProperty" }>) {
	const record = model.record(op.controlId);
	if (record.control.type.toLowerCase() !== "tabstrip") fail("UFE005", "タブの操作は TabStrip にだけ行えます。");
	const raw = record.raw;
	if (raw.tabs !== undefined && raw.tabs !== null && !Array.isArray(raw.tabs)) fail("UFE005", "tabs は配列でなければなりません。");
	const tabs = Array.isArray(raw.tabs) ? raw.tabs : undefined;
	const count = tabs?.length ?? 0;
	if (op.type === "addTab") {
		const tab: JsonObject = {};
		for (const [key, value] of Object.entries(op.tab as SpecTab)) {
			const contract = editableTabProperties()[key];
			if (!contract) fail("UFE003", `タブに ${key} という項目はありません。`);
			checkValue(key, contract!, value);
			tab[key] = value as JsonValue;
		}
		const selected = selectedAfterAdd(raw.selectedIndex, count);
		if (tabs) tabs.push(tab);
		else raw.tabs = [tab];
		if (selected !== undefined) raw.selectedIndex = selected;
		return;
	}
	if (!tabs) return fail("UFE004", "tabs 配列がありません。");
	if (op.index < 0 || op.index >= count) fail("UFE003", "タブのインデックスが範囲外です。");
	if (op.type === "removeTab") {
		const selected = selectedAfterRemove(raw.selectedIndex, op.index, count);
		if (selected !== undefined) raw.selectedIndex = selected;
		tabs.splice(op.index, 1);
	} else if (op.type === "reorderTab") {
		if (op.toIndex < 0 || op.toIndex >= count) fail("UFE003", "タブのインデックスが範囲外です。");
		const selected = selectedAfterReorder(raw.selectedIndex, op.index, op.toIndex);
		if (selected !== undefined) raw.selectedIndex = selected;
		const [moved] = tabs.splice(op.index, 1);
		tabs.splice(op.toIndex, 0, moved);
	} else {
		const contract = editableTabProperties()[op.field];
		if (!contract) fail("UFE005", `タブの ${op.field} は編集できません。`);
		checkValue(op.field, contract!, op.value);
		const tab = tabs[op.index];
		if (!isObject(tab)) fail("UFE005", "タブがオブジェクトではありません。");
		(tab as JsonObject)[op.field] = op.value as JsonValue;
	}
}

// --- productivity (spec/edit/productivity.go) -------------------------------

function selected(model: Model, ids: string[]): Geometry[] {
	let parentId: string | undefined;
	const states = ids.map((id, index) => {
		const record = model.record(id);
		if (record.control.type.toLowerCase() === "page" || !isBuiltIn(record)) fail("UFE005", "Page やカスタム コントロールは対象にできません。");
		const parent = record.control.parentId ?? "";
		if (index === 0) parentId = parent;
		else if (parent !== parentId) fail("UFE005", "同じ親を持つコントロールを選択してください。");
		return geometry(model, record);
	});
	checkProductivityParent(model, parentId ?? "");
	return states;
}

function checkProductivityParent(model: Model, parentId: string) {
	if (!parentId) return;
	const parent = model.find(parentId);
	if (!parent) return fail("UFE004", "親コントロールが見つかりません。");
	const type = parent.control.type.toLowerCase();
	if (!isBuiltIn(parent) || (type !== "frame" && type !== "page")) fail("UFE005", "フォーム・Frame・Page の直下のコントロールを選択してください。");
}

function applyProductivity(model: Model, op: SemanticOperation) {
	switch (op.type) {
		case "moveControls": {
			const states = selected(
				model,
				op.moves.map((m) => m.controlId),
			);
			const dl = stable(op.moves[0].left - states[0].left);
			const dt = stable(op.moves[0].top - states[0].top);
			op.moves.slice(1).forEach((move, i) => {
				const s = states[i + 1];
				if (Math.abs(stable(move.left - s.left) - dl) > tolerance || Math.abs(stable(move.top - s.top) - dt) > tolerance)
					fail("UFE003", "選択したコントロールは同じ量だけ移動しなければなりません。");
			});
			let [minL, maxL, minT, maxT] = [-Infinity, Infinity, -Infinity, Infinity];
			for (const s of states) {
				if (!(s.width >= 1 && s.height >= 1) || s.width > s.parentWidth || s.height > s.parentHeight) fail("UFE003", "親コンテナーに収まらないコントロールがあります。");
				if (s.left < 0 || s.top < 0 || s.left + s.width > s.parentWidth || s.top + s.height > s.parentHeight)
					fail("UFE003", "すでに親コンテナーからはみ出しているコントロールがあります。");
				minL = Math.max(minL, -s.left);
				maxL = Math.min(maxL, s.parentWidth - s.left - s.width);
				minT = Math.max(minT, -s.top);
				maxT = Math.min(maxT, s.parentHeight - s.top - s.height);
			}
			const left = Math.max(minL, Math.min(dl, maxL));
			const top = Math.max(minT, Math.min(dt, maxT));
			for (const s of states) {
				s.left += left;
				s.top += top;
				setGeometryFields(model, s);
			}
			return;
		}
		case "alignControls": {
			const states = selected(model, op.controlIds);
			const anchor = states.find((s) => s.record.control.id === op.anchorId);
			if (!anchor) return fail("UFE003", "基準のコントロールが選択に含まれていません。");
			for (const s of states) {
				if (s === anchor) continue;
				const [oldLeft, oldTop] = [s.left, s.top];
				if (op.alignment === "left") s.left = anchor.left;
				if (op.alignment === "center") s.left = anchor.left + (anchor.width - s.width) / 2;
				if (op.alignment === "right") s.left = anchor.left + anchor.width - s.width;
				if (op.alignment === "top") s.top = anchor.top;
				if (op.alignment === "middle") s.top = anchor.top + (anchor.height - s.height) / 2;
				if (op.alignment === "bottom") s.top = anchor.top + anchor.height - s.height;
				s.left = oldLeft + stable(s.left - oldLeft);
				s.top = oldTop + stable(s.top - oldTop);
				const error = checkGeometry(model, s);
				if (error) fail("UFE003", error);
			}
			states.forEach((s) => setGeometryFields(model, s));
			return;
		}
		case "distributeControls": {
			const states = selected(model, op.controlIds);
			const order = model.children(states[0].record.control.parentId ?? "");
			const horizontal = op.axis === "horizontal";
			const ordered = order
				.filter((r) => op.controlIds.includes(r.control.id))
				.map((r) => states.find((s) => s.record === r)!)
				.sort((a, b) => (horizontal ? a.left - b.left : a.top - b.top));
			const pos = (s: Geometry) => (horizontal ? s.left : s.top);
			const size = (s: Geometry) => (horizontal ? s.width : s.height);
			const first = pos(ordered[0]);
			const last = ordered[ordered.length - 1];
			const total = ordered.reduce((sum, s) => sum + size(s), 0);
			let gap = (pos(last) + size(last) - first - total) / (ordered.length - 1);
			if (gap < -tolerance) fail("UFE003", "重ならずに並べる余白がありません。");
			gap = Math.max(0, gap);
			let preceding = 0;
			ordered.forEach((s, index) => {
				if (index > 0 && index < ordered.length - 1) {
					const displacement = stable(first + preceding + gap * index - pos(s));
					if (horizontal) s.left += displacement;
					else s.top += displacement;
				}
				const error = checkGeometry(model, s);
				if (error) fail("UFE003", error);
				preceding += size(s);
			});
			ordered.forEach((s) => setGeometryFields(model, s));
			return;
		}
		case "sameSizeControls": {
			const states = selected(model, op.controlIds);
			const anchor = states.find((s) => s.record.control.id === op.anchorId);
			if (!anchor) return fail("UFE003", "基準のコントロールが選択に含まれていません。");
			const [w, h] = [anchor.width, anchor.height];
			for (const s of states) {
				if (op.dimension !== "height") s.width = w;
				if (op.dimension !== "width") s.height = h;
				const error = checkGeometry(model, s);
				if (error) fail("UFE003", error);
			}
			states.forEach((s) => setGeometryFields(model, s));
			return;
		}
		case "removeControls": {
			selected(model, op.controlIds);
			for (const id of op.controlIds) if (model.find(id)) removeSubtree(model, id, op.cascade);
			return;
		}
		case "reorderControls": {
			const states = selected(model, op.controlIds);
			const current = model.children(states[0].record.control.parentId ?? "");
			const chosen = new Set(op.controlIds);
			let final = [...current];
			const isChosen = (r: ControlRecord) => chosen.has(r.control.id);
			if (op.command === "front") final = [...current.filter((r) => !isChosen(r)), ...current.filter(isChosen)];
			if (op.command === "back") final = [...current.filter(isChosen), ...current.filter((r) => !isChosen(r))];
			if (op.command === "forward")
				for (let i = final.length - 2; i >= 0; i--) if (isChosen(final[i]) && !isChosen(final[i + 1])) [final[i], final[i + 1]] = [final[i + 1], final[i]];
			if (op.command === "backward")
				for (let i = 1; i < final.length; i++) if (isChosen(final[i]) && !isChosen(final[i - 1])) [final[i], final[i - 1]] = [final[i - 1], final[i]];
			if (final.every((r, i) => r === current[i])) return;
			const base = Math.min(0, ...final.map((r) => r.control.zIndex));
			final.forEach((r, index) => {
				if (r.control.zIndex !== base + index) r.raw.zIndex = base + index;
			});
			model.reindex();
			return;
		}
		case "setTabOrder": {
			checkProductivityParent(model, op.parentId);
			const eligible = model.records.filter(
				(r) => (r.control.parentId ?? "") === op.parentId && r.control.type.toLowerCase() !== "page" && editableControlProperties(r.control.type, r.control.progId).tabIndex,
			);
			if (eligible.length !== op.controlIds.length || !op.controlIds.every((id) => eligible.some((r) => r.control.id === id)))
				fail("UFE003", "タブ順にはこの親の直下のコントロールをちょうど一度ずつ含めてください。");
			op.controlIds.forEach((id, index) => {
				model.record(id).raw.tabIndex = index;
			});
			return;
		}
		case "pasteControls":
			paste(model, op);
			return;
	}
	fail("UFE003", "未対応の操作です。");
}

// --- clipboard (spec/edit/clipboard.go) -------------------------------------

/** Copies the selected sibling roots and their descendants as a FormSpec snapshot. */
export function copyControls(text: string, controlIds: string[]): { snapshot: ClipboardSnapshot; hasPicturePaths: boolean } {
	const parsed = parseSpec(text);
	if (!parsed.ok) return fail("UFE002", parsed.message);
	const model = new Model(structuredClone(parsed.spec.root));
	if (!controlIds.length) fail("UFE003", "コピーするコントロールを選択してください。");
	const states = selected(model, controlIds);
	const parentId = states[0].record.control.parentId ?? "";
	const roots = model.children(parentId).filter((r) => controlIds.includes(r.control.id));
	const included = new Set(roots.map((r) => r.control.id));
	for (let changed = true; changed; ) {
		changed = false;
		for (const r of model.records)
			if (r.control.parentId && included.has(r.control.parentId) && !included.has(r.control.id)) {
				if (!isBuiltIn(r) && r.control.type.toLowerCase() !== "page") fail("UFE005", "カスタム コントロールはコピーできません。");
				included.add(r.control.id);
				changed = true;
			}
	}
	const flat: JsonObject[] = [];
	let hasPicturePaths = false;
	for (const r of [...roots, ...model.records.filter((r) => included.has(r.control.id) && !roots.includes(r))]) {
		const copy = structuredClone(r.raw);
		delete copy.controls;
		if (typeof copy.id !== "string") copy.id = r.control.id;
		if (roots.includes(r)) delete copy.parentId;
		else copy.parentId = r.control.parentId!;
		if (copy.zIndex === undefined) copy.zIndex = r.control.zIndex;
		if (isObject(copy.picture) && typeof copy.picture.path === "string") hasPicturePaths = true;
		flat.push(copy);
	}
	const snapshot: JsonObject = { ...parsed.spec.root, controls: flat };
	delete snapshot.warnings;
	return { snapshot: { format: "json", text: JSON.stringify(snapshot), controlIds: roots.map((r) => r.control.id) }, hasPicturePaths };
}

function paste(model: Model, op: Extract<SemanticOperation, { type: "pasteControls" }>) {
	let copied: JsonObject;
	try {
		copied = JSON.parse(op.snapshot.text) as JsonObject;
	} catch {
		return fail("UFE003", "クリップボードの内容が正しくありません。");
	}
	const source = new Model(copied);
	const roots = new Set(op.snapshot.controlIds);
	if (op.parentId) checkProductivityParent(model, op.parentId);
	const allocation = new Map(op.allocations.map((a) => [a.sourceId, a]));
	if (allocation.size !== source.records.length) fail("UFE003", "貼り付けるコントロールの割り当てが不足しています。");
	const names = new Set([String(model.form.name), ...model.records.map((r) => r.control.name)].map((n) => n.toLowerCase()));
	const ids = new Set(model.records.map((r) => r.control.id));
	for (const a of op.allocations) {
		if (ids.has(a.id) || names.has(a.name.toLowerCase())) fail("UFE008", "貼り付けるコントロールの ID か名前が既存のものと重複しています。");
		ids.add(a.id);
		names.add(a.name.toLowerCase());
	}
	const [parentWidth, parentHeight] = parentBounds(model, op.parentId);
	const rootStates = source.records.filter((r) => roots.has(r.control.id)).map((r) => geometry(source, r));
	let [minL, maxL, minT, maxT] = [-Infinity, Infinity, -Infinity, Infinity];
	const desired = new Map<string, { left: number; top: number }>();
	for (const s of rootStates) {
		if (!(s.width >= 1 && s.height >= 1) || s.width > parentWidth || s.height > parentHeight) fail("UFE003", "貼り付け先に収まらないコントロールがあります。");
		const left = s.left + op.offset.left;
		const top = s.top + op.offset.top;
		desired.set(s.record.control.id, { left, top });
		minL = Math.max(minL, -left);
		maxL = Math.min(maxL, parentWidth - left - s.width);
		minT = Math.max(minT, -top);
		maxT = Math.min(maxT, parentHeight - top - s.height);
	}
	if (minL > maxL || minT > maxT) fail("UFE003", "貼り付けるコントロールが貼り付け先に収まりません。");
	const dl = Math.max(minL, Math.min(0, maxL));
	const dt = Math.max(minT, Math.min(0, maxT));
	let maxZ = Math.max(-1, ...model.children(op.parentId).map((r) => r.control.zIndex));
	const rootOrder = source.children("").filter((r) => roots.has(r.control.id));
	const rootZ = new Map(rootOrder.map((r) => [r.control.id, ++maxZ]));
	for (const r of source.records) {
		const a = allocation.get(r.control.id);
		if (!a) return fail("UFE003", "貼り付けるコントロールの割り当てが不足しています。");
		const raw = structuredClone(r.raw);
		delete raw.controls;
		raw.id = a.id;
		raw.name = a.name;
		if (roots.has(r.control.id)) {
			if (op.parentId) raw.parentId = op.parentId;
			else delete raw.parentId;
			const d = desired.get(r.control.id)!;
			const s = rootStates.find((x) => x.record === r)!;
			if (d.left + dl !== s.left) raw.left = d.left + dl;
			if (d.top + dt !== s.top) raw.top = d.top + dt;
			raw.zIndex = rootZ.get(r.control.id)!;
		} else {
			const parent = allocation.get(r.control.parentId ?? "");
			if (!parent) return fail("UFE003", "クリップボードの親子関係が正しくありません。");
			raw.parentId = parent.id;
		}
		const ordered = newControlObject(raw as Record<string, unknown>, controlKeyOrder);
		model.append(ordered);
	}
	model.reindex();
}

// ---------------------------------------------------------------------------

export interface ApplyResult {
	text: string;
	/** IDs to select after the edit (paste / duplicate). */
	selectionIds?: string[];
}

/**
 * Applies a transaction. The edit is rejected when it introduces a validation
 * error that the original document did not have; nothing is changed then.
 */
export function applyOperations(text: string, operations: SemanticOperation[]): ApplyResult {
	const parsed = parseSpec(text);
	if (!parsed.ok) return fail("UFE002", parsed.message);
	return applyToSpec(parsed.spec, operations);
}

/** The font a control is drawn with: the typed field, else the legacy property bag's. */
function fontOf(raw: JsonObject): unknown {
	return raw.font ?? (isObject(raw.properties) ? raw.properties.font : undefined);
}
const multiLineOf = (raw: JsonObject) =>
	raw.multiLine === true || (isObject(raw.properties) && raw.properties.multiLine === true);
const sizing = (raw: JsonObject) => JSON.stringify([raw.height ?? null, fontOf(raw) ?? null]);

/**
 * A TextBox / ComboBox whose height or font the edit changed (or that it added) is
 * raised to the height its font needs, so that MSForms does not clip the text.
 * Controls the edit left alone keep what was written.
 */
function fitEditHeights(model: Model, before: Map<string, string>) {
	for (const record of model.records) {
		const { raw, control } = record;
		if (before.get(control.id) === sizing(raw)) continue;
		const minimum = minimumEditHeight(control.type, fontOf(raw), multiLineOf(raw));
		if (minimum === undefined) continue;
		const height = typeof raw.height === "number" ? raw.height : defaults[control.type.toLowerCase()]?.[1];
		if (height !== undefined && height < minimum) raw.height = minimum;
	}
}

export function applyToSpec(spec: ParsedSpec, operations: SemanticOperation[]): ApplyResult {
	const before = errorKeys(spec.issues);
	const model = new Model(structuredClone(spec.root));
	const sized = new Map(model.records.map((r) => [r.control.id, sizing(r.raw)]));
	for (const op of operations) applyOne(model, op);
	fitEditHeights(model, sized);
	const after = validateSpec(model.root).filter((i) => i.severity === "error" && !before.has(`${i.code}\0${i.field}`));
	if (after.length) fail(after[0].code, after[0].message);
	const result: ApplyResult = { text: formatJson(model.root, spec.text) };
	const last = operations[operations.length - 1];
	if (last?.type === "pasteControls") result.selectionIds = last.allocations.filter((a) => last.snapshot.controlIds.includes(a.sourceId)).map((a) => a.id);
	return result;
}
