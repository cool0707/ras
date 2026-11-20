// FormSpec v1 (forms/<Name>.form.json): parsing, normalisation to the flat
// control list, validation (the UFVxxx rules) and Property Grid metadata.
//
// Ported from xlflow's internal/vba/userforms/spec and spec/edit
// (https://github.com/harumiWeb/xlflow, MIT License, Copyright (c) 2026 harumiWeb).

import {
	colorError,
	controlProperties,
	documentProperties,
	editableControlProperties,
	editableFormProperties,
	editableTabProperties,
	formObservedProperties,
	formProperties,
	lookupControlContract,
	lookupControlContractByProgId,
	lookupProperty,
	propertyDescriptor,
	tabProperties,
	commonControlProperties,
	controlContracts,
} from "./contract";
import type { PropertyContract, ValueType } from "./contract";
import { minimumEditHeight } from "./editHeight";
import type { JsonNode, JsonValue } from "./jsonSource";
import { JsonSyntaxError, nodeAtPath, offsetToPosition, parseJson } from "./jsonSource";
import type {
	DesignerProblem,
	PreviewDocument,
	PropertyGridData,
	PropertyState,
	PropertyTarget,
	SpecControl,
} from "./protocol";

export type JsonObject = { [key: string]: JsonValue };
export const FORM_KIND = "ras.userform";
export const ACCEPTED_KINDS = ["ras.userform", "xlflow.userform"];

export interface ValidationIssue {
	code: string;
	severity: "error" | "warning";
	message: string;
	field: string;
	suggestion?: string;
}

export const isObject = (value: unknown): value is JsonObject =>
	typeof value === "object" && value !== null && !Array.isArray(value);

/** A control of the flat model together with the JSON object it came from. */
export interface ControlRecord {
	control: SpecControl & { zIndex: number };
	raw: JsonObject;
	/** Field path of the raw object, e.g. `controls[2]` or `controls[0].controls[1]`. */
	path: string;
	/** True when the control sits in a legacy nested `controls` array. */
	nested: boolean;
}

export interface ParsedSpec {
	text: string;
	root: JsonObject;
	node: JsonNode;
	form: JsonObject;
	records: ControlRecord[];
	issues: ValidationIssue[];
}

export type ParseResult = { ok: true; spec: ParsedSpec } | { ok: false; message: string; offset?: number };

/** Parses the source and checks just enough shape for the Designer to render it. */
export function parseSpec(text: string): ParseResult {
	let parsed: { value: JsonValue; node: JsonNode };
	try {
		parsed = parseJson(text);
	} catch (error) {
		if (error instanceof JsonSyntaxError) {
			const at = offsetToPosition(text, error.offset);
			return { ok: false, message: `JSON の構文エラー（${at.line + 1} 行 ${at.character + 1} 列）: ${error.message}`, offset: error.offset };
		}
		throw error;
	}
	const root = parsed.value;
	if (!isObject(root)) return { ok: false, message: "FormSpec のルートはオブジェクトでなければなりません。" };
	if (!isObject(root.form)) return { ok: false, message: "form オブジェクトがありません。" };
	if (typeof root.form.name !== "string" || !root.form.name.trim()) return { ok: false, message: "form.name がありません。" };
	if (!Array.isArray(root.controls)) return { ok: false, message: "controls 配列がありません。" };
	const records = flattenControls(root.controls);
	return { ok: true, spec: { text, root, node: parsed.node, form: root.form, records, issues: validateSpec(root) } };
}

/** Flattens legacy nested `controls`, fills IDs and zIndex like NormalizeFormSpec. */
export function flattenControls(controls: JsonValue[]): ControlRecord[] {
	const used = new Set<string>();
	let next = 1;
	const uniqueId = (name: string) => {
		const base = (name.trim() || "control").replace(/ /g, "_").toLowerCase();
		if (!used.has(base)) {
			used.add(base);
			return base;
		}
		for (;;) {
			const id = `${base}_${String(next++).padStart(3, "0")}`;
			if (!used.has(id)) {
				used.add(id);
				return id;
			}
		}
	};
	const result: ControlRecord[] = [];
	const visit = (items: JsonValue[], path: string, inherited: string | undefined, nested: boolean) => {
		items.forEach((item, index) => {
			if (!isObject(item)) return;
			const name = typeof item.name === "string" ? item.name : "";
			const rawId = typeof item.id === "string" ? item.id.trim() : "";
			const id = rawId || uniqueId(name);
			if (rawId) used.add(rawId);
			const parentId = typeof item.parentId === "string" && item.parentId.trim() ? item.parentId.trim() : inherited;
			const control = { ...(item as unknown as SpecControl), id, name, type: typeof item.type === "string" ? item.type : "" } as SpecControl & {
				zIndex: number;
			};
			delete (control as { controls?: unknown }).controls;
			if (parentId) control.parentId = parentId;
			else delete control.parentId;
			control.zIndex = typeof item.zIndex === "number" ? item.zIndex : index;
			result.push({ control, raw: item, path: `${path}[${index}]`, nested });
			if (Array.isArray(item.controls)) visit(item.controls, `${path}[${index}].controls`, id, true);
		});
	};
	visit(controls, "controls", undefined, false);
	return result;
}

export function previewDocument(spec: ParsedSpec): PreviewDocument {
	const form = spec.form;
	const pick = <T>(value: unknown, guard: (v: unknown) => boolean): T | undefined => (guard(value) ? (value as T) : undefined);
	const num = (v: unknown) => typeof v === "number" && Number.isFinite(v);
	return {
		form: {
			name: String(form.name),
			caption: pick<string>(form.caption, (v) => typeof v === "string"),
			width: pick<number>(form.width, num),
			height: pick<number>(form.height, num),
			font: pick(form.font, isObject),
			foreColor: pick(form.foreColor, (v) => colorError(v) === undefined),
			backColor: pick(form.backColor, (v) => colorError(v) === undefined),
			borderStyle: pick<number>(form.borderStyle, num),
			borderColor: pick(form.borderColor, (v) => colorError(v) === undefined),
			specialEffect: pick<number>(form.specialEffect, num),
		},
		controls: spec.records.map(({ control }) => {
			const copy = { ...control };
			for (const key of ["left", "top", "width", "height"] as const) if (!num(copy[key])) delete copy[key];
			for (const key of ["foreColor", "backColor"] as const) if (copy[key] !== undefined && colorError(copy[key])) delete copy[key];
			if (copy.font !== undefined && !isObject(copy.font)) delete copy.font;
			if (copy.list !== undefined && !valueMatchesType(copy.list, "listItems")) delete copy.list;
			return copy;
		}),
	};
}

/** Validation issues with their source range, for the Problems panel. */
export function problems(spec: ParsedSpec, issues: ValidationIssue[] = spec.issues): DesignerProblem[] {
	return issues.map((issue) => {
		const node = nodeAtPath(spec.node, issue.field);
		const start = node?.keyStart ?? node?.start ?? 0;
		const end = node?.end ?? start;
		return {
			code: issue.code,
			severity: issue.severity,
			message: issue.message,
			field: issue.field,
			...(issue.suggestion ? { suggestion: issue.suggestion } : {}),
			range: { start: offsetToPosition(spec.text, start), end: offsetToPosition(spec.text, end) },
		};
	});
}

// ---------------------------------------------------------------------------
// Validation (spec.go validateRawFormSpec and friends)
// ---------------------------------------------------------------------------

const issue = (code: string, severity: "error" | "warning", message: string, field: string, suggestion?: string): ValidationIssue => ({
	code,
	severity,
	message,
	field,
	...(suggestion ? { suggestion } : {}),
});

const isNumber = (v: unknown) => typeof v === "number" && Number.isFinite(v);
const isInteger = (v: unknown) => isNumber(v) && Number.isInteger(v);

function valueMatchesType(value: unknown, type: ValueType): boolean {
	if (value === null) return true;
	switch (type) {
		case "any":
			return true;
		case "string":
			return typeof value === "string";
		case "number":
			return isNumber(value);
		case "integer":
			return isInteger(value);
		case "boolean":
			return typeof value === "boolean";
		case "stringArray":
			return Array.isArray(value) && value.every((item) => typeof item === "string");
		case "listItems":
			return Array.isArray(value) && value.every((item) => typeof item === "string" || (Array.isArray(item) && item.every((column) => typeof column === "string")));
		case "object":
			return isObject(value);
		case "objectArray":
			return Array.isArray(value) && value.every(isObject);
	}
}

const typeLabel: Record<ValueType, string> = {
	string: "文字列",
	number: "数値",
	integer: "整数",
	boolean: "真偽値",
	any: "任意の値",
	stringArray: "文字列の配列",
	listItems: "文字列か、文字列の配列（各列）の配列",
	object: "オブジェクト",
	objectArray: "オブジェクトの配列",
};

const join = (base: string, key: string) => (base ? `${base}.${key}` : key);

function lookupRaw(root: JsonObject, key: string): [JsonValue, boolean] {
	if (Object.hasOwn(root, key)) return [root[key], true];
	const found = Object.keys(root).find((k) => k.trim().toLowerCase() === key.toLowerCase());
	return found === undefined ? [null, false] : [root[found], true];
}

function supportSuggestion(level: string): string | undefined {
	switch (level) {
		case "best-effort":
			return "ベストエフォートの項目です。生成されたフォームを Excel で確認してください。";
		case "observed-only":
			return "観測値として扱われ、ビルドでの再現は保証されません。";
		case "custom/unchecked":
			return "カスタム コントロールは受け付けますが、種類固有の内容は検証できません。";
	}
	return undefined;
}

function validateObject(root: JsonObject, properties: Record<string, PropertyContract>, path: string): ValidationIssue[] {
	const issues: ValidationIssue[] = [];
	for (const key of Object.keys(root).sort()) {
		const value = root[key];
		const field = join(path, key);
		const contract = lookupProperty(properties, key);
		if (!contract) {
			if (path === "form" && key.toLowerCase() === "build") {
				issues.push(issue("UFV001", "error", "form.build はスキーマから削除されました。", field, "form.build.caption は form.caption に、clientWidth/clientHeight は form.width/form.height に移してください。"));
				continue;
			}
			issues.push(issue("UFV001", "error", `${field} は FormSpec で定義されていない項目です。`, field, "項目を削除するか、カスタムの値は properties の下に移してください。"));
			continue;
		}
		if (!valueMatchesType(value, contract.valueType)) {
			issues.push(issue("UFV002", "error", `${field} は${typeLabel[contract.valueType]}でなければなりません。`, field));
			continue;
		}
		if (value !== null && contract.allowedValues?.length && !contract.allowedValues.includes(String(value)))
			issues.push(issue("UFV003", "error", `${field} は次のいずれかでなければなりません: ${contract.allowedValues.join(", ")}`, field));
		if (typeof value === "number" && ((contract.minimum !== undefined && value < contract.minimum) || (contract.maximum !== undefined && value > contract.maximum)))
			issues.push(issue("UFV021", "error", `${field} は ${contract.minimum}〜${contract.maximum} でなければなりません。`, field));
		if (value !== null && contract.children && isObject(value)) issues.push(...validateObject(value, contract.children, field));
		if (contract.support !== "supported" && contract.support !== "snapshot-only")
			issues.push(issue("UFV013", "warning", `${field} のサポートは ${contract.support} です。`, field, supportSuggestion(contract.support)));
	}
	for (const key of Object.keys(properties).sort()) {
		if (properties[key].required && !lookupRaw(root, key)[1]) issues.push(issue("UFV004", "error", `${join(path, key)} は必須です。`, join(path, key)));
	}
	return issues;
}

function validFormDimension(value: number): boolean {
	return Number.isFinite(value) && value >= 0 && Math.round((value * 2540) / 72) <= 2 ** 31 - 1;
}

function validateFormDimensions(form: JsonObject): ValidationIssue[] {
	const issues: ValidationIssue[] = [];
	const check = (object: JsonObject, key: string, path: string) => {
		const [value, present] = lookupRaw(object, key);
		if (present && isNumber(value) && !validFormDimension(value as number))
			issues.push(issue("UFV016", "error", `${path} は 0 以上の有限の値で、HIMETRIC に変換できる範囲でなければなりません。`, path));
	};
	check(form, "width", "form.width");
	check(form, "height", "form.height");
	const [observed] = lookupRaw(form, "observed");
	if (isObject(observed))
		for (const key of ["width", "height", "insideWidth", "insideHeight", "clientWidth", "clientHeight"]) check(observed, key, `form.observed.${key}`);
	return issues;
}

function validateFont(font: unknown, path: string): ValidationIssue[] {
	if (!isObject(font)) return [];
	const issues: ValidationIssue[] = [];
	const size = font.size;
	if (typeof size === "number" && !(size > 0 && size <= 3276.75))
		issues.push(issue("UFV021", "error", "フォント サイズは 0 より大きく 3276.75 ポイント以下でなければなりません。", `${path}.size`));
	const charset = font.charset;
	if (typeof charset === "number" && !(charset >= 0 && charset <= 255))
		issues.push(issue("UFV021", "error", "フォントの charset は 0〜255 の整数でなければなりません。", `${path}.charset`));
	const name = font.name;
	if (typeof name === "string" && (new TextEncoder().encode(name).length > 255 || name.includes("\0")))
		issues.push(issue("UFV021", "error", "フォント名は 255 バイト以下で、NUL を含めることはできません。", `${path}.name`));
	return issues;
}

function validateTyped(object: JsonObject, path: string, properties: Record<string, PropertyContract>, type?: string): ValidationIssue[] {
	const issues = validateFont(object.font, `${path}.font`);
	for (const name of ["foreColor", "backColor"]) {
		if (object[name] === undefined || object[name] === null) continue;
		if (!properties[name]) continue;
		const error = colorError(object[name]);
		if (error) issues.push(issue("UFV021", "error", error, `${path}.${name}`));
	}
	if (type && isObject(object.properties)) {
		for (const name of ["font", "foreColor", "backColor", "borderColor", "textAlign", "borderStyle", "pictureAlignment", "pictureSizeMode"]) {
			if (object[name] === undefined || object[name] === null) continue;
			const [legacy, present] = lookupRaw(object.properties, name);
			if (present && !legacyEqual(name, object[name], legacy))
				issues.push(issue("UFV021", "error", "型付きのプロパティと properties の値が矛盾しています。", `${path}.${name}`));
		}
	}
	return issues;
}

function legacyEqual(name: string, typed: JsonValue, legacy: JsonValue): boolean {
	if ((name === "foreColor" || name === "backColor" || name === "borderColor") && isObject(typed) && typeof legacy === "number") {
		const ole = colorToOle(typed);
		if (ole === undefined) return false;
		return legacy === ole || legacy === (ole > 0x7fffffff ? ole - 2 ** 32 : ole);
	}
	return JSON.stringify(typed) === JSON.stringify(legacy);
}

export function colorToOle(color: JsonObject): number | undefined {
	if (colorError(color)) return undefined;
	const value = String(color.value);
	if (color.kind === "rgb") {
		const n = parseInt(value.slice(1), 16);
		return ((n & 255) << 16) | (n & 0xff00) | (n >> 16);
	}
	if (color.kind === "system") return 0x80000000 + (systemColorIndex(value) ?? 0);
	return parseInt(value.slice(2), 16);
}
function systemColorIndex(name: string): number | undefined {
	const table = ["scrollBar", "desktop", "activeCaption", "inactiveCaption", "menu", "window", "windowFrame", "menuText", "windowText", "captionText", "activeBorder", "inactiveBorder", "appWorkspace", "highlight", "highlightText", "buttonFace", "buttonShadow", "grayText", "buttonText", "inactiveCaptionText", "buttonHighlight", "buttonDarkShadow", "buttonLight", "infoText", "infoBackground", "", "hotLight", "gradientActiveCaption", "gradientInactiveCaption", "menuHighlight", "menuBar"];
	const index = table.indexOf(name);
	return index < 0 ? undefined : index;
}

interface RawRef {
	path: string;
	id: string;
	parentId: string;
	name: string;
	type: string;
	progId: string;
	selectedIndex?: number;
	childrenKnown: boolean;
}

function validateControls(items: JsonValue, path: string, inherited: string): [RawRef[], ValidationIssue[]] {
	if (!Array.isArray(items)) return [[], []];
	const refs: RawRef[] = [];
	const issues: ValidationIssue[] = [];
	items.forEach((item, index) => {
		const itemPath = `${path}[${index}]`;
		if (!isObject(item)) {
			issues.push(issue("UFV002", "error", `${itemPath} はオブジェクトでなければなりません。`, itemPath));
			return;
		}
		const [r, i] = validateControl(item, itemPath, inherited);
		refs.push(...r);
		issues.push(...i);
	});
	return [refs, issues];
}

const str = (value: unknown) => (typeof value === "string" ? value.trim() : "");

function validateControl(control: JsonObject, path: string, inherited: string): [RawRef[], ValidationIssue[]] {
	const type = str(control.type);
	const progId = str(control.progId);
	const issues = validateControlProperties(control, type, progId, path);
	const ref: RawRef = {
		path,
		id: str(control.id),
		parentId: str(control.parentId) || inherited,
		name: str(control.name),
		type,
		progId,
		childrenKnown: false,
	};
	const [selected, hasSelected] = lookupRaw(control, "selectedIndex");
	if (hasSelected && isInteger(selected)) ref.selectedIndex = selected as number;
	const refs = [ref];
	if (isObject(control.observed)) issues.push(...validateObservedControl(control.observed, type, progId, `${path}.observed`));
	if (control.controls !== undefined) {
		ref.childrenKnown = true;
		issues.push(issue("UFV013", "warning", `${path}.controls は旧形式の入れ子のコントロール構造です。`, `${path}.controls`, "parentId を使ったフラットな controls 配列を使ってください。"));
		const [childRefs, childIssues] = validateControls(control.controls, `${path}.controls`, ref.id);
		refs.push(...childRefs);
		issues.push(...childIssues);
	}
	return [refs, issues];
}

function genericCustom(properties: Record<string, PropertyContract>) {
	const make = (valueType: ValueType, description: string): PropertyContract => ({ valueType, required: false, support: "custom/unchecked", description, authoring: false });
	properties.caption = make("string", "Generic custom control caption.");
	properties.text = make("string", "Generic custom control text.");
	properties.value = make("any", "Generic custom control value.");
}

function knownControlPropertyName(name: string): boolean {
	return (
		lookupProperty(commonControlProperties, name) !== undefined ||
		Object.values(controlContracts).some((c) => lookupProperty(c.properties, name) !== undefined)
	);
}

function markUnsupported(issues: ValidationIssue[], type: string) {
	for (const item of issues) {
		if (item.code !== "UFV001") continue;
		const name = item.field.slice(item.field.lastIndexOf(".") + 1);
		if (!knownControlPropertyName(name)) continue;
		item.code = "UFV005";
		item.message = `${item.field} はコントロールの種類 ${type} では使えません。`;
		item.suggestion = "プロパティを削除するか、それを持つ種類のコントロールを使ってください。";
	}
}

function validateControlProperties(control: JsonObject, type: string, progId: string, path: string): ValidationIssue[] {
	const builtIn = lookupControlContract(type);
	let allowed: Record<string, PropertyContract> = { ...commonControlProperties };
	if (builtIn) allowed = controlProperties(type);
	else if (progId && !lookupControlContractByProgId(progId)) genericCustom(allowed);
	const issues = validateObject(control, allowed, path);
	const [picture, hasPicture] = lookupRaw(control, "picture");
	if (hasPicture) issues.push(...validatePicture(picture, type, `${path}.picture`));
	const [tabs, hasTabs] = lookupRaw(control, "tabs");
	if (hasTabs) issues.push(...validateTabs(tabs, `${path}.tabs`));
	issues.push(...validateSelectedIndex(control, type, path));
	if (builtIn) {
		markUnsupported(issues, type);
		issues.push(...validateTyped(control, path, allowed, type));
		for (const name of ["textAlign", "borderStyle", "pictureAlignment", "pictureSizeMode"]) {
			const contract = allowed[name];
			const value = control[name];
			if (contract && value !== undefined && value !== null && contract.allowedValues && !contract.allowedValues.includes(String(value)))
				issues.push(issue("UFV021", "error", "値がプロパティの列挙範囲の外にあります。", `${path}.${name}`));
		}
		if (typeof control.passwordChar === "string" && control.passwordChar.length > 1)
			issues.push(issue("UFV021", "error", "passwordChar は 1 文字にしてください。", `${path}.passwordChar`));
		// AddItem and List(row, column) fill at most 10 columns of a list not bound to cells.
		if (Array.isArray(control.list))
			control.list.forEach((item, index) => {
				if (Array.isArray(item) && (item.length < 1 || item.length > 10))
					issues.push(issue("UFV021", "error", "list の項目の列は 1〜10 個にしてください。", `${path}.list[${index}]`));
			});
		if (typeof control.columnWidths === "string" && control.columnWidths.trim() && !/^\s*(-1|[\d.]+\s*(pt|in|cm)?)?\s*(;\s*(-1|[\d.]+\s*(pt|in|cm)?)?\s*)*$/i.test(control.columnWidths))
			issues.push(issue("UFV021", "error", "columnWidths は ; で区切った幅（pt、in、cm の単位を付けられます）にしてください。", `${path}.columnWidths`));
		// MSForms refuses AddItem on a list that comes from cells.
		if (typeof control.rowSource === "string" && control.rowSource && Array.isArray(control.list) && control.list.length > 0)
			issues.push(issue("UFV021", "error", "rowSource を指定したコントロールには list を書けません（項目はセル範囲から読み込まれます）。", `${path}.list`));
		if ((type === "SpinButton" || type === "ScrollBar") && typeof control.value === "number") {
			const min = typeof control.min === "number" ? control.min : 0;
			const max = typeof control.max === "number" ? control.max : type === "SpinButton" ? 100 : 32767;
			const [lo, hi] = [Math.min(min, max), Math.max(min, max)];
			if (control.value < lo || control.value > hi)
				issues.push(issue("UFV021", "error", `value は min と max の間（${lo}〜${hi}）にしてください。`, `${path}.value`));
		}
		const font = control.font ?? (isObject(control.properties) ? control.properties.font : undefined);
		const multiLine = control.multiLine === true || (isObject(control.properties) && control.properties.multiLine === true);
		const minimum = minimumEditHeight(type, font, multiLine);
		if (minimum !== undefined && typeof control.height === "number" && control.height < minimum)
			issues.push(
				issue(
					"UFV022",
					"warning",
					`${path}.height（${control.height} pt）はフォントに対して低すぎます。Excel では文字が欠け、端に線が残ります。`,
					`${path}.height`,
					`${minimum} pt 以上にしてください（デザイナーでフォントやサイズを変えると自動で調整されます）。`,
				),
			);
	}
	if (type && progId) {
		const byProgId = lookupControlContractByProgId(progId);
		if (byProgId && byProgId.type.toLowerCase() !== type.toLowerCase())
			issues.push(issue("UFV012", "error", `${path}.progId "${progId}" は ${byProgId.type} 用で、${type} 用ではありません。`, `${path}.progId`, "type に合う ProgID を使うか、type を ProgID に合わせてください。"));
	}
	if (type && !builtIn) {
		const unsupported = control.unsupported;
		if (type.toLowerCase() === "control" && !progId && Array.isArray(unsupported) && unsupported.includes("controlType"))
			issues.push(issue("UFV015", "warning", `${path} は種類を復元できなかったコントロールのスナップショット用の代替です。`, `${path}.type`));
		else if (!progId)
			issues.push(issue("UFV006", "error", `${path}.type "${type}" は対応する組み込みコントロールの種類ではありません。`, `${path}.type`, "組み込みの種類を使うか、カスタムの progId を指定してください。"));
		else if (!lookupControlContractByProgId(progId))
			issues.push(issue("UFV014", "warning", `${path} はカスタム コントロール "${type}"（ProgID "${progId}"）で、内容を検証できません。`, `${path}.progId`));
	}
	return issues;
}

function validatePicture(value: JsonValue, type: string, path: string): ValidationIssue[] {
	if (type.toLowerCase() !== "image") return [issue("UFV005", "error", `${path} はコントロールの種類 ${type} では使えません。`, path)];
	if (!isObject(value)) return [issue("UFV005", "error", `${path} はオブジェクトでなければなりません。`, path, "{ \"path\": ... } または { \"remove\": true } を使ってください。")];
	const issues: ValidationIssue[] = [];
	const keys = Object.keys(value).map((k) => k.trim().toLowerCase());
	for (const key of Object.keys(value))
		if (!["path", "remove"].includes(key.trim().toLowerCase())) issues.push(issue("UFV001", "error", `${path}.${key} は picture では定義されていない項目です。`, `${path}.${key}`));
	const hasPath = keys.includes("path");
	const hasRemove = keys.includes("remove");
	if (hasPath && hasRemove) issues.push(issue("UFV005", "error", `${path} には path と remove のどちらか一方だけを指定します。`, path));
	else if (hasPath) {
		if (typeof value.path !== "string" || !value.path.trim()) issues.push(issue("UFV005", "error", `${path}.path は空でない文字列でなければなりません。`, `${path}.path`));
	} else if (hasRemove) {
		if (value.remove !== true) issues.push(issue("UFV005", "error", `${path}.remove は true でなければなりません。`, `${path}.remove`));
	} else issues.push(issue("UFV005", "error", `${path} には path または remove: true が必要です。`, path));
	return issues;
}

function validateTabs(value: JsonValue, path: string): ValidationIssue[] {
	if (!Array.isArray(value)) return [];
	const issues: ValidationIssue[] = [];
	const seen: string[] = [];
	value.forEach((item, index) => {
		if (!isObject(item)) return;
		const itemPath = `${path}[${index}]`;
		issues.push(...validateObject(item, tabProperties, itemPath));
		const name = str(item.name);
		if (!name) return;
		const first = seen.find((s) => s.toLowerCase() === name.toLowerCase());
		if (first) {
			issues.push(issue("UFV019", "error", `${itemPath}.name はこの TabStrip の前のタブ名 "${first}" と重複しています。`, `${itemPath}.name`, "TabStrip 内で大文字小文字を区別せず一意な名前を使ってください。"));
			return;
		}
		seen.push(name);
	});
	return issues;
}

function validateSelectedIndex(values: JsonObject, type: string, path: string): ValidationIssue[] {
	const t = type.toLowerCase();
	if (!["multipage", "tabstrip", "combobox", "listbox"].includes(t)) return [];
	const [value, present] = lookupRaw(values, "selectedIndex");
	if (!present || !isInteger(value)) return [];
	const index = value as number;
	let invalid = index < -1;
	if (t === "tabstrip") {
		const [tabs] = lookupRaw(values, "tabs");
		if (Array.isArray(tabs)) invalid = (tabs.length === 0 && index !== -1) || (tabs.length > 0 && (index < 0 || index >= tabs.length));
	}
	if (t === "combobox" || t === "listbox") {
		// A missing list counts as empty: the generated code could not select an item that is not there.
		const [list] = lookupRaw(values, "list");
		invalid = index < -1 || index >= (Array.isArray(list) ? list.length : 0);
	}
	return invalid ? [selectedIndexIssue(`${path}.selectedIndex`)] : [];
}

const selectedIndexIssue = (field: string) =>
	issue("UFV020", "error", `${field} は -1 か、有効な 0 から始まるインデックスでなければなりません。`, field, "選択なしは -1、それ以外はコレクションの範囲内のインデックスを使ってください。");

function validateObservedControl(observed: JsonObject, type: string, progId: string, path: string): ValidationIssue[] {
	const allowed: Record<string, PropertyContract> = {};
	for (const key of ["left", "top", "width", "height", "tabIndex", "enabled", "visible", "unsupported", "properties"])
		allowed[key] = commonControlProperties[key];
	const builtIn = lookupControlContract(type);
	if (builtIn) Object.assign(allowed, builtIn.properties);
	else if (progId && !lookupControlContractByProgId(progId)) genericCustom(allowed);
	const issues = validateObject(observed, allowed, path);
	if (builtIn) markUnsupported(issues, type);
	return issues;
}

function parentAllowsChild(parent: RawRef, child: RawRef): boolean | undefined {
	const contractOf = (ref: RawRef) => (ref.progId && lookupControlContractByProgId(ref.progId)) || lookupControlContract(ref.type);
	const parentContract = contractOf(parent);
	const childContract = contractOf(child);
	if (!parentContract) return childContract?.allowedParentTypes ? false : undefined;
	if (!parentContract.canContainChildren) return false;
	if (parentContract.allowedChildTypes && (!childContract || !parentContract.allowedChildTypes.includes(childContract.type))) return false;
	if (childContract?.allowedParentTypes && !childContract.allowedParentTypes.includes(parentContract.type)) return false;
	return true;
}

function validateStructure(controls: RawRef[]): ValidationIssue[] {
	const issues: ValidationIssue[] = [];
	const ids = new Map<string, RawRef>();
	for (const control of controls) {
		if (!control.id) continue;
		const existing = ids.get(control.id);
		if (existing) {
			issues.push(issue("UFV007", "error", `${control.path}.id "${control.id}" が重複しています（最初は ${existing.path}.id）。`, `${control.path}.id`, "コントロールごとに一意の id を使ってください。"));
			continue;
		}
		ids.set(control.id, control);
	}
	const parentById = new Map<string, string>();
	const fieldById = new Map<string, string>();
	for (const control of controls) {
		const field = `${control.path}.parentId`;
		if (!control.parentId) {
			if (control.type.toLowerCase() === "page")
				issues.push(issue("UFV011", "error", `${field} は Page に必須です。Page は MultiPage に属さなければなりません。`, field, "parentId に親の MultiPage を指定してください。"));
			continue;
		}
		if (control.id && control.parentId === control.id) {
			issues.push(issue("UFV009", "error", `${field} が自分自身を参照しています。`, field));
			continue;
		}
		const parent = ids.get(control.parentId);
		if (!parent) {
			issues.push(issue("UFV008", "error", `${field} "${control.parentId}" が見つかりません。`, field, "既存のコンテナー コントロールの id を指定してください。"));
			continue;
		}
		if (parentAllowsChild(parent, control) === false) {
			const contract = (parent.progId && lookupControlContractByProgId(parent.progId)) || lookupControlContract(parent.type);
			issues.push(
				contract && !contract.canContainChildren
					? issue("UFV011", "error", `${field} "${parent.id}" はコンテナーではないコントロール "${parent.name}" を参照しています。`, field)
					: issue("UFV011", "error", `${field} の親 "${parent.name}" は ${control.type} を含められません。`, field, "MultiPage には Page だけを、Page は MultiPage の下にだけ置けます。"),
			);
		}
		if (control.id) {
			parentById.set(control.id, control.parentId);
			fieldById.set(control.id, field);
		}
	}
	const pageCount = new Map<string, number>();
	const known = new Set<string>();
	for (const control of controls) {
		if (control.type.toLowerCase() === "multipage" && control.childrenKnown) known.add(control.id);
		if (control.type.toLowerCase() === "page" && control.parentId) {
			pageCount.set(control.parentId, (pageCount.get(control.parentId) ?? 0) + 1);
			known.add(control.parentId);
		}
	}
	for (const control of controls) {
		if (control.type.toLowerCase() !== "multipage" || !known.has(control.id) || control.selectedIndex === undefined) continue;
		const count = pageCount.get(control.id) ?? 0;
		const index = control.selectedIndex;
		if ((count === 0 && index !== -1) || (count > 0 && (index < 0 || index >= count))) issues.push(selectedIndexIssue(`${control.path}.selectedIndex`));
	}
	const reported = new Set<string>();
	for (const id of [...parentById.keys()].sort()) {
		if (reported.has(id)) continue;
		const path: string[] = [];
		const seen = new Map<string, number>();
		for (let current: string | undefined = id; current; current = parentById.get(current)) {
			const start = seen.get(current);
			if (start !== undefined) {
				const cycle = path.slice(start);
				cycle.forEach((c) => reported.add(c));
				issues.push(issue("UFV010", "error", `"${cycle[0]}" の parentId の連鎖が循環しています。`, fieldById.get(cycle[0]) ?? "controls"));
				break;
			}
			if (reported.has(current)) break;
			seen.set(current, path.length);
			path.push(current);
		}
	}
	return issues;
}

function validateCapturedWarnings(value: JsonValue | undefined): ValidationIssue[] {
	if (!Array.isArray(value)) return [];
	const issues: ValidationIssue[] = [];
	value.forEach((entry, index) => {
		if (!isObject(entry)) return;
		const path = `warnings[${index}]`;
		const props: Record<string, PropertyContract> = {
			code: { valueType: "string", required: false, support: "supported", description: "", authoring: false },
			message: { valueType: "string", required: false, support: "supported", description: "", authoring: false },
			control: { valueType: "string", required: false, support: "supported", description: "", authoring: false },
		};
		const entryIssues = validateObject(entry, props, path);
		issues.push(...entryIssues);
		if (entryIssues.some((i) => i.severity === "error")) return;
		const message = typeof entry.message === "string" ? entry.message : "";
		if (message.trim()) issues.push(issue(typeof entry.code === "string" ? entry.code : "", "warning", message, `${path}.message`));
	});
	return issues;
}

export function validateSpec(root: JsonObject): ValidationIssue[] {
	const issues = validateObject(root, documentProperties, "");
	const form = root.form;
	if (isObject(form)) {
		issues.push(...validateObject(form, formProperties, "form"));
		if (isObject(form.observed)) issues.push(...validateObject(form.observed, formObservedProperties, "form.observed"));
		issues.push(...validateFormDimensions(form));
		issues.push(...validateTyped(form, "form", formProperties));
		// The Tag of the designer block: [MS-OVBA] allows 130 characters in quotes, on one line.
		if (typeof form.tag === "string" && (/[\u0000-\u001f\u007f]/.test(form.tag) || [...form.tag].length > 130))
			issues.push(issue("UFV021", "error", "tag は制御文字を含まない 130 文字以内の文字列にしてください。", "form.tag"));
	}
	if (root.controls !== undefined) {
		const [refs, controlIssues] = validateControls(root.controls, "controls", "");
		issues.push(...controlIssues, ...validateStructure(refs));
	}
	issues.push(...validateCapturedWarnings(root.warnings));
	// Remove exact duplicates (the same rule can fire through two paths).
	const seen = new Set<string>();
	return issues.filter((i) => {
		const key = `${i.code}\0${i.field}\0${i.message}`;
		if (seen.has(key)) return false;
		seen.add(key);
		return true;
	});
}

export const errorKeys = (issues: ValidationIssue[]) => new Set(issues.filter((i) => i.severity === "error").map((i) => `${i.code}\0${i.field}`));

// ---------------------------------------------------------------------------
// Property Grid metadata (spec/edit/properties.go PropertyGrid)
// ---------------------------------------------------------------------------

function propertyTarget(properties: Record<string, PropertyContract>, object: JsonObject): PropertyTarget {
	const target: PropertyTarget = { descriptors: [], values: {} };
	for (const name of Object.keys(properties).sort()) {
		const contract = properties[name];
		let state: PropertyState = { present: Object.hasOwn(object, name), value: (object[name] ?? null) as PropertyState["value"] };
		if (!state.present && isObject(object.properties)) {
			const [legacy, present] = lookupRaw(object.properties, name);
			if (present && (name === "textAlign" || name === "borderStyle" || name === "pictureAlignment" || name === "pictureSizeMode"))
				state = { present: true, legacy: true, value: legacy as PropertyState["value"] };
		}
		target.descriptors.push(propertyDescriptor(name, contract));
		target.values[name] = state;
	}
	return target;
}

export function propertyGrid(spec: ParsedSpec): PropertyGridData {
	const result: PropertyGridData = { form: propertyTarget(editableFormProperties(), spec.form), controls: {} };
	for (const record of spec.records) {
		if (typeof record.raw.id !== "string") continue; // Inferred IDs are not editable targets.
		const { control } = record;
		result.controls[control.id] = propertyTarget(editableControlProperties(control.type, control.progId), record.raw);
		if (control.type.toLowerCase() === "tabstrip") {
			const tabs = Array.isArray(record.raw.tabs) ? record.raw.tabs : [];
			result.tabs ??= {};
			result.tabs[control.id] = tabs.filter(isObject).map((tab) => propertyTarget(editableTabProperties(), tab));
		}
	}
	return result;
}

/** A new, empty form in the minimal authoring shape. */
export function newFormSpec(name: string): JsonObject {
	return {
		schemaVersion: 1,
		kind: FORM_KIND,
		basis: "designer",
		form: { name, caption: name, width: 240, height: 180 },
		controls: [],
	};
}
