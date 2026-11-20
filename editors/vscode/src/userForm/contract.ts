// The UserForm FormSpec v1 property contract, as served by xlflow's language
// server (internal/vba/userforms/spec/contract.go, appearance.go and
// spec/edit/property_editors.go), turned into static tables.
//
// Ported from xlflow (https://github.com/harumiWeb/xlflow),
// MIT License, Copyright (c) 2026 harumiWeb.

import type { PropertyDescriptor } from "./protocol";

export type ValueType = "string" | "number" | "integer" | "boolean" | "any" | "stringArray" | "object" | "objectArray";
export type SupportLevel = "supported" | "best-effort" | "observed-only" | "snapshot-only" | "custom/unchecked";

export interface PropertyContract {
	valueType: ValueType;
	required: boolean;
	support: SupportLevel;
	description: string;
	authoring: boolean;
	allowedValues?: string[];
	editor?: "font" | "color" | "enum" | "picture" | "list";
	children?: Record<string, PropertyContract>;
	minimum?: number;
	maximum?: number;
	exclusiveMinimum?: boolean;
}

export interface ControlContract {
	type: string;
	progId: string;
	canContainChildren: boolean;
	allowedChildTypes?: string[];
	allowedParentTypes?: string[];
	excludedCommon?: string[];
	properties: Record<string, PropertyContract>;
}

const p = (
	valueType: ValueType,
	required: boolean,
	support: SupportLevel,
	description: string,
	authoring: boolean,
	allowedValues?: string[],
): PropertyContract => ({ valueType, required, support, description, authoring, ...(allowedValues ? { allowedValues } : {}) });

export const documentProperties: Record<string, PropertyContract> = {
	schemaVersion: p("integer", true, "supported", "UserForm spec schema version.", false, ["1"]),
	kind: p("string", true, "supported", "Document kind discriminator.", false, ["ras.userform", "xlflow.userform"]),
	basis: p("string", true, "supported", "Inspection/build basis for the spec.", false, ["designer"]),
	coordinateSystem: p("string", false, "supported", "Coordinate basis used by geometry fields.", false, ["points", "parent-relative"]),
	form: p("object", true, "supported", "Top-level UserForm authoring and observed metadata.", false),
	controls: p("objectArray", true, "supported", "Flat control list.", false),
	warnings: p("objectArray", false, "snapshot-only", "Snapshot-oriented warnings persisted with the spec.", false),
};

const font: PropertyContract = {
	...p("object", false, "supported", "Partial MSForms font; size is in points and charset is numeric.", true),
	editor: "font",
	children: {
		bold: p("boolean", false, "supported", "Font style.", true),
		italic: p("boolean", false, "supported", "Font style.", true),
		underline: p("boolean", false, "supported", "Font style.", true),
		strikeThrough: p("boolean", false, "supported", "Font style.", true),
		name: p("string", false, "supported", "Font family name.", true),
		size: { ...p("number", false, "supported", "Font size in points, greater than zero.", true), minimum: 0, maximum: 3276.75, exclusiveMinimum: true },
		charset: { ...p("integer", false, "supported", "Windows charset number; unlisted values remain valid.", true, []), minimum: 0, maximum: 255 },
	},
};
const color: PropertyContract = {
	...p("object", false, "supported", "RGB, system or raw OLE color.", true),
	editor: "color",
	children: {
		kind: p("string", true, "supported", "Color representation.", true, ["rgb", "system", "ole"]),
		value: p("string", true, "supported", "#RRGGBB, system color name, or 0xXXXXXXXX.", true),
	},
};
const borderStyle: PropertyContract = { ...p("integer", false, "supported", "MSForms border style.", true, ["0", "1"]), editor: "enum" };
const textAlign: PropertyContract = { ...p("integer", false, "supported", "Text alignment: 1 Left, 2 Center, 3 Right.", true, ["1", "2", "3"]), editor: "enum" };

export const formProperties: Record<string, PropertyContract> = {
	name: p("string", true, "supported", "VBA UserForm component name.", true),
	caption: p("string", false, "supported", "UserForm caption.", true),
	width: p("number", false, "supported", "Persisted UserForm client content width in points.", true),
	height: p("number", false, "supported", "Persisted UserForm client content height in points.", true),
	observed: p("object", false, "snapshot-only", "Observed form state captured from Excel.", false),
	font,
	foreColor: color,
	backColor: color,
	borderStyle,
};

export const formObservedProperties: Record<string, PropertyContract> = Object.fromEntries(
	["caption", "width", "height", "insideWidth", "insideHeight", "clientWidth", "clientHeight"].map((name) => [
		name,
		p(name === "caption" ? "string" : "number", false, "snapshot-only", `Observed ${name}.`, false),
	]),
);

export const commonControlProperties: Record<string, PropertyContract> = {
	id: p("string", true, "supported", "Stable control identifier used by parentId references.", true),
	name: p("string", true, "supported", "VBA control name.", true),
	type: p("string", true, "supported", "Canonical control type.", true),
	progId: p("string", false, "supported", "COM ProgID used when creating the control.", false),
	parentId: p("string", false, "supported", "ID of the containing control.", true),
	zIndex: p("integer", false, "supported", "Sibling ordering hint for Designer rebuild.", true),
	left: p("number", false, "supported", "Left coordinate relative to the parent container.", true),
	top: p("number", false, "supported", "Top coordinate relative to the parent container.", true),
	width: p("number", false, "supported", "Control width in Designer points.", true),
	height: p("number", false, "supported", "Control height in Designer points.", true),
	tabIndex: p("integer", false, "supported", "Control tab order index.", true),
	enabled: p("boolean", false, "supported", "Whether the control is enabled.", true),
	visible: p("boolean", false, "supported", "Whether the control is visible.", true),
	tag: p("string", false, "supported", "Opaque metadata associated with the control.", true),
	controlTipText: p("string", false, "supported", "Tooltip shown for the control.", true),
	observed: p("object", false, "snapshot-only", "Snapshot state captured from Excel for review.", false),
	controls: p("objectArray", false, "snapshot-only", "Legacy nested input accepted and normalized to flat controls.", false),
	properties: p("object", false, "custom/unchecked", "Unchecked custom property bag.", false),
	unsupported: p("stringArray", false, "snapshot-only", "Properties observed during inspection that were not modeled.", false),
};

export const tabProperties: Record<string, PropertyContract> = {
	name: p("string", true, "supported", "TabStrip tab name; unique within the TabStrip, ignoring case.", true),
	caption: p("string", false, "supported", "Text displayed on the tab.", true),
	controlTipText: p("string", false, "supported", "Tooltip shown for the tab.", true),
	tag: p("string", false, "supported", "Opaque metadata associated with the tab.", true),
	accelerator: p("string", false, "supported", "Keyboard accelerator for the tab.", true),
	enabled: p("boolean", false, "supported", "Whether the tab is enabled.", true),
	visible: p("boolean", false, "supported", "Whether the tab is visible.", true),
};

export const pictureActionProperties: Record<string, PropertyContract> = {
	path: p("string", false, "supported", "Project-root-relative BMP or JPEG asset path, with / separators.", true),
	remove: p("boolean", false, "supported", "No picture; the same as omitting the field.", true, ["true"]),
};

const typed = (...entries: string[]) => {
	const result: Record<string, PropertyContract> = {};
	for (let i = 0; i + 1 < entries.length; i += 2)
		result[entries[i]] = p(entries[i] === "value" ? "any" : "string", false, "supported", entries[i + 1], true);
	return result;
};
const listControl = (): Record<string, PropertyContract> => ({
	...typed("text", "Displayed text.", "value", "Current control value."),
	list: { ...p("stringArray", false, "supported", "List items; generated code in UserForm_Initialize fills the list.", true), editor: "list" },
	selectedIndex: p("integer", false, "supported", "Zero-based selected item of list, or -1 for none; set by generated code in UserForm_Initialize.", true),
});
const control = (type: string, canContainChildren: boolean, properties: Record<string, PropertyContract>): ControlContract => ({
	type,
	progId: `Forms.${type}.1`,
	canContainChildren,
	properties,
});

function buildControls(): Record<string, ControlContract> {
	const controls: Record<string, ControlContract> = {
		label: control("Label", false, typed("caption", "Text displayed by the label.")),
		textbox: control("TextBox", false, typed("text", "TextBox text.", "value", "TextBox value.")),
		combobox: control("ComboBox", false, listControl()),
		listbox: control("ListBox", false, listControl()),
		commandbutton: control("CommandButton", false, typed("caption", "Button caption.")),
		checkbox: control("CheckBox", false, typed("caption", "CheckBox caption.", "value", "CheckBox checked state.")),
		optionbutton: control("OptionButton", false, typed("caption", "OptionButton caption.", "value", "OptionButton selected state.")),
		togglebutton: control("ToggleButton", false, typed("caption", "ToggleButton caption.", "value", "ToggleButton selected state.")),
		spinbutton: control("SpinButton", false, { value: p("integer", false, "supported", "SpinButton value.", true) }),
		scrollbar: control("ScrollBar", false, { value: p("integer", false, "supported", "ScrollBar value.", true) }),
		image: control("Image", false, {
			picture: { ...p("object", false, "supported", "Picture embedded into the Image control.", true), editor: "picture", children: pictureActionProperties },
			pictureAlignment: { ...p("integer", false, "supported", "Picture alignment: TopLeft, TopRight, Center, BottomLeft, BottomRight.", true, ["0", "1", "2", "3", "4"]), editor: "enum" },
			pictureSizeMode: { ...p("integer", false, "supported", "Picture size mode: Clip, Stretch, Zoom.", true, ["0", "1", "3"]), editor: "enum" },
		}),
		frame: control("Frame", true, typed("caption", "Frame caption.")),
		multipage: {
			...control("MultiPage", true, {
				selectedIndex: p("integer", false, "supported", "Zero-based selected Page index, or -1 when no Page is selected.", true),
			}),
			allowedChildTypes: ["Page"],
		},
		page: {
			...control(
				"Page",
				true,
				typed(
					"caption",
					"Text displayed on the page tab.",
					"controlTipText",
					"Tooltip shown for the page.",
					"tag",
					"Opaque metadata associated with the page.",
					"accelerator",
					"Keyboard accelerator for the page.",
				),
			),
			allowedParentTypes: ["MultiPage"],
			excludedCommon: ["left", "top", "width", "height"],
		},
		tabstrip: control("TabStrip", false, {
			selectedIndex: p("integer", false, "supported", "Zero-based selected Tab index, or -1 when no Tab is selected.", true),
			tabs: p("objectArray", false, "supported", "Standalone TabStrip tabs. An explicit empty array removes all tabs.", true),
		}),
	};
	for (const contract of Object.values(controls)) {
		const t = contract.type;
		if (!["Page", "Image", "SpinButton", "ScrollBar"].includes(t)) contract.properties.font = font;
		if (t !== "Page") contract.properties.backColor = color;
		if (t !== "Image" && t !== "Page") contract.properties.foreColor = color;
		if (["Label", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton"].includes(t)) contract.properties.textAlign = textAlign;
		if (["Frame", "TextBox", "ComboBox", "ListBox", "Label", "Image"].includes(t)) contract.properties.borderStyle = borderStyle;
	}
	return controls;
}

export const controlContracts = buildControls();

export function lookupControlContract(type: unknown): ControlContract | undefined {
	return typeof type === "string" ? controlContracts[type.trim().toLowerCase()] : undefined;
}

export function lookupControlContractByProgId(progId: unknown): ControlContract | undefined {
	if (typeof progId !== "string" || !progId.trim()) return undefined;
	const wanted = progId.trim().toLowerCase();
	return Object.values(controlContracts).find((c) => c.progId.toLowerCase() === wanted);
}

/** Common and type-specific properties of a control type, with exclusions applied. */
export function controlProperties(type: unknown): Record<string, PropertyContract> {
	const result: Record<string, PropertyContract> = { ...commonControlProperties };
	const contract = lookupControlContract(type);
	if (!contract) return result;
	for (const name of contract.excludedCommon ?? []) delete result[name];
	return { ...result, ...contract.properties };
}

export function lookupProperty(properties: Record<string, PropertyContract>, name: string): PropertyContract | undefined {
	if (Object.hasOwn(properties, name)) return properties[name];
	const key = name.trim().toLowerCase();
	const found = Object.keys(properties).find((candidate) => candidate.toLowerCase() === key);
	return found ? properties[found] : undefined;
}

const editorPropertyNames = ["font", "foreColor", "backColor", "picture", "textAlign", "borderStyle", "pictureAlignment", "pictureSizeMode"];
const scalarTypes: ValueType[] = ["string", "number", "integer", "boolean", "any"];
export const isEditorProperty = (name: string) => editorPropertyNames.includes(name);
const isScalar = (contract: PropertyContract) => scalarTypes.includes(contract.valueType);

/** Fields the Property Grid offers for the form (scalar fields plus typed editors). */
export function editableFormProperties(): Record<string, PropertyContract> {
	return Object.fromEntries(
		Object.entries(formProperties).filter(([name, c]) => c.authoring && (isScalar(c) || isEditorProperty(name))),
	);
}

/** Fields the Property Grid offers for a control (scalar fields plus typed editors). */
export function editableControlProperties(type: string, progId?: string): Record<string, PropertyContract> {
	const typedEditors = !progId || lookupControlContractByProgId(progId)?.type.toLowerCase() === type.toLowerCase();
	return Object.fromEntries(
		Object.entries(controlProperties(type)).filter(
			([name, c]) =>
				!["id", "type", "parentId", "zIndex"].includes(name) &&
				c.authoring &&
				(isEditorProperty(name) || c.editor === "list" ? typedEditors : isScalar(c)),
		),
	);
}

export function editableTabProperties(): Record<string, PropertyContract> {
	return { ...tabProperties };
}

export const systemColors: Record<string, number> = {
	scrollBar: 0,
	desktop: 1,
	activeCaption: 2,
	inactiveCaption: 3,
	menu: 4,
	window: 5,
	windowFrame: 6,
	menuText: 7,
	windowText: 8,
	captionText: 9,
	activeBorder: 10,
	inactiveBorder: 11,
	appWorkspace: 12,
	highlight: 13,
	highlightText: 14,
	buttonFace: 15,
	buttonShadow: 16,
	grayText: 17,
	buttonText: 18,
	inactiveCaptionText: 19,
	buttonHighlight: 20,
	buttonDarkShadow: 21,
	buttonLight: 22,
	infoText: 23,
	infoBackground: 24,
	hotLight: 26,
	gradientActiveCaption: 27,
	gradientInactiveCaption: 28,
	menuHighlight: 29,
	menuBar: 30,
};

/** Validates a `{ kind, value }` color; returns an error message or undefined. */
export function colorError(value: unknown): string | undefined {
	if (typeof value !== "object" || value === null || Array.isArray(value)) return "色は { kind, value } のオブジェクトで指定します。";
	const { kind, value: text } = value as { kind?: unknown; value?: unknown };
	if (kind === "rgb" && typeof text === "string" && /^#[0-9A-Fa-f]{6}$/.test(text)) return;
	if (kind === "system" && typeof text === "string" && Object.hasOwn(systemColors, text)) return;
	if (kind === "ole" && typeof text === "string" && /^0x[0-9A-Fa-f]{8}$/.test(text)) return;
	return `invalid ${String(kind)} color ${JSON.stringify(text)}`;
}

const labels: Record<string, string[]> = {
	textAlign: ["Left", "Center", "Right"],
	borderStyle: ["None", "Single"],
	pictureAlignment: ["TopLeft", "TopRight", "Center", "BottomLeft", "BottomRight"],
	pictureSizeMode: ["Clip", "Stretch", "Zoom"],
};
const charsets: [number, string][] = [
	[0, "ANSI"],
	[1, "Default"],
	[2, "Symbol"],
	[128, "Shift JIS"],
	[129, "Hangul"],
	[134, "GB2312"],
	[136, "Big5"],
	[161, "Greek"],
	[162, "Turkish"],
	[177, "Hebrew"],
	[178, "Arabic"],
	[186, "Baltic"],
	[204, "Cyrillic"],
	[222, "Thai"],
	[238, "East Europe"],
	[255, "OEM"],
];

export function propertyDescriptor(field: string, contract: PropertyContract): PropertyDescriptor {
	const descriptor: PropertyDescriptor = {
		field,
		valueType: contract.valueType as PropertyDescriptor["valueType"],
		required: contract.required,
		nullable: contract.editor === "picture" ? false : !contract.required,
	};
	if (contract.allowedValues?.length) descriptor.allowedValues = [...contract.allowedValues];
	if (contract.editor) descriptor.editor = contract.editor;
	else if (contract.valueType === "boolean") descriptor.editor = "boolean";
	else if (contract.allowedValues?.length) descriptor.editor = "enum";
	if (contract.minimum !== undefined) descriptor.minimum = contract.minimum;
	if (contract.maximum !== undefined) descriptor.maximum = contract.maximum;
	if (contract.exclusiveMinimum) descriptor.exclusiveMinimum = true;
	const names = labels[field];
	if (names && names.length === descriptor.allowedValues?.length)
		descriptor.suggestions = descriptor.allowedValues.map((value, index) => ({ value, label: names[index] }));
	if (contract.children) {
		descriptor.children = Object.keys(contract.children)
			.sort()
			.map((name) => {
				const child = propertyDescriptor(name, contract.children![name]);
				if (contract.editor === "font" && name === "charset") {
					delete child.allowedValues;
					if (child.editor === "enum") delete child.editor;
					child.suggestions = charsets.map(([value, label]) => ({ value, label }));
				}
				if (contract.editor === "color" && name === "value")
					child.suggestions = Object.keys(systemColors).map((value) => ({ value, label: value }));
				return child;
			});
	}
	return descriptor;
}
