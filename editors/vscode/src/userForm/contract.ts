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
	editor?: "font" | "color" | "enum" | "picture" | "list" | "toggle";
	children?: Record<string, PropertyContract>;
	minimum?: number;
	maximum?: number;
	exclusiveMinimum?: boolean;
	/** Names of allowedValues, in order (the VBA constants without their prefix). */
	valueLabels?: string[];
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

/** Caption and Value of a CheckBox, OptionButton or ToggleButton; Value is edited as True, False or Null. */
const toggleControl = (caption: string, value: string): Record<string, PropertyContract> => {
	const properties = typed("caption", caption, "value", value);
	properties.value = { ...properties.value, editor: "toggle" };
	return properties;
};

// The plain properties of the controls (crates/ras-form/src/scalars.rs has the same
// table): the VBA property of the same name, and its values. Omitted, each is what
// Excel gives a new control.
type Scalar = [name: string, types: string[], contract: PropertyContract, labels?: string[]];
const EDITS = ["TextBox", "ComboBox"];
const FOCUSABLE = ["TextBox", "ComboBox", "ListBox", "CommandButton", "CheckBox", "OptionButton", "ToggleButton", "SpinButton", "ScrollBar", "MultiPage", "TabStrip"];
const BOUND = ["TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton", "SpinButton", "ScrollBar"];
const flag = (description: string) => p("boolean", false, "supported", description, true);
const choice = (description: string, values: number[]) =>
	p("integer", false, "supported", description, true, values.map(String));
const integer = (description: string, minimum: number, maximum: number): PropertyContract => ({
	...p("integer", false, "supported", description, true),
	minimum,
	maximum,
});
const pointsUpTo = (description: string, maximum: number): PropertyContract => ({
	...p("number", false, "supported", description, true),
	minimum: 0,
	maximum,
});
const effects = [0, 1, 2, 3, 6];
const effectLabels = ["Flat", "Raised", "Sunken", "Etched", "Bump"];
const pointers = [...Array(16).keys()];
const int32 = [-2147483648, 2147483647] as const;
export const scalarProperties: Scalar[] = [
	["locked", ["TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton", "CommandButton"], flag("The value cannot be changed by the user.")],
	["backStyle", ["TextBox", "ComboBox", "CheckBox", "OptionButton", "ToggleButton", "CommandButton", "Label"], choice("Background: Transparent or Opaque.", [0, 1]), ["Transparent", "Opaque"]],
	["integralHeight", ["TextBox", "ListBox"], flag("Shows only complete lines.")],
	["matchRequired", ["ComboBox"], flag("The text has to match an item.")],
	["alignment", ["CheckBox", "OptionButton"], choice("Caption at the Left or the Right of the box.", [0, 1]), ["Left", "Right"]],
	["imeMode", ["TextBox", "ComboBox", "ListBox"], choice("IME mode.", [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]), ["NoControl", "On", "Off", "Disable", "Hiragana", "Katakana", "KatakanaHalf", "AlphaFull", "Alpha", "HangulFull", "Hangul"]],
	["dragBehavior", EDITS, choice("Drag and drop of text.", [0, 1]), ["Disabled", "Enabled"]],
	["enterKeyBehavior", ["TextBox"], flag("Enter adds a new line.")],
	["enterFieldBehavior", EDITS, choice("On entering: select all, or keep the last selection.", [0, 1]), ["SelectAll", "RecallSelection"]],
	["tabKeyBehavior", ["TextBox"], flag("Tab inserts a tab character.")],
	["wordWrap", ["TextBox", "CheckBox", "OptionButton", "ToggleButton", "CommandButton", "Label"], flag("Long lines wrap.")],
	["selectionMargin", EDITS, flag("A margin to select whole lines.")],
	["autoWordSelect", EDITS, flag("Selection extends by whole words.")],
	["autoSize", ["TextBox", "ComboBox", "CheckBox", "OptionButton", "ToggleButton", "CommandButton", "Label"], flag("The size follows the contents.")],
	["hideSelection", EDITS, flag("The selection is hidden without the focus.")],
	["autoTab", EDITS, flag("The focus moves on at MaxLength characters.")],
	["multiLine", ["TextBox"], flag("Several lines of text.")],
	["maxLength", EDITS, integer("Most characters; 0 for no limit.", 0, 2147483647)],
	["scrollBars", ["TextBox"], choice("Scroll bars.", [0, 1, 2, 3]), ["None", "Horizontal", "Vertical", "Both"]],
	["passwordChar", ["TextBox"], p("string", false, "supported", "Character shown for each character typed; empty for the text.", true)],
	["style", ["ComboBox"], choice("DropDownCombo (text can be typed) or DropDownList (an item has to be chosen).", [0, 2]), ["DropDownCombo", "DropDownList"]],
	["listStyle", ["ComboBox", "ListBox"], choice("List style: Plain, Option (a box beside each item).", [0, 1]), ["Plain", "Option"]],
	["matchEntry", ["ComboBox", "ListBox"], choice("Typing selects an item: FirstLetter, Complete, None.", [0, 1, 2]), ["FirstLetter", "Complete", "None"]],
	["showDropButtonWhen", ["ComboBox"], choice("When the drop button shows.", [0, 1, 2]), ["Never", "Focus", "Always"]],
	["dropButtonStyle", ["ComboBox"], choice("Symbol on the drop button.", [0, 1, 2, 3]), ["Plain", "Arrow", "Ellipsis", "Reduce"]],
	["listRows", ["ComboBox"], integer("Rows shown in the drop-down list.", 0, 32767)],
	["listWidth", ["ComboBox"], pointsUpTo("Width of the drop-down list in points; 0 for the control's width.", 72000)],
	["multiSelect", ["ListBox"], choice("Selection: Single, Multi, Extended.", [0, 1, 2]), ["Single", "Multi", "Extended"]],
	["tripleState", ["CheckBox", "OptionButton", "ToggleButton"], flag("A click also gives Null (not for an OptionButton).")],
	["groupName", ["OptionButton"], p("string", false, "supported", "Option buttons of the same group exclude each other.", true)],
	["specialEffect", ["TextBox", "ComboBox", "ListBox", "Label", "Image", "Frame"], choice("3D look of the edge.", effects), effectLabels],
	["specialEffect", ["CheckBox", "OptionButton"], choice("3D look of the box.", [0, 2]), ["Flat", "Sunken"]],
	["borderColor", ["TextBox", "ComboBox", "ListBox", "Label", "Image", "Frame"], { ...color, description: "Color of a Single border." }],
	["cycle", ["Frame"], choice("Tab from the last control: to the next form's controls, or back to this frame's first.", [0, 2]), ["AllForms", "CurrentForm"]],
	["mousePointer", ["Label", "TextBox", "ComboBox", "ListBox", "CommandButton", "CheckBox", "OptionButton", "ToggleButton", "SpinButton", "ScrollBar", "Image", "Frame", "TabStrip"], choice("Mouse pointer over the control.", pointers), ["Default", "Arrow", "Cross", "IBeam", "SizeNESW", "SizeNS", "SizeNWSE", "SizeWE", "UpArrow", "HourGlass", "NoDrop", "AppStarting", "Help", "SizeAll", "SizeAll", "Hand"]],
	["default", ["CommandButton"], flag("Enter clicks this button.")],
	["cancel", ["CommandButton"], flag("Esc clicks this button.")],
	["takeFocusOnClick", ["CommandButton"], flag("Clicking moves the focus to the button.")],
	["tabStop", FOCUSABLE, flag("Tab moves the focus to the control.")],
	["controlSource", BOUND, p("string", false, "supported", "Cell the value is read from when the form loads and written to when it changes (\"Sheet1!B2\").", true)],
	["rowSource", ["ComboBox", "ListBox"], p("string", false, "supported", "Cells the items come from (\"Sheet1!A2:B10\"); list cannot be written as well.", true)],
	["min", ["SpinButton", "ScrollBar"], integer("Least value.", ...int32)],
	["max", ["SpinButton", "ScrollBar"], integer("Greatest value (100 for a SpinButton, 32767 for a ScrollBar when omitted).", ...int32)],
	["smallChange", ["SpinButton", "ScrollBar"], integer("Change of a click on an arrow.", ...int32)],
	["largeChange", ["ScrollBar"], integer("Change of a click on the track.", ...int32)],
	["orientation", ["SpinButton", "ScrollBar"], choice("Auto (by the shape), Vertical or Horizontal.", [-1, 0, 1]), ["Auto", "Vertical", "Horizontal"]],
	["proportionalThumb", ["ScrollBar"], flag("The thumb's size shows the visible part.")],
	["delay", ["SpinButton", "ScrollBar"], integer("Milliseconds between repeated changes while held.", 0, 2147483647)],
	["style", ["TabStrip", "MultiPage"], choice("Tabs, Buttons or None.", [0, 1, 2]), ["Tabs", "Buttons", "None"]],
	["tabOrientation", ["TabStrip", "MultiPage"], choice("Side of the tabs.", [0, 1, 2, 3]), ["Top", "Bottom", "Left", "Right"]],
	["multiRow", ["TabStrip", "MultiPage"], flag("Tabs in several rows.")],
	["tabFixedWidth", ["TabStrip", "MultiPage"], pointsUpTo("Width of every tab in points; 0 for each tab's own.", 7200)],
	["tabFixedHeight", ["TabStrip", "MultiPage"], pointsUpTo("Height of the tabs in points; 0 for the font's.", 7200)],
];
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
		checkbox: control("CheckBox", false, toggleControl("CheckBox caption.", "CheckBox checked state: true, false or \"Null\".")),
		optionbutton: control("OptionButton", false, toggleControl("OptionButton caption.", "OptionButton selected state: true, false or \"Null\".")),
		togglebutton: control("ToggleButton", false, toggleControl("ToggleButton caption.", "ToggleButton pressed state: true, false or \"Null\".")),
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
		for (const [name, types, scalar, valueLabels] of scalarProperties)
			if (types.includes(t)) contract.properties[name] = valueLabels ? { ...scalar, valueLabels } : scalar;
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

const editorPropertyNames = ["font", "foreColor", "backColor", "borderColor", "picture", "textAlign", "borderStyle", "pictureAlignment", "pictureSizeMode"];
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
	const names = contract.valueLabels ?? labels[field];
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
