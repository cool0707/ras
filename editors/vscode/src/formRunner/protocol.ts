// Debug run of a user form (spec §14.10.7): what the debug process (the run time's
// `ras-forms.js`), the extension and the form's panel say to each other.
//
// The debug process connects to a TCP server of the extension and sends one JSON object
// per line; the extension keeps the last state of each form, draws it in a webview panel
// and forwards what the user does there. Nothing here imports `vscode`, so that the panel
// (a browser bundle) and the tests can use it.

/** Property values of a control or the form, by their VBA names (`Caption`, `Text`, `BackColor` …). */
export type ControlState = Record<string, unknown> & { Font?: FontState };

export interface FontState {
	Name?: string;
	Size?: number;
	Bold?: boolean;
	Italic?: boolean;
	Underline?: boolean;
	Strikethrough?: boolean;
}

export interface ControlSnapshot {
	name: string;
	/** `TextBox`, `CommandButton`, `Page` … */
	type: string;
	/** Name of the container (`Frame`, `Page`, `MultiPage`), "" for the form. */
	parent: string;
	state: ControlState;
	/** Rows of a ComboBox / ListBox, each a list of column texts. */
	list?: string[][];
	listIndex?: number;
	selected?: boolean[];
	tabs?: ControlState[];
	/** Data URI of the control's picture (an Image's, a Label's, a button's; a Frame's or Page's background). */
	picture?: string | null;
	/** Data URI of the icon or cursor of MousePointer Custom (99). */
	mouseIcon?: string | null;
	/** `seq` of the last `input` of the panel that the form's code has seen (TextBox / ComboBox). */
	inputSeq?: number;
}

/** From the debug process. */
export type RuntimeMessage =
	| { t: "show"; id: number; name: string; modal: boolean; form: ControlState; controls: ControlSnapshot[] }
	| { t: "update"; id: number; form?: ControlState; controls: ControlSnapshot[]; focus?: string; removed?: string[] }
	| { t: "hide"; id: number }
	| { t: "unload"; id: number }
	| { t: "error"; id: number; message: string }
	| { t: "blocked"; id: number };

/** From the panel (the extension adds the form's `id` before passing it on). */
export type UserAction =
	| { t: "focus"; control: string }
	| { t: "input"; control: string; text: string; seq: number; selStart?: number }
	| { t: "click"; control?: string; key?: boolean }
	| { t: "dblclick"; control?: string }
	| { t: "select"; control: string; index: number; selected?: boolean[] }
	| { t: "page"; control: string; index: number }
	| { t: "spin"; control: string; delta: number }
	| { t: "scroll"; control: string; value: number }
	| { t: "close" };

/** What the extension posts to the panel. */
export type PanelMessage =
	| { type: "state"; state: FormState; focus?: string }
	| { type: "error"; message: string }
	| { type: "blocked" }
	| { type: "ended" };

/** Everything the panel needs to draw a form. */
export interface FormState {
	name: string;
	modal: boolean;
	form: ControlState;
	/** In the order of the form's design (the z-order). */
	controls: ControlSnapshot[];
}

export function initialState(msg: Extract<RuntimeMessage, { t: "show" }>): FormState {
	return { name: msg.name, modal: msg.modal, form: { ...msg.form }, controls: msg.controls.map((c) => ({ ...c })) };
}

/**
 * Applies an `update` to a form's state. Controls are replaced whole; their order stays. Controls
 * the form's code added (`Controls.Add`) come last, on top of the others; `removed` names those
 * it removed (`Controls.Remove`).
 */
export function applyUpdate(state: FormState, msg: Extract<RuntimeMessage, { t: "update" }>): FormState {
	const changed = new Map(msg.controls.map((c) => [c.name.toLowerCase(), c]));
	const removed = new Set((msg.removed ?? []).map((n) => n.toLowerCase()));
	const known = new Set(state.controls.map((c) => c.name.toLowerCase()));
	const controls = state.controls.filter((c) => !removed.has(c.name.toLowerCase())).map((c) => changed.get(c.name.toLowerCase()) ?? c);
	for (const c of msg.controls) if (!known.has(c.name.toLowerCase())) controls.push(c);
	return { ...state, form: msg.form ? { ...msg.form } : state.form, controls };
}

/** Splits a stream into lines; returns the complete lines and keeps the rest. */
export class LineReader {
	private buffer = "";

	push(chunk: string): string[] {
		this.buffer += chunk;
		const lines = this.buffer.split("\n");
		this.buffer = lines.pop() ?? "";
		return lines.map((l) => l.trim()).filter((l) => l !== "");
	}
}

// Windows' default system colors (OLE colors 0x80000000 + index), so that the form looks
// as it does in Excel whatever the editor's theme.
const SYSTEM_COLORS = [
	"#c8c8c8", "#000000", "#99b4d1", "#bfcddb", "#f0f0f0", "#ffffff", "#646464", "#000000", "#000000", "#000000", "#b4b4b4",
	"#f4f7fc", "#ababab", "#0078d7", "#ffffff", "#f0f0f0", "#a0a0a0", "#6d6d6d", "#000000", "#000000", "#ffffff", "#696969",
	"#e3e3e3", "#000000", "#ffffe1", "#000000", "#0066cc", "#b9d1ea", "#d7e4f2", "#3399ff", "#f0f0f0",
];

/** An OLE color (`BackColor`) as a CSS color; `undefined` for a value that is not one. */
export function oleToCss(value: unknown): string | undefined {
	if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
	const n = value >>> 0;
	if (n >= 0x80000000) return SYSTEM_COLORS[n - 0x80000000];
	if (n > 0xffffff) return undefined;
	const hex = (x: number) => x.toString(16).padStart(2, "0");
	return `#${hex(n & 255)}${hex((n >>> 8) & 255)}${hex((n >>> 16) & 255)}`;
}

/** The CSS cursor of an fmMousePointer ("" for Default: the control's own); Custom (99) shows `icon`. */
export function cursorOf(pointer: unknown, icon?: string | null): string {
	if (pointer === 99) return icon ? `url("${icon}"), auto` : "";
	const cursors: Record<number, string> = {
		1: "default",
		2: "crosshair",
		3: "text",
		6: "nesw-resize",
		7: "ns-resize",
		8: "nwse-resize",
		9: "ew-resize",
		10: "n-resize",
		11: "wait",
		12: "no-drop",
		13: "progress",
		14: "help",
		15: "move",
	};
	return typeof pointer === "number" ? (cursors[pointer] ?? "") : "";
}

/** Points (the unit of MSForms) → CSS pixels. */
export function px(points: unknown): string {
	const n = typeof points === "number" && Number.isFinite(points) ? points : 0;
	return `${(n * 96) / 72}px`;
}

export interface Insets {
	left: number;
	top: number;
	right: number;
	bottom: number;
}

/**
 * Offset of a container's content from its edges, in points (the Designer draws them the same
 * way): the tabs of a MultiPage take 18pt on the side of its TabOrientation.
 */
export function contentInsets(type: string, tabOrientation?: unknown): Insets {
	if (type === "MultiPage") {
		const side = (["top", "bottom", "left", "right"] as const)[typeof tabOrientation === "number" ? tabOrientation : 0] ?? "top";
		return { left: 1.5, top: 1.5, right: 1.5, bottom: 1.5, [side]: 18 };
	}
	if (type === "Frame") return { left: 0.75, top: 0.75, right: 0.75, bottom: 0.75 };
	return { left: 0, top: 0, right: 0, bottom: 0 };
}
