// The panel of a user form shown by a debug run (spec §14.10.7). It draws the state the debug
// process reports (through the extension) and reports what the user does; the form's code
// decides everything else (values, focus, closing), so the panel keeps no state of its own.
import "./formRunner.css";
import { contentInsets, oleToCss, px } from "../../src/formRunner/protocol";
import type { ControlSnapshot, ControlState, FontState, FormState, PanelMessage, UserAction } from "../../src/formRunner/protocol";

declare function acquireVsCodeApi(): { postMessage(message: unknown): void };
const vscode = acquireVsCodeApi();

function send(action: UserAction): void {
	vscode.postMessage({ type: "action", action });
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, className?: string, parent?: HTMLElement): HTMLElementTagNameMap[K] {
	const node = document.createElement(tag);
	if (className) node.className = className;
	parent?.appendChild(node);
	return node;
}

// ------------------------------------------------------------------ the window

const app = document.getElementById("app") as HTMLElement;
const frame = el("div", "form-window", app);
const titleBar = el("div", "form-title", frame);
const titleText = el("span", "form-caption", titleBar);
const closeButton = el("button", "form-close", titleBar);
closeButton.type = "button";
closeButton.title = "閉じる";
closeButton.textContent = "✕";
closeButton.addEventListener("click", () => send({ t: "close" }));
const client = el("div", "form-client", frame);
client.addEventListener("click", (e) => {
	if (e.target === client) send({ t: "click" });
});
client.addEventListener("dblclick", (e) => {
	if (e.target === client) send({ t: "dblclick" });
});
const banners = el("div", "banners", app);
const note = el("div", "note", app);
note.textContent = "デバッグ実行中のフォームです（RAS）。操作はフォームのコードで処理されます。";

function showBanner(text: string, kind: "error" | "info"): void {
	const banner = el("div", `banner ${kind}`, banners);
	el("span", "banner-text", banner).textContent = text;
	const dismiss = el("button", "banner-close", banner);
	dismiss.type = "button";
	dismiss.textContent = "✕";
	dismiss.title = "閉じる";
	dismiss.addEventListener("click", () => banner.remove());
	while (banners.children.length > 3) banners.firstElementChild?.remove();
	if (kind === "info") setTimeout(() => banner.remove(), 2500);
}

// ------------------------------------------------------------------ controls

interface View {
	type: string;
	root: HTMLElement;
	/** Where the children of a container go. */
	content?: HTMLElement;
	/** The element that takes the focus. */
	focus?: HTMLElement;
	update(c: ControlSnapshot, enabled: boolean): void;
}

let state: FormState | undefined;
const views = new Map<string, View>();
/**
 * Inputs sent per control. While the form's code has not seen the last one, the text the
 * panel shows is newer than the state's (typing goes on during a slow `Change` handler).
 */
let inputSeq = 0;
const sentInput = new Map<string, number>();

function sendInput(control: string, text: string, selStart?: number): void {
	inputSeq++;
	sentInput.set(control.toLowerCase(), inputSeq);
	send({ t: "input", control, text, seq: inputSeq, selStart });
}

/** The state's text of a control is not behind what the user typed. */
function textIsCurrent(c: ControlSnapshot): boolean {
	return (c.inputSeq ?? 0) >= (sentInput.get(c.name.toLowerCase()) ?? 0);
}
/** Set while the panel moves the focus itself, so that the move is not reported back. */
let applyingFocus = false;

const str = (v: unknown) => (v === null || v === undefined ? "" : String(v));
const bool = (v: unknown) => v === true;
const num = (v: unknown, fallback = 0) => (typeof v === "number" && Number.isFinite(v) ? v : fallback);

function fontCss(node: HTMLElement, font: FontState | undefined): void {
	if (!font) return;
	node.style.fontFamily = `${JSON.stringify(font.Name ?? "MS UI Gothic")}, "Yu Gothic UI", "Meiryo UI", "Segoe UI", sans-serif`;
	node.style.fontSize = `${num(font.Size, 9)}pt`;
	node.style.fontWeight = font.Bold ? "bold" : "normal";
	node.style.fontStyle = font.Italic ? "italic" : "normal";
	node.style.textDecoration = [font.Underline ? "underline" : "", font.Strikethrough ? "line-through" : ""].filter(Boolean).join(" ") || "none";
}

function colors(node: HTMLElement, s: ControlState, transparent = false): void {
	node.style.color = oleToCss(s.ForeColor) ?? "";
	node.style.backgroundColor = transparent ? "transparent" : (oleToCss(s.BackColor) ?? "");
}

const ALIGN = ["left", "left", "center", "right"];

function focusable(view: View, node: HTMLElement, name: string): void {
	view.focus = node;
	node.addEventListener("focus", () => {
		if (!applyingFocus) send({ t: "focus", control: name });
	});
}

function create(c: ControlSnapshot): View {
	const name = c.name;
	const root = el("div", `ctl ctl-${c.type.toLowerCase()}`);
	root.dataset.name = name;
	root.addEventListener("dblclick", (e) => {
		if ((e.target as HTMLElement).closest(".ctl") === root) send({ t: "dblclick", control: name });
	});
	const view: View = { type: c.type, root, update: () => undefined };
	switch (c.type) {
		case "Label": {
			const text = el("span", "label-text", root);
			root.addEventListener("click", () => send({ t: "click", control: name }));
			view.update = (s) => {
				text.textContent = str(s.state.Caption);
				root.style.whiteSpace = s.state.WordWrap === false ? "pre" : "pre-wrap";
				root.style.textAlign = ALIGN[num(s.state.TextAlign, 1)] ?? "left";
				colors(root, s.state, num(s.state.BackStyle, 1) === 0);
				root.classList.toggle("bordered", num(s.state.BorderStyle) === 1);
			};
			break;
		}
		case "TextBox": {
			let input: HTMLInputElement | HTMLTextAreaElement | undefined;
			let multi: boolean | undefined;
			view.update = (s, enabled) => {
				const isMulti = bool(s.state.MultiLine);
				if (input === undefined || multi !== isMulti) {
					input?.remove();
					multi = isMulti;
					input = el(isMulti ? "textarea" : "input", "field", root);
					const field = input;
					focusable(view, field, name);
					field.addEventListener("input", () => sendInput(name, field.value, field.selectionStart ?? undefined));
				}
				if (input instanceof HTMLInputElement) input.type = str(s.state.PasswordChar) ? "password" : "text";
				const text = str(s.state.Text);
				if (input.value !== text && textIsCurrent(s)) input.value = text;
				input.readOnly = bool(s.state.Locked);
				input.disabled = !enabled;
				const max = num(s.state.MaxLength);
				if (max > 0) input.maxLength = max;
				else input.removeAttribute("maxlength");
				input.style.textAlign = ALIGN[num(s.state.TextAlign, 1)] ?? "left";
				colors(input, s.state, num(s.state.BackStyle, 1) === 0);
			};
			break;
		}
		case "ComboBox": {
			let select: HTMLSelectElement | undefined;
			let input: HTMLInputElement | undefined;
			const listId = `list-${Math.random().toString(36).slice(2)}`;
			const datalist = el("datalist", "", root);
			datalist.id = listId;
			// The drop button of MSForms, drawn over the field (clicks go through to it).
			const symbol = el("span", "combo-chevron", el("span", "combo-arrow", root));
			view.update = (s, enabled) => {
				const list = (s.list ?? []).map((row) => row.join("  "));
				const dropDownList = num(s.state.Style) === 2;
				const button = num(s.state.ShowDropButtonWhen, 2);
				root.classList.toggle("drop-never", button === 0);
				root.classList.toggle("drop-focus", button === 1);
				// DropButtonStyle: Plain, Arrow, Ellipsis, Reduce.
				const style = num(s.state.DropButtonStyle, 1);
				symbol.className = ["", "combo-chevron", "combo-ellipsis", "combo-reduce"][style] ?? "combo-chevron";
				symbol.textContent = style === 2 ? "…" : "";
				if (dropDownList) {
					input?.remove();
					input = undefined;
					if (select === undefined) {
						const field = el("select", "field", root);
						select = field;
						focusable(view, field, name);
						field.addEventListener("change", () => send({ t: "select", control: name, index: field.selectedIndex }));
					}
					fillOptions(select, list);
					select.selectedIndex = s.listIndex ?? -1;
					select.disabled = !enabled || bool(s.state.Locked);
					colors(select, s.state);
				} else {
					select?.remove();
					select = undefined;
					if (input === undefined) {
						const field = el("input", "field", root);
						input = field;
						field.setAttribute("list", listId);
						focusable(view, field, name);
						field.addEventListener("input", (e) => {
							// Picking from the list replaces the whole text.
							const picked = (e as InputEvent).inputType === "insertReplacementText" || (e as InputEvent).inputType === undefined;
							const index = picked ? (state?.controls.find((x) => x.name === name)?.list ?? []).findIndex((row) => row.join("  ") === field.value || row[0] === field.value) : -1;
							if (index >= 0) send({ t: "select", control: name, index });
							else sendInput(name, field.value);
						});
					}
					while (datalist.firstChild) datalist.firstChild.remove();
					for (const item of list) el("option", "", datalist).value = item;
					const text = str(s.state.Text);
					if (input.value !== text && textIsCurrent(s)) input.value = text;
					input.readOnly = bool(s.state.Locked);
					input.disabled = !enabled;
					colors(input, s.state);
				}
			};
			break;
		}
		case "ListBox": {
			const select = el("select", "field", root);
			focusable(view, select, name);
			select.addEventListener("change", () => {
				const options = [...select.options];
				if (select.multiple) send({ t: "select", control: name, index: select.selectedIndex, selected: options.map((o) => o.selected) });
				else send({ t: "select", control: name, index: select.selectedIndex });
			});
			view.update = (s, enabled) => {
				const list = (s.list ?? []).map((row) => row.join("  "));
				select.multiple = num(s.state.MultiSelect) !== 0;
				select.size = Math.max(2, list.length);
				fillOptions(select, list);
				const selected = s.selected ?? [];
				[...select.options].forEach((o, i) => (o.selected = !!selected[i]));
				if (!select.multiple && !selected.some(Boolean)) select.selectedIndex = -1;
				select.disabled = !enabled || bool(s.state.Locked);
				colors(select, s.state);
			};
			break;
		}
		case "CommandButton":
		case "ToggleButton": {
			const button = el("button", "button", root);
			button.type = "button";
			focusable(view, button, name);
			button.addEventListener("click", () => send({ t: "click", control: name }));
			view.update = (s, enabled) => {
				button.textContent = str(s.state.Caption);
				button.disabled = !enabled;
				button.classList.toggle("default", bool(s.state.Default));
				if (c.type === "ToggleButton") {
					button.classList.toggle("pressed", s.state.Value === true);
					button.classList.toggle("mixed", s.state.Value === null);
				}
				colors(button, s.state, num(s.state.BackStyle, 1) === 0);
			};
			break;
		}
		case "CheckBox":
		case "OptionButton": {
			const label = el("label", "toggle", root);
			const box = el("input", "", label);
			box.type = c.type === "CheckBox" ? "checkbox" : "radio";
			const caption = el("span", "", label);
			focusable(view, box, name);
			// The form's code sets the value; the box shows it when the state comes back.
			box.addEventListener("click", (e) => {
				e.preventDefault();
				send({ t: "click", control: name });
			});
			view.update = (s, enabled) => {
				caption.textContent = str(s.state.Caption);
				box.checked = s.state.Value === true;
				label.classList.toggle("null", s.state.Value === null);
				box.disabled = !enabled;
				label.classList.toggle("disabled", !enabled);
				colors(root, s.state, true);
			};
			break;
		}
		case "SpinButton": {
			const up = el("button", "spin up", root);
			const down = el("button", "spin down", root);
			up.type = down.type = "button";
			focusable(view, up, name);
			down.addEventListener("focus", () => {
				if (!applyingFocus) send({ t: "focus", control: name });
			});
			up.addEventListener("click", () => send({ t: "spin", control: name, delta: 1 }));
			down.addEventListener("click", () => send({ t: "spin", control: name, delta: -1 }));
			view.update = (s, enabled) => {
				const horizontal = num(s.state.Orientation, -1) === 1 || (num(s.state.Orientation, -1) === -1 && num(s.state.Width) > num(s.state.Height));
				root.classList.toggle("horizontal", horizontal);
				up.textContent = horizontal ? "▶" : "▲";
				down.textContent = horizontal ? "◀" : "▼";
				up.disabled = down.disabled = !enabled;
				colors(root, s.state);
			};
			break;
		}
		case "ScrollBar": {
			const range = el("input", "range", root);
			range.type = "range";
			focusable(view, range, name);
			range.addEventListener("input", () => send({ t: "scroll", control: name, value: Number(range.value) }));
			view.update = (s, enabled) => {
				const lo = Math.min(num(s.state.Min), num(s.state.Max));
				const hi = Math.max(num(s.state.Min), num(s.state.Max));
				range.min = String(lo);
				range.max = String(hi);
				range.step = String(Math.max(1, num(s.state.SmallChange, 1)));
				range.value = String(num(s.state.Value));
				range.disabled = !enabled;
				const vertical = num(s.state.Orientation, -1) === 0 || (num(s.state.Orientation, -1) === -1 && num(s.state.Height) > num(s.state.Width));
				root.classList.toggle("vertical", vertical);
				range.style.width = vertical ? px(s.state.Height) : "100%";
			};
			break;
		}
		case "Image": {
			const img = el("img", "picture", root);
			img.alt = "";
			img.draggable = false;
			root.addEventListener("click", () => send({ t: "click", control: name }));
			view.update = (s) => {
				if (s.picture) img.src = s.picture;
				else img.removeAttribute("src");
				img.style.display = s.picture ? "" : "none";
				const mode = num(s.state.PictureSizeMode);
				img.style.objectFit = mode === 1 ? "fill" : mode === 3 ? "contain" : "none";
				img.style.objectPosition = ["left top", "right top", "center", "left bottom", "right bottom"][num(s.state.PictureAlignment, 2)] ?? "center";
				colors(root, s.state);
				root.classList.toggle("bordered", num(s.state.BorderStyle, 1) === 1);
			};
			break;
		}
		case "Frame": {
			const caption = el("span", "frame-caption", root);
			const content = el("div", "content", root);
			view.content = content;
			root.addEventListener("click", (e) => {
				if (e.target === root || e.target === content) send({ t: "click", control: name });
			});
			view.update = (s) => {
				caption.textContent = str(s.state.Caption);
				caption.style.display = str(s.state.Caption) ? "" : "none";
				root.classList.toggle("has-caption", str(s.state.Caption) !== "");
				colors(root, s.state);
				caption.style.backgroundColor = root.style.backgroundColor;
				const inset = contentInsets("Frame");
				content.style.left = px(inset.left);
				content.style.top = px(inset.top);
			};
			break;
		}
		case "MultiPage":
		case "TabStrip": {
			const tabs = el("div", "tabs", root);
			const content = el("div", "content", root);
			if (c.type === "MultiPage") view.content = content;
			view.update = (s, enabled) => {
				const pages = c.type === "MultiPage" ? (state?.controls ?? []).filter((x) => x.parent === name && x.type === "Page").map((p) => p.state) : (s.tabs ?? []);
				const current = num(s.state.Value);
				while (tabs.firstChild) tabs.firstChild.remove();
				pages.forEach((page, index) => {
					if (page.Visible === false) return;
					const tab = el("button", index === current ? "tab active" : "tab", tabs);
					tab.type = "button";
					el("span", "tab-text", tab).textContent = str(page.Caption) || str(page.Name);
					tab.disabled = !enabled || page.Enabled === false;
					tab.tabIndex = index === current ? 0 : -1;
					tab.addEventListener("focus", () => {
						if (!applyingFocus) send({ t: "focus", control: name });
					});
					tab.addEventListener("click", () => send({ t: "page", control: name, index }));
					if (index === current) view.focus = tab;
				});
				colors(root, s.state);
				const inset = contentInsets(c.type, s.state.TabOrientation);
				content.style.left = px(inset.left);
				content.style.top = px(inset.top);
				content.style.right = px(inset.right);
				content.style.bottom = px(inset.bottom);
			};
			break;
		}
		case "Page": {
			const content = el("div", "content", root);
			view.content = content;
			view.update = (s) => colors(root, s.state);
			break;
		}
		default:
			root.textContent = `${c.type}: ${name}`;
			root.classList.add("unknown");
	}
	return view;
}

/** The plain properties that change how any control looks (formRunner.css draws them). */
function look(root: HTMLElement, s: ControlState, type: string): void {
	const effect = s.SpecialEffect;
	if (typeof effect === "number") root.dataset.effect = String(effect);
	else delete root.dataset.effect;
	root.classList.toggle("caption-left", s.Alignment === 0);
	root.classList.toggle("wrap", s.WordWrap === true);
	root.classList.toggle("no-wrap", s.WordWrap === false);
	const border = oleToCss(s.BorderColor);
	if (border) root.style.setProperty("--border-color", border);
	else root.style.removeProperty("--border-color");
	if (type === "MultiPage" || type === "TabStrip") {
		for (const n of [0, 1, 2]) root.classList.toggle(`tab-style-${n}`, num(s.Style) === n);
		for (const n of [1, 2, 3]) root.classList.toggle(`tab-orient-${n}`, num(s.TabOrientation) === n);
	}
}

function fillOptions(select: HTMLSelectElement, items: string[]): void {
	const current = [...select.options].map((o) => o.text);
	if (current.length === items.length && current.every((t, i) => t === items[i])) return;
	while (select.firstChild) select.firstChild.remove();
	for (const item of items) el("option", "", select).text = item;
}

// ------------------------------------------------------------------ drawing a state

function render(next: FormState, focus: string | undefined): void {
	state = next;
	const form = next.form;
	titleText.textContent = str(form.Caption);
	client.style.width = px(form.Width);
	client.style.height = px(form.Height);
	colors(client, form);
	fontCss(client, form.Font as FontState | undefined);
	frame.classList.toggle("disabled", form.Enabled === false);

	const byName = new Map(next.controls.map((c) => [c.name.toLowerCase(), c]));
	for (const [key, view] of views) {
		const c = byName.get(key);
		if (!c || c.type !== view.type) {
			view.root.remove();
			views.delete(key);
		}
	}
	const enabledOf = (c: ControlSnapshot): boolean => {
		for (let x: ControlSnapshot | undefined = c, guard = 0; x && guard < 64; x = x.parent ? byName.get(x.parent.toLowerCase()) : undefined, guard++) {
			if (x.state.Enabled === false) return false;
		}
		return form.Enabled !== false;
	};
	// Containers first, so that their children have a place to go.
	for (const c of next.controls) {
		const key = c.name.toLowerCase();
		let view = views.get(key);
		if (!view) {
			view = create(c);
			views.set(key, view);
		}
	}
	for (const c of next.controls) {
		const view = views.get(c.name.toLowerCase()) as View;
		const parent = c.parent ? views.get(c.parent.toLowerCase()) : undefined;
		const home = parent?.content ?? client;
		if (view.root.parentElement !== home) home.appendChild(view.root);
		const s = c.state;
		if (c.type === "Page") {
			const owner = parent ? byName.get(c.parent.toLowerCase()) : undefined;
			const pages = next.controls.filter((x) => x.parent === c.parent && x.type === "Page");
			view.root.style.display = owner && num(owner.state.Value) === pages.indexOf(c) && s.Visible !== false ? "" : "none";
		} else {
			view.root.style.left = px(s.Left);
			view.root.style.top = px(s.Top);
			view.root.style.width = px(s.Width);
			view.root.style.height = px(s.Height);
			view.root.style.display = s.Visible === false ? "none" : "";
		}
		view.root.title = str(s.ControlTipText);
		fontCss(view.root, s.Font);
		view.update(c, enabledOf(c));
		look(view.root, s, c.type);
		if (view.focus) view.focus.tabIndex = s.TabStop === false ? -1 : num(s.TabIndex) + 1;
	}
	if (focus) {
		const target = views.get(focus.toLowerCase())?.focus;
		if (target && document.activeElement !== target) {
			applyingFocus = true;
			try {
				target.focus();
			} finally {
				applyingFocus = false;
			}
		}
	}
}

// Enter presses the default button and Esc the cancel button, as in a dialog box of Windows.
document.addEventListener("keydown", (e) => {
	if (!state || (e.key !== "Enter" && e.key !== "Escape")) return;
	const target = e.target as HTMLElement;
	if (e.key === "Enter" && (target instanceof HTMLTextAreaElement || target instanceof HTMLButtonElement)) return;
	const property = e.key === "Enter" ? "Default" : "Cancel";
	const button = state.controls.find((c) => c.type === "CommandButton" && c.state[property] === true && c.state.Visible !== false && c.state.Enabled !== false);
	if (!button) return;
	e.preventDefault();
	send({ t: "click", control: button.name, key: true });
});

window.addEventListener("message", (event: MessageEvent<PanelMessage>) => {
	const message = event.data;
	switch (message.type) {
		case "state":
			render(message.state, message.focus);
			break;
		case "error":
			showBanner(message.message, "error");
			break;
		case "blocked":
			showBanner("モーダル フォームが表示されている間は、ほかのフォームは操作できません。", "info");
			break;
		case "ended":
			showBanner("プログラムは終了しました。", "info");
			break;
	}
});

vscode.postMessage({ type: "ready" });
