import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { applyOperations, copyControls } from "../../src/userForm/engine";
import { bodyIndent, eventsForType, eventTarget, findHandler, handlerText, jsName, rasParams } from "../../src/userForm/events";
import { formatJson, minimalEdit } from "../../src/userForm/jsonSource";
import { columnLayout, itemColumns, listToText, parseColumnWidths, textColumnIndex, textToList } from "../../src/userForm/listItems";
import { designerDocument, toggleState } from "../../src/userForm/model";
import { isStructuralOperation, isWebviewMessage } from "../../src/userForm/protocol";
import { createControl } from "../../src/userForm/toolbox";
import type { SemanticOperation } from "../../src/userForm/protocol";
import { pasteAllocations } from "../../src/userForm/clipboard";
import { newFormSpec, parseSpec, previewDocument, problems, propertyGrid } from "../../src/userForm/spec";

type Json = Record<string, any>;

const form = (controls: Json[] = [], formFields: Json = {}) =>
	formatJson({ ...newFormSpec("Main"), form: { name: "Main", caption: "Main", width: 240, height: 180, ...formFields }, controls });

const apply = (text: string, ...operations: SemanticOperation[]) => applyOperations(text, operations).text;
const json = (text: string): Json => JSON.parse(text);
const control = (text: string, id: string): Json => json(text).controls.find((c: Json) => c.id === id);
const rejects = (text: string, operations: SemanticOperation[], code?: string) =>
	assert.throws(() => applyOperations(text, operations), (error: any) => (code ? error.code === code : true));

const label = { id: "l1", name: "Label1", type: "Label", caption: "Label1", left: 8, top: 8, width: 72, height: 18 };
const frame = { id: "f1", name: "Frame1", type: "Frame", caption: "Frame1", left: 8, top: 40, width: 144, height: 108 };

describe("FormSpec editing", () => {
	it("keeps the source formatting and only changes what an edit touches", () => {
		const text = form([label]);
		const after = apply(text, { type: "moveControl", controlId: "l1", left: 16, top: 24 });
		assert.equal(after, text.replace('"left": 8,\n      "top": 8', '"left": 16,\n      "top": 24'));
		const crlf = text.replace(/\n/g, "\r\n").replace(/ {2}/g, "\t");
		const edited = apply(crlf, { type: "resizeControl", controlId: "l1", width: 80, height: 20 });
		assert.ok(edited.includes('\r\n\t\t\t"width": 80,'));
		assert.deepEqual(minimalEdit("abcdef", "abXYef"), { start: 2, end: 4, text: "XY" });
	});

	it("sets form geometry and properties, merging partial fonts", () => {
		let text = form();
		text = apply(text, { type: "setGeometry", width: 300, height: 200 });
		text = apply(text, { type: "setFormProperty", field: "font", value: { name: "Meiryo UI", size: 9 } });
		text = apply(text, { type: "setFormProperty", field: "font", value: { bold: true } });
		text = apply(text, { type: "setFormProperty", field: "backColor", value: { kind: "system", value: "buttonFace" } });
		const result = json(text).form;
		assert.equal(result.width, 300);
		assert.deepEqual(result.font, { name: "Meiryo UI", size: 9, bold: true });
		assert.deepEqual(result.backColor, { kind: "system", value: "buttonFace" });
		rejects(text, [{ type: "setFormProperty", field: "foreColor", value: { kind: "rgb", value: "red" } }], "UFV021");
		rejects(text, [{ type: "setFormProperty", field: "observed", value: "x" }], "UFE005");
	});

	it("creates labels, check boxes and option buttons with a transparent background", () => {
		const parsed = parseSpec(form());
		assert.ok(parsed.ok);
		const document = designerDocument(previewDocument(parsed.spec));
		for (const type of ["Label", "CheckBox", "OptionButton"] as const) {
			const created = createControl(document, type, 8, 8, true, () => "x")!;
			assert.equal(created.backStyle, 0);
			const op = { type: "addControl", control: created } as const;
			assert.ok(isStructuralOperation(op));
			assert.equal(control(apply(form(), op), created.id).backStyle, 0);
		}
		for (const type of ["CommandButton", "ToggleButton", "TextBox", "ComboBox"] as const)
			assert.equal(Object.hasOwn(createControl(document, type, 8, 8, true, () => "x")!, "backStyle"), false);
		assert.equal(isStructuralOperation({ type: "addControl", control: { id: "t", name: "TextBox1", type: "TextBox", left: 0, top: 0, width: 72, height: 18, backStyle: 0 } }), false);
	});

	it("adds, edits and removes controls", () => {
		let text = form([frame]);
		text = apply(text, { type: "addControl", control: { id: "c1", parentId: "f1", name: "CommandButton1", type: "CommandButton", caption: "OK", left: 4, top: 4, width: 72, height: 24 } });
		assert.deepEqual(Object.keys(control(text, "c1")), ["id", "parentId", "name", "type", "caption", "left", "top", "width", "height"]);
		text = apply(text, { type: "setControlProperty", controlId: "c1", field: "caption", value: "Run" });
		text = apply(text, { type: "setControlProperty", controlId: "f1", field: "borderStyle", value: 1 });
		assert.equal(control(text, "c1").caption, "Run");
		rejects(text, [{ type: "setControlProperty", controlId: "c1", field: "textAlign", value: 2 }], "UFE005");
		rejects(text, [{ type: "setControlProperty", controlId: "f1", field: "borderStyle", value: 7 }], "UFE003");
		rejects(text, [{ type: "removeControl", controlId: "f1", cascade: false }], "UFE007");
		text = apply(text, { type: "removeControl", controlId: "f1", cascade: true });
		assert.deepEqual(json(text).controls, []);
	});

	it("edits the list properties of a ComboBox and a ListBox", () => {
		const combo = { id: "c1", name: "ComboBox1", type: "ComboBox", left: 8, top: 8, width: 120, height: 18 };
		const list = { id: "b1", name: "ListBox1", type: "ListBox", left: 8, top: 40, width: 120, height: 72 };
		let text = form([combo, list]);
		text = apply(
			text,
			{ type: "setControlProperty", controlId: "c1", field: "style", value: 2 },
			{ type: "setControlProperty", controlId: "c1", field: "showDropButtonWhen", value: 1 },
			{ type: "setControlProperty", controlId: "c1", field: "listRows", value: 12 },
			{ type: "setControlProperty", controlId: "b1", field: "multiSelect", value: 1 },
			{ type: "setControlProperty", controlId: "b1", field: "listStyle", value: 1 },
		);
		assert.deepEqual(
			[control(text, "c1").style, control(text, "c1").showDropButtonWhen, control(text, "c1").listRows],
			[2, 1, 12],
		);
		assert.deepEqual([control(text, "b1").multiSelect, control(text, "b1").listStyle], [1, 1]);
		rejects(text, [{ type: "setControlProperty", controlId: "c1", field: "style", value: 1 }], "UFE003");
		rejects(text, [{ type: "setControlProperty", controlId: "c1", field: "multiSelect", value: 1 }], "UFE005");
		rejects(text, [{ type: "setControlProperty", controlId: "b1", field: "showDropButtonWhen", value: 0 }], "UFE005");
		const grid = propertyGrid((parseSpec(text) as any).spec);
		const style = grid.controls.c1.descriptors.find((d) => d.field === "style");
		assert.deepEqual(style?.suggestions, [
			{ value: "0", label: "DropDownCombo" },
			{ value: "2", label: "DropDownList" },
		]);
		assert.equal(previewDocument((parseSpec(text) as any).spec).controls[0].style, 2);
	});

	it("edits the plain properties of each control type", () => {
		const controls = [
			{ id: "t1", name: "TextBox1", type: "TextBox", left: 8, top: 8, width: 120, height: 18 },
			{ id: "b1", name: "OK", type: "CommandButton", left: 8, top: 30, width: 72, height: 24 },
			{ id: "s1", name: "Spin1", type: "SpinButton", left: 140, top: 8, width: 18, height: 36 },
			{ id: "x1", name: "Tabs1", type: "TabStrip", left: 8, top: 60, width: 200, height: 60 },
			{ id: "k1", name: "Check1", type: "CheckBox", left: 140, top: 50, width: 72, height: 18 },
		];
		let text = form(controls);
		const set = (controlId: string, field: string, value: unknown): SemanticOperation => ({ type: "setControlProperty", controlId, field, value });
		text = apply(
			text,
			set("t1", "locked", true),
			set("t1", "passwordChar", "*"),
			set("t1", "specialEffect", 0),
			set("t1", "borderColor", { kind: "rgb", value: "#FF0000" }),
			set("b1", "default", true),
			set("b1", "takeFocusOnClick", false),
			set("s1", "min", 10),
			set("s1", "max", 20),
			set("s1", "value", 15),
			set("x1", "style", 1),
			set("x1", "tabOrientation", 1),
			set("k1", "alignment", 0),
		);
		assert.deepEqual(
			[control(text, "t1").locked, control(text, "t1").passwordChar, control(text, "b1").default, control(text, "s1").value, control(text, "x1").style],
			[true, "*", true, 15, 1],
		);
		// Values outside what the property takes, and properties of other types, are refused.
		rejects(text, [set("t1", "specialEffect", 4)], "UFE003");
		rejects(text, [set("s1", "value", 30)], "UFV021");
		rejects(text, [set("t1", "passwordChar", "**")], "UFV021");
		rejects(text, [set("b1", "multiLine", true)], "UFE005");
		rejects(text, [set("k1", "style", 2)], "UFE005");
		// The grid labels values by their VBA constants, per control type.
		const grid = propertyGrid((parseSpec(text) as any).spec);
		const labels = (id: string, field: string) =>
			grid.controls[id].descriptors.find((d) => d.field === field)?.suggestions?.map((s) => s.label);
		assert.deepEqual(labels("x1", "style"), ["Tabs", "Buttons", "None"]);
		assert.deepEqual(labels("k1", "alignment"), ["Left", "Right"]);
		assert.deepEqual(labels("t1", "specialEffect"), ["Flat", "Raised", "Sunken", "Etched", "Bump"]);
	});

	it("keeps the pitch of a pulled font through font edits", () => {
		const pulled = { ...label, font: { name: "MS UI Gothic", size: 9, charset: 128, pitchAndFamily: 2 } };
		let text = form([pulled]);
		assert.deepEqual(problems((parseSpec(text) as any).spec).filter((p: any) => p.severity === "error"), []);
		text = apply(text, { type: "setControlProperty", controlId: "l1", field: "font", value: { size: 12 } });
		assert.deepEqual(control(text, "l1").font, { name: "MS UI Gothic", size: 12, charset: 128, pitchAndFamily: 2 });
		// The font editor does not offer it.
		const grid = propertyGrid((parseSpec(text) as any).spec);
		const font = grid.controls.l1.descriptors.find((d) => d.field === "font");
		assert.ok(font?.children?.some((c) => c.field === "charset") && !font.children.some((c) => c.field === "pitchAndFamily"));
		rejects(text, [{ type: "setControlProperty", controlId: "l1", field: "font", value: { pitchAndFamily: 300 } }], "UFV021");
	});

	it("edits TripleState and a Null value of the toggles", () => {
		const controls = [
			{ id: "k1", name: "Check1", type: "CheckBox", left: 8, top: 8, width: 72, height: 18 },
			{ id: "o1", name: "Option1", type: "OptionButton", left: 8, top: 30, width: 72, height: 18, value: "True" },
		];
		let text = form(controls);
		const set = (controlId: string, field: string, value: unknown): SemanticOperation => ({ type: "setControlProperty", controlId, field, value });
		text = apply(text, set("k1", "tripleState", true), set("k1", "value", "Null"), set("o1", "tripleState", true));
		assert.deepEqual([control(text, "k1").tripleState, control(text, "k1").value, control(text, "o1").tripleState], [true, "Null", true]);
		rejects(text, [set("l1", "tripleState", true)]);
		const grid = propertyGrid((parseSpec(text) as any).spec);
		const descriptor = (id: string, field: string) => grid.controls[id].descriptors.find((d) => d.field === field);
		assert.equal(descriptor("k1", "value")?.editor, "toggle");
		assert.equal(descriptor("k1", "tripleState")?.editor, "boolean");
		// What a value stands for, however it was written.
		assert.deepEqual(
			[true, "True", -1, "1", false, "False", 0, undefined, "Null", "null"].map(toggleState),
			["on", "on", "on", "on", "off", "off", "off", "off", "null", "null"],
		);
	});

	it("binds controls to cells with controlSource and rowSource", () => {
		const controls = [
			{ id: "t1", name: "Text1", type: "TextBox", left: 8, top: 8, width: 120, height: 18 },
			{ id: "b1", name: "List1", type: "ListBox", left: 8, top: 30, width: 120, height: 60 },
		];
		let text = form(controls);
		const set = (controlId: string, field: string, value: unknown): SemanticOperation => ({ type: "setControlProperty", controlId, field, value });
		text = apply(text, set("t1", "controlSource", "Sheet1!B2"), set("b1", "rowSource", "Sheet1!A2:A9"), set("b1", "controlSource", "Sheet1!C2"));
		assert.deepEqual([control(text, "t1").controlSource, control(text, "b1").rowSource], ["Sheet1!B2", "Sheet1!A2:A9"]);
		rejects(text, [set("t1", "rowSource", "A1:A3")], "UFE005");
		rejects(text, [set("l1", "controlSource", "A1")]);
		// Items cannot be written for a list that comes from cells.
		rejects(text, [set("b1", "list", ["a", "b"])], "UFV021");
	});

	it("edits the properties of the form itself", () => {
		let text = form([label]);
		const set = (field: string, value: unknown): SemanticOperation => ({ type: "setFormProperty", field, value });
		text = apply(text, set("startUpPosition", 0), set("showModal", false), set("tag", "メモ"), set("zoom", 150), set("scrollBars", 3));
		const f = json(text).form;
		assert.deepEqual([f.startUpPosition, f.showModal, f.tag, f.zoom, f.scrollBars], [0, false, "メモ", 150, 3]);
		rejects(text, [set("zoom", 5)], "UFV021");
		rejects(text, [set("startUpPosition", 4)], "UFE003");
		rejects(text, [set("tag", "a\nb")], "UFV021");
		rejects(text, [set("mousePointer", 4)], "UFE003");
	});

	it("raises a TextBox to the height its font needs, but only where the edit touched", () => {
		const low = { id: "t0", name: "Untouched", type: "TextBox", left: 8, top: 60, width: 120, height: 12 };
		let text = form([{ id: "t1", name: "TextBox1", type: "TextBox", left: 8, top: 8, width: 120, height: 18 }, low]);
		// 9pt MS UI Gothic fits in 18pt; 16pt Meiryo UI needs 16 × 1.33 + 4.5 = 25.78 → 26.25pt (35 px).
		text = apply(text, { type: "setControlProperty", controlId: "t1", field: "font", value: { name: "Meiryo UI", size: 16 } });
		assert.equal(control(text, "t1").height, 26.25);
		text = apply(text, { type: "resizeControl", controlId: "t1", width: 120, height: 10 });
		assert.equal(control(text, "t1").height, 26.25);
		text = apply(text, { type: "resizeControl", controlId: "t1", width: 120, height: 40 });
		assert.equal(control(text, "t1").height, 40);
		// A control the edit did not touch keeps its height; validation warns about it.
		assert.equal(control(text, "t0").height, 12);
		const warnings = problems((parseSpec(text) as any).spec).filter((p) => p.code === "UFV022");
		assert.deepEqual(
			warnings.map((p) => [p.severity, p.field]),
			[["warning", "controls[1].height"]],
		);
		// Labels are not edit fields.
		text = apply(form([label]), { type: "setControlProperty", controlId: "l1", field: "font", value: { name: "Meiryo UI", size: 20 } });
		assert.equal(control(text, "l1").height, 18);
	});

	it("warns about a Single border with an effect, which MSForms cannot draw together", () => {
		const both = { id: "t1", name: "TextBox1", type: "TextBox", left: 8, top: 8, width: 120, height: 18, borderStyle: 1, specialEffect: 2 };
		const warnings = (text: string) =>
			problems((parseSpec(text) as any).spec)
				.filter((p) => p.code === "UFV023")
				.map((p) => [p.severity, p.field]);
		assert.deepEqual(warnings(form([both])), [["warning", "controls[0].specialEffect"]]);
		assert.deepEqual(warnings(form([], { borderStyle: 1, specialEffect: 3 })), [["warning", "form.specialEffect"]]);
		// One of them, or a flat effect: no warning.
		assert.deepEqual(warnings(form([{ ...both, specialEffect: 0 }])), []);
		assert.deepEqual(warnings(form([{ ...both, borderStyle: 0 }])), []);
	});

	it("moves controls between parents and refuses invalid containers", () => {
		let text = form([frame, label]);
		text = apply(text, { type: "setParent", controlId: "l1", parentId: "f1", left: 2, top: 3 });
		assert.equal(control(text, "l1").parentId, "f1");
		assert.equal(control(text, "l1").left, 2);
		rejects(text, [{ type: "setParent", controlId: "f1", parentId: "l1", left: 0, top: 0 }]);
		text = apply(text, { type: "setParent", controlId: "l1", parentId: "", left: 9, top: 9 });
		assert.equal(control(text, "l1").parentId, undefined);
	});

	it("creates a MultiPage with two pages in one transaction and edits the page collection", () => {
		let text = form();
		text = apply(
			text,
			{ type: "addControl", control: { id: "mp", name: "MultiPage1", type: "MultiPage", left: 0, top: 0, width: 200, height: 150, selectedIndex: -1 } },
			{ type: "addPage", control: { id: "p1", parentId: "mp", name: "Page1", type: "Page", caption: "Page1" } },
			{ type: "addPage", control: { id: "p2", parentId: "mp", name: "Page2", type: "Page", caption: "Page2" } },
		);
		assert.equal(control(text, "mp").selectedIndex, 0);
		assert.deepEqual([control(text, "p1").zIndex, control(text, "p2").zIndex], [0, 1]);
		text = apply(text, { type: "reorderPage", controlId: "p2", index: 0 });
		assert.deepEqual([control(text, "p2").zIndex, control(text, "p1").zIndex], [0, 1]);
		assert.equal(control(text, "mp").selectedIndex, 1);
		text = apply(text, { type: "addControl", control: { id: "t1", parentId: "p1", name: "TextBox1", type: "TextBox", left: 0, top: 0, width: 50, height: 18 } });
		text = apply(text, { type: "removePage", controlId: "p1", cascade: true });
		assert.equal(control(text, "t1"), undefined);
		assert.equal(control(text, "mp").selectedIndex, 0);
		assert.equal(parseSpec(text).ok && (parseSpec(text) as any).spec.issues.length, 0);
	});

	it("edits TabStrip tabs and keeps selectedIndex valid", () => {
		let text = form([{ id: "ts", name: "TabStrip1", type: "TabStrip", left: 0, top: 0, width: 200, height: 40, selectedIndex: 0, tabs: [{ name: "Tab1" }, { name: "Tab2" }] }]);
		text = apply(text, { type: "addTab", controlId: "ts", tab: { name: "Tab3", caption: "Three" } });
		text = apply(text, { type: "reorderTab", controlId: "ts", index: 0, toIndex: 2 });
		assert.deepEqual(control(text, "ts").tabs.map((t: Json) => t.name), ["Tab2", "Tab3", "Tab1"]);
		assert.equal(control(text, "ts").selectedIndex, 2);
		text = apply(text, { type: "setTabProperty", controlId: "ts", index: 1, field: "enabled", value: false });
		text = apply(text, { type: "removeTab", controlId: "ts", index: 2 });
		assert.equal(control(text, "ts").selectedIndex, 0);
		// An edit that introduces a validation error is rejected.
		rejects(text, [{ type: "setControlProperty", controlId: "ts", field: "selectedIndex", value: 5 }], "UFV020");
		rejects(text, [{ type: "addTab", controlId: "ts", tab: { name: "tab2" } }], "UFV019");
	});

	it("sets and removes the picture of an Image", () => {
		const image = { id: "img", name: "Image1", type: "Image", left: 8, top: 8, width: 72, height: 72 };
		let text = form([image, label]);
		text = apply(text, { type: "setControlProperty", controlId: "img", field: "picture", value: { path: "forms/assets/logo.bmp" } });
		assert.deepEqual(control(text, "img").picture, { path: "forms/assets/logo.bmp" });
		rejects(text, [{ type: "setControlProperty", controlId: "img", field: "picture", value: { path: "" } }], "UFV005");
		rejects(text, [{ type: "setControlProperty", controlId: "img", field: "picture", value: { file: "a.bmp" } as never }], "UFE003");
		// A Label, the buttons, Frames, Pages and the form have pictures too; a TextBox has none.
		const withLabel = apply(text, { type: "setControlProperty", controlId: "l1", field: "picture", value: { path: "a.bmp" } }, { type: "setControlProperty", controlId: "l1", field: "picturePosition", value: 12 });
		assert.deepEqual([control(withLabel, "l1").picture, control(withLabel, "l1").picturePosition], [{ path: "a.bmp" }, 12]);
		const withForm = apply(text, { type: "setFormProperty", field: "picture", value: { path: "a.bmp" } }, { type: "setFormProperty", field: "pictureTiling", value: true }, { type: "setFormProperty", field: "mouseIcon", value: { path: "hand.cur" } });
		assert.deepEqual([json(withForm).form.picture, json(withForm).form.pictureTiling, json(withForm).form.mouseIcon], [{ path: "a.bmp" }, true, { path: "hand.cur" }]);
		const withIcon = apply(text, { type: "setControlProperty", controlId: "l1", field: "mouseIcon", value: { path: "hand.cur" } }, { type: "setControlProperty", controlId: "l1", field: "mousePointer", value: 99 });
		assert.deepEqual([control(withIcon, "l1").mouseIcon, control(withIcon, "l1").mousePointer], [{ path: "hand.cur" }, 99]);
		const box = apply(text, { type: "addControl", control: { id: "t1", name: "Text1", type: "TextBox", left: 8, top: 90, width: 72, height: 18 } as never });
		rejects(box, [{ type: "setControlProperty", controlId: "t1", field: "picture", value: { path: "a.bmp" } }], "UFE005");
		// Removing writes no picture field at all, and restores the source text.
		const removed = apply(text, { type: "setControlProperty", controlId: "img", field: "picture", value: { remove: true } });
		assert.equal(removed, form([image, label]));
		assert.equal(apply(removed, { type: "setControlProperty", controlId: "img", field: "picture", value: { remove: true } }), removed);
	});

	it("edits the list of a ComboBox / ListBox and keeps selectedIndex valid", () => {
		const combo = { id: "cb", name: "ComboBox1", type: "ComboBox", left: 8, top: 8, width: 120, height: 18 };
		let text = form([combo, label]);
		text = apply(text, { type: "setControlProperty", controlId: "cb", field: "list", value: ["Apple", "Banana", ""] });
		text = apply(text, { type: "setControlProperty", controlId: "cb", field: "selectedIndex", value: 2 });
		assert.deepEqual(control(text, "cb").list, ["Apple", "Banana", ""]);
		rejects(text, [{ type: "setControlProperty", controlId: "cb", field: "selectedIndex", value: 3 }], "UFV020");
		rejects(text, [{ type: "setControlProperty", controlId: "cb", field: "selectedIndex", value: -2 }], "UFV020");
		rejects(text, [{ type: "setControlProperty", controlId: "cb", field: "list", value: [1] as never }], "UFE003");
		rejects(text, [{ type: "setControlProperty", controlId: "cb", field: "list", value: "Apple" }], "UFE003");
		rejects(text, [{ type: "setControlProperty", controlId: "l1", field: "list", value: ["a"] }], "UFE005");
		// A selection the shorter list no longer has becomes "none".
		text = apply(text, { type: "setControlProperty", controlId: "cb", field: "list", value: ["Apple", "Banana"] });
		assert.equal(control(text, "cb").selectedIndex, -1);
		text = apply(text, { type: "setControlProperty", controlId: "cb", field: "selectedIndex", value: 1 });
		text = apply(text, { type: "setControlProperty", controlId: "cb", field: "list", value: ["Apple", "Cherry", "Banana"] });
		assert.equal(control(text, "cb").selectedIndex, 1);
		// null removes the list, and with it the selection.
		text = apply(text, { type: "setControlProperty", controlId: "cb", field: "list", value: null });
		assert.equal("list" in control(text, "cb"), false);
		assert.equal(control(text, "cb").selectedIndex, -1);
		rejects(text, [{ type: "setControlProperty", controlId: "cb", field: "selectedIndex", value: 0 }], "UFV020");
		text = apply(text, { type: "setControlProperty", controlId: "cb", field: "list", value: ["Apple"] });
		text = apply(text, { type: "setControlProperty", controlId: "cb", field: "selectedIndex", value: 0 });
		assert.equal(control(text, "cb").selectedIndex, 0);
	});

	it("converts list items to and from the one-item-per-line text", () => {
		assert.equal(listToText(["a", "", "b"]), "a\n\nb");
		assert.equal(listToText(undefined), "");
		assert.equal(listToText("a"), "");
		assert.deepEqual(textToList("a\r\n\nb \n\n"), ["a", "", "b "]);
		assert.equal(textToList(""), null);
		assert.equal(textToList("\n\n"), null);
		assert.deepEqual(textToList(listToText(["x", "y"])), ["x", "y"]);
		// The columns of an item are separated by tabs (as Excel copies a range).
		assert.equal(listToText([["001", "山田"], "002", ["003", ""]]), "001\t山田\n002\n003\t");
		assert.deepEqual(textToList("001\t山田\n002\n003\t"), [["001", "山田"], "002", ["003", ""]]);
		assert.equal(listToText([["a", 1]]), "");
	});

	it("lays out the columns of a list as MSForms does", () => {
		assert.deepEqual(parseColumnWidths("72 pt;;1 in;2.54cm;-1;x"), [72, undefined, 72, 72, undefined, undefined]);
		assert.deepEqual(parseColumnWidths(""), []);
		// One column takes the whole width; the open columns share what is left (72pt each at least).
		assert.deepEqual(columnLayout(undefined, undefined, 120, 3), [120]);
		assert.deepEqual(columnLayout(3, "30;60", 200, 3), [30, 60, 110]);
		assert.deepEqual(columnLayout(3, "30", 100, 3), [30, 72, 72]);
		assert.deepEqual(columnLayout(-1, "", 150, 2), [75, 75]);
		// TextColumn: -1 the first column wider than 0, 0 the row number, n the n-th.
		assert.equal(textColumnIndex(-1, [0, 72]), 1);
		assert.equal(textColumnIndex(undefined, [72, 72]), 0);
		assert.equal(textColumnIndex(0, [72]), -1);
		assert.equal(textColumnIndex(2, [72, 72]), 1);
		assert.deepEqual(itemColumns("a"), ["a"]);
	});

	it("edits a list of several columns", () => {
		let text = form([{ id: "lb", name: "List1", type: "ListBox", left: 8, top: 8, width: 120, height: 60 }]);
		const set = (field: string, value: unknown): SemanticOperation => ({ type: "setControlProperty", controlId: "lb", field, value });
		text = apply(text, set("list", [["1", "a"], "2"]), set("columnCount", 2), set("columnWidths", "36 pt;"), set("boundColumn", 2), set("textColumn", 1), set("columnHeads", true));
		const lb = control(text, "lb");
		assert.deepEqual([lb.list, lb.columnCount, lb.columnWidths, lb.boundColumn, lb.columnHeads], [[["1", "a"], "2"], 2, "36 pt;", 2, true]);
		rejects(text, [set("list", [["1", 2]])], "UFE003");
		rejects(text, [set("columnWidths", "wide")], "UFV021");
		rejects(text, [set("list", [["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11"]])], "UFV021");
		rejects(text, [set("columnCount", -2)]);
		const edit = (operation: unknown) => isWebviewMessage({ type: "edit", requestId: 1, version: 1, operations: [operation] });
		assert.equal(edit({ type: "setControlProperty", controlId: "lb", field: "list", value: [["a", "b"], "c"] }), true);
		assert.equal(edit({ type: "setControlProperty", controlId: "lb", field: "list", value: [["a", 1]] }), false);
	});

	it("accepts list and picture edits from the webview", () => {
		const edit = (operation: unknown) => isWebviewMessage({ type: "edit", requestId: 1, version: 1, operations: [operation] });
		assert.equal(edit({ type: "setControlProperty", controlId: "cb", field: "list", value: ["a", "b"] }), true);
		assert.equal(edit({ type: "setControlProperty", controlId: "cb", field: "list", value: null }), true);
		assert.equal(edit({ type: "setControlProperty", controlId: "cb", field: "list", value: ["a", 1] }), false);
		assert.equal(edit({ type: "setControlProperty", controlId: "cb", field: "caption", value: ["a"] }), false);
		assert.equal(edit({ type: "setFormProperty", field: "list", value: ["a"] }), false);
		assert.equal(edit({ type: "setControlProperty", controlId: "img", field: "picture", value: { remove: true } }), true);
		assert.equal(isWebviewMessage({ type: "picturePick", requestId: 1, version: 1, controlId: "img", field: "picture" }), true);
		assert.equal(isWebviewMessage({ type: "picturePick", requestId: 1, version: 1, controlId: "", field: "mouseIcon" }), true);
		assert.equal(isWebviewMessage({ type: "picturePick", requestId: 1, version: 1, controlId: "img", field: "caption" }), false);
		assert.equal(isWebviewMessage({ type: "pictureCommit", requestId: 1, version: 1, controlId: "img", field: "picture", token: "t" }), true);
		assert.equal(isWebviewMessage({ type: "pictureCommit", requestId: 1, version: 1, controlId: "img", field: "picture" }), false);
	});

	it("sets tag and controlTipText on any control", () => {
		let text = form([label, frame]);
		text = apply(text, { type: "setControlProperty", controlId: "l1", field: "tag", value: "note" });
		text = apply(text, { type: "setControlProperty", controlId: "f1", field: "controlTipText", value: "ヒント" });
		assert.equal(control(text, "l1").tag, "note");
		assert.equal(control(text, "f1").controlTipText, "ヒント");
		rejects(text, [{ type: "setControlProperty", controlId: "l1", field: "tag", value: 1 }], "UFE003");
		rejects(text, [{ type: "setControlProperty", controlId: "l1", field: "accelerator", value: "A" }], "UFE005");
	});

	it("aligns, distributes, sizes and reorders sibling controls", () => {
		const a = { id: "a", name: "A", type: "Label", left: 10, top: 10, width: 40, height: 18 };
		const b = { id: "b", name: "B", type: "Label", left: 60, top: 30, width: 20, height: 18 };
		const c = { id: "c", name: "C", type: "Label", left: 150, top: 50, width: 30, height: 30 };
		let text = form([a, b, c]);
		text = apply(text, { type: "alignControls", controlIds: ["a", "b", "c"], anchorId: "a", alignment: "top" });
		assert.deepEqual(["a", "b", "c"].map((id) => control(text, id).top), [10, 10, 10]);
		text = apply(text, { type: "distributeControls", controlIds: ["a", "b", "c"], axis: "horizontal" });
		assert.equal(control(text, "b").left, 90);
		text = apply(text, { type: "sameSizeControls", controlIds: ["a", "c"], anchorId: "c", dimension: "size" });
		assert.deepEqual([control(text, "a").width, control(text, "a").height], [30, 30]);
		text = apply(text, { type: "reorderControls", controlIds: ["a"], command: "front" });
		assert.deepEqual(["b", "c", "a"].map((id) => control(text, id).zIndex), [0, 1, 2]);
		text = apply(text, { type: "moveControls", moves: [{ controlId: "a", left: 0, top: 0 }, { controlId: "b", left: 80, top: 0 }] });
		assert.deepEqual([control(text, "a").left, control(text, "b").left], [0, 80]);
		text = apply(text, { type: "setTabOrder", parentId: "", controlIds: ["c", "a", "b"] });
		assert.deepEqual(["c", "a", "b"].map((id) => control(text, id).tabIndex), [0, 1, 2]);
		rejects(text, [{ type: "alignControls", controlIds: ["a", "c"], anchorId: "a", alignment: "right" }, { type: "setTabOrder", parentId: "", controlIds: ["c"] }]);
	});

	it("copies a subtree and pastes it with fresh names, IDs and an offset", () => {
		const child = { id: "x", parentId: "f1", name: "Label1", type: "Label", left: 0, top: 0, width: 20, height: 18 };
		const text = form([frame, child]);
		const { snapshot } = copyControls(text, ["f1"]);
		assert.deepEqual(snapshot.controlIds, ["f1"]);
		const parsed = parseSpec(text);
		assert.ok(parsed.ok);
		const target = designerDocument(previewDocument(parsed.spec));
		const copied = JSON.parse(snapshot.text).controls.map((c: Json) => ({ id: c.id, type: c.type }));
		const allocations = pasteAllocations({ controls: copied }, target);
		assert.deepEqual(allocations.map((a) => a.name), ["Frame2", "Label2"]);
		const after = applyOperations(text, [{ type: "pasteControls", snapshot, allocations, parentId: "", offset: { left: 8, top: 8 } }]);
		const pasted = json(after.text).controls.slice(2);
		assert.equal(pasted[0].left, 16);
		assert.equal(pasted[1].parentId, pasted[0].id);
		assert.deepEqual(after.selectionIds, [pasted[0].id]);
	});
});

describe("FormSpec validation and metadata", () => {
	const codes = (controls: Json[], extra: Json = {}) => {
		const parsed = parseSpec(formatJson({ ...newFormSpec("Main"), ...extra, controls }));
		assert.ok(parsed.ok);
		return parsed.spec.issues.map((i) => i.code).sort();
	};

	it("accepts new forms and only the RAS document kind", () => {
		assert.deepEqual(codes([]), []);
		assert.deepEqual(codes([], { kind: "xlflow.userform" }), ["UFV003"]);
		assert.deepEqual(codes([], { kind: "other" }), ["UFV003"]);
	});

	it("reports the structural rules", () => {
		assert.deepEqual(codes([label, { ...label, name: "Other" }]), ["UFV007"]);
		assert.deepEqual(codes([{ ...label, parentId: "missing" }]), ["UFV008"]);
		assert.deepEqual(codes([{ id: "p", name: "P", type: "Page" }]), ["UFV011"]);
		assert.deepEqual(codes([{ ...label, parentId: "l1" }]), ["UFV009"]);
		assert.deepEqual(
			codes([
				{ ...frame, id: "a", name: "A", parentId: "b" },
				{ ...frame, id: "b", name: "B", parentId: "a" },
			]),
			["UFV010"],
		);
		assert.deepEqual(codes([{ ...label, bogus: 1 }]), ["UFV001"]);
		assert.deepEqual(codes([{ ...label, tabs: [] }]), ["UFV005"]);
		assert.deepEqual(codes([{ id: "s", name: "S", type: "SpinButton", font: { size: 9 } }]), ["UFV005"]);
		assert.deepEqual(codes([{ ...label, font: { size: 0 } }]), ["UFV021"]);
		assert.deepEqual(codes([{ ...label, progId: "Forms.TextBox.1" }]), ["UFV012"]);
		assert.deepEqual(codes([{ id: "x", name: "X", type: "Gizmo" }]), ["UFV006"]);
		assert.deepEqual(codes([{ id: "x", name: "X", type: "Gizmo", progId: "Acme.Gizmo.1" }]), ["UFV014"]);
	});

	it("accepts list and selectedIndex on ComboBox / ListBox and checks the range", () => {
		const list = (type: string, fields: Json) => codes([{ id: "c", name: "C", type, ...fields }]);
		for (const type of ["ComboBox", "ListBox"]) {
			assert.deepEqual(list(type, { list: ["a"] }), []);
			assert.deepEqual(list(type, { list: ["a", "b"], selectedIndex: 1 }), []);
			assert.deepEqual(list(type, { list: ["a", "b"], selectedIndex: -1 }), []);
			assert.deepEqual(list(type, { list: ["a", "b"], selectedIndex: 2 }), ["UFV020"]);
			assert.deepEqual(list(type, { list: ["a", "b"], selectedIndex: -2 }), ["UFV020"]);
			// With an empty list only -1 is valid.
			assert.deepEqual(list(type, { list: [], selectedIndex: -1 }), []);
			assert.deepEqual(list(type, { list: [], selectedIndex: 0 }), ["UFV020"]);
			// A missing list counts as empty.
			assert.deepEqual(list(type, { selectedIndex: -1 }), []);
			assert.deepEqual(list(type, { selectedIndex: 0 }), ["UFV020"]);
			assert.deepEqual(list(type, { selectedIndex: 5 }), ["UFV020"]);
			assert.deepEqual(list(type, { selectedIndex: -2 }), ["UFV020"]);
			assert.deepEqual(list(type, { list: null, selectedIndex: 0 }), ["UFV020"]);
			assert.deepEqual(list(type, { selectedIndex: 1.5 }), ["UFV002"]);
			assert.deepEqual(list(type, { list: ["a", 1] }), ["UFV002"]);
		}
		assert.deepEqual(list("TextBox", { list: ["a"] }), ["UFV005"]);
		const parsed = parseSpec(form([{ id: "c", name: "C", type: "ListBox", list: ["a"], selectedIndex: 3 }]));
		assert.ok(parsed.ok);
		assert.equal(parsed.spec.issues[0].field, "controls[0].selectedIndex");
		// A list that is not an array of strings does not reach the canvas.
		const broken = parseSpec(form([{ id: "c", name: "C", type: "ListBox", list: "a" }]));
		assert.ok(broken.ok);
		assert.equal(previewDocument(broken.spec).controls[0].list, undefined);
	});

	it("accepts tag and controlTipText on every control type", () => {
		const extra = { tag: "t", controlTipText: "tip" };
		assert.deepEqual(codes([{ ...label, ...extra }, { ...frame, ...extra }, { id: "img", name: "I", type: "Image", ...extra }]), []);
		assert.deepEqual(codes([{ id: "mp", name: "M", type: "MultiPage", ...extra }, { id: "pg", parentId: "mp", name: "P", type: "Page", ...extra }]), []);
		assert.deepEqual(codes([{ ...label, tag: 1 }]), ["UFV002"]);
		assert.deepEqual(codes([{ ...label, accelerator: "A" }]), ["UFV005"]);
	});

	it("validates the picture of an Image", () => {
		const image = (picture: unknown) => codes([{ id: "img", name: "I", type: "Image", picture }]);
		assert.deepEqual(image({ path: "forms/assets/logo.bmp" }), []);
		assert.deepEqual(image({ remove: true }), []);
		assert.deepEqual(image({ path: "" }), ["UFV005"]);
		assert.deepEqual(image({ path: "a.bmp", remove: true }), ["UFV005"]);
		assert.deepEqual(codes([{ ...label, picture: { path: "a.bmp" } }]), []);
		assert.deepEqual(codes([{ id: "t", name: "T", type: "TextBox", picture: { path: "a.bmp" } }]), ["UFV005"]);
		assert.deepEqual(codes([{ ...label, mouseIcon: { path: "" } }]), ["UFV005"]);
	});

	it("locates problems in the source", () => {
		const text = form([{ ...label, bogus: 1 }]);
		const parsed = parseSpec(text);
		assert.ok(parsed.ok);
		const [problem] = problems(parsed.spec);
		assert.equal(problem.field, "controls[0].bogus");
		assert.equal(text.split("\n")[problem.range!.start.line].trim(), '"bogus": 1');
		assert.equal(parseSpec('{"form": ').ok, false);
		assert.equal(parseSpec('{"a": 1, "a": 2}').ok, false);
	});

	it("serves typed property metadata per control type", () => {
		const parsed = parseSpec(
			form([
				label,
				{ id: "img", name: "Image1", type: "Image", picture: { path: "logo.bmp" } },
				{ id: "mp", name: "M", type: "MultiPage", selectedIndex: 0 },
				{ id: "pg", parentId: "mp", name: "P", type: "Page" },
				{ id: "lb", name: "ListBox1", type: "ListBox", list: ["a", "b"], selectedIndex: 1 },
				{ id: "odd", name: "Odd", type: "Image", progId: "Forms.Label.1" },
			]),
		);
		assert.ok(parsed.ok);
		const grid = propertyGrid(parsed.spec);
		const fields = (id: string) => grid.controls[id].descriptors.map((d) => d.field);
		assert.ok(fields("l1").includes("font") && fields("l1").includes("textAlign"));
		assert.ok(!fields("img").includes("font") && fields("img").includes("picture") && fields("img").includes("pictureSizeMode"));
		assert.ok(fields("l1").includes("picture") && fields("l1").includes("mouseIcon") && !fields("l1").includes("list") && !fields("odd").includes("picture"));
		assert.ok(!fields("pg").includes("left") && fields("pg").includes("controlTipText"));
		// tag and controlTipText are common to every control; accelerator is not.
		for (const id of ["l1", "img", "mp", "pg", "lb"]) assert.ok(fields(id).includes("tag") && fields(id).includes("controlTipText"), id);
		assert.ok(!fields("l1").includes("accelerator") && fields("pg").includes("accelerator"));
		const descriptor = (id: string, field: string) => grid.controls[id].descriptors.find((d) => d.field === field);
		assert.deepEqual(descriptor("img", "picture")?.children?.map((d) => d.field), ["path", "remove"]);
		assert.deepEqual(grid.controls.img.values.picture, { present: true, value: { path: "logo.bmp" } });
		assert.deepEqual([descriptor("lb", "list")?.valueType, descriptor("lb", "list")?.editor], ["listItems", "list"]);
		assert.deepEqual(grid.controls.lb.values.list.value, ["a", "b"]);
		assert.equal(descriptor("lb", "selectedIndex")?.valueType, "integer");
		assert.deepEqual(grid.form.descriptors.map((d) => d.field), [
			"backColor", "borderColor", "borderStyle", "caption", "cycle", "drawBuffer", "enabled", "font", "foreColor", "height", "helpContextId",
			"keepScrollBarsVisible", "mouseIcon", "mousePointer", "name", "picture", "pictureAlignment", "pictureSizeMode", "pictureTiling", "rightToLeft", "scrollBars", "scrollHeight", "scrollLeft", "scrollTop", "scrollWidth",
			"showModal", "specialEffect", "startUpPosition", "tag", "whatsThisButton", "whatsThisHelp", "width", "zoom",
		]);
		const formField = (field: string) => grid.form.descriptors.find((d) => d.field === field);
		assert.deepEqual(formField("startUpPosition")?.suggestions?.map((s) => s.label), ["Manual", "CenterOwner", "CenterScreen", "WindowsDefault"]);
		assert.deepEqual(formField("mousePointer")?.allowedValues?.slice(0, 5), ["0", "1", "2", "3", "6"]);
		assert.equal(grid.controls.l1.values.caption.value, "Label1");
		assert.equal(grid.controls.l1.values.font.present, false);
	});
});

describe("event handlers", () => {
	it("resolves targets and default events from the catalogue", () => {
		const controls = [
			{ id: "mp", name: "Steps", type: "MultiPage" },
			{ id: "p", parentId: "mp", name: "P1", type: "Page" },
			{ id: "b", name: "OkButton", type: "CommandButton" },
		];
		assert.deepEqual(eventTarget(controls, undefined), { type: "UserForm", objectName: "UserForm" });
		assert.deepEqual(eventTarget(controls, "p"), { type: "MultiPage", objectName: "Steps" });
		const click = eventsForType("CommandButton", "OkButton").find((e) => e.isDefault);
		assert.equal(click?.signature, "Private Sub OkButton_Click()");
		assert.equal(eventsForType("UserForm", "UserForm").filter((e) => e.isDefault).length, 1);
	});

	it("finds existing handlers and appends new ones", () => {
		const code = '// form\n\n@on("OkButton.Click")\nfunction ok() {\n\tmsgBox(1);\n}\n';
		assert.deepEqual(findHandler(code, "okbutton", "click"), { line: 2, bodyLine: 4 });
		assert.equal(findHandler(code, "OkButton", "DblClick"), undefined);
		assert.equal(bodyIndent(code, "    "), "\t");
		const init = eventsForType("UserForm", "UserForm").find((e) => e.name === "Initialize")!;
		const appended = code + handlerText(code, "UserForm", init, "\t", "\n");
		assert.equal(appended.endsWith('\n}\n\n@on("UserForm.Initialize")\nfunction userFormInitialize() {\n\t\n}\n'), true, appended);
		assert.deepEqual(findHandler(appended, "UserForm", "Initialize"), { line: 7, bodyLine: 9 });
		const exit = eventsForType("TextBox", "txtName").find((e) => e.name === "Exit")!;
		assert.equal(handlerText("", "txtName", exit, "    ", "\r\n"), '@on("txtName.Exit")\r\nfunction txtNameExit(cancel: MSForms.ReturnBoolean) {\r\n    \r\n}\r\n');
		// A name that is taken gets a number.
		const taken = "function txtNameExit() {\n}\n";
		assert.equal(handlerText(taken, "txtName", exit, "    ", "\n").includes("function txtNameExit2("), true);
	});

	it("writes event parameters as RAS", () => {
		assert.equal(jsName("KeyCode"), "keyCode");
		assert.equal(jsName("RGB"), "rgb");
		assert.equal(rasParams("Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)"), "byref cancel: short, byref closeMode: short");
		assert.equal(rasParams("Private Sub B_Click()"), "");
	});
});
