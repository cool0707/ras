import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { applyOperations, copyControls } from "../../src/userForm/engine";
import { bodyIndent, eventsForType, eventTarget, findHandler, handlerText } from "../../src/userForm/events";
import { formatJson, minimalEdit } from "../../src/userForm/jsonSource";
import { listToText, textToList } from "../../src/userForm/listItems";
import { designerDocument } from "../../src/userForm/model";
import { isWebviewMessage } from "../../src/userForm/protocol";
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
		rejects(text, [{ type: "setControlProperty", controlId: "l1", field: "picture", value: { path: "a.bmp" } }], "UFE005");
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
	});

	it("accepts list and picture edits from the webview", () => {
		const edit = (operation: unknown) => isWebviewMessage({ type: "edit", requestId: 1, version: 1, operations: [operation] });
		assert.equal(edit({ type: "setControlProperty", controlId: "cb", field: "list", value: ["a", "b"] }), true);
		assert.equal(edit({ type: "setControlProperty", controlId: "cb", field: "list", value: null }), true);
		assert.equal(edit({ type: "setControlProperty", controlId: "cb", field: "list", value: ["a", 1] }), false);
		assert.equal(edit({ type: "setControlProperty", controlId: "cb", field: "caption", value: ["a"] }), false);
		assert.equal(edit({ type: "setFormProperty", field: "list", value: ["a"] }), false);
		assert.equal(edit({ type: "setControlProperty", controlId: "img", field: "picture", value: { remove: true } }), true);
		assert.equal(isWebviewMessage({ type: "picturePick", requestId: 1, version: 1, controlId: "img" }), true);
		assert.equal(isWebviewMessage({ type: "pictureCommit", requestId: 1, version: 1, controlId: "img", token: "t" }), true);
		assert.equal(isWebviewMessage({ type: "pictureCommit", requestId: 1, version: 1, controlId: "img" }), false);
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

	it("accepts new forms and both document kinds", () => {
		assert.deepEqual(codes([]), []);
		assert.deepEqual(codes([], { kind: "xlflow.userform" }), []);
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
		assert.deepEqual(codes([{ ...label, picture: { path: "a.bmp" } }]), ["UFV005"]);
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
		assert.ok(!fields("l1").includes("picture") && !fields("l1").includes("list") && !fields("odd").includes("picture"));
		assert.ok(!fields("pg").includes("left") && fields("pg").includes("controlTipText"));
		// tag and controlTipText are common to every control; accelerator is not.
		for (const id of ["l1", "img", "mp", "pg", "lb"]) assert.ok(fields(id).includes("tag") && fields(id).includes("controlTipText"), id);
		assert.ok(!fields("l1").includes("accelerator") && fields("pg").includes("accelerator"));
		const descriptor = (id: string, field: string) => grid.controls[id].descriptors.find((d) => d.field === field);
		assert.deepEqual(descriptor("img", "picture")?.children?.map((d) => d.field), ["path", "remove"]);
		assert.deepEqual(grid.controls.img.values.picture, { present: true, value: { path: "logo.bmp" } });
		assert.deepEqual([descriptor("lb", "list")?.valueType, descriptor("lb", "list")?.editor], ["stringArray", "list"]);
		assert.deepEqual(grid.controls.lb.values.list.value, ["a", "b"]);
		assert.equal(descriptor("lb", "selectedIndex")?.valueType, "integer");
		assert.deepEqual(grid.form.descriptors.map((d) => d.field), ["backColor", "borderStyle", "caption", "font", "foreColor", "height", "name", "width"]);
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
		const code = "Option Explicit\n\nPrivate Sub OkButton_Click()\n\tMsgBox 1\nEnd Sub\n";
		assert.deepEqual(findHandler(code, "okbutton", "click"), { line: 2, bodyLine: 3 });
		assert.equal(findHandler(code, "OkButton", "DblClick"), undefined);
		assert.equal(bodyIndent(code, "    "), "\t");
		const appended = code + handlerText(code, "Private Sub UserForm_Initialize()", "\t", "\n");
		assert.equal(appended.endsWith("\nEnd Sub\n\nPrivate Sub UserForm_Initialize()\n\t\nEnd Sub\n"), true);
		assert.equal(handlerText("", "Private Sub A_B()", "    ", "\r\n"), "Private Sub A_B()\r\n    \r\nEnd Sub\r\n");
	});
});
