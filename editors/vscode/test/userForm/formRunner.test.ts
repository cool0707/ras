import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { applyUpdate, contentInsets, cursorOf, initialState, LineReader, oleToCss, px } from "../../src/formRunner/protocol";
import type { RuntimeMessage } from "../../src/formRunner/protocol";

// The panel of a form being debugged (spec §14.10.7): what the extension keeps of the state
// the debug process reports, and how the panel turns MSForms values into CSS.

const show: Extract<RuntimeMessage, { t: "show" }> = {
	t: "show",
	id: 1,
	name: "frmInput",
	modal: true,
	form: { Caption: "入力", Width: 240, Height: 180 },
	controls: [
		{ name: "txtName", type: "TextBox", parent: "", state: { Text: "" } },
		{ name: "fraSize", type: "Frame", parent: "", state: { Caption: "サイズ" } },
		{ name: "optS", type: "OptionButton", parent: "fraSize", state: { Value: true } },
	],
};

describe("form runner protocol", () => {
	it("keeps the controls in their order and replaces the updated ones", () => {
		const state = initialState(show);
		const next = applyUpdate(state, { t: "update", id: 1, controls: [{ name: "OPTS", type: "OptionButton", parent: "fraSize", state: { Value: false } }], focus: "txtName" });
		assert.deepEqual(
			next.controls.map((c) => c.name),
			["txtName", "fraSize", "OPTS"],
		);
		assert.equal(next.controls[2].state.Value, false);
		assert.equal(next.form, state.form, "the form is kept when the update has none");
		assert.equal(state.controls[2].state.Value, true, "the previous state is not changed");
		const captioned = applyUpdate(next, { t: "update", id: 1, form: { Caption: "入力 2" }, controls: [] });
		assert.equal(captioned.form.Caption, "入力 2");
	});

	it("adds the controls the form's code added and drops the removed ones", () => {
		const state = initialState(show);
		const added = applyUpdate(state, { t: "update", id: 1, controls: [{ name: "txtNameError", type: "Label", parent: "", state: { Caption: "必須です", BackStyle: 0 } }] });
		assert.deepEqual(
			added.controls.map((c) => c.name),
			["txtName", "fraSize", "optS", "txtNameError"],
		);
		const removed = applyUpdate(added, { t: "update", id: 1, controls: [], removed: ["TXTNAMEERROR"] });
		assert.deepEqual(
			removed.controls.map((c) => c.name),
			["txtName", "fraSize", "optS"],
		);
	});

	it("splits the stream into lines", () => {
		const reader = new LineReader();
		assert.deepEqual(reader.push('{"t":"hide"'), []);
		assert.deepEqual(reader.push(',"id":1}\n\n{"t":'), ['{"t":"hide","id":1}']);
		assert.deepEqual(reader.push('"unload","id":1}\r\n'), ['{"t":"unload","id":1}']);
	});

	it("turns OLE colors into CSS colors", () => {
		assert.equal(oleToCss(0x0000ff), "#ff0000", "OLE colors are BGR");
		assert.equal(oleToCss(0x203040), "#403020");
		assert.equal(oleToCss(0x8000000f), "#f0f0f0", "buttonFace");
		assert.equal(oleToCss(0x80000005), "#ffffff", "window");
		assert.equal(oleToCss(-2147483633), "#f0f0f0", "a negative Long is the same system color");
		assert.equal(oleToCss(0x1000000), undefined);
		assert.equal(oleToCss("red"), undefined);
	});

	it("measures in points", () => {
		assert.equal(px(72), "96px");
		assert.equal(px(undefined), "0px");
		assert.deepEqual(contentInsets("MultiPage"), { left: 1.5, top: 18, right: 1.5, bottom: 1.5 });
		assert.deepEqual(contentInsets("MultiPage", 1), { left: 1.5, top: 1.5, right: 1.5, bottom: 18 });
		assert.deepEqual(contentInsets("MultiPage", 2), { left: 18, top: 1.5, right: 1.5, bottom: 1.5 });
		assert.deepEqual(contentInsets("MultiPage", 3), { left: 1.5, top: 1.5, right: 18, bottom: 1.5 });
		assert.deepEqual(contentInsets("Label"), { left: 0, top: 0, right: 0, bottom: 0 });
		assert.deepEqual([0, 3, 11, 99, undefined].map((n) => cursorOf(n)), ["", "text", "wait", "", ""]);
		assert.equal(cursorOf(99, "data:x"), 'url("data:x"), auto');
	});
});
