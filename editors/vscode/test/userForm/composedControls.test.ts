import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { formatJson } from "../../src/userForm/jsonSource";
import { progressBarLayout } from "../../src/userForm/progressBar";
import { newFormSpec, parseSpec, previewDocument, propertyGrid } from "../../src/userForm/spec";
import { createControl } from "../../src/userForm/toolbox";
import { designerDocument } from "../../src/userForm/model";

type Json = Record<string, any>;
const parse = (controls: Json[]) => {
	const parsed = parseSpec(formatJson({ ...newFormSpec("Main"), controls }));
	assert.ok(parsed.ok);
	return parsed.spec;
};

describe("progress bar", () => {
	it("is laid out as ras-form writes it", () => {
		const layout = progressBarLayout({ width: 240, height: 24, value: 56 } as any);
		assert.deepEqual(layout.track, [0, 19.5, 240, 4.5]);
		assert.deepEqual(layout.bar, [0, 19.5, 134.25, 4.5]);
		assert.deepEqual(layout.caption, [0, 0, 204, 16.5]);
		assert.deepEqual(layout.value, [204, 0, 36, 16.5]);
		assert.equal(layout.text, "56%");
		const hidden = progressBarLayout({ width: 120, height: 24, value: 150, max: 200, showValue: false } as any);
		assert.equal(hidden.text, "75%");
		assert.deepEqual(hidden.caption, [0, 0, 120, 16.5]);
	});

	it("is checked and edited like a built-in control", () => {
		const bar = { id: "pb", name: "pbUpload", type: "ProgressBar", left: 12, top: 12, width: 240, height: 24, caption: "Upload progress", value: 56, barColor: { kind: "rgb", value: "#000000" } };
		const spec = parse([bar]);
		assert.deepEqual(spec.issues, []);
		const fields = propertyGrid(spec).controls.pb.descriptors.map((d) => d.field);
		for (const field of ["caption", "value", "min", "max", "showValue", "barHeight", "barColor", "trackColor", "valueColor", "foreColor", "font"]) assert.ok(fields.includes(field), field);
		assert.ok(!fields.includes("backColor") && !fields.includes("tabIndex"));
		const codes = (controls: Json[]) => parse(controls).issues.map((i) => `${i.code} ${i.field}`);
		assert.deepEqual(codes([{ ...bar, value: 101 }]), ["UFV024 controls[0].value"]);
		assert.deepEqual(codes([{ ...bar, min: 5, max: 5 }]), ["UFV024 controls[0].max"]);
		assert.deepEqual(codes([{ ...bar, trackColor: { kind: "rgb", value: "red" } }]), ["UFV024 controls[0].trackColor"]);
		assert.deepEqual(codes([bar, { id: "l", name: "pbUpload_Bar", type: "Label", left: 0, top: 60 }]), ["UFV024 controls[1].name"]);
	});

	it("is in the Toolbox", () => {
		const document = designerDocument(previewDocument(parse([])));
		const made = createControl(document as any, "ProgressBar", 12, 12, false, () => "1");
		assert.equal(made?.name, "ProgressBar1");
		assert.equal(made?.caption, "ProgressBar1");
		assert.deepEqual([made?.width, made?.height], [200, 24]);
	});
});

describe("file picker", () => {
	const picker = { id: "fp", name: "fpInput", type: "FilePicker", left: 12, top: 12, width: 240, height: 18, title: "ブックを選択", filters: [{ description: "Excel", extensions: "*.xlsx" }], filterIndex: 1 };

	it("is checked and edited like a built-in control", () => {
		const spec = parse([picker]);
		assert.deepEqual(spec.issues, []);
		const fields = propertyGrid(spec).controls.fp.descriptors.map((d) => d.field);
		for (const field of ["text", "dialogType", "title", "buttonName", "allowMultiSelect", "initialFileName", "filterIndex", "buttonCaption", "buttonWidth", "locked", "foreColor", "backColor", "font", "tabIndex"]) assert.ok(fields.includes(field), field);
		const codes = (controls: Json[]) => parse(controls).issues.map((i) => `${i.code} ${i.field}`);
		assert.deepEqual(codes([{ ...picker, dialogType: 4 }]), ["UFV025 controls[0].filters"]);
		assert.deepEqual(codes([{ ...picker, filterIndex: 2 }]), ["UFV025 controls[0].filterIndex"]);
		assert.deepEqual(codes([{ ...picker, filters: [{ description: "x" }] }]), ["UFV025 controls[0].filters"]);
		assert.deepEqual(codes([picker, { id: "b", name: "fpInput_Button", type: "CommandButton", left: 0, top: 60 }]), ["UFV025 controls[1].name"]);
		assert.ok(codes([{ ...picker, height: 8 }]).includes("UFV022 controls[0].height"));
	});

	it("is in the Toolbox", () => {
		const document = designerDocument(previewDocument(parse([])));
		const made = createControl(document as any, "FilePicker", 12, 12, false, () => "1");
		assert.equal(made?.name, "FilePicker1");
		assert.deepEqual([made?.width, made?.height], [200, 18]);
	});
});
