import assert from "node:assert/strict";
import { test } from "node:test";
import { macrosOf, moduleName, newMacroName, sheetControlEvents } from "../src/sheetEvents";

test("a control on a sheet raises GotFocus and LostFocus, not what a form raises", () => {
	const names = sheetControlEvents("TextBox", "TextBox1").map((e) => e.name);
	assert.ok(names.includes("Change") && names.includes("KeyDown"));
	assert.ok(names.includes("GotFocus") && names.includes("LostFocus"));
	assert.ok(!names.includes("Enter") && !names.includes("AfterUpdate"));
	const click = sheetControlEvents("CommandButton", "CommandButton1").find((e) => e.isDefault);
	assert.equal(click?.name, "Click");
	assert.equal(click?.signature, "Private Sub CommandButton1_Click()");
	assert.deepEqual(sheetControlEvents("Unknown", "X"), []);
});

test("the module of a file of src/ is named as the compiler names it", () => {
	assert.equal(moduleName("/p/src/report.ras"), "Report");
	assert.equal(moduleName("/p/src/sub/mod_util.ras"), "ModUtil");
	assert.equal(moduleName("/p/src/2nd.ras"), "Mod2nd");
});

test("the macros of a file are its exported functions without parameters and @macro", () => {
	const text = "export function go() {\n}\nfunction hidden() {}\n@macro\nfunction marked() {\n}\nexport function withArg(a: int) {}\nexport function 実行() {}\n";
	assert.deepEqual(
		macrosOf(text).map((m) => [m.name, m.line]),
		[["go", 0], ["marked", 4], ["実行", 7]],
	);
});

test("a new macro is named after its object, unused in the file", () => {
	assert.equal(newMacroName("", "Chart 1", undefined), "chart1Click");
	assert.equal(newMacroName("", "ボタン 1", undefined), "ボタン1Click");
	assert.equal(newMacroName("", "", 1025), "shape1025Click");
	assert.equal(newMacroName("export function chart1Click() {}\n", "Chart 1", undefined), "chart1Click2");
	assert.equal(newMacroName("export function chart1ClickX() {}\n", "Chart 1", undefined), "chart1Click");
});
