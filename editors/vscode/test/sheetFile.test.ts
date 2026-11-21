import assert from "node:assert/strict";
import { test } from "node:test";
import { sheetOfRasFile } from "../src/sheetFile";

test("the file of a sheet is found by its name", () => {
	assert.equal(sheetOfRasFile("Sheet1Events.ras", "// nothing yet\n"), "Sheet1");
	assert.equal(sheetOfRasFile("shtSalesEvents2.ras", ""), "shtSales");
	assert.equal(sheetOfRasFile("ThisWorkbookEvents.ras", ""), undefined);
});

test("otherwise by its first sheet event handler", () => {
	assert.equal(sheetOfRasFile("report.ras", '@on("Workbook.Open")\nfunction a() {}\n@on(\'Sheet2.Change\')\nfunction b(t: Excel.Range) {}\n'), "Sheet2");
	assert.equal(sheetOfRasFile("report.ras", '@on("Workbook.SheetChange")\nfunction a() {}\n'), undefined);
	assert.equal(sheetOfRasFile("util.ras", "export function f() {}\n"), undefined);
});
