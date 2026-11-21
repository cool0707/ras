import assert from "node:assert/strict";
import { test } from "node:test";
import { sheetOfRasFile } from "../src/sheetFile";

test("a file of objects/ is the module of the sheet it is named after", () => {
	assert.equal(sheetOfRasFile("/work/book/objects/Sheet1.ras"), "Sheet1");
	assert.equal(sheetOfRasFile("C:\\work\\book\\objects\\shtSales.ras"), "shtSales");
	assert.equal(sheetOfRasFile("objects/売上.ras"), "売上");
});

test("the workbook's file and the files of other folders are not", () => {
	assert.equal(sheetOfRasFile("/work/book/objects/ThisWorkbook.ras"), undefined);
	assert.equal(sheetOfRasFile("/work/book/src/Sheet1.ras"), undefined);
	assert.equal(sheetOfRasFile("/work/book/objects/sub/Sheet1.ras"), undefined);
	assert.equal(sheetOfRasFile("/work/book/objects/types.d.ras"), undefined);
	assert.equal(sheetOfRasFile("Sheet1.ras"), undefined);
});
