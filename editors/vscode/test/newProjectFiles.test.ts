import assert from "node:assert/strict";
import { test } from "node:test";
import { isNewProjectFolder, projectBookName, setWorkbook, workbookFileNameError } from "../src/newProjectFiles";

test("an empty folder, or one with only hidden entries and workbooks, is where a project starts", () => {
	assert.equal(isNewProjectFolder([]), true);
	assert.equal(isNewProjectFolder([".git", ".vscode"]), true);
	assert.equal(isNewProjectFolder(["売上.xlsx", "Tool.XLSM", "~$Tool.xlsm", ".git"]), true);
	assert.equal(isNewProjectFolder(["README.md"]), false);
	assert.equal(isNewProjectFolder(["src", "Book.xlsm"]), false);
});

test("the new workbook is named after the folder", () => {
	assert.equal(projectBookName("売上集計"), "売上集計.xlsm");
	assert.equal(projectBookName(""), "Book.xlsm");
	assert.equal(projectBookName(".hidden"), "Book.xlsm");
});

test("only macro-enabled workbook names are accepted", () => {
	assert.equal(workbookFileNameError("Book.xlsm"), undefined);
	assert.equal(workbookFileNameError("Tool.xlam"), undefined);
	assert.notEqual(workbookFileNameError("Book.xlsx"), undefined);
	assert.notEqual(workbookFileNameError("sub/Book.xlsm"), undefined);
	assert.notEqual(workbookFileNameError(""), undefined);
});

test("the workbook line of ras.toml is set in place of the one ras init comments out", () => {
	const init = '[project]\nname = "x"\n# workbook = "Book.xlsm"    # `ras pack` / 同期の対象ブック\n\n[target]\nhost = "excel"\n';
	assert.equal(setWorkbook(init, "売上.xlsm"), '[project]\nname = "x"\nworkbook = "売上.xlsm"\n\n[target]\nhost = "excel"\n');
	assert.equal(setWorkbook('[project]\nname = "x"\n', "B.xlsm"), '[project]\nworkbook = "B.xlsm"\nname = "x"\n');
});
