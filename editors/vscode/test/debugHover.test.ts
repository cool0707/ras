import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { hoverExpression } from "../src/debugHover";

/** The hover at the `|` of `text`. */
function at(text: string) {
	const column = text.indexOf("|");
	return hoverExpression(text.slice(0, column) + text.slice(column + 1), column);
}

describe("debug hovers", () => {
	it("rewrites a member chain into a runtime read", () => {
		const hover = at("    const sheetName: string = sheet.na|me;");
		assert.deepEqual(hover, {
			start: 30,
			end: 40,
			expression: '$rasDebugRead([() => sheet$, () => sheet], "sheet", ["name"])',
		});
	});

	it("reads only up to the hovered name", () => {
		assert.equal(at("x = s|heet.name;")?.expression, '$rasDebugRead([() => sheet$, () => sheet], "sheet", [])');
		assert.equal(at("x = a.b|.c;")?.expression, '$rasDebugRead([() => a$, () => a], "a", ["b"])');
	});

	it("reads the receiver of an extension method", () => {
		assert.equal(at("return this.ro|ws.count;")?.expression, '$rasDebugRead([() => $this, () => this], "this", ["rows"])');
	});

	it("leaves calls and chains through calls alone", () => {
		assert.equal(at("ws.del|ete();"), undefined);
		assert.equal(at("msg|Box(x);"), undefined);
		assert.equal(at("ws.cells(1, 1).val|ue = 3;"), undefined);
		assert.equal(at("xs[0].na|me"), undefined);
		assert.equal(at('"abc".len|gth'), undefined);
	});

	it("ignores strings, comments and numbers", () => {
		assert.equal(at('msgBox("sheet.na|me");'), undefined);
		assert.equal(at("// sheet.na|me"), undefined);
		assert.equal(at("x = 1e|5;"), undefined);
	});
});
