// Grammar test: tokenizes test/fixtures/sample.ras with the same engine VS Code
// uses (vscode-textmate + vscode-oniguruma) and asserts the scopes of the RAS
// additions to the TypeScript grammar.   Run with: npm test

import assert from "node:assert/strict";
import fs from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import oniguruma from "vscode-oniguruma";
import textmate from "vscode-textmate";
import { buildGrammar, OUTPUT_PATH } from "../scripts/build-grammar.mjs";

const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const manifest = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));
const grammarContribution = manifest.contributes.grammars[0];
const grammarPath = path.join(root, grammarContribution.path);
const samplePath = path.join(root, "test", "fixtures", "sample.ras");

async function tokenizeSample() {
	const wasm = fs.readFileSync(require.resolve("vscode-oniguruma/release/onig.wasm"));
	await oniguruma.loadWASM(wasm.buffer.slice(wasm.byteOffset, wasm.byteOffset + wasm.byteLength));
	const registry = new textmate.Registry({
		onigLib: Promise.resolve({
			createOnigScanner: (patterns) => new oniguruma.OnigScanner(patterns),
			createOnigString: (text) => new oniguruma.OnigString(text),
		}),
		loadGrammar: async (scopeName) =>
			scopeName === "source.ras" ? textmate.parseRawGrammar(fs.readFileSync(grammarPath, "utf8"), grammarPath) : null,
	});
	const grammar = await registry.loadGrammar("source.ras");
	assert.ok(grammar, "grammar source.ras failed to load");

	const lines = [];
	let ruleStack = textmate.INITIAL;
	for (const text of fs.readFileSync(samplePath, "utf8").split(/\r?\n/)) {
		const result = grammar.tokenizeLine(text, ruleStack);
		ruleStack = result.ruleStack;
		lines.push({
			text,
			tokens: result.tokens.map((token) => ({
				start: token.startIndex,
				end: token.endIndex,
				text: text.slice(token.startIndex, token.endIndex),
				scopes: token.scopes,
			})),
		});
	}
	return lines;
}

const lines = await tokenizeSample();

/** The single sample line that equals `snippet` (ignoring indentation) or, failing that, contains it. */
function line(snippet) {
	const exact = lines.filter((candidate) => candidate.text.trim() === snippet.trim());
	if (exact.length === 1) return exact[0];
	const found = lines.filter((candidate) => candidate.text.includes(snippet));
	assert.equal(found.length, 1, `expected exactly one sample line containing ${JSON.stringify(snippet)}, found ${found.length}`);
	return found[0];
}

/** Scopes of the `occurrence`-th token whose text is exactly `tokenText` on the line containing `snippet`. */
function scopesOf(snippet, tokenText, occurrence = 0) {
	const target = line(snippet);
	const token = target.tokens.filter((candidate) => candidate.text === tokenText)[occurrence];
	assert.ok(token, `no token ${JSON.stringify(tokenText)} (#${occurrence}) on line ${JSON.stringify(target.text)}; tokens: ${JSON.stringify(target.tokens.map((t) => t.text))}`);
	return token.scopes;
}

function expectScopes(snippet, tokenText, expected, { occurrence = 0, not = [] } = {}) {
	const scopes = scopesOf(snippet, tokenText, occurrence);
	for (const scope of [expected].flat()) {
		assert.ok(scopes.includes(scope), `${JSON.stringify(tokenText)} in ${JSON.stringify(snippet)}: expected scope ${scope}, got ${scopes.join(" ")}`);
	}
	for (const scope of not) {
		assert.ok(!scopes.some((actual) => actual.startsWith(scope)), `${JSON.stringify(tokenText)} in ${JSON.stringify(snippet)}: unexpected scope ${scope} in ${scopes.join(" ")}`);
	}
}

/**
 * Asserts that `literal` on the line containing `snippet` is covered only by
 * numeric-constant tokens (i.e. is not split into a number and an identifier)
 * and that its type suffix, if any, is scoped as the suffix.
 */
function expectNumber(snippet, literal, { kind = "decimal", suffix } = {}) {
	const target = line(snippet);
	const start = target.text.indexOf(literal);
	assert.ok(start >= 0, `${JSON.stringify(literal)} not found on line ${JSON.stringify(target.text)}`);
	const end = start + literal.length;
	const covering = target.tokens.filter((token) => token.start < end && token.end > start);
	assert.equal(covering[0].start, start, `${literal}: a token straddles the start of the literal`);
	assert.equal(covering.at(-1).end, end, `${literal}: a token straddles the end of the literal`);
	for (const token of covering) {
		assert.ok(token.scopes.includes(`constant.numeric.${kind}.ras`), `${literal}: token ${JSON.stringify(token.text)} is not constant.numeric.${kind}.ras (${token.scopes.join(" ")})`);
	}
	const suffixTokens = covering.filter((token) => token.scopes.includes("storage.type.numeric.ras"));
	assert.deepEqual(suffixTokens.map((token) => token.text), suffix ? [suffix] : [], `${literal}: type suffix tokens`);
}

test("generated grammar is up to date with scripts/build-grammar.mjs", () => {
	assert.equal(grammarPath, OUTPUT_PATH);
	assert.equal(fs.readFileSync(grammarPath, "utf8"), buildGrammar(), "run `npm run build:grammar`");
});

test("manifest wires the grammar and language configuration", () => {
	const grammar = JSON.parse(fs.readFileSync(grammarPath, "utf8"));
	assert.equal(grammar.scopeName, "source.ras");
	assert.equal(grammarContribution.scopeName, "source.ras");
	assert.equal(grammarContribution.language, "ras");
	const language = manifest.contributes.languages.find((candidate) => candidate.id === "ras");
	assert.deepEqual(language.extensions, [".ras"]);
	// The language configuration must be strict JSON so every tool can read it.
	JSON.parse(fs.readFileSync(path.join(root, language.configuration), "utf8"));
});

test("every scope is renamed from .ts to .ras", () => {
	for (const { text, tokens } of lines) {
		for (const token of tokens) {
			assert.equal(token.scopes[0], "source.ras");
			for (const scope of token.scopes) {
				assert.ok(!/\.tsx?$/.test(scope), `scope ${scope} on line ${JSON.stringify(text)}`);
			}
		}
	}
});

test("byref / byval parameter modifiers", () => {
	expectScopes("function calculate(", "byref", "storage.modifier.ras");
	expectScopes("function calculate(", "outResult", "variable.parameter.ras");
	expectScopes("function calculate(", "function", "storage.type.function.ras");
	expectScopes("function calculate(", "calculate", "entity.name.function.ras");
	expectScopes("function inc(", "byref", "storage.modifier.ras");
	expectScopes("function inc(", "this", "variable.language.this.ras");
	expectScopes("function scale(", "byval", "storage.modifier.ras");
	expectScopes("function scale(", "factor", "variable.parameter.ras");
	expectScopes("function scale(", "byref", "storage.modifier.ras");
	expectScopes("function scale(", "target", "variable.parameter.ras");
	// Arrow function parameters.
	expectScopes("const adder =", "byref", "storage.modifier.ras");
	expectScopes("const adder =", "acc", "variable.parameter.ras");
});

test("inline / noinline function modifiers", () => {
	expectScopes("inline function clamp(", "inline", ["storage.modifier.ras", "meta.function.ras"]);
	expectScopes("inline function clamp(", "function", "storage.type.function.ras");
	expectScopes("inline function clamp(", "clamp", "entity.name.function.ras");
	expectScopes("export inline function half(", "export", "keyword.control.export.ras");
	expectScopes("export inline function half(", "inline", "storage.modifier.ras");
	expectScopes("export inline function half(", "function", "storage.type.function.ras");
	expectScopes("export inline function half(", "half", "entity.name.function.ras");
	expectScopes("noinline function slowPath(", "noinline", "storage.modifier.ras");
	expectScopes("noinline function slowPath(", "slowPath", "entity.name.function.ras");
});

test("resume statement keyword", () => {
	expectScopes("        resume;", "resume", "keyword.control.ras");
	expectScopes("} catch {", "catch", "keyword.control.trycatch.ras");
	// An identifier that merely starts with `resume` is not the keyword.
	expectScopes("resumeScreen();", "resumeScreen", "entity.name.function.ras", { not: ["keyword.control"] });
});

test("using declarations", () => {
	expectScopes("using _ = fastMode();", "using", "storage.type.ras");
	expectScopes("using _ = fastMode();", "_", "variable.other.constant.ras");
	expectScopes("using _ = fastMode();", "fastMode", "entity.name.function.ras");
	expectScopes("using guard = fastMode(", "using", "storage.type.ras");
	expectScopes("using guard = fastMode(", "guard", "variable.other.constant.ras");
});

test("primitive types", () => {
	const primitives = ["byte", "short", "int", "long", "float", "double", "currency", "decimal", "ptr", "date", "string", "boolean", "any", "object", "wstr", "astr"];
	for (const type of primitives) {
		const name = `t${type[0].toUpperCase()}${type.slice(1)}`;
		expectScopes(`let ${name}: ${type};`, type, "support.type.primitive.ras");
	}
	expectScopes("noinline function slowPath(", "void", "support.type.primitive.ras");
	expectScopes("function calculate(", "double", "support.type.primitive.ras");
	expectScopes("function calculate(", "boolean", "support.type.primitive.ras");
	expectScopes("const narrowed = total as int;", "int", "support.type.primitive.ras");
	// `number` is not a RAS type: it must not look like a valid primitive.
	expectScopes("let tNumber: number;", "number", "entity.name.type.ras", { not: ["support.type"] });
	// Host types are ordinary (qualified) type names.
	expectScopes("let tRange: Excel.Range;", "Excel", "entity.name.type.module.ras");
	expectScopes("let tRange: Excel.Range;", "Range", "entity.name.type.ras");
});

test("numeric literals and type suffixes", () => {
	expectNumber("const nInt = 123;", "123");
	expectNumber("const nLong = 123L;", "123L", { suffix: "L" });
	expectNumber("const nBig = 10L;", "10L", { suffix: "L" });
	expectNumber("const nDouble = 1.5;", "1.5");
	expectNumber("const nExp = 1e3;", "1e3");
	expectNumber("const nDoubleSuffix = 2d;", "2d", { suffix: "d" });
	expectNumber("const nFloat = 1.5f;", "1.5f", { suffix: "f" });
	expectNumber("const nDecimal = 1.5m;", "1.5m", { suffix: "m" });
	expectNumber("const nHex = 0xFF;", "0xFF", { kind: "hex" });
	expectNumber("const nHexLong = 0xFFL;", "0xFFL", { kind: "hex", suffix: "L" });
	expectNumber("const nColumn = 0cAA;", "0cAA", { kind: "column" });
	expectNumber("const nColumnL = 0cAL;", "0cAL", { kind: "column" });
	expectNumber("const nSum =", "300");
	expectNumber("const nSum =", "200");
	expectNumber("const nSum =", "0.25f", { suffix: "f" });
	// No identifier token may be left over after a literal.
	for (const snippet of ["const nLong = 123L;", "const nFloat = 1.5f;", "const nDoubleSuffix = 2d;", "const nDecimal = 1.5m;"]) {
		const stray = line(snippet).tokens.filter((token) => /^[A-Za-z]$/.test(token.text) && token.scopes.some((scope) => scope.startsWith("variable.")));
		assert.deepEqual(stray, [], `${snippet}: suffix tokenized as an identifier`);
	}
});

test("decorators on functions", () => {
	expectScopes("@fast", "@", ["punctuation.decorator.ras", "meta.decorator.ras"]);
	expectScopes("@fast", "fast", ["entity.name.function.decorator.ras", "meta.decorator.ras"]);
	expectScopes("export function rebuildReport()", "export", "keyword.control.export.ras", { not: ["meta.decorator"] });
	expectScopes("export function rebuildReport()", "function", "storage.type.function.ras");
	expectScopes("export function rebuildReport()", "rebuildReport", "entity.name.function.ras");

	expectScopes('@on("Workbook.Open")', "@", "punctuation.decorator.ras");
	expectScopes('@on("Workbook.Open")', "on", "entity.name.function.decorator.ras");
	expectScopes('@on("Workbook.Open")', "Workbook.Open", ["string.quoted.double.ras", "meta.decorator.ras"]);
	expectScopes("export function setup()", "function", "storage.type.function.ras", { not: ["meta.decorator"] });
	expectScopes("export function setup()", "setup", "entity.name.function.ras");

	// Arguments containing whitespace do not end the decorator early.
	expectScopes('@on("Sheet1.Change", { reentrant: true })', "on", "entity.name.function.decorator.ras");
	expectScopes('@on("Sheet1.Change", { reentrant: true })', "true", ["constant.language.boolean.true.ras", "meta.decorator.ras"]);
	expectScopes("function onSheet1Change(", "function", "storage.type.function.ras", { not: ["meta.decorator"] });
	expectScopes("function onSheet1Change(", "onSheet1Change", "entity.name.function.ras");

	expectScopes("@macro", "macro", "entity.name.function.decorator.ras");
	expectScopes("function runReport()", "function", "storage.type.function.ras", { not: ["meta.decorator"] });

	// Decorator and declaration on one line.
	expectScopes("@fast export function sameLine()", "fast", "entity.name.function.decorator.ras");
	expectScopes("@fast export function sameLine()", "export", "keyword.control.export.ras", { not: ["meta.decorator"] });
	expectScopes("@fast export function sameLine()", "function", "storage.type.function.ras");
	expectScopes("@fast export function sameLine()", "sameLine", "entity.name.function.ras");
});

test("decorators on loops and statements", () => {
	expectScopes("    @bulk", "@", "punctuation.decorator.ras");
	expectScopes("    @bulk", "bulk", "entity.name.function.decorator.ras");
	expectScopes("for (let r = 2; r <= last; r++) {", "for", "keyword.control.loop.ras", { not: ["meta.decorator"] });
	expectScopes("for (let r = 2; r <= last; r++) {", "let", "storage.type.ras");

	expectScopes("    @nodoevents", "nodoevents", "entity.name.function.decorator.ras");
	expectScopes("while (pending()) {", "while", "keyword.control.loop.ras", { not: ["meta.decorator"] });

	expectScopes("@nobulk for (const v of values)", "nobulk", "entity.name.function.decorator.ras");
	expectScopes("@nobulk for (const v of values)", "for", "keyword.control.loop.ras", { not: ["meta.decorator"] });
	expectScopes("@nobulk for (const v of values)", "of", "keyword.operator.expression.of.ras");
});

test("declare module", () => {
	expectScopes('declare module "user32" {', "declare", "storage.modifier.ras");
	expectScopes('declare module "user32" {', "module", "storage.type.namespace.ras");
	expectScopes('declare module "user32" {', "user32", "string.quoted.double.ras");
	expectScopes("function GetForegroundWindow(): ptr;", "function", "storage.type.function.ras");
	expectScopes("function GetForegroundWindow(): ptr;", "GetForegroundWindow", "entity.name.function.ras");
	expectScopes("function GetForegroundWindow(): ptr;", "ptr", "support.type.primitive.ras");
	expectScopes('@alias("GetWindowTextW")', "alias", "entity.name.function.decorator.ras");
	expectScopes("function GetWindowText(", "function", "storage.type.function.ras", { not: ["meta.decorator"] });
	expectScopes("function GetWindowText(", "lpString", "variable.parameter.ras");
	expectScopes("function GetWindowText(", "wstr", "support.type.primitive.ras");
	expectScopes("function GetTickCount64(): long;", "long", "support.type.primitive.ras");
	expectScopes("function GetTickCount64(): long;", " 64bit only", "comment.line.double-slash.ras");
});

test("declare struct", () => {
	expectScopes("declare struct RECT {", "declare", ["storage.modifier.ras", "meta.struct.ras"]);
	expectScopes("declare struct RECT {", "struct", "storage.type.struct.ras");
	expectScopes("declare struct RECT {", "RECT", "entity.name.type.struct.ras");
	expectScopes("    left: int;", "left", ["variable.object.property.ras", "meta.struct.ras"]);
	expectScopes("    left: int;", "int", "support.type.primitive.ras");
	expectScopes("    bottom: int;", "bottom", "variable.object.property.ras");
	// The struct body ends at its closing brace.
	expectScopes("export function maxOf<T extends Comparable>", "export", "keyword.control.export.ras", { not: ["meta.struct"] });
});

test("TypeScript constructs RAS reuses still tokenize", () => {
	expectScopes("export function maxOf<T extends Comparable>", "T", "entity.name.type.ras");
	expectScopes("export function maxOf<T extends Comparable>", "extends", "storage.modifier.ras");
	expectScopes("export function padLeft(", "this", "variable.language.this.ras");
	expectScopes("const stack = new Stack<Excel.Range>();", "new", "keyword.operator.new.ras");
	expectScopes(" * @param col", "param", ["storage.type.class.jsdoc", "comment.block.documentation.ras"]);
	expectScopes("status(`処理中", "nInt", ["meta.template.expression.ras", "string.template.ras"]);
	expectScopes("const doubled = arr.map(x => x * base);", "=>", "storage.type.function.arrow.ras");
	// JavaScript runtime names get no special treatment in RAS.
	expectScopes("return this.length >= width", "length", "variable.other.property.ras", { not: ["support."] });
});

test("Excel formulas in fx templates", () => {
	const sum = "const total = fx`SUM(";
	expectScopes(sum, "fx", "entity.name.function.tagged-template.ras");
	expectScopes(sum, "`", "punctuation.definition.string.template.begin.ras");
	// The formula is code (meta.embedded), not a string: VS Code suggests while typing.
	expectScopes(sum, "SUM", ["support.function.formula.ras", "meta.embedded.formula.ras"], { not: ["string.template"] });
	expectScopes(sum, "Sheet1", "entity.name.type.sheet.formula.ras");
	expectScopes(sum, "$A$1", "variable.other.constant.reference.formula.ras");
	expectScopes(sum, ":", "keyword.operator.range.formula.ras");
	expectScopes(sum, "B3", "variable.other.constant.reference.formula.ras");
	expectScopes(sum, "#N/A", "constant.language.error.formula.ras");
	expectScopes(sum, "&", "keyword.operator.concatenation.formula.ras");
	expectScopes(sum, '""', ["constant.character.escape.formula.ras", "string.quoted.double.formula.ras"]);

	const lookup = "const looked = ws.fx`VLOOKUP(";
	expectScopes(lookup, "ws", "variable.other.object.ras");
	expectScopes(lookup, "fx", "entity.name.function.tagged-template.ras");
	expectScopes(lookup, "VLOOKUP", "support.function.formula.ras");
	expectScopes(lookup, "key", ["meta.template.expression.ras", "variable.other.readwrite.ras"]);
	expectScopes(lookup, "'Price list'", "entity.name.type.sheet.formula.ras");
	expectScopes(lookup, "2", "constant.numeric.formula.ras");
	expectScopes(lookup, "FALSE", "constant.language.boolean.formula.ras");
	expectScopes(lookup, ";", "punctuation.terminator.statement.ras", { not: ["meta.embedded.formula"] });

	// Any other tag is an ordinary template.
	expectScopes("const plain = tag`SUM(A1)`;", "SUM(A1)", "string.template.ras", { not: ["meta.embedded.formula"] });
});
