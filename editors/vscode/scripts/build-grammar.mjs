#!/usr/bin/env node
// Derives syntaxes/ras.tmLanguage.json from the vendored, unmodified upstream
// TypeScript TextMate grammar (third_party/TypeScript-TmLanguage/TypeScript.tmLanguage).
//
//   node scripts/build-grammar.mjs           regenerate syntaxes/ras.tmLanguage.json
//   node scripts/build-grammar.mjs --check   exit 1 if the generated file is out of date
//
// Every RAS-specific change to the grammar lives in this file. Each edit names the
// upstream text it expects and throws if upstream no longer contains it, so a
// refreshed upstream snapshot cannot silently drop a RAS modification.
// Keep the "what was modified" list in THIRD_PARTY_NOTICES.md in sync with this file.

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parse as parsePlist } from "plist";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

export const UPSTREAM_REPOSITORY = "https://github.com/microsoft/TypeScript-TmLanguage";
export const UPSTREAM_COMMIT = "48f608692aa6d6ad7bd65b478187906c798234a8";
export const UPSTREAM_PATH = path.join(root, "third_party", "TypeScript-TmLanguage", "TypeScript.tmLanguage");
export const OUTPUT_PATH = path.join(root, "syntaxes", "ras.tmLanguage.json");

// Building blocks the upstream grammar uses for identifier boundaries (its YAML
// source calls them startOfIdentifier / endOfIdentifier / identifier).
const START_OF_IDENTIFIER = "(?<![_$[:alnum:]])(?:(?<=\\.\\.\\.)|(?<!\\.))";
const END_OF_IDENTIFIER = "(?![_$[:alnum:]])(?:(?=\\.\\.\\.)|(?!\\.))";
const IDENTIFIER = "[_$[:alpha:]][_$[:alnum:]]*";

// §2.1 / §11.2 of the language specification. `number` is deliberately absent.
const PRIMITIVE_TYPES = [
	"byte", "short", "int", "long", "float", "double", "currency", "decimal",
	"ptr", "date", "string", "boolean", "any", "cellValue", "object", "void", "wstr", "astr",
];

/** Replaces `from` with `to` in `holder[key]`, requiring exactly `count` occurrences. */
function replaceIn(holder, key, from, to, count = 1) {
	const text = holder[key];
	if (typeof text !== "string") {
		throw new Error(`build-grammar: expected a string at "${key}"`);
	}
	const found = text.split(from).length - 1;
	if (found !== count) {
		throw new Error(`build-grammar: expected ${count} occurrence(s) of ${JSON.stringify(from)} in "${key}", found ${found}. Upstream changed; review this edit.`);
	}
	holder[key] = text.split(from).join(to);
}

function rule(grammar, name) {
	const found = grammar.repository[name];
	if (!found) {
		throw new Error(`build-grammar: upstream rule "${name}" not found`);
	}
	return found;
}

function walk(node, visit) {
	if (Array.isArray(node)) {
		for (const item of node) walk(item, visit);
	} else if (node !== null && typeof node === "object") {
		visit(node);
		for (const value of Object.values(node)) walk(value, visit);
	}
}

/** Deletes a repository rule together with every `{ include: "#name" }` that refers to it. */
function removeRule(grammar, name) {
	rule(grammar, name);
	delete grammar.repository[name];
	const include = `#${name}`;
	walk(grammar, (node) => {
		if (Array.isArray(node.patterns)) {
			node.patterns = node.patterns.filter((pattern) => pattern.include !== include);
		}
	});
}

// --- RAS modifications -------------------------------------------------------

/** `byref` / `byval` parameter modifiers (spec §3), wherever upstream accepts `readonly` on a parameter. */
function addParameterModifiers(grammar) {
	const withOverride = "(override|public|private|protected|readonly)\\s+)?";
	const withOverrideRas = "(override|public|private|protected|readonly|byref|byval)\\s+)?";
	const sequence = "(override|public|protected|private|readonly)";
	const sequenceRas = "(override|public|protected|private|readonly|byref|byval)";

	// Function / method parameters.
	const parameterName = rule(grammar, "parameter-name").patterns;
	replaceIn(parameterName[0], "match", sequence, sequenceRas, 2);
	replaceIn(parameterName[1], "match", withOverride, withOverrideRas);
	replaceIn(parameterName[2], "match", withOverride, withOverrideRas);

	// Arrow function parameters: `(byref x: int) => ...`.
	const arrowParameters = rule(grammar, "expression-inside-possibly-arrow-parens").patterns;
	replaceIn(arrowParameters[5], "match", sequence, sequenceRas, 2);
	replaceIn(arrowParameters[6], "match", withOverride, withOverrideRas);
	replaceIn(arrowParameters[7], "match", withOverride, withOverrideRas);

	// Function types: `(byref x: int) => void`.
	const typeParameters = rule(grammar, "type-paren-or-function-parameters").patterns;
	const inType = "(public|private|protected|readonly)\\s+)?";
	const inTypeRas = "(public|private|protected|readonly|byref|byval)\\s+)?";
	replaceIn(typeParameters[0], "match", inType, inTypeRas);
	replaceIn(typeParameters[1], "match", inType, inTypeRas);
}

/** `inline function` / `noinline function`, also after `export` (spec §8.1). */
function addInlineModifiers(grammar) {
	const declaration = rule(grammar, "function-declaration");
	replaceIn(declaration, "begin", "(?:(async)\\s+)?(function\\b)", "(?:(inline|noinline)\\s+)?(?:(async)\\s+)?(function\\b)");
	const upstream = declaration.beginCaptures;
	if (Object.keys(upstream).join() !== "1,2,3,4,5,6") {
		throw new Error("build-grammar: function-declaration captures changed upstream; review addInlineModifiers");
	}
	declaration.beginCaptures = {
		1: upstream[1], // export
		2: upstream[2], // declare
		3: { name: "storage.modifier.ras" }, // inline | noinline
		4: upstream[3], // async
		5: upstream[4], // function
		6: upstream[5], // *
		7: upstream[6], // name
	};
}

/** `resume;` inside `catch { }` (spec §5.2). */
function addResumeKeyword(grammar) {
	const patterns = rule(grammar, "control-statement").patterns;
	const tryCatch = patterns.findIndex((pattern) => pattern.name === "keyword.control.trycatch.ts");
	if (tryCatch < 0) {
		throw new Error("build-grammar: try/catch keyword pattern not found in control-statement");
	}
	patterns.splice(tryCatch + 1, 0, {
		name: "keyword.control.ras",
		match: `${START_OF_IDENTIFIER}(resume)${END_OF_IDENTIFIER}`,
	});
}

/** RAS primitive types (spec §2.1, §11.2). `number` falls through to a plain type name. */
function replacePrimitiveTypes(grammar) {
	replaceIn(
		rule(grammar, "type-primitive"), "match",
		"(string|number|bigint|boolean|symbol|any|void|never|unknown)",
		`(${PRIMITIVE_TYPES.join("|")})`,
	);
	// `object` is a primitive in RAS (VBA `Object`), not a built-in literal type.
	replaceIn(
		rule(grammar, "type-builtin-literals"), "match",
		"(this|true|false|undefined|null|object)",
		"(this|true|false|undefined|null)",
	);
}

/**
 * Numeric literals with C#-style type suffixes (spec §2.2): `123L`, `1.5f`, `2d`, `1.5m`,
 * and the column numbers `0cAA`.
 * The bigint suffix `n` does not exist in RAS and is dropped.
 */
function replaceNumericLiterals(grammar) {
	const numeric = rule(grammar, "numeric-literal");
	const upstreamNames = numeric.patterns.map((pattern) => pattern.name ?? pattern.captures?.[0]?.name);
	if (upstreamNames.join() !== "constant.numeric.hex.ts,constant.numeric.binary.ts,constant.numeric.octal.ts,constant.numeric.decimal.ts") {
		throw new Error("build-grammar: numeric-literal changed upstream; review replaceNumericLiterals");
	}
	const [hex, binary, octal, decimal] = numeric.patterns;
	const suffix = { name: "storage.type.numeric.ras" };
	const period = decimal.captures[1];

	// Only `L` is unambiguous after hex digits (`d` and `f` are digits there).
	for (const pattern of [hex, binary, octal]) {
		replaceIn(pattern, "match", "(n)?", "([Ll])?");
		pattern.captures = { 1: suffix };
	}

	const real = "([FfDdMm])?";
	decimal.match = [
		"(?x)",
		"(?<!\\$)(?:",
		`  (?:\\b[0-9][0-9_]*(\\.)[0-9][0-9_]*[eE][+-]?[0-9][0-9_]*${real}\\b)| # 1.1E+3f`,
		`  (?:\\b[0-9][0-9_]*(\\.)[eE][+-]?[0-9][0-9_]*${real}\\b)|             # 1.E+3f`,
		`  (?:\\B(\\.)[0-9][0-9_]*[eE][+-]?[0-9][0-9_]*${real}\\b)|             # .1E+3f`,
		`  (?:\\b[0-9][0-9_]*[eE][+-]?[0-9][0-9_]*${real}\\b)|                 # 1E+3f`,
		`  (?:\\b[0-9][0-9_]*(\\.)[0-9][0-9_]*${real}\\b)|                      # 1.5f 1.5d 1.5m`,
		"  (?:\\b[0-9][0-9_]*(\\.)\\B)|                                        # 1.",
		`  (?:\\B(\\.)[0-9][0-9_]*${real}\\b)|                                  # .5f`,
		"  (?:\\b[0-9][0-9_]*([LlFfDdMm])?\\b(?!\\.))                           # 1 123L 2d",
		")(?!\\$)",
	].join("\n");
	decimal.captures = {
		0: decimal.captures[0],
		1: period, 2: suffix,
		3: period, 4: suffix,
		5: period, 6: suffix,
		7: suffix,
		8: period, 9: suffix,
		10: period,
		11: period, 12: suffix,
		13: suffix,
	};

	// `0cAA` — an Excel column number in letters (spec §2.2). No suffix: `L` is a column letter.
	numeric.patterns.unshift({
		name: "constant.numeric.column.ts",
		match: "\\b(?<!\\$)0(?:c|C)[A-Za-z]+\\b(?!\\$)",
	});
}

/**
 * Decorators (`@fast`, `@on("Workbook.Open")`, `@bulk` ...). Upstream already
 * accepts a decorator wherever a declaration or statement may start, so `@bulk`
 * in front of `for` / `while` needs no structural change. RAS decorators are
 * compiler directives, so the name gets its own scope instead of being coloured
 * as an ordinary variable. Qualified names (`@a.b`) keep the upstream behaviour.
 */
function scopeDecoratorNames(grammar) {
	const decorator = rule(grammar, "decorator");
	replaceIn(decorator, "begin", "\\@", `(\\@)(?:(${IDENTIFIER})(?![_$[:alnum:]]|\\s*\\.))?`);
	decorator.beginCaptures = {
		1: decorator.beginCaptures[0],
		2: { name: "entity.name.function.decorator.ras" },
	};
}

/** `declare struct Name { field: type; }` (spec §11.3), modelled on upstream's interface declaration. */
function addStructDeclaration(grammar) {
	const interfaceDeclaration = rule(grammar, "interface-declaration");
	grammar.repository["struct-declaration"] = {
		name: "meta.struct.ras",
		begin: `${START_OF_IDENTIFIER}(?:(\\bexport)\\s+)?(?:(\\bdeclare)\\s+)?\\b(struct)\\b(?=\\s+[_$[:alpha:]])`,
		beginCaptures: {
			1: interfaceDeclaration.beginCaptures[1], // export
			2: interfaceDeclaration.beginCaptures[2], // declare
			3: { name: "storage.type.struct.ras" },
		},
		end: "(?<=\\})",
		patterns: [
			{ include: "#comment" },
			{ match: IDENTIFIER, captures: { 0: { name: "entity.name.type.struct.ras" } } },
			{ include: "#class-or-interface-body" },
		],
	};
	const declarations = rule(grammar, "declaration").patterns;
	const afterInterface = declarations.findIndex((pattern) => pattern.include === "#interface-declaration") + 1;
	if (afterInterface === 0) {
		throw new Error("build-grammar: #interface-declaration not found in declaration");
	}
	declarations.splice(afterInterface, 0, { include: "#struct-declaration" });
}

/**
 * Excel formulas in `` fx`SUM(A1:A3)` `` and `` ws.fx`A1` `` (spec §2.7). The formula is scoped
 * `meta.embedded.formula.ras`, so VS Code treats it as code (suggestions while typing), not as
 * a string; its own strings are strings again. `${…}` is a RAS expression as in any template.
 * Added in front of upstream's tagged templates, so it wins over them for the tag `fx`.
 */
function addFormulaTemplate(grammar) {
	const template = rule(grammar, "template").patterns;
	if (template[0]?.include !== "#template-call") {
		throw new Error("build-grammar: template no longer starts with #template-call; review addFormulaTemplate");
	}
	const scope = (name) => ({ name: `${name}.formula.ts` });
	grammar.repository["formula-template"] = {
		begin: "(?:(\\??\\.)\\s*)?(?<![_$[:alnum:]])(fx)\\s*(`)",
		beginCaptures: {
			1: { name: "punctuation.accessor.ts" },
			2: { name: "entity.name.function.tagged-template.ts" },
			3: { name: "string.template.ts punctuation.definition.string.template.begin.ts" },
		},
		end: "`",
		endCaptures: { 0: { name: "string.template.ts punctuation.definition.string.template.end.ts" } },
		contentName: "meta.embedded.formula.ts",
		patterns: [
			{ include: "#template-substitution-element" },
			{ include: "#string-character-escape" },
			{ include: "#formula" },
		],
	};
	grammar.repository.formula = {
		patterns: [
			// "text", with "" for a quote
			{
				name: "string.quoted.double.formula.ts",
				begin: '"',
				beginCaptures: { 0: { name: "punctuation.definition.string.begin.formula.ts" } },
				end: '"(?!")',
				endCaptures: { 0: { name: "punctuation.definition.string.end.formula.ts" } },
				patterns: [{ name: "constant.character.escape.formula.ts", match: '""' }],
			},
			// 'Sheet name'! and Sheet1! (also [Book.xlsx]Sheet1!)
			{
				match: "('(?:[^']|'')*'|(?:\\[[^\\]`]+\\])?[_\\p{L}][_.\\p{L}\\p{N}]*)(!)",
				captures: { 1: scope("entity.name.type.sheet"), 2: scope("punctuation.separator.sheet") },
			},
			{
				name: "constant.language.error.formula.ts",
				match: "#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\\?|NUM!|N/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!)",
			},
			// SUM( / _xlfn.XLOOKUP( / STDEV.S(
			{
				match: "(?i)(?<![_.\\p{L}\\p{N}$])((?:_xlfn\\.|_xlws\\.)?[_\\p{L}][_.\\p{L}\\p{N}]*)(?=\\s*\\()",
				captures: { 1: scope("support.function") },
			},
			{
				name: "constant.language.boolean.formula.ts",
				match: "(?i)(?<![_.\\p{L}\\p{N}$])(?:TRUE|FALSE)(?![_.\\p{L}\\p{N}(])",
			},
			// A1, $B$2, A:C, 1:3, R1C1, R[-1]C[2]
			{
				name: "variable.other.constant.reference.formula.ts",
				match: "(?i)(?<![_.\\p{L}\\p{N}$])(?:\\$?[A-Z]{1,3}\\$?[0-9]+|\\$?[A-Z]{1,3}(?=\\s*:\\s*\\$?[A-Z]{1,3}(?![_.\\p{L}\\p{N}]))|(?<=:)\\s*\\$?[A-Z]{1,3}|R(?:\\[-?[0-9]+\\]|[0-9]+)?C(?:\\[-?[0-9]+\\]|[0-9]+)?)(?![_.\\p{L}\\p{N}(])",
			},
			{
				name: "constant.numeric.formula.ts",
				match: "(?<![_.\\p{L}\\p{N}$])(?:[0-9]+(?:\\.[0-9]*)?|\\.[0-9]+)(?:[eE][+-]?[0-9]+)?",
			},
			// Structured references: Table1[Column], [@Column], [#Headers]
			{
				name: "meta.structured-reference.formula.ts",
				begin: "\\[",
				beginCaptures: { 0: scope("punctuation.definition.structured-reference.begin") },
				end: "\\]",
				endCaptures: { 0: scope("punctuation.definition.structured-reference.end") },
				patterns: [
					{ match: "\\[", name: "punctuation.definition.structured-reference.begin.formula.ts" },
					{ match: "\\]", name: "punctuation.definition.structured-reference.end.formula.ts" },
					{ match: "[@#][^\\]\\[,]*", name: "variable.language.structured-reference.formula.ts" },
					{ match: "[^\\]\\[,@#`$\\\\]+", name: "variable.other.property.formula.ts" },
				],
			},
			// Defined names, LET's and LAMBDA's names
			{ name: "variable.other.readwrite.formula.ts", match: "[_\\p{L}][_.\\p{L}\\p{N}]*" },
			{ name: "keyword.operator.comparison.formula.ts", match: "<>|<=|>=|=|<|>" },
			{ name: "keyword.operator.arithmetic.formula.ts", match: "[-+*/^%]" },
			{ name: "keyword.operator.concatenation.formula.ts", match: "&" },
			{ name: "keyword.operator.range.formula.ts", match: ":" },
			// implicit intersection, spilled range
			{ name: "keyword.operator.reference.formula.ts", match: "[@#]" },
			{ name: "meta.brace.round.ts", match: "[()]" },
			{ name: "meta.brace.curly.formula.ts", match: "[{}]" },
			{ name: "punctuation.separator.formula.ts", match: "[,;]" },
		],
	};
	template.unshift({ include: "#formula-template" });
	// Upstream's tagged-template call starts at the object of `ws.fx` (earlier than the `.`
	// above): it must not take `fx`.
	const call = rule(grammar, "template-call").patterns[0];
	const chain = `(?:${IDENTIFIER}\\s*\\??\\.\\s*)*`;
	replaceIn(call, "begin", "(?=(([_$[:alpha:]]", `(?!${chain}(?:\\??\\.\\s*)?fx\\s*\`)(?=(([_$[:alpha:]]`);
}

/** `.ts` → `.ras` on every scope name. */
function renameScopes(grammar) {
	const rename = (scopes) => scopes.replace(/\.ts(?=\s|$)/g, ".ras");
	walk(grammar.repository, (node) => {
		for (const key of ["name", "contentName"]) {
			if (typeof node[key] === "string") node[key] = rename(node[key]);
		}
	});
	walk(grammar.repository, (node) => {
		for (const key of ["name", "contentName"]) {
			if (typeof node[key] === "string" && /\.tsx?(\s|$)/.test(node[key])) {
				throw new Error(`build-grammar: scope not renamed: ${node[key]}`);
			}
		}
		if (typeof node.include === "string" && node.include.startsWith("source.ts")) {
			throw new Error(`build-grammar: include of ${node.include} not rewritten`);
		}
	});
}

export function buildGrammar() {
	const upstream = parsePlist(fs.readFileSync(UPSTREAM_PATH, "utf8"));
	if (upstream.scopeName !== "source.ts") {
		throw new Error(`build-grammar: unexpected upstream scopeName ${upstream.scopeName}`);
	}

	addParameterModifiers(upstream);
	addInlineModifiers(upstream);
	addResumeKeyword(upstream);
	replacePrimitiveTypes(upstream);
	replaceNumericLiterals(upstream);
	scopeDecoratorNames(upstream);
	addStructDeclaration(upstream);
	addFormulaTemplate(upstream);
	// JavaScript, DOM and Node.js library names (Math, JSON, console, document,
	// process, `.length` ...). RAS has its own standard library; its members are
	// coloured by the language server's semantic tokens instead.
	removeRule(upstream, "support-objects");
	renameScopes(upstream);

	const grammar = {
		information_for_contributors: [
			"This file is generated by scripts/build-grammar.mjs. Do not edit it by hand; change the script and run `npm run build:grammar`.",
			"It is derived from the TypeScript grammar of microsoft/TypeScript-TmLanguage (MIT License).",
			"See THIRD_PARTY_NOTICES.md for the licence and the list of modifications.",
		],
		version: `${UPSTREAM_REPOSITORY}/commit/${UPSTREAM_COMMIT}`,
		name: "RAS",
		scopeName: "source.ras",
		fileTypes: ["ras"],
		patterns: upstream.patterns,
		repository: upstream.repository,
	};
	return JSON.stringify(grammar, null, "\t") + "\n";
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
	const generated = buildGrammar();
	if (process.argv.includes("--check")) {
		const current = fs.existsSync(OUTPUT_PATH) ? fs.readFileSync(OUTPUT_PATH, "utf8") : "";
		if (current !== generated) {
			console.error("syntaxes/ras.tmLanguage.json is out of date. Run `npm run build:grammar`.");
			process.exit(1);
		}
		console.log("syntaxes/ras.tmLanguage.json is up to date.");
	} else {
		fs.mkdirSync(path.dirname(OUTPUT_PATH), { recursive: true });
		fs.writeFileSync(OUTPUT_PATH, generated);
		console.log(`Wrote ${path.relative(root, OUTPUT_PATH)} (${generated.length} bytes)`);
	}
}
