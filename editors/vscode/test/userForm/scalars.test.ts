import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { describe, it } from "node:test";
import { scalarProperties } from "../../src/userForm/contract";

// The table of plain properties exists twice: crates/ras-form/src/scalars.rs (generation,
// projection, validation) and contract.ts (the Designer). They have to name the same
// properties for the same control types.
describe("plain properties", () => {
	it("are the same as crates/ras-form/src/scalars.rs", () => {
		const rust = fs.readFileSync(path.join("..", "..", "crates", "ras-form", "src", "scalars.rs"), "utf8");
		const lists = new Map<string, string[]>();
		for (const m of rust.matchAll(/const (\w+): &\[&str\] = &\[([^\]]*)\];/g))
			lists.set(m[1], [...m[2].matchAll(/"(\w+)"/g)].map((t) => t[1]));
		const pairs = (entries: [string, string[]][]) =>
			entries.flatMap(([name, types]) => types.map((t) => `${name}:${t}`)).sort();
		const fromRust: [string, string[]][] = [...rust.matchAll(/\bs\(\s*"(\w+)",\s*(&\[[^\]]*\]|\w+),/g)].map((m) => [
			m[1],
			m[2].startsWith("&[") ? [...m[2].matchAll(/"(\w+)"/g)].map((t) => t[1]) : (lists.get(m[2]) ?? [`?${m[2]}`]),
		]);
		assert.ok(fromRust.length > 40, "the Rust table was not found");
		assert.deepEqual(pairs(scalarProperties.map(([name, types]) => [name, types])), pairs(fromRust));
	});
});
