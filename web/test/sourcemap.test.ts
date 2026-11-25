// The debug build's source maps for the browser's DevTools (spec §14.10.9): the runtime inlines
// each module's map with the text of its `.ras` files (`npm test`; this file does not need web/pkg).
import assert from "node:assert/strict";
import { test } from "node:test";
import { MemFs } from "../src/runtime/memfs";
import { inlineSourceMap, NodeWorld } from "../src/runtime/node";

const SRC = "export function run() {\n    debugger;\n    console.log(\"日本語\");\n}\n";

function project(): MemFs {
    const fs = new MemFs("/project");
    fs.put("src/main.ras", SRC);
    fs.put(
        "build/debug/main.js",
        '// Generated\n"use strict";\nexports.run = run;\nfunction run() {\n    debugger;\n    console.log("日本語");\n}\n//# sourceMappingURL=main.js.map\n',
    );
    fs.put(
        "build/debug/main.js.map",
        JSON.stringify({ version: 3, file: "main.js", sources: ["../../src/main.ras"], names: [], mappings: ";;;AAAA;IACI;IACA" }),
    );
    return fs;
}

function decode(sourceMap: string) {
    const m = /^\n\/\/# sourceMappingURL=data:application\/json;charset=utf-8;base64,(.+)$/.exec(sourceMap);
    assert.ok(m, sourceMap);
    return JSON.parse(Buffer.from(m[1], "base64").toString("utf8"));
}

test("the source map is inlined with the text of the .ras files", () => {
    const fs = project();
    const text = fs.readFileSync("/project/build/debug/main.js", "utf8") as string;
    const { body, sourceMap } = inlineSourceMap(text, "/project/build/debug/main.js", fs);
    assert.ok(!body.includes("sourceMappingURL"), body);
    assert.ok(body.endsWith("}\n"));
    const map = decode(sourceMap);
    assert.deepEqual(map.sources, ["ras://project/src/main.ras"]);
    assert.deepEqual(map.sourcesContent, [SRC]);
    assert.equal(map.mappings, ";;;AAAA;IACI;IACA");
});

test("the text of the standard library comes from the map", () => {
    const fs = project();
    fs.put("build/debug/ras_collections.js", "x;\n//# sourceMappingURL=ras_collections.js.map\n");
    const lib = "export class List {}\n";
    fs.put(
        "build/debug/ras_collections.js.map",
        JSON.stringify({ version: 3, sources: ["../../<stdlib>/ras_collections.ras"], sourcesContent: [lib], names: [], mappings: "AAAA" }),
    );
    const text = fs.readFileSync("/project/build/debug/ras_collections.js", "utf8") as string;
    const map = decode(inlineSourceMap(text, "/project/build/debug/ras_collections.js", fs).sourceMap);
    assert.deepEqual(map.sources, ["ras://stdlib/ras_collections.ras"]);
    assert.deepEqual(map.sourcesContent, [lib]);
});

test("modules without a map are left as they are", () => {
    const fs = project();
    const text = '"use strict";\nexports.x = 1;\n';
    assert.deepEqual(inlineSourceMap(text, "/project/build/debug/ras-runtime.js", fs), { body: text, sourceMap: "" });
    // A map that is not there: the comment goes (it would name a URL that is not there either).
    const missing = inlineSourceMap("x;\n//# sourceMappingURL=gone.js.map\n", "/project/build/debug/gone.js", fs);
    assert.deepEqual(missing, { body: "x;\n", sourceMap: "" });
});

test("modules with an inlined map still load and run", () => {
    const fs = project();
    let out = "";
    const world = new NodeWorld({ fs, out: (_s, t) => (out += t), env: {}, argv: [], extra: {} });
    const m = world.require("./build/debug/main.js", "/project") as { run(): void };
    // `debugger;` does nothing without a debugger attached.
    const log = console.log;
    console.log = (...a: unknown[]) => (out += `${a.join(" ")}\n`);
    try {
        m.run();
    } finally {
        console.log = log;
    }
    assert.equal(out, "日本語\n");
});
