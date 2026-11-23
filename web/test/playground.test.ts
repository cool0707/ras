// The playground's compiler and runs, without a browser: the samples are compiled with the wasm
// module (web/pkg, `npm run wasm`) and run on its stand-in for Excel, as the page does
// (`npm test`, which bundles this file with esbuild and runs it with node --test).
import assert from "node:assert/strict";
import { readFileSync, readdirSync, statSync } from "node:fs";
import { join } from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { Book, compile, diffText, explain, format, initSync } from "../pkg/ras_wasm.js";
import { type RunResult, runJob, type WasmBookClass } from "../src/runtime/run";

const web = join(fileURLToPath(import.meta.url), "..", "..");
initSync({ module: readFileSync(join(web, "pkg", "ras_wasm_bg.wasm")) });

function project(dir: string) {
    const files: { path: string; text: string }[] = [];
    const walk = (d: string, rel: string) => {
        for (const n of readdirSync(d)) {
            const p = join(d, n);
            if (statSync(p).isDirectory()) walk(p, `${rel}${n}/`);
            else if (n.endsWith(".ras")) files.push({ path: `${rel}${n}`, text: readFileSync(p, "utf8") });
        }
    };
    for (const sub of ["src", "types", "objects", "test"]) {
        try {
            walk(join(dir, sub), `${sub}/`);
        } catch {
            // not in this project
        }
    }
    return { files, toml: readFileSync(join(dir, "ras.toml"), "utf8") };
}

function build(dir: string) {
    const p = project(dir);
    const out = JSON.parse(compile(JSON.stringify({ files: p.files, toml: p.toml, sheets: ["Sheet1"], js: true })));
    return { ...p, out };
}

const book = () => Book.create().toBytes();

function run(dir: string, kind: "run" | "test", target: string, ask?: Parameters<typeof runJob>[3]): { r: RunResult; before: Uint8Array } {
    const { files, out } = build(dir);
    const before = book();
    const r = runJob(Book as unknown as WasmBookClass, { kind, js: out.js, files: [...files.map((f) => ({ path: f.path, data: f.text })), { path: "Book.xlsx", data: before }], target }, undefined, ask);
    return { r, before };
}

const hello = join(web, "samples", "hello");
const dialogs = join(web, "samples", "dialogs");

test("compiles the sample to VBA without the tests", () => {
    const { out } = build(hello);
    assert.deepEqual(out.diagnostics.filter((d: { severity: string }) => d.severity === "error"), []);
    const names = out.vba.map((f: { fileName: string }) => f.fileName);
    assert.ok(names.includes("Main.bas"), names.join());
    assert.ok(!names.some((n: string) => n.startsWith("T_")), names.join());
    assert.ok(out.tests);
    assert.deepEqual(out.entries, ["Main.makeReport"]);
});

test("reports errors at their positions, with the explanation of the code", () => {
    const out = JSON.parse(compile(JSON.stringify({ files: [{ path: "src/main.ras", text: 'export function f() {\n    let x: int = "a";\n}\n' }], toml: '[project]\nname = "P"\n' })));
    const e = out.diagnostics.find((d: { severity: string }) => d.severity === "error");
    assert.equal(e.path, "src/main.ras");
    assert.equal(e.line, 1);
    assert.match(explain(e.code) ?? "", new RegExp(e.code));
    assert.deepEqual(out.vba, []);
});

test("formats, and leaves what does not parse", () => {
    assert.equal(format("export function f() {\nreturn;\n}\n"), "export function f() {\n    return;\n}\n");
    assert.equal(format("export function ("), undefined);
});

test("runs a function on the stand-in and reports the cells it changed", () => {
    const { r, before } = run(hello, "run", "Main.makeReport");
    assert.equal(r.exitCode, 0, JSON.stringify(r.output));
    const printed = r.output.map((o) => o.text).join("");
    assert.match(printed, /合計: 2060 円/);
    const saved = r.saved.find((s) => s.path === "build/debug/workbooks/Book.xlsx");
    assert.ok(saved);
    assert.match(diffText(before, saved.data, 100), /D5\t（空） → =SUM\(D2:D4\)（2060）/);
});

test("runs the tests of a test file", () => {
    const { r } = run(hello, "test", "test/main.test.ras");
    const results = r.events.filter((e) => (e as { type: string }).type === "test") as { status: string; name: string[] }[];
    assert.equal(results.length, 3);
    assert.ok(results.every((t) => t.status === "passed"), JSON.stringify(results));
    assert.equal(r.exitCode, 0);
});

test("asks the dialogs, and reports a run-time error at its line of RAS", () => {
    const asked: string[] = [];
    const { r } = run(dialogs, "run", "Main.ask", (d) => {
        asked.push(d.kind);
        return d.kind === "inputBox" ? "山田" : 6;
    });
    assert.deepEqual(asked, ["inputBox", "msgBox"]);
    assert.match(r.output.map((o) => o.text).join(""), /書き込みました/);
    const failed = run(dialogs, "run", "Main.oops").r;
    assert.equal(failed.exitCode, 1);
    const error = (failed.report as { error: { code: number; file: string; line: number } }).error;
    assert.equal(error.code, 9);
    assert.equal(error.file, "src/main.ras");
    assert.equal(error.line, 22);
});

test("writes the VBA into a workbook (.xlsx → .xlsm)", async () => {
    const { packBook } = await import("../pkg/ras_wasm.js");
    const p = project(hello);
    const packed = packBook(JSON.stringify({ files: p.files, toml: p.toml }), book(), "Book.xlsx");
    assert.equal(packed.name, "Book.xlsm");
    const report = JSON.parse(packed.report);
    assert.ok(report.created.includes("Main"), packed.report);
    assert.ok(packed.bytes.length > 1000);
});

test("edits the workbook as the page of the book tab posts (ras book edit)", async () => {
    const { diffText, editBook, viewPage } = await import("../pkg/ras_wasm.js");
    const before = book();
    const page = viewPage(before, "Book.xlsx", undefined, "", true);
    assert.match(page, /"editable":true/);
    assert.match(page, /"host":"parent"/);
    const after = editBook(before, JSON.stringify([{ op: "set", range: "B2", text: "100" }, { op: "format", range: "B2", set: { bold: true } }]));
    assert.match(diffText(before, after, 10), /B2/);
    assert.throws(() => editBook(before, JSON.stringify({ op: "nope", range: "A1" })));
});
