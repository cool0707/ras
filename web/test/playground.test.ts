// The playground's compiler and runs, without a browser: the samples are compiled with the wasm
// module (web/pkg, `npm run wasm`) and run on its stand-in for Excel, as the page does
// (`npm test`, which bundles this file with esbuild and runs it with node --test).
import assert from "node:assert/strict";
import { readFileSync, readdirSync, statSync } from "node:fs";
import { join } from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { Book, compile, diffText, explain, format, generateTest, initSync } from "../pkg/ras_wasm.js";
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

/** `workbook` of the project's ras.toml (the samples' are named after the project: Hello.xlsx). */
const workbookOf = (toml: string) => /^\s*workbook\s*=\s*"([^"]+)"/m.exec(toml)![1];

function run(dir: string, kind: "run" | "test", target: string, ask?: Parameters<typeof runJob>[3]): { r: RunResult; before: Uint8Array } {
    const { files, toml, out } = build(dir);
    const before = book();
    const r = runJob(Book as unknown as WasmBookClass, { kind, js: out.js, files: [...files.map((f) => ({ path: f.path, data: f.text })), { path: workbookOf(toml), data: before }], target }, undefined, ask);
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
    const saved = r.saved.find((s) => s.path === "build/debug/workbooks/Hello.xlsx");
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

test("generates the skeleton of a test, which compiles and runs", () => {
    const { files, toml, out } = build(hello);
    assert.deepEqual(
        out.testTargets.map((t: { path: string; name: string; testPath: string }) => [t.path, t.name, t.testPath]),
        [["src/main.ras", "makeReport", "test/main.test.ras"], ["src/main.ras", "sumTo", "test/main.test.ras"]],
    );
    const input = (fs: { path: string; text: string }[]) => JSON.stringify({ files: fs, toml, sheets: ["Sheet1"], js: true });
    // Into the test file of the sample: at the end of its describe("sumTo").
    const added = JSON.parse(generateTest(input(files), "src/main.ras", "sumTo"));
    assert.equal(added.path, "test/main.test.ras");
    assert.equal(added.created, false);
    assert.match(added.text.split("\n")[added.line], /^ {4}test\("sumTo の戻り値", \(\) => \{$/);
    // Without the test file: a new one.
    const others = files.filter((f) => f.path !== "test/main.test.ras");
    const created = JSON.parse(generateTest(input(others), "src/main.ras", "makeReport"));
    assert.equal(created.created, true);
    assert.ok(created.text.startsWith('import { makeReport } from "../src/main";\n'), created.text);
    const next = [...others, { path: created.path, text: created.text }];
    const compiled = JSON.parse(compile(input(next)));
    assert.deepEqual(compiled.diagnostics.filter((d: { severity: string }) => d.severity === "error"), []);
    const r = runJob(Book as unknown as WasmBookClass, { kind: "test", js: compiled.js, files: [...next.map((f) => ({ path: f.path, data: f.text })), { path: workbookOf(toml), data: book() }], target: created.path });
    const results = r.events.filter((e) => (e as { type: string }).type === "test") as { status: string }[];
    assert.deepEqual(results.map((t) => t.status), ["passed"]);
    assert.deepEqual(JSON.parse(generateTest(input(files), "src/main.ras", "nothing")).error !== undefined, true);
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
    const packed = packBook(JSON.stringify({ files: p.files, toml: p.toml }), book(), "Hello.xlsx");
    assert.equal(packed.name, "Hello.xlsm");
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

/** A project with a form (forms/frmInput.form.json and its code-behind) that a function shows. */
async function formProject() {
    const { newFormSpec } = await import("../pkg/ras_wasm.js");
    const spec = JSON.parse(newFormSpec("frmInput"));
    spec.controls = [
        { id: "c1", name: "txtName", type: "TextBox", left: 12, top: 12, width: 150, height: 18 },
        { id: "c2", name: "btnOK", type: "CommandButton", caption: "OK", left: 12, top: 50, width: 72, height: 24 },
    ];
    return {
        toml: '[project]\nname = "P"\nworkbook = "Book.xlsx"\n',
        files: [
            { path: "src/main.ras", text: 'export function go() {\n    frmInput.show();\n    debug.print(Sheet1.range("A1").value);\n}\n' },
            { path: "forms/frmInput.form.json", text: JSON.stringify(spec, null, 2) },
            { path: "forms/frmInput.ras", text: '@on("btnOK.Click")\nfunction ok() {\n    Sheet1.range("A1").value = this.txtName.text;\n    unload(this);\n}\n' },
        ],
    };
}

test("compiles, checks and writes the forms of forms/", async () => {
    const { packBook } = await import("../pkg/ras_wasm.js");
    const p = await formProject();
    const out = JSON.parse(compile(JSON.stringify({ files: p.files, toml: p.toml, sheets: ["Sheet1"], js: true })));
    assert.deepEqual(out.diagnostics, []);
    assert.ok(out.entries.includes("frmInput.show"), out.entries.join(", "));
    assert.ok(out.vba.some((v: { name: string }) => v.name === "frmInput"));
    // A problem of a form's design is reported at its place in the .form.json.
    const broken = p.files.map((f) => (f.path.endsWith(".form.json") ? { ...f, text: f.text.replace('"name": "txtName"', '"name": "btnOK"') } : f));
    const bad = JSON.parse(compile(JSON.stringify({ files: broken, toml: p.toml, sheets: ["Sheet1"], js: true })));
    assert.ok(bad.diagnostics.some((d: { path: string; severity: string }) => d.path === "forms/frmInput.form.json" && d.severity === "error"), JSON.stringify(bad.diagnostics));
    const packed = packBook(JSON.stringify({ files: p.files, toml: p.toml }), book(), "Book.xlsx");
    assert.ok(JSON.parse(packed.report).created.includes("frmInput"), packed.report);
});

test("shows a form through the page's host and runs its handlers", async () => {
    const p = await formProject();
    const out = JSON.parse(compile(JSON.stringify({ files: p.files, toml: p.toml, sheets: ["Sheet1"], js: true })));
    const sent: { t: string; name?: string }[] = [];
    const actions = [
        { t: "input", control: "txtName", text: "山田", seq: 1 },
        { t: "click", control: "btnOK" },
    ];
    const forms = {
        send: (m: unknown) => void sent.push(m as { t: string }),
        receive: () => {
            const a = actions.shift();
            return a ? { ...a, id: (sent.find((m) => m.t === "show") as { id?: number } | undefined)?.id } : { t: "disconnected" };
        },
    };
    const r = runJob(Book as unknown as WasmBookClass, { kind: "run", js: out.js, files: [...p.files.map((f) => ({ path: f.path, data: f.text })), { path: "Book.xlsx", data: book() }], target: "Main.go" }, undefined, undefined, forms);
    assert.equal(r.exitCode, 0, JSON.stringify(r.output));
    assert.equal(sent[0].t, "show");
    assert.equal(sent[0].name, "frmInput");
    assert.ok(sent.some((m) => m.t === "unload"));
    assert.match(r.output.map((o) => o.text).join(""), /山田/);
});

test("reads the ActiveX controls of the sheets for @on(\"<control>.<event>\")", async () => {
    const { sheetControls } = await import("../pkg/ras_wasm.js");
    assert.deepEqual(JSON.parse(sheetControls(book())), []);
    assert.equal(sheetControls(new Uint8Array([1, 2, 3])), "null");
    const files = [{ path: "objects/Sheet1.ras", text: '@on("CommandButton1.Click")\nfunction c() {\n}\n' }];
    const toml = '[project]\nname = "P"\n';
    const known = JSON.parse(compile(JSON.stringify({ files, toml, sheets: ["Sheet1"], sheetControls: [{ name: "Sheet1", controls: [{ name: "CommandButton1", type: "CommandButton" }] }] })));
    assert.deepEqual(known.diagnostics, []);
    // Another name's Click is a shape's or a form control's (its macro); its other events are no control's.
    const shape = JSON.parse(compile(JSON.stringify({ files, toml, sheets: ["Sheet1"], sheetControls: [] })));
    assert.deepEqual(shape.diagnostics, []);
    const other = [{ path: "objects/Sheet1.ras", text: '@on("CommandButton1.MouseDown")\nfunction c() {\n}\n' }];
    const missing = JSON.parse(compile(JSON.stringify({ files: other, toml, sheets: ["Sheet1"], sheetControls: [] })));
    assert.ok(missing.diagnostics.some((d: { code: string }) => d.code === "E0613"), JSON.stringify(missing.diagnostics));
});
