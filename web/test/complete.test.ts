// The completion, the hovers and the parameter information of RAS in the playground's editor:
// the compiler's answers (web/pkg, `npm run wasm`) as CodeMirror shows them (`npm test`).
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { type Completion, CompletionContext, type CompletionResult, type CompletionSource } from "@codemirror/autocomplete";
import { ensureSyntaxTree } from "@codemirror/language";
import { EditorState } from "@codemirror/state";
import type { EditorView } from "@codemirror/view";
import { Analysis, initSync } from "../pkg/ras_wasm.js";
import { type Item, rasCompletions } from "../src/complete";
import { rasLanguage } from "../src/formula";
import { type HoverInfo, openParen, type SignatureInfo } from "../src/hover";

const web = join(fileURLToPath(import.meta.url), "..", "..");
initSync({ module: readFileSync(join(web, "pkg", "ras_wasm_bg.wasm")) });

const UTIL = "/** 2 倍にします。 */\nexport function twice(n: int): int {\n    return n * 2;\n}\n";

/** The completion at `|` of `text` (src/main.ras, beside src/util.ras), as the editor asks. */
function complete(text: string, explicit = false): { state: EditorState; pos: number; result: CompletionResult | null } {
    const pos = text.indexOf("|");
    const doc = text.replace("|", "");
    const input = JSON.stringify({ files: [{ path: "src/main.ras", text: doc }, { path: "src/util.ras", text: UTIL }], toml: '[project]\nname = "P"\n' });
    const analysis = new Analysis(input);
    const source = rasCompletions((line, col, namedOnly) => JSON.parse(analysis.completions("src/main.ras", line, col, namedOnly)) as Item[]);
    const state = EditorState.create({ doc, extensions: [rasLanguage(source)] });
    ensureSyntaxTree(state, doc.length, 5000);
    let result: CompletionResult | null = null;
    for (const s of state.languageDataAt<CompletionSource>("autocomplete", pos, -1)) {
        const r = s(new CompletionContext(state, pos, explicit));
        if (r && !(r instanceof Promise)) result = r;
    }
    analysis.free();
    return { state, pos, result };
}

/** The text after accepting the option `label`. */
function accept(c: ReturnType<typeof complete>, label: string): string {
    const o = c.result!.options.find((x) => x.label === label);
    assert.ok(o, `${label} is offered`);
    if (typeof o.apply !== "function") {
        return c.state.update({ changes: { from: c.result!.from, to: c.pos, insert: o.apply ?? o.label } }).state.doc.toString();
    }
    let state = c.state;
    const view = { state, dispatch: (spec: Parameters<EditorState["update"]>[0]) => void (state = state.update(spec).state) };
    (o.apply as (view: EditorView, c: Completion, from: number, to: number) => void)(view as unknown as EditorView, o, c.result!.from, c.pos);
    return state.doc.toString();
}

const labels = (r: CompletionResult | null) => (r?.options ?? []).map((o) => o.label);

test("after a dot: the members of the type, while the next line follows", () => {
    const c = complete('export function f() {\n    const words = ["a"];\n    words.|\n    msgBox("x");\n}\n');
    assert.ok(labels(c.result).includes("map") && labels(c.result).includes("length"), labels(c.result).join());
    assert.ok(!labels(c.result).includes("msgBox"));
    assert.equal(c.result!.from, c.pos);
    assert.equal(c.result!.options.find((o) => o.label === "map")!.type, "method");
    // The word written so far is replaced.
    const typed = complete('export function f() {\n    const words = ["a"];\n    words.le|\n}\n');
    assert.equal(typed.result!.from, typed.pos - 2);
    assert.match(accept(typed, "length"), /words\.length\n/);
});

test("a word: the names in scope, and the exports of other modules with their import", () => {
    const c = complete("export function f(count: int) {\n    let total = 0;\n    t|\n}\n");
    for (const want of ["total", "count", "f", "msgBox", "twice"]) assert.ok(labels(c.result).includes(want), want);
    const twice = c.result!.options.find((o) => o.label === "twice")!;
    assert.equal(twice.type, "function");
    assert.match(twice.detail ?? "", /インポート/);
    const text = accept(c, "twice");
    assert.match(text, /^import \{ twice \} from "[^"]*util";\n/);
    assert.match(text, /\n    twice\n\}\n$/);
});

test("nothing where no name is being written", () => {
    assert.equal(complete("export function f() {\n    |\n}\n").result, null);
    assert.equal(complete("export function f() {\n    let x = 1.|\n}\n").result, null);
    assert.equal(complete("export function f() {\n    let x = 12|\n}\n").result, null);
    // … unless asked for (Ctrl+Space).
    assert.ok(labels(complete("export function f() {\n    |\n}\n", true).result).includes("msgBox"));
});

test("after ( and , : the names of the arguments", () => {
    const c = complete("function area(width: int, height: int): int {\n    return width * height;\n}\nexport function f() {\n    area(|\n}\n");
    assert.deepEqual(labels(c.result).sort(), ["height:", "width:"]);
    // The next keystroke asks again, for every name.
    assert.equal(c.result!.validFor, undefined);
    assert.match(accept(c, "width:"), /area\(width: \n/);
});

test("in the formula of an fx template: Excel's functions", () => {
    const c = complete("export function f() {\n    const v = Sheet1.fx`1 + VLO|`;\n}\n");
    assert.ok(labels(c.result).includes("VLOOKUP") && !labels(c.result).includes("msgBox"), labels(c.result).slice(0, 5).join());
    assert.match(accept(c, "VLOOKUP"), /fx`1 \+ VLOOKUP\(`/);
});

test("the call the cursor is in is found in the text", () => {
    const at = (text: string) => openParen(EditorState.create({ doc: text.replace("|", "") }), text.indexOf("|"));
    assert.equal(at("f(1, |"), 1);
    assert.equal(at("f(g(1), |"), 1);
    assert.equal(at("f(g(1, |"), 3);
    assert.equal(at('f(")", |'), 1);
    assert.equal(at("f(1); // (\n|"), -1);
    assert.equal(at("f(1)|"), -1);
    assert.equal(at("ws.fx`SUM(A1, |`"), 9);
});

test("hovers and parameter information of a call", () => {
    const text = "/**\n * 面積を返します。\n * @param width 幅\n */\nfunction area(width: int, height: int): int {\n    return width * height;\n}\nexport function f() {\n    area(1, \n}\n";
    const analysis = new Analysis(JSON.stringify({ files: [{ path: "src/main.ras", text }], toml: '[project]\nname = "P"\n' }));
    const hover = JSON.parse(analysis.hover("src/main.ras", 8, 5)) as HoverInfo;
    assert.match(hover.markdown, /area\(width: int, height: int\): int[\s\S]*面積を返します。/);
    assert.deepEqual([hover.line, hover.col, hover.endCol], [8, 4, 8]);
    const help = JSON.parse(analysis.signatureHelp("src/main.ras", 8, 12)) as SignatureInfo;
    const sig = help.signatures[help.active];
    assert.equal(sig.active, 1);
    assert.equal(sig.label.slice(sig.params[1].from, sig.params[1].to), "height: int");
    assert.equal(sig.params[0].doc, "幅");
    assert.equal(analysis.signatureHelp("src/main.ras", 7, 0), "null");
    analysis.free();
});
