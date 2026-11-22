// The playground: the files of a project in an editor, the VBA they compile to, the problems the
// compiler finds, and runs and tests of the program on the stand-in for Excel, all in the browser
// (the compiler and the stand-in are crates/ras-wasm).

import { autocompletion, closeBrackets, closeBracketsKeymap, completionKeymap } from "@codemirror/autocomplete";
import { defaultKeymap, history as undoHistory, historyKeymap, indentWithTab } from "@codemirror/commands";
import { javascript } from "@codemirror/lang-javascript";
import { bracketMatching, defaultHighlightStyle, foldGutter, HighlightStyle, indentOnInput, indentUnit, StreamLanguage, syntaxHighlighting } from "@codemirror/language";
import { vb } from "@codemirror/legacy-modes/mode/vb";
import { type Diagnostic as LintDiagnostic, lintGutter, setDiagnostics } from "@codemirror/lint";
import { highlightSelectionMatches, searchKeymap } from "@codemirror/search";
import { EditorState, type Extension } from "@codemirror/state";
import { drawSelection, EditorView, highlightActiveLine, highlightActiveLineGutter, keymap, lineNumbers } from "@codemirror/view";
import { tags } from "@lezer/highlight";
import init, { Book, compile, diffText, encodeVba, explain, format, packBook, viewPage } from "../pkg/ras_wasm.js";
import type { Dialog, RunJob, RunResult } from "./runtime/run";
import samples from "virtual:samples";
import type { FromWorker, ToWorker } from "./worker";

interface ProjectFile {
    path: string;
    text: string;
}
interface Diag {
    path: string | null;
    line: number;
    col: number;
    endLine: number;
    endCol: number;
    severity: "error" | "warning" | "info" | "hint";
    code: string;
    message: string;
}
interface Compiled {
    diagnostics: Diag[];
    /** `text`: the code pane's (LF); `file`: the exported file, with its header (CRLF). */
    vba: { name: string; fileName: string; text: string; file: string }[];
    js: { path: string; text: string }[] | null;
    entries: string[];
    tests: boolean;
    codepage: number;
    error: string | null;
}
interface Workbook {
    name: string;
    bytes: Uint8Array;
}

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const STORE = "ras-playground:v1";

// ---------------------------------------------------------------------------- state

let files: ProjectFile[] = [];
let active = "";
let sampleId = "";
let workbook: Workbook;
/** The workbook as the last run left it, and what it was before. */
let lastRun: { before: Workbook; after: Workbook } | null = null;
let compiled: Compiled | null = null;
let wasmModule: WebAssembly.Module;
const states = new Map<string, EditorState>();
let sheetsCache: { bytes: Uint8Array; sheets: string[] } | null = null;

function newWorkbook(): Workbook {
    const b = Book.create();
    try {
        return { name: "Book.xlsx", bytes: b.toBytes() };
    } finally {
        b.free();
    }
}

/** The code names of the sheets of the workbook (what `Sheet1` and objects/Sheet1.ras refer to). */
function sheetNames(): string[] {
    if (sheetsCache && sheetsCache.bytes === workbook.bytes) return sheetsCache.sheets;
    const sheets: string[] = [];
    try {
        const b = Book.fromBytes(workbook.bytes);
        try {
            const n = JSON.parse(b.call("sheetCount", "[]")) as number;
            for (let i = 0; i < n; i++) sheets.push((JSON.parse(b.call("sheetInfo", `[${i}]`)) as { codeName: string }).codeName);
        } finally {
            b.free();
        }
    } catch {
        sheets.push("Sheet1");
    }
    sheetsCache = { bytes: workbook.bytes, sheets };
    return sheets;
}

function toml(): string {
    return files.find((f) => f.path === "ras.toml")?.text ?? '[project]\nname = "Playground"\nworkbook = "Book.xlsx"\n';
}

/** `workbook` of ras.toml: where the program finds the workbook. */
function workbookPath(): string | null {
    const m = /^\s*workbook\s*=\s*"([^"]+)"/m.exec(toml());
    return m ? m[1] : null;
}

/** Points `workbook` of ras.toml at `name`. */
function setWorkbookPath(name: string) {
    const f = files.find((x) => x.path === "ras.toml");
    if (!f) return;
    const line = `workbook = "${name.replace(/"/g, "")}"`;
    if (/^\s*workbook\s*=.*$/m.test(f.text)) f.text = f.text.replace(/^\s*workbook\s*=.*$/m, line);
    else if (/^\[project\][^\n]*\n/m.test(f.text)) f.text = f.text.replace(/^\[project\][^\n]*\n/m, (h) => `${h}${line}\n`);
    else f.text = `[project]\n${line}\n${f.text}`;
    states.delete("ras.toml");
}

// ---------------------------------------------------------------------------- saving and sharing

function toBase64(bytes: Uint8Array): string {
    let s = "";
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    return btoa(s);
}
function fromBase64(text: string): Uint8Array {
    const s = atob(text);
    const out = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
    return out;
}

function save() {
    try {
        const book = workbook.bytes.length < 1_500_000 ? { name: workbook.name, data: toBase64(workbook.bytes) } : null;
        localStorage.setItem(STORE, JSON.stringify({ files, active, sample: sampleId, book }));
    } catch {
        // Without storage the project lasts as long as the page.
    }
}

function restore(): boolean {
    try {
        const saved = JSON.parse(localStorage.getItem(STORE) || "null");
        if (!saved || !Array.isArray(saved.files) || !saved.files.length) return false;
        files = saved.files;
        active = saved.active;
        sampleId = saved.sample || "";
        if (saved.book) workbook = { name: saved.book.name, bytes: fromBase64(saved.book.data) };
        return true;
    } catch {
        return false;
    }
}

async function deflate(text: string): Promise<string> {
    const stream = new Blob([text]).stream().pipeThrough(new CompressionStream("deflate-raw"));
    const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
    return toBase64(bytes).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
async function inflate(text: string): Promise<string> {
    const bytes = fromBase64(text.replace(/-/g, "+").replace(/_/g, "/"));
    const stream = new Blob([bytes as BlobPart]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Response(stream).text();
}

async function shareUrl(): Promise<string> {
    const data = await deflate(JSON.stringify({ files, active }));
    const url = new URL(location.href);
    url.hash = `code=${data}`;
    return url.toString();
}

async function fromHash(): Promise<boolean> {
    const m = /^#code=(.+)$/.exec(location.hash);
    if (!m) return false;
    try {
        const p = JSON.parse(await inflate(m[1]));
        if (!Array.isArray(p.files)) return false;
        files = p.files.filter((f: ProjectFile) => typeof f.path === "string" && typeof f.text === "string");
        active = p.active;
        sampleId = "";
        return files.length > 0;
    } catch {
        toast("URL のプロジェクトを読み取れませんでした");
        return false;
    }
}

function loadSample(id: string) {
    const s = samples.find((x) => x.id === id) || samples[0];
    files = s.files.map((f) => ({ ...f }));
    sampleId = s.id;
    active = files.find((f) => f.path.startsWith("src/"))?.path || files[0].path;
    workbook = newWorkbook();
    lastRun = null;
    states.clear();
}

// ---------------------------------------------------------------------------- editors

const highlight = HighlightStyle.define([
    { tag: tags.keyword, color: "var(--hl-keyword, #7c3aed)" },
    { tag: [tags.string, tags.special(tags.string)], color: "var(--hl-string, #a1460a)" },
    { tag: tags.number, color: "var(--hl-number, #0b7a75)" },
    { tag: tags.comment, color: "var(--hl-comment, #6a7a72)", fontStyle: "italic" },
    { tag: [tags.typeName, tags.className], color: "var(--hl-type, #1f6feb)" },
    { tag: tags.function(tags.variableName), color: "var(--hl-fn, #b4235a)" },
    { tag: tags.meta, color: "var(--hl-meta, #9d5d00)" },
]);

const theme = EditorView.theme({
    "&": { color: "var(--text)", backgroundColor: "var(--code-bg)" },
    ".cm-content": { caretColor: "var(--text)" },
    ".cm-gutters": { backgroundColor: "var(--panel)", color: "var(--muted)", borderRight: "1px solid var(--line)" },
    ".cm-activeLine": { backgroundColor: "color-mix(in srgb, var(--accent-soft) 45%, transparent)" },
    ".cm-activeLineGutter": { backgroundColor: "var(--accent-soft)" },
    "&.cm-focused .cm-selectionBackground, .cm-selectionBackground, ::selection": { backgroundColor: "var(--selection) !important" },
    ".cm-tooltip": { backgroundColor: "var(--panel)", color: "var(--text)", border: "1px solid var(--line)" },
});

const common: Extension[] = [
    lineNumbers(),
    highlightActiveLineGutter(),
    foldGutter(),
    drawSelection(),
    indentUnit.of("    "),
    EditorState.tabSize.of(4),
    bracketMatching(),
    highlightSelectionMatches(),
    syntaxHighlighting(highlight),
    syntaxHighlighting(defaultHighlightStyle, { fallback: true }),
    theme,
];

function language(path: string): Extension {
    if (path.endsWith(".toml")) return [];
    // RAS is close enough to TypeScript for its highlighting.
    return javascript({ typescript: true });
}

function editorState(f: ProjectFile): EditorState {
    return EditorState.create({
        doc: f.text,
        extensions: [
            common,
            undoHistory(),
            highlightActiveLine(),
            indentOnInput(),
            closeBrackets(),
            autocompletion(),
            lintGutter(),
            language(f.path),
            keymap.of([
                { key: "Mod-Enter", run: () => (run("run"), true) },
                { key: "Shift-Alt-f", run: () => (formatActive(), true) },
                indentWithTab,
                ...closeBracketsKeymap,
                ...defaultKeymap,
                ...searchKeymap,
                ...historyKeymap,
                ...completionKeymap,
            ]),
            EditorView.updateListener.of((u) => {
                if (!u.docChanged) return;
                const file = files.find((x) => x.path === active);
                if (file) file.text = u.state.doc.toString();
                scheduleCompile();
            }),
        ],
    });
}

let editor: EditorView;
let vbaView: EditorView;

function openFile(path: string, at?: { line: number; col: number }) {
    const f = files.find((x) => x.path === path);
    if (!f) return;
    if (active && editor) states.set(active, editor.state);
    active = path;
    let state = states.get(path);
    if (!state || state.doc.toString() !== f.text) state = editorState(f);
    editor.setState(state);
    applyDiagnostics();
    if (at) {
        const doc = editor.state.doc;
        const line = doc.line(Math.min(Math.max(at.line + 1, 1), doc.lines));
        const pos = Math.min(line.from + at.col, line.to);
        editor.dispatch({ selection: { anchor: pos }, scrollIntoView: true });
    }
    editor.focus();
    renderFiles();
    save();
}

function renderFiles() {
    const nav = $("files");
    nav.textContent = "";
    const errors = new Set((compiled?.diagnostics || []).filter((d) => d.severity === "error").map((d) => d.path));
    for (const f of files) {
        const b = document.createElement("button");
        b.className = `file${errors.has(f.path) ? " has-error" : ""}`;
        b.setAttribute("aria-selected", String(f.path === active));
        b.title = f.path === "ras.toml" ? "プロジェクトの設定" : f.path;
        b.append(f.path);
        if (f.path !== "ras.toml") {
            const x = document.createElement("span");
            x.className = "x";
            x.textContent = "×";
            x.title = `${f.path} を削除`;
            x.onclick = (e) => {
                e.stopPropagation();
                if (!confirm(`${f.path} を削除しますか？`)) return;
                files = files.filter((y) => y !== f);
                states.delete(f.path);
                if (active === f.path) openFile(files.find((y) => y.path.startsWith("src/"))?.path || files[0].path);
                renderFiles();
                scheduleCompile(0);
            };
            b.append(x);
        }
        b.onclick = () => openFile(f.path);
        b.ondblclick = () => renameFile(f);
        nav.append(b);
    }
    const add = document.createElement("button");
    add.textContent = "＋";
    add.title = "ファイルを追加（src/〜.ras、objects/Sheet1.ras、test/〜.test.ras、types/〜.d.ras）";
    add.onclick = addFile;
    nav.append(add);
}

function validPath(p: string): string | null {
    const path = p.trim().replace(/\\/g, "/").replace(/^\/+/, "");
    if (!/^(src|objects|test|types)\/[^/].*\.ras$/.test(path)) {
        toast("src/・objects/・test/・types/ の下の .ras ファイルにしてください（例: src/util.ras）");
        return null;
    }
    if (files.some((f) => f.path.toLowerCase() === path.toLowerCase())) {
        toast(`${path} はもうあります`);
        return null;
    }
    return path;
}

function addFile() {
    const p = prompt("追加するファイル（src/〜.ras が標準モジュール、objects/Sheet1.ras がシートのモジュール、test/〜.test.ras がテスト）", "src/util.ras");
    if (!p) return;
    const path = validPath(p);
    if (!path) return;
    const text = path.startsWith("test/")
        ? 'describe("", () => {\n    test("", () => {\n        expect(1 + 1).toBe(2);\n    });\n});\n'
        : path.startsWith("objects/")
          ? '@on("Worksheet.Change")\nfunction onChange(target: Excel.Range) {\n    debug.print(target.address);\n}\n'
          : "export function example() {\n}\n";
    files.push({ path, text });
    openFile(path);
    scheduleCompile(0);
}

function renameFile(f: ProjectFile) {
    if (f.path === "ras.toml") return;
    const p = prompt("新しい名前", f.path);
    if (!p || p === f.path) return;
    const path = validPath(p);
    if (!path) return;
    states.delete(f.path);
    const wasActive = active === f.path;
    f.path = path;
    if (wasActive) {
        active = "";
        openFile(path);
    }
    renderFiles();
    scheduleCompile(0);
}

function formatActive() {
    const f = files.find((x) => x.path === active);
    if (!f || !f.path.endsWith(".ras")) return;
    const out = format(f.text);
    if (out === undefined) {
        toast("構文エラーがあるため整形できません");
        return;
    }
    if (out !== f.text) editor.dispatch({ changes: { from: 0, to: editor.state.doc.length, insert: out } });
}

// ---------------------------------------------------------------------------- compiling

let compileTimer = 0;
function scheduleCompile(delay = 250) {
    clearTimeout(compileTimer);
    compileTimer = window.setTimeout(() => {
        runCompile();
        save();
    }, delay);
}

function runCompile() {
    const started = performance.now();
    const input = {
        files: files.filter((f) => f.path.endsWith(".ras")),
        toml: toml(),
        sheets: sheetNames(),
        js: true,
    };
    compiled = JSON.parse(compile(JSON.stringify(input))) as Compiled;
    const ms = performance.now() - started;
    if (compiled.error) {
        compiled.diagnostics = [{ path: "ras.toml", line: 0, col: 0, endLine: 0, endCol: 0, severity: "error", code: "toml", message: compiled.error }];
    }
    const errors = compiled.diagnostics.filter((d) => d.severity === "error").length;
    const warnings = compiled.diagnostics.filter((d) => d.severity === "warning").length;
    const status = $("status");
    status.innerHTML = "";
    const e = document.createElement("span");
    e.className = errors ? "e" : "";
    e.textContent = `エラー ${errors}`;
    const w = document.createElement("span");
    w.className = warnings ? "w" : "";
    w.textContent = `警告 ${warnings}`;
    status.append(`コンパイル ${ms.toFixed(0)} ms · `, e, " · ", w);
    const count = $("problem-count");
    count.textContent = compiled.diagnostics.length ? String(compiled.diagnostics.length) : "";
    count.className = `count${errors ? " e" : ""}`;
    applyDiagnostics();
    renderVba();
    renderProblems();
    renderEntries();
    renderFiles();
}

function position(doc: EditorState["doc"], line: number, col: number): number {
    if (line + 1 > doc.lines) return doc.length;
    const l = doc.line(line + 1);
    return Math.min(l.from + col, l.to);
}

function applyDiagnostics() {
    if (!editor || !compiled) return;
    const doc = editor.state.doc;
    const list: LintDiagnostic[] = compiled.diagnostics
        .filter((d) => d.path === active)
        .map((d) => {
            const from = position(doc, d.line, d.col);
            let to = position(doc, d.endLine, d.endCol);
            if (to <= from) to = Math.min(from + 1, doc.length);
            return { from, to, severity: d.severity === "hint" ? "hint" : d.severity, message: `${d.message} [${d.code}]`, source: "ras" };
        });
    editor.dispatch(setDiagnostics(editor.state, list));
}

let vbaSelected = "";
function renderVba() {
    const select = $<HTMLSelectElement>("vba-module");
    const list = compiled?.vba || [];
    const prev = vbaSelected || select.value;
    select.textContent = "";
    // The program's modules first, then the run time's (RAS_*).
    const sorted = [...list].sort((a, b) => Number(a.name.startsWith("RAS_")) - Number(b.name.startsWith("RAS_")));
    for (const f of sorted) {
        const o = document.createElement("option");
        o.value = f.fileName;
        o.textContent = f.fileName;
        select.append(o);
    }
    let text: string;
    if (!list.length) {
        text = compiled && compiled.diagnostics.some((d) => d.severity === "error") ? "' エラーがあるため VBA を生成していません（「問題」を参照）" : "' 生成されたモジュールはありません";
        select.disabled = true;
    } else {
        select.disabled = false;
        const fallback = sorted.find((f) => active.startsWith("src/") && f.fileName.toLowerCase().startsWith(active.slice(4).replace(/\.ras$/, "").toLowerCase())) || sorted[0];
        const chosen = sorted.find((f) => f.fileName === prev) || fallback;
        select.value = chosen.fileName;
        text = chosen.text;
    }
    if (vbaView.state.doc.toString() !== text) vbaView.dispatch({ changes: { from: 0, to: vbaView.state.doc.length, insert: text } });
}

function renderProblems() {
    const ul = $("problems");
    ul.textContent = "";
    const list = compiled?.diagnostics || [];
    if (!list.length) {
        const li = document.createElement("li");
        li.className = "empty";
        li.textContent = "問題はありません";
        ul.append(li);
        $("explain").hidden = true;
        return;
    }
    const names: Record<string, string> = { error: "エラー", warning: "警告", info: "情報", hint: "ヒント" };
    for (const d of list) {
        const li = document.createElement("li");
        li.className = d.severity;
        const sev = document.createElement("span");
        sev.className = "sev";
        sev.textContent = names[d.severity] || d.severity;
        const msg = document.createElement("span");
        msg.textContent = d.message;
        const where = document.createElement("span");
        where.className = "where";
        where.textContent = `${d.path ?? ""}${d.path ? `:${d.line + 1}:${d.col + 1}` : ""}  ${d.code}`;
        li.append(sev, msg, where);
        li.onclick = () => {
            for (const x of ul.querySelectorAll(".selected")) x.classList.remove("selected");
            li.classList.add("selected");
            const text = explain(d.code);
            const box = $("explain");
            box.hidden = !text;
            box.textContent = text || "";
            if (d.path && files.some((f) => f.path === d.path)) openFile(d.path, { line: d.line, col: d.col });
        };
        ul.append(li);
    }
}

function renderEntries() {
    const select = $<HTMLSelectElement>("entry");
    const prev = select.value;
    select.textContent = "";
    const entries = compiled?.entries || [];
    for (const e of entries) {
        const o = document.createElement("option");
        o.value = e;
        o.textContent = e;
        select.append(o);
    }
    if (entries.includes(prev)) select.value = prev;
    const blocked = !compiled || !compiled.js;
    select.disabled = blocked || !entries.length;
    $<HTMLButtonElement>("run").disabled = running !== null || blocked || !entries.length;
    $<HTMLButtonElement>("test").disabled = running !== null || blocked || !compiled?.tests;
}

// ---------------------------------------------------------------------------- running

let running: Worker | null = null;

function showTab(name: string) {
    for (const t of document.querySelectorAll<HTMLButtonElement>('[role="tab"]')) t.setAttribute("aria-selected", String(t.dataset.tab === name));
    for (const b of document.querySelectorAll<HTMLElement>(".tab-body")) b.hidden = b.dataset.body !== name;
    if (name === "book") renderBook();
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
}

/** Shows a dialog of the program and returns the answer (`ask` of the worker). */
function showDialog(d: Dialog): Promise<number | string | false> {
    const dialog = $<HTMLDialogElement>("dialog");
    const input = $<HTMLInputElement>("dialog-input");
    const buttons = $("dialog-buttons");
    buttons.textContent = "";
    $("dialog-title").textContent = d.title || "Microsoft Excel";
    $("dialog-prompt").textContent = d.prompt;
    return new Promise((resolve) => {
        const done = (v: number | string | false) => {
            dialog.onclose = null;
            dialog.close();
            resolve(v);
        };
        const add = (label: string, value: () => number | string | false, focus: boolean) => {
            const b = el("button", undefined, label);
            b.type = "button";
            b.onclick = () => done(value());
            buttons.append(b);
            if (focus) setTimeout(() => b.focus(), 0);
            return b;
        };
        if (d.kind === "msgBox") {
            input.hidden = true;
            const sets: [string, number][][] = [
                [["OK", 1]],
                [["OK", 1], ["キャンセル", 2]],
                [["中止", 3], ["再試行", 4], ["無視", 5]],
                [["はい", 6], ["いいえ", 7], ["キャンセル", 2]],
                [["はい", 6], ["いいえ", 7]],
                [["再試行", 4], ["キャンセル", 2]],
            ];
            const set = sets[d.buttons & 7] || sets[0];
            const def = Math.min((d.buttons >> 8) & 3, set.length - 1);
            set.forEach(([label, v], i) => add(label, () => v, i === def));
            // Escape: the cancelling button (or OK, as Windows does).
            const cancel = set.find(([, v]) => v === 2) || (set.length === 1 ? set[0] : null);
            dialog.oncancel = (e) => {
                e.preventDefault();
                if (cancel) done(cancel[1]);
            };
        } else {
            input.hidden = false;
            input.value = d.default ?? "";
            setTimeout(() => input.select(), 0);
            const ok = add("OK", () => input.value, false);
            ok.classList.add("primary");
            add("キャンセル", () => false, false);
            input.onkeydown = (e) => {
                if (e.key === "Enter") {
                    e.preventDefault();
                    done(input.value);
                }
            };
            dialog.oncancel = (e) => {
                e.preventDefault();
                done(false);
            };
        }
        dialog.showModal();
    });
}

function download(name: string, data: string | Uint8Array, type: string) {
    const url = URL.createObjectURL(new Blob([data as BlobPart], { type }));
    const a = el("a");
    a.href = url;
    a.download = name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

function stopRun() {
    if (!running) return;
    running.terminate();
    running = null;
    $("stop").hidden = true;
    renderEntries();
    const out = $("output");
    out.append(el("pre", "err", "実行を止めました"));
}

async function run(kind: "run" | "test") {
    if (running || !compiled || !compiled.js) return;
    const js = compiled.js;
    const entry = $<HTMLSelectElement>("entry").value;
    const testFiles = files.filter((f) => /\.test\.ras$/.test(f.path)).map((f) => f.path);
    if (kind === "run" && !entry) return;
    if (kind === "test" && !testFiles.length) return;
    showTab("output");
    const out = $("output");
    out.textContent = "";
    const bookPath = workbookPath();
    const before: Workbook = { name: bookPath ? bookPath.split("/").pop()! : workbook.name, bytes: workbook.bytes };
    const projectFiles: RunJob["files"] = files.map((f) => ({ path: f.path, data: f.text }));
    if (bookPath) projectFiles.push({ path: bookPath, data: workbook.bytes });
    const targets = kind === "run" ? [entry] : testFiles;
    $("stop").hidden = false;
    let failures = 0;
    let passed = 0;
    const startedAll = performance.now();
    for (const target of targets) {
        const head = el("div", "result", kind === "run" ? `${target} を実行中…` : `${target}`);
        out.append(head);
        const console_ = el("pre");
        const tests = el("ul", "tests");
        if (kind === "test") out.append(tests);
        const job: RunJob = { kind, js, files: projectFiles, target };
        const result = await new Promise<{ result: RunResult; ms: number } | null>((resolve) => {
            const worker = new Worker("worker.js", { type: "module" });
            running = worker;
            renderEntries();
            const shared = typeof SharedArrayBuffer !== "undefined" && crossOriginIsolated ? new SharedArrayBuffer(8 + 65536) : null;
            worker.onmessage = async (e: MessageEvent<FromWorker>) => {
                const m = e.data;
                if (m.type === "output") {
                    if (kind === "run") {
                        if (!console_.isConnected) out.append(console_);
                        const span = el("span", m.stream === "stderr" ? "stderr" : undefined, m.text);
                        console_.append(span);
                    }
                } else if (m.type === "dialog" && shared) {
                    const answer = await showDialog(m.dialog);
                    const bytes = new TextEncoder().encode(JSON.stringify({ value: answer }));
                    const state = new Int32Array(shared, 0, 2);
                    new Uint8Array(shared, 8).set(bytes.subarray(0, 65536));
                    state[1] = Math.min(bytes.length, 65536);
                    Atomics.store(state, 0, 1);
                    Atomics.notify(state, 0);
                } else if (m.type === "done") {
                    worker.terminate();
                    resolve({ result: m.result, ms: m.ms });
                } else if (m.type === "failed") {
                    worker.terminate();
                    out.append(el("pre", "err", m.message));
                    resolve(null);
                }
            };
            worker.onerror = (e) => {
                worker.terminate();
                out.append(el("pre", "err", `ワーカーのエラー: ${e.message}`));
                resolve(null);
            };
            worker.postMessage({ type: "run", module: wasmModule, job, shared } satisfies ToWorker);
            if (!shared && kind === "run") {
                // (Without SharedArrayBuffer the dialogs get the answers `ras run` gives.)
            }
        });
        if (running === null) {
            // Stopped.
            head.textContent = `${target}: 止めました`;
            break;
        }
        running = null;
        if (!result) {
            failures++;
            head.textContent = `${target}: 実行できませんでした`;
            head.className = "result ng";
            continue;
        }
        if (kind === "run") showRunResult(result.result, result.ms, head, before);
        else {
            const r = showTestResult(result.result, tests, head);
            failures += r.failed;
            passed += r.passed;
        }
    }
    $("stop").hidden = true;
    renderEntries();
    if (kind === "test" && running === null) {
        const sum = el("div", `result ${failures ? "ng" : "ok"}`, `${failures ? "失敗あり" : "すべて成功"}: 成功 ${passed}、失敗 ${failures}（${((performance.now() - startedAll) / 1000).toFixed(2)} 秒）`);
        out.prepend(sum);
    }
}

function rasLink(file: string | null, line: number): HTMLElement | string {
    if (!file || !line || !files.some((f) => f.path === file)) return file ? `${file}:${line}` : "";
    const a = el("a", undefined, `${file}:${line}`);
    a.onclick = () => openFile(file, { line: line - 1, col: 0 });
    return a;
}

interface Report {
    ok: boolean;
    error: { code: number; description: string; file: string | null; line: number; stack: string } | null;
    workbooks: { name: string; file: string; source: string | null }[];
    dialogs: { answered: boolean }[];
}

function showRunResult(r: RunResult, ms: number, head: HTMLElement, before: Workbook) {
    const out = $("output");
    const report = r.report as Report | null;
    const ok = r.exitCode === 0 && (!report || report.ok);
    head.className = `result ${ok ? "ok" : "ng"}`;
    head.textContent = `${report ? (report as unknown as { entry: string }).entry : ""}: ${ok ? "正常に終了しました" : "エラーで終了しました"}（${(ms / 1000).toFixed(2)} 秒）`;
    if (report && report.error) {
        const e = report.error;
        const box = el("pre", "err");
        box.append(`実行時エラー ${e.code}: ${e.description}`);
        if (e.file) box.append("（", rasLink(e.file, e.line), "）");
        const frames = String(e.stack || "")
            .split("\n")
            .filter((l) => /^\s+at /.test(l));
        if (frames.length > 1) box.append(`\n${frames.join("\n")}`);
        out.append(box);
    }
    const unanswered = report ? report.dialogs.filter((d) => !d.answered).length : 0;
    if (unanswered && !crossOriginIsolated) out.append(el("p", "hint", `ダイアログ ${unanswered} 個に既定の応答をしました（このブラウザでは、実行中にダイアログを表示できません）`));
    // The workbook the program saved (the project's workbook, changed).
    const saved = report?.workbooks.find((w) => w.source && workbookPath() && w.source === workbookPath()) || report?.workbooks[0];
    const file = saved && r.saved.find((s) => s.path === saved.file);
    if (file) {
        const after: Workbook = { name: saved.name, bytes: file.data };
        lastRun = { before, after };
        out.append(el("h3", undefined, `変更されたセル（${before.name} → 実行後）`));
        let text = "";
        try {
            text = diffText(before.bytes, after.bytes, 200);
        } catch (e) {
            text = `差分を取れませんでした: ${(e as Error).message}`;
        }
        out.append(el("pre", undefined, text || "セルの変更はありません"));
        const view = el("a", undefined, "ブックで見る →");
        view.onclick = () => showTab("book");
        out.append(view);
    } else {
        lastRun = null;
    }
}

interface TestEvent {
    type: string;
    name?: string[];
    status?: string;
    message?: string;
    failure?: { kind: string; message?: string; description?: string; code?: number; location: { file: string; line: number } | null };
    failures?: number;
}

function showTestResult(r: RunResult, list: HTMLElement, head: HTMLElement): { passed: number; failed: number } {
    let passed = 0;
    let failed = 0;
    for (const e of r.events as TestEvent[]) {
        if (e.type === "fatal") {
            failed++;
            list.append(el("li", "failed", e.message));
        }
        if (e.type !== "test") continue;
        const li = el("li", e.status, (e.name || []).join(" > "));
        if (e.status === "passed") passed++;
        if (e.status === "failed") {
            failed++;
            const f = e.failure;
            if (f) {
                const box = el("div", "failure");
                box.append(f.message || `実行時エラー ${f.code}: ${f.description}`);
                if (f.location) box.append("（", rasLink(f.location.file, f.location.line), "）");
                li.append(box);
            }
        }
        list.append(li);
    }
    const errs = r.output.filter((o) => o.stream === "stderr").map((o) => o.text).join("");
    if (!r.events.length && errs) list.after(el("pre", "stderr", errs));
    head.className = `result ${failed ? "ng" : "ok"}`;
    return { passed, failed };
}

// ---------------------------------------------------------------------------- the workbook

function renderBook() {
    const shown = lastRun ? lastRun.after : workbook;
    $("book-name").textContent = lastRun ? `${lastRun.after.name}（実行後。変更されたセルに印が付きます）` : `${workbook.name}（実行前）`;
    let html: string;
    try {
        html = lastRun ? viewPage(shown.bytes, shown.name, lastRun.before.bytes, "実行前") : viewPage(shown.bytes, shown.name, undefined, "");
    } catch (e) {
        html = `<p style="font-family:sans-serif">ブックを表示できません: ${String((e as Error).message).replace(/</g, "&lt;")}</p>`;
    }
    $<HTMLIFrameElement>("book").srcdoc = html;
}

async function readWorkbook(file: File) {
    const bytes = new Uint8Array(await file.arrayBuffer());
    try {
        Book.fromBytes(bytes).free();
    } catch (e) {
        toast(`ブックを読み込めません: ${(e as Error).message}`);
        return;
    }
    workbook = { name: file.name, bytes };
    lastRun = null;
    setWorkbookPath(file.name);
    if (active === "ras.toml") {
        active = "";
        openFile("ras.toml");
    }
    renderBook();
    scheduleCompile(0);
    toast(`${file.name} を読み込みました。プログラムはこのブックで動きます（シート: ${sheetNames().join(", ")}）`);
}

/** `ras pack`: the VBA written into the workbook (the one read into the page, or a new one), saved as .xlsm. */
function packWorkbook() {
    const input = { files: files.filter((f) => f.path.endsWith(".ras")), toml: toml() };
    let packed;
    try {
        packed = packBook(JSON.stringify(input), workbook.bytes, workbook.name);
    } catch (e) {
        toast(String((e as Error).message || e).split("\n")[0]);
        showTab("problems");
        return;
    }
    try {
        const report = JSON.parse(packed.report) as { created: string[]; updated: string[]; removed: string[]; replaced: string[]; warnings: string[] };
        download(packed.name, packed.bytes, "application/vnd.ms-excel.sheet.macroEnabled.12");
        const parts = [`${packed.name} に VBA を書き込みました（作成 ${report.created.length} / 更新 ${report.updated.length} / 削除 ${report.removed.length}）`];
        if (report.replaced.length) parts.push(`手書きの同名モジュールを置き換えました: ${report.replaced.join(", ")}`);
        toast(parts.join("。"));
    } finally {
        packed.free();
    }
}

// ---------------------------------------------------------------------------- the page

let toastTimer = 0;
function toast(text: string) {
    const t = $("toast");
    t.textContent = text;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = window.setTimeout(() => (t.hidden = true), 4000);
}

/** The page is cross-origin isolated through sw.js (for the dialogs); reloads once to be. */
async function isolate() {
    if (crossOriginIsolated || !("serviceWorker" in navigator) || !window.isSecureContext) return;
    try {
        const reg = await navigator.serviceWorker.register("sw.js");
        if (navigator.serviceWorker.controller) return;
        if (sessionStorage.getItem("ras-isolate")) return;
        sessionStorage.setItem("ras-isolate", "1");
        await navigator.serviceWorker.ready;
        if (reg.active) location.reload();
    } catch {
        // The dialogs then get the default answers.
    }
}

async function main() {
    void isolate();
    $("editor").innerHTML = '<div class="loading">コンパイラーを読み込んでいます…</div>';
    const response = fetch("ras_wasm_bg.wasm");
    wasmModule = await (WebAssembly.compileStreaming ? WebAssembly.compileStreaming(response) : response.then((r) => r.arrayBuffer()).then((b) => WebAssembly.compile(b)));
    await init({ module_or_path: wasmModule });
    $("editor").innerHTML = "";

    const sample = $<HTMLSelectElement>("sample");
    const blank = el("option", undefined, "サンプルを開く…");
    blank.value = "";
    sample.append(blank);
    for (const s of samples) {
        const o = el("option", undefined, s.title);
        o.value = s.id;
        sample.append(o);
    }
    workbook = newWorkbook();
    if (!(await fromHash()) && !restore()) loadSample(samples[0].id);
    if (!files.some((f) => f.path === active)) active = files[0].path;
    sample.value = sampleId;
    sample.onchange = () => {
        if (!sample.value) return;
        if (!confirm("サンプルを開くと、いまのファイルは置き換わります。開きますか？")) {
            sample.value = sampleId;
            return;
        }
        loadSample(sample.value);
        history.replaceState(null, "", location.pathname + location.search);
        const a = active;
        active = "";
        openFile(a);
        renderBook();
        runCompile();
        showTab("vba");
    };

    editor = new EditorView({ parent: $("editor") });
    vbaView = new EditorView({
        parent: $("vba"),
        state: EditorState.create({ doc: "", extensions: [common, StreamLanguage.define(vb), EditorState.readOnly.of(true), EditorView.editable.of(false)] }),
    });
    const first = active;
    active = "";
    openFile(first);
    runCompile();

    for (const t of document.querySelectorAll<HTMLButtonElement>('[role="tab"]')) t.onclick = () => showTab(t.dataset.tab!);
    $<HTMLSelectElement>("vba-module").onchange = (e) => {
        vbaSelected = (e.target as HTMLSelectElement).value;
        renderVba();
    };
    $("vba-download").onclick = () => {
        const f = compiled?.vba.find((x) => x.fileName === $<HTMLSelectElement>("vba-module").value);
        if (!f) return;
        // A .bas / .cls is the file the VBE imports (with its Attribute header). The module of a
        // sheet or of ThisWorkbook cannot be imported: its code is pasted into the module.
        // Both in the code page of ras.toml (Shift_JIS by default), as `ras build` writes them.
        const codepage = compiled?.codepage || 932;
        if (f.fileName.endsWith(".doccls")) download(`${f.name}.txt`, encodeVba(f.text.replace(/\n/g, "\r\n"), codepage), "text/plain");
        else download(f.fileName, encodeVba(f.file, codepage), "text/plain");
    };
    $("pack").onclick = packWorkbook;
    $("run").onclick = () => run("run");
    $("test").onclick = () => run("test");
    $("stop").onclick = stopRun;
    $("format").onclick = formatActive;
    $("share").onclick = async () => {
        const url = await shareUrl();
        history.replaceState(null, "", url);
        try {
            await navigator.clipboard.writeText(url);
            toast("このプロジェクトを開く URL をコピーしました（ブックは含みません）");
        } catch {
            toast("アドレスバーの URL でこのプロジェクトを開けます（ブックは含みません）");
        }
    };
    $<HTMLInputElement>("book-file").onchange = (e) => {
        const input = e.target as HTMLInputElement;
        const f = input.files && input.files[0];
        if (f) void readWorkbook(f);
        input.value = "";
    };
    $("book-new").onclick = () => {
        workbook = newWorkbook();
        lastRun = null;
        setWorkbookPath(workbook.name);
        if (active === "ras.toml") {
            active = "";
            openFile("ras.toml");
        }
        renderBook();
        scheduleCompile(0);
    };
    $("book-download").onclick = () => {
        const b = lastRun ? lastRun.after : workbook;
        const type = b.name.endsWith(".xlsm") ? "application/vnd.ms-excel.sheet.macroEnabled.12" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        download(b.name, b.bytes, type);
    };
    document.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && (e.ctrlKey || e.metaKey) && !editor.hasFocus) {
            e.preventDefault();
            void run("run");
        }
    });
}

main().catch((e) => {
    document.body.innerHTML = "";
    const p = el("pre", undefined, `プレイグラウンドを起動できませんでした。\n\n${(e && e.stack) || e}`);
    p.style.padding = "16px";
    document.body.append(p);
});
