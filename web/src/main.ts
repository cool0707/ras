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
import init, { Book, compile, diffText, editBook, encodeVba, explain, format, generateTest, newFormSpec, packBook, sheetControls, viewPage, viewRunPage } from "../pkg/ras_wasm.js";
import type { EventDescriptor } from "../../editors/vscode/src/userForm/events";
import { bodyIndent, findHandler } from "../../editors/vscode/src/userForm/events";
import type { RuntimeMessage } from "../../editors/vscode/src/formRunner/protocol";
import { isShapeMacro, macrosOf, moduleName, newMacroName, shapeHandlerText, shapeMacro, sheetControlEvents } from "../../editors/vscode/src/sheetEvents";
import "@vscode/codicons/dist/codicon.css";
import { codeBehindPath, type DesignerHooks, FormDesigner, openHandler } from "./designer";
import { Explorer } from "./explorer";
import { FormPanels } from "./forms";
import { rasLanguage } from "./formula";
import type { Dialog, RunJob, RunResult, ServeReply } from "./runtime/run";
import samples from "virtual:samples";
import type { FromWorker, ToWorker } from "./worker";
import { versioned } from "./versioned";
import { showTestLenses, testLenses } from "./testLens";
import { rasIcons } from "./icons";
import { moduleFileNameError, newModuleSource, vbaModuleName } from "../../editors/vscode/src/moduleNames";

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
    /** The functions "テストを生成" is shown above (`line` from 0) and their test files. */
    testTargets: { path: string; name: string; line: number; testPath: string }[];
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
/** The files open in the editor's tabs, in their order. */
let openTabs: string[] = [];
let sampleId = "";
let workbook: Workbook;
/** The workbook as the last run left it, and what it was before. */
let lastRun: { before: Workbook; after: Workbook } | null = null;
/** The book tab shows the workbook before the run (where it is edited), not the one after it. */
let showBefore = false;
/** The bytes the workbook had before the edits made on the book tab (`redo`: after those undone). */
const bookHistory: { undo: Uint8Array[]; redo: Uint8Array[] } = { undo: [], redo: [] };
let compiled: Compiled | null = null;
let wasmModule: WebAssembly.Module;
const states = new Map<string, EditorState>();
let sheetsCache: { bytes: Uint8Array; sheets: string[] } | null = null;
let controlsCache: { bytes: Uint8Array; controls: unknown } | null = null;
/** The forms shown as their JSON text rather than in the Designer. */
const formsAsText = new Set<string>();
let designer: FormDesigner;
let formPanels: FormPanels;
let explorer: Explorer;

/** An empty workbook, named as `workbook` of ras.toml names it (the project's name, `Hello.xlsx`). */
function newWorkbook(): Workbook {
    const b = Book.create();
    try {
        return { name: workbookPath()?.split("/").pop() || "Book.xlsx", bytes: b.toBytes() };
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

/** The ActiveX controls on the sheets of the workbook (what `@on("CommandButton1.Click")` refers to). */
function workbookControls(): unknown {
    if (controlsCache && controlsCache.bytes === workbook.bytes) return controlsCache.controls;
    let controls: unknown = null;
    try {
        controls = JSON.parse(sheetControls(workbook.bytes));
    } catch {
        controls = null;
    }
    controlsCache = { bytes: workbook.bytes, controls };
    return controls;
}

const isForm = (path: string) => /^forms\/[^/]+\.form\.json$/i.test(path);

/** The files the compiler reads: the .ras files and the forms' designs. */
function sourceFiles(): ProjectFile[] {
    return files.filter((f) => f.path.endsWith(".ras") || isForm(f.path));
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
        localStorage.setItem(STORE, JSON.stringify({ files, active, open: openTabs, sample: sampleId, book }));
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
        openTabs = Array.isArray(saved.open) ? saved.open.filter((p: unknown) => typeof p === "string") : [];
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
        openTabs = [];
        sampleId = "";
        return files.length > 0;
    } catch {
        toast("URL のプロジェクトを読み取れませんでした");
        return false;
    }
}

/** A new project, as `ras init --workbook Playground.xlsx` makes it: ras.toml, src/main.ras and an empty workbook. */
const NEW_PROJECT: ProjectFile[] = [
    { path: "ras.toml", text: '[project]\nname = "Playground"\nworkbook = "Playground.xlsx"\n\n[build]\nopt-level = 1\n' },
    {
        path: "src/main.ras",
        text: "/**\n * サンプルのマクロ。Excel の「マクロ」ダイアログから `hello` を実行できる。\n */\nexport function hello() {\n    const names = [\"RAS\", \"VBA\"];\n    const greeting = names.map(n => `Hello, ${n}!`).join(\"\\n\");\n    msgBox(greeting);\n}\n",
    },
];

function newProject() {
    files = NEW_PROJECT.map((f) => ({ ...f }));
    sampleId = "";
    active = "src/main.ras";
    openTabs = [];
    workbook = newWorkbook();
    lastRun = null;
    clearBookHistory();
    states.clear();
}

function loadSample(id: string) {
    const s = samples.find((x) => x.id === id) || samples[0];
    files = s.files.map((f) => ({ ...f }));
    sampleId = s.id;
    active = files.find((f) => f.path.startsWith("src/"))?.path || files[0].path;
    openTabs = [];
    workbook = newWorkbook();
    lastRun = null;
    clearBookHistory();
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
    // Excel formulas (`fx`…``): references and sheets.
    { tag: tags.special(tags.variableName), color: "var(--hl-ref, #0b6bcb)" },
    { tag: tags.namespace, color: "var(--hl-type, #1f6feb)" },
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
    if (path.endsWith(".json")) return javascript();
    // RAS is close enough to TypeScript for its highlighting; the formulas of `fx` templates
    // are Excel's (spec §2.7).
    return rasLanguage();
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
            lenses,
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
                if (file && isForm(file.path)) designer.changed(file.path);
                scheduleCompile();
            }),
        ],
    });
}

let editor: EditorView;
let vbaView: EditorView;

/** "テストを生成" above the exported functions of src/ (spec §14.12.10). */
const lenses = testLenses((name) => generateTestOf(active, name));

function openFile(path: string, at?: { line: number; col: number }) {
    const f = files.find((x) => x.path === path);
    if (!f) return;
    if (active && editor && !$("editor").hidden) states.set(active, editor.state);
    if (!openTabs.includes(path)) openTabs.splice(active && openTabs.includes(active) ? openTabs.indexOf(active) + 1 : openTabs.length, 0, path);
    active = path;
    explorer.reveal(path);
    $("editor-empty").hidden = true;
    // A form is shown in the Designer (its JSON when asked for, or at a place of its text).
    if (isForm(path) && at) formsAsText.add(path);
    if (isForm(path) && !formsAsText.has(path)) {
        $("editor").hidden = true;
        designer.show(path);
        renderFiles();
        save();
        return;
    }
    designer.hide();
    $("editor").hidden = false;
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

function removeFile(f: ProjectFile) {
    files = files.filter((y) => y !== f);
    states.delete(f.path);
    formsAsText.delete(f.path);
    closeTab(f.path);
    renderFiles();
    scheduleCompile(0);
}

/** The tab of a file closed (the file stays in the project); the one next to it is opened. */
function closeTab(path: string) {
    const i = openTabs.indexOf(path);
    if (i < 0) return;
    openTabs.splice(i, 1);
    if (active !== path) {
        renderFiles();
        save();
        return;
    }
    if (!$("editor").hidden) states.set(path, editor.state);
    const next = openTabs[Math.min(i, openTabs.length - 1)];
    if (next) {
        active = "";
        openFile(next);
        return;
    }
    // No file open: the editor is empty, as in VS Code.
    active = "";
    designer.hide();
    $("editor").hidden = true;
    $("editor-empty").hidden = false;
    renderFiles();
    save();
}

/** The project's files (in the explorer) and the files open (the editor's tabs). */
function renderFiles() {
    openTabs = openTabs.filter((p) => files.some((f) => f.path === p));
    explorer.render();
    renderTabs();
}

function errorCounts(): Map<string, number> {
    const counts = new Map<string, number>();
    for (const d of compiled?.diagnostics || []) if (d.severity === "error" && d.path) counts.set(d.path, (counts.get(d.path) || 0) + 1);
    return counts;
}

function renderTabs() {
    const nav = $("files");
    nav.textContent = "";
    const errors = errorCounts();
    for (const path of openTabs) {
        const f = files.find((x) => x.path === path)!;
        const b = document.createElement("button");
        b.className = `file${errors.has(f.path) ? " has-error" : ""}`;
        b.setAttribute("aria-selected", String(f.path === active));
        b.title = f.path === "ras.toml" ? "プロジェクトの設定" : isForm(f.path) ? `${f.path}（フォーム。開いているときにもう一度クリックすると、デザイナーと JSON を切り替えます）` : f.path;
        const name = f.path.slice(f.path.lastIndexOf("/") + 1);
        b.append(name);
        // The folder, when another tab has the same name (src/main.ras and objects/main.ras).
        if (openTabs.some((p) => p !== path && p.slice(p.lastIndexOf("/") + 1) === name) && path.includes("/")) {
            const dir = document.createElement("span");
            dir.className = "dir";
            dir.textContent = path.slice(0, path.lastIndexOf("/"));
            b.append(dir);
        }
        const x = document.createElement("span");
        x.className = "x codicon codicon-close";
        x.title = "閉じる";
        x.setAttribute("aria-label", `${f.path} を閉じる`);
        x.onclick = (e) => {
            e.stopPropagation();
            closeTab(f.path);
        };
        b.append(x);
        b.onclick = () => {
            // The form open: between the Designer and its JSON.
            if (f.path === active && isForm(f.path)) {
                if (formsAsText.has(f.path)) formsAsText.delete(f.path);
                else formsAsText.add(f.path);
            }
            openFile(f.path);
        };
        // The middle button closes the tab.
        b.onauxclick = (e) => {
            if (e.button !== 1) return;
            e.preventDefault();
            closeTab(f.path);
        };
        nav.append(b);
        if (f.path === active) b.scrollIntoView({ block: "nearest", inline: "nearest" });
    }
}

/** `ras form new`: a new form, opened in the Designer. */
function addForm(name: string): boolean {
    const taken = (n: string) => files.some((f) => f.path.toLowerCase().startsWith(`forms/${n.toLowerCase()}.`));
    if (!/^[A-Za-z][A-Za-z0-9_]{0,30}$/.test(name)) {
        toast(`${name} はフォーム名に使えません（英字で始まり、英数字と _ からなる 31 文字以内の名前にしてください）`);
        return false;
    }
    if (taken(name)) {
        toast(`フォーム ${name} はもうあります（フォーム名は大文字小文字を区別しません）`);
        return false;
    }
    const spec = `forms/${name}.form.json`;
    files.push({ path: spec, text: newFormSpec(name) });
    files.push({
        path: `forms/${name}.ras`,
        text: `// フォーム ${name} のコードビハインド。\`this\` がフォームで、コントロールは \`this.<名前>\` で参照します。\n// イベントは @on("<コントロール名>.<イベント>") を付けた関数で受け取ります（フォーム自身は UserForm）。\n// デザイナーでコントロールをダブルクリックすると、そのイベントのハンドラーをここに追加します。\n// 表示するには、ほかのモジュールから ${name}.show() を呼びます（「実行」の ${name}.show でも表示できます）。\n`,
    });
    openFile(spec);
    scheduleCompile(0);
    return true;
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

/** The text a new file starts with. */
function templateOf(path: string): string {
    return path.startsWith("test/")
        ? 'describe("", () => {\n    test("", () => {\n        expect(1 + 1).toBe(2);\n    });\n});\n'
        : path.startsWith("objects/")
          ? '@on("Worksheet.Change")\nfunction onChange(target: Excel.Range) {\n    debug.print(target.address);\n}\n'
          : path.startsWith("types/")
            ? ""
            : "export function example() {\n}\n";
}

/**
 * Why `path` (src/…/<name>.ras) cannot be a new standard module: its name, or a module or
 * form of the project with the same VBA name. Undefined when it can.
 */
function moduleNameError(path: string): string | undefined {
    const stem = path.slice(path.lastIndexOf("/") + 1).replace(/\.ras$/i, "");
    const error = moduleFileNameError(stem);
    if (error) return error;
    if (files.some((f) => f.path.toLowerCase() === path.toLowerCase())) return `${path} はもうあります`;
    const vbaName = vbaModuleName(stem);
    const taken = files.some((f) => {
        const m = /^src\/(?:.*\/)?([^/]+)\.ras$/i.exec(f.path) ?? /^forms\/([^/]+)\.form\.json$/i.exec(f.path);
        if (!m || f.path.endsWith(".d.ras")) return false;
        return (f.path.startsWith("forms/") ? m[1] : vbaModuleName(m[1])).toLowerCase() === vbaName.toLowerCase();
    });
    return taken ? `VBA のモジュール名 ${vbaName} は、ほかのモジュールかフォームが使っています` : undefined;
}

/** A file named in the explorer: a .ras file, a standard module (`module`), or a form (its name typed in forms/). */
function createFile(typed: string, module = false): boolean {
    if (module) {
        const path = /\.ras$/i.test(typed) ? typed : `${typed}.ras`;
        const error = moduleNameError(path);
        if (error) {
            toast(error);
            return false;
        }
        const stem = path.slice(path.lastIndexOf("/") + 1, -".ras".length);
        files.push({ path, text: newModuleSource(vbaModuleName(stem)) });
        openFile(path);
        scheduleCompile(0);
        return true;
    }
    if (/^forms\//i.test(typed)) return addForm(typed.slice(6).replace(/(\.form\.json|\.json|\.ras)$/i, ""));
    // The extension may be left out (util → util.ras).
    const path = validPath(/\.[^/]*$/.test(typed) ? typed : `${typed}.ras`);
    if (!path) return false;
    files.push({ path, text: templateOf(path) });
    openFile(path);
    scheduleCompile(0);
    return true;
}

function renameFile(f: ProjectFile, to: string): boolean {
    if (to === f.path) return true;
    // A name differing only in case is not taken by the file itself.
    const others = files;
    files = files.filter((x) => x !== f);
    const path = validPath(to);
    files = others;
    if (!path) return false;
    states.delete(f.path);
    const wasActive = active === f.path;
    openTabs = openTabs.map((p) => (p === f.path ? path : p));
    f.path = path;
    if (wasActive) {
        active = "";
        openFile(path);
    }
    renderFiles();
    scheduleCompile(0);
    return true;
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

/** The project as `compile` and `generateTest` take it. */
function compileInput() {
    return {
        files: sourceFiles(),
        toml: toml(),
        sheets: sheetNames(),
        sheetControls: workbookControls() ?? undefined,
        js: true,
    };
}

/**
 * Adds the skeleton of a test of the function `name` of the file `path` to its test file
 * (`test/<module>.test.ras`, created when there is none) and opens it at the test.
 */
function generateTestOf(path: string, name: string) {
    const r = JSON.parse(generateTest(JSON.stringify(compileInput()), path, name)) as
        | { path: string; created: boolean; text: string; line: number; error?: undefined }
        | { error: string };
    if (r.error !== undefined) {
        toast(r.error);
        return;
    }
    const existing = files.find((f) => f.path === r.path);
    if (existing) existing.text = r.text;
    else files.push({ path: r.path, text: r.text });
    states.delete(r.path);
    openFile(r.path, { line: r.line, col: 4 });
    scheduleCompile(0);
    toast(r.created ? `${r.path} を作成して ${name} のテストを追加しました` : `${r.path} に ${name} のテストを追加しました`);
}

function runCompile() {
    const started = performance.now();
    const input = compileInput();
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
    showTestLenses(editor, (compiled.testTargets || []).filter((t) => t.path === active));
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

/** Hands the run (waiting on `shared`) a value as JSON: [0] 1 when given, [1] its length, then its bytes. */
function answer(shared: SharedArrayBuffer, value: unknown) {
    const bytes = new TextEncoder().encode(JSON.stringify(value));
    const state = new Int32Array(shared, 0, 2);
    new Uint8Array(shared, 8).set(bytes.subarray(0, 65536));
    state[1] = Math.min(bytes.length, 65536);
    Atomics.store(state, 0, 1);
    Atomics.notify(state, 0);
}

function stopRun() {
    if (!running) return;
    running.terminate();
    formPanels.end();
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
    // (The forms and dialogs are the run's now: the run mode of the book tab ends.)
    if (bookRun) endBookRun("実行モードを終了しました。");
    showTab("output");
    const out = $("output");
    out.textContent = "";
    const bookPath = workbookPath();
    const before: Workbook = { name: bookPath ? bookPath.split("/").pop()! : workbook.name, bytes: workbook.bytes };
    const projectFiles = runFiles();
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
            const worker = new Worker(versioned("worker.js"), { type: "module" });
            running = worker;
            renderEntries();
            const shared = typeof SharedArrayBuffer !== "undefined" && crossOriginIsolated ? new SharedArrayBuffer(8 + 65536) : null;
            // The forms the run shows are drawn here; what the user does on them is handed back.
            const forms = kind === "run" && shared ? new SharedArrayBuffer(8 + 65536) : null;
            if (forms) formPanels = new FormPanels((action) => answer(forms, action));
            worker.onmessage = async (e: MessageEvent<FromWorker>) => {
                const m = e.data;
                if (m.type === "output") {
                    if (kind === "run") {
                        if (!console_.isConnected) out.append(console_);
                        const span = el("span", m.stream === "stderr" ? "stderr" : undefined, m.text);
                        console_.append(span);
                    }
                } else if (m.type === "dialog" && shared) {
                    answer(shared, { value: await showDialog(m.dialog) });
                } else if (m.type === "form") {
                    formPanels.receive(m.message as RuntimeMessage);
                } else if (m.type === "formWait") {
                    formPanels.wait();
                } else if (m.type === "formPoll") {
                    formPanels.poll();
                } else if (m.type === "done") {
                    worker.terminate();
                    formPanels.end();
                    resolve({ result: m.result, ms: m.ms });
                } else if (m.type === "failed") {
                    worker.terminate();
                    formPanels.end();
                    out.append(el("pre", "err", m.message));
                    resolve(null);
                }
            };
            worker.onerror = (e) => {
                worker.terminate();
                formPanels.end();
                out.append(el("pre", "err", `ワーカーのエラー: ${e.message}`));
                resolve(null);
            };
            worker.postMessage({ type: "run", module: wasmModule, job, shared, forms } satisfies ToWorker);
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
        if (kind === "run") {
            showRunResult(result.result, result.ms, head, before);
            if (!crossOriginIsolated && out.textContent?.includes("ユーザーフォームを表示できません")) out.prepend(notIsolatedHint());
        } else {
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

/** The project's files as a run has them (the workbook with them). */
function runFiles(): RunJob["files"] {
    const bookPath = workbookPath();
    const projectFiles: RunJob["files"] = files.map((f) => ({ path: f.path, data: f.text }));
    if (bookPath) projectFiles.push({ path: bookPath, data: workbook.bytes });
    return projectFiles;
}

/**
 * Why a run could not show its user forms in this page: it is not cross-origin isolated (sw.js
 * did not serve it, as after a hard reload or with the DevTools' "Disable cache"), and how to be.
 */
function notIsolatedHint(): HTMLElement {
    const box = el("div", "result ng", "このページはいまクロスオリジン分離されていないため、ユーザーフォームを表示できません（MsgBox / InputBox も既定の応答になります）。スーパーリロードや開発者ツールの「キャッシュを無効にする」の後に起こります。");
    const reload = el("button", undefined, "再読み込み");
    reload.onclick = () => location.reload();
    box.append(" ", reload);
    return box;
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
        showBefore = false;
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
    // (Another workbook than the one the run mode started with: the run mode ends.)
    if (bookRun && bookRun.workbook !== workbook) endBookRun("ブックが変わったため、実行モードを終了しました。", false);
    const after = lastRun && !showBefore ? lastRun : null;
    $("book-which").hidden = !lastRun;
    for (const b of document.querySelectorAll<HTMLButtonElement>("#book-which button")) b.setAttribute("aria-pressed", String((b.dataset.which === "after") === !!after));
    $("book-name").textContent = after
        ? `${after.after.name}（実行後。変更されたセルに印が付きます。編集は「実行前」で）`
        : bookRun
          ? `${workbook.name}（実行モード。シートでの操作がプログラムのイベントを起こします。ブックは変わりません）`
          : `${workbook.name}（実行前。編集すると、プログラムはこのブックで動きます）`;
    let html: string;
    try {
        if (bookRun && !after) {
            showBookPage(viewRunPage(bookRun.bytes, workbook.name, workbook.bytes, bookRun.sheet ?? undefined, bookRun.controls ?? undefined));
            return;
        }
        // The workbook before the run is edited on the page: it posts its edits here (bookMessage).
        html = after ? viewPage(after.after.bytes, after.after.name, after.before.bytes, "実行前", false) : viewPage(workbook.bytes, workbook.name, undefined, "", true);
    } catch (e) {
        html = `<p style="font-family:sans-serif">ブックを表示できません: ${String((e as Error).message).replace(/</g, "&lt;")}</p>`;
    }
    showBookPage(html);
}

/** The page drawn again, made under the one shown: it takes its place when it is drawn. */
let incomingBook: HTMLIFrameElement | null = null;

function showBookPage(html: string) {
    const shown = $<HTMLIFrameElement>("book");
    incomingBook?.remove();
    incomingBook = null;
    // (Nothing to keep on the screen: a page not drawn yet, or the tab hidden.)
    if (!shown.srcdoc || !shown.offsetParent) {
        shown.srcdoc = html;
        return;
    }
    const next = shown.cloneNode(false) as HTMLIFrameElement;
    next.removeAttribute("id");
    next.classList.add("incoming");
    next.addEventListener(
        "load",
        () => {
            if (incomingBook !== next) return;
            incomingBook = null;
            const focused = document.activeElement === shown;
            shown.remove();
            next.id = "book";
            next.classList.remove("incoming");
            if (focused) next.contentWindow?.focus();
        },
        { once: true },
    );
    next.srcdoc = html;
    incomingBook = next;
    shown.after(next);
}

/** The window of the book tab's page that posted `source`: the one shown, or the one drawn under it. */
function bookFrameOf(source: MessageEventSource | null): Window | null {
    const shown = $<HTMLIFrameElement>("book").contentWindow;
    if (source && source === shown) return shown;
    if (source && incomingBook && source === incomingBook.contentWindow) return incomingBook.contentWindow;
    return null;
}

function clearBookHistory() {
    bookHistory.undo = [];
    bookHistory.redo = [];
}

/** The workbook changed on the book tab: the run's result is of the workbook it was before. */
function bookEdited(bytes: Uint8Array) {
    workbook = { name: workbook.name, bytes };
    lastRun = null;
    showBefore = false;
    save();
    renderBook();
    // (Sheets may have been added, renamed or deleted: their code names are what objects/ refers to.)
    scheduleCompile(0);
}

/** What the page of the book tab posts (as `{ rasView: message }`): edits, undo and redo, links. */
function bookMessage(m: { type?: string; ops?: unknown; url?: string }, frame: Window) {
    const fail = (message: string) => frame.postMessage({ type: "error", message }, "*");
    if (m.type === "link") {
        if (typeof m.url === "string" && /^(https?:\/\/|mailto:)/i.test(m.url)) window.open(m.url, "_blank", "noopener");
        return;
    }
    if (m.type === "ready") {
        // The objects of the workbook before the run have events: their handlers are in objects/.
        // (The playground has no run mode: the page shows no switch for it.)
        if (!(lastRun && !showBefore)) {
            frame.postMessage({ type: "events", on: true, run: !!bookRun, session: bookRun?.id }, "*");
            // What the program reported when the page was drawn again.
            if (bookRun && bookRun.report.length) fail(bookRun.report.splice(0).join("\n"));
        }
        return;
    }
    if (m.type === "mode") {
        switchBookMode((m as { run?: unknown }).run === true, frame).catch((e) => fail(String((e as Error).message || e)));
        return;
    }
    if (m.type === "run") {
        bookRunEvent(m as unknown as { id: number; [key: string]: unknown }, frame).catch((e) => fail(String((e as Error).message || e)));
        return;
    }
    if (bookRun) {
        fail("実行モードでは編集できません。リボンの右端のスイッチで「編集」モードに切り替えてください。");
        return;
    }
    if (m.type === "event") {
        objectEvent(m as unknown as ObjectEvent).catch((e) => fail(String((e as Error).message || e)));
        return;
    }
    if (m.type === "undo" || m.type === "redo") {
        const [from, to] = m.type === "redo" ? [bookHistory.redo, bookHistory.undo] : [bookHistory.undo, bookHistory.redo];
        const bytes = from.pop();
        if (!bytes) {
            fail(m.type === "redo" ? "やり直す操作はありません。" : "元に戻す操作はありません。");
            return;
        }
        to.push(workbook.bytes);
        bookEdited(bytes);
        return;
    }
    if (m.type === "sheet") {
        sheetMessage(m as unknown as SheetMessage, frame);
        return;
    }
    if (m.type !== "edit") return;
    try {
        // (An edit that changed nothing, the same value typed again, is not a step to undo.)
        editWorkbook(m.ops ?? []);
    } catch (e) {
        fail(String((e as Error).message || e));
    }
}

// ---------------------------------------------------------------------------- the run mode of the book tab

/**
 * The run mode of the book tab (the switch at the right end of its ribbon): the program runs in
 * a worker as `ras-main.js --serve` (`ServeSession`), and what is done on the sheet raises its
 * events. The page shows the workbook as the program has it; the workbook itself is not changed.
 */
interface BookRun {
    id: number;
    worker: Worker;
    /** The workbook it started with (another one ends it). */
    workbook: Workbook;
    /** The project's workbook as the program has it, and its controls (JSON). */
    bytes: Uint8Array;
    controls: string | null;
    /** The sheet the program made active. */
    sheet: string | null;
    waiting: Map<number, (r: Extract<FromWorker, { type: "reply" }>) => void>;
    next: number;
    /** The errors to show on the page drawn again. */
    report: string[];
    /** Where its output is shown (the 実行結果 tab). */
    console: HTMLElement;
}

let bookRun: BookRun | null = null;
let bookRunIds = 0;

/** The page's switch: the run mode on (the program started) or off. */
async function switchBookMode(on: boolean, frame: Window) {
    const idle = () => frame.postMessage({ type: "idle" }, "*");
    if (!on) {
        if (bookRun) endBookRun("実行モードを終了しました（プログラムが変えたブックは残りません）。");
        else idle();
        return;
    }
    if (bookRun) return idle();
    if (running) throw new Error("プログラムの実行中は、実行モードにできません。");
    if (!compiled || !compiled.js) throw new Error("コンパイルエラーがあるため、実行モードにできません（「問題」タブを見てください）。");
    if (!workbookPath()) throw new Error("実行モードは、ras.toml の project.workbook のブックで使えます。");
    const out = $("output");
    out.textContent = "";
    const head = el("div", "result", "実行モード（ブック）を開始しています…");
    const console_ = el("pre");
    out.append(head, console_);
    const shared = typeof SharedArrayBuffer !== "undefined" && crossOriginIsolated ? new SharedArrayBuffer(8 + 65536) : null;
    const forms = shared ? new SharedArrayBuffer(8 + 65536) : null;
    if (forms) formPanels = new FormPanels((action) => answer(forms, action));
    const worker = new Worker(versioned("worker.js"), { type: "module" });
    const session: BookRun = { id: ++bookRunIds, worker, workbook, bytes: workbook.bytes, controls: null, sheet: null, waiting: new Map(), next: 1, report: [], console: console_ };
    const started = new Promise<Extract<FromWorker, { type: "reply" }>>((resolve, reject) => {
        session.waiting.set(0, resolve);
        worker.onmessage = async (e: MessageEvent<FromWorker>) => {
            const m = e.data;
            if (m.type === "output") {
                session.console.append(el("span", m.stream === "stderr" ? "stderr" : undefined, m.text));
            } else if (m.type === "dialog" && shared) {
                answer(shared, { value: await showDialog(m.dialog) });
            } else if (m.type === "form") {
                formPanels.receive(m.message as RuntimeMessage);
            } else if (m.type === "formWait") {
                formPanels.wait();
            } else if (m.type === "formPoll") {
                formPanels.poll();
            } else if (m.type === "reply") {
                const done = session.waiting.get(m.id);
                session.waiting.delete(m.id);
                done?.(m);
            } else if (m.type === "failed") {
                reject(new Error(m.message));
                if (bookRun === session) endBookRun(`実行モードのプログラムが止まりました: ${m.message}`);
            }
        };
        worker.onerror = (e) => {
            reject(new Error(`ワーカーのエラー: ${e.message}`));
            if (bookRun === session) endBookRun(`実行モードのプログラムが止まりました: ${e.message}`);
        };
    });
    worker.postMessage({ type: "serve", module: wasmModule, job: { kind: "serve", js: compiled.js, files: runFiles(), target: "" }, shared, forms } satisfies ToWorker);
    let r: Extract<FromWorker, { type: "reply" }>;
    try {
        r = await started;
    } catch (e) {
        worker.terminate();
        formPanels.end();
        head.textContent = "実行モードにできませんでした";
        head.className = "result ng";
        throw e;
    }
    const errors = runReport(r.reply);
    if (!r.reply.ok || r.exited) {
        worker.terminate();
        formPanels.end();
        head.textContent = "実行モードにできませんでした";
        head.className = "result ng";
        throw new Error(errors.join("\n") || "プログラムが終了しました。");
    }
    head.textContent = "実行モード（ブック）: シートでの操作がプログラムのイベントを起こします";
    head.className = "result ok";
    if (r.snapshot) {
        session.bytes = r.snapshot.bytes;
        session.controls = r.snapshot.controls;
    }
    session.sheet = r.reply.where?.sheet ?? null;
    session.report = errors;
    // (Another workbook, or the page shown otherwise, by now.)
    if (workbook !== session.workbook || (lastRun && !showBefore)) {
        worker.terminate();
        formPanels.end();
        return;
    }
    bookRun = session;
    renderBook();
}

/** What the program reported: the error that ended a request, those of its handlers. */
function runReport(reply: ServeReply): string[] {
    const at = (file: string | null, line: number) => (file ? `（${file}:${line}）` : "");
    const list: string[] = [];
    if (!reply.ok && reply.error) list.push(`実行時エラー ${reply.error.code ? `${reply.error.code}: ` : ""}${reply.error.description}${at(reply.error.file, reply.error.line)}`);
    for (const e of reply.errors ?? []) list.push(`${e.handler}: 実行時エラー ${e.code}: ${e.description}${at(e.file, e.line)}`);
    if (reply.snapshotError) list.push(`ブックを書き出せませんでした: ${reply.snapshotError}`);
    // (The 実行結果 tab has them as the program printed them.)
    return list;
}

/** What was done on the sheet in run mode: its events raised, the page drawn again when the workbook changed. */
async function bookRunEvent(m: { id: number; [key: string]: unknown }, frame: Window) {
    const session = bookRun;
    if (!session) throw new Error("実行モードではありません。");
    const { type: _type, id, ...request } = m;
    const n = session.next++;
    const r = await new Promise<Extract<FromWorker, { type: "reply" }>>((resolve) => {
        session.waiting.set(n, resolve);
        session.worker.postMessage({ type: "request", id: n, request } satisfies ToWorker);
    });
    if (bookRun !== session) return;
    const errors = runReport(r.reply);
    if (r.exited) {
        endBookRun(`実行モードのプログラムが終了しました。編集モードに戻ります。${errors.length ? `\n${errors.join("\n")}` : ""}`);
        return;
    }
    const sheet = r.reply.where?.sheet ?? session.sheet;
    if (r.snapshot || sheet !== session.sheet) {
        if (r.snapshot) {
            session.bytes = r.snapshot.bytes;
            session.controls = r.snapshot.controls;
        }
        session.sheet = sheet;
        session.report.push(...errors);
        renderBook();
        return;
    }
    frame.postMessage({ type: "runDone", id, cancel: r.reply.cancel === true }, "*");
    if (errors.length) frame.postMessage({ type: "error", message: errors.join("\n") }, "*");
}

/** The run mode ends: the program is stopped, the workbook shown again (`render`). */
function endBookRun(message: string, render = true) {
    const session = bookRun;
    if (!session) return;
    bookRun = null;
    session.worker.terminate();
    formPanels.end();
    session.console.append(el("span", undefined, `${message}\n`));
    if (render) renderBook();
}

/** What the tabs of the book tab's page ask for: a sheet inserted, removed, renamed or copied. */
interface SheetMessage {
    action: "add" | "copy" | "remove" | "rename";
    sheet: string;
    name?: string;
    codeName?: string;
    color?: string | null;
    /** A copy's place ("シートの移動またはコピー"): before this sheet, or after it. */
    before?: string;
    after?: string;
}

function sheetMessage(m: SheetMessage, frame: Window) {
    // A removed sheet's file in objects/ goes with it.
    const file = m.action === "remove" && m.codeName ? files.find((f) => f.path.toLowerCase() === `objects/${m.codeName}.ras`.toLowerCase()) : undefined;
    if (m.action === "remove" && !confirm(`シート「${m.sheet}」を削除しますか？${file ? `（${file.path} も削除します）` : ""}`)) {
        frame.postMessage({ type: "idle" }, "*");
        return;
    }
    const ops: unknown[] = [{ op: "sheet", action: m.action, sheet: m.sheet, name: m.name ?? "", codeName: m.action === "add" ? (m.codeName ?? "") : "", before: m.before ?? "", after: m.after ?? "" }];
    if (m.action === "add" && m.name && m.color) ops.push({ op: "tabColor", sheet: m.name, color: m.color });
    try {
        editWorkbook(ops);
    } catch (e) {
        frame.postMessage({ type: "error", message: String((e as Error).message || e) }, "*");
        return;
    }
    if (file) removeFile(file);
}

/** Applies edits to the workbook before the run, as a step to undo. */
function editWorkbook(ops: unknown) {
    const bytes = editBook(workbook.bytes, JSON.stringify(ops));
    if (!sameBytes(bytes, workbook.bytes)) {
        bookHistory.undo.push(workbook.bytes);
        if (bookHistory.undo.length > 50) bookHistory.undo.shift();
        bookHistory.redo = [];
    }
    bookEdited(bytes);
}

// ---------------------------------------------------------------------------- the events of the sheets' objects

/** What the book tab posts for the events of a shape or a control (`choose`: the menu's "…"). */
interface ObjectEvent {
    sheet: string;
    codeName: string;
    object: string;
    /** A form control's VML shape. */
    shape?: number;
    /** An ActiveX control: its Microsoft Forms type (`CommandButton`; empty for another control). */
    activex?: string;
    /** The macro assigned to it (`Sheet1.button1Click`). */
    macro: string;
    /** What runs it, as the page names it: "クリック", "選択" (a list or combo box), "変更" (a spin button, a scroll bar). */
    eventName?: string;
    choose: boolean;
}

/** The file of a sheet: `objects/<CodeName>.ras`, as it is spelled in the project if it is there. */
function sheetFile(codeName: string): string {
    const want = `objects/${codeName}.ras`;
    return files.find((f) => f.path.toLowerCase() === want.toLowerCase())?.path ?? want;
}

/** The files that hold macros (the sheets' and the standard modules), with their modules. */
function macroFiles(): { file: string; module: string }[] {
    return files
        .filter((f) => (/^objects\/[^/]+\.ras$/.test(f.path) && !/^objects\/ThisWorkbook\.ras$/i.test(f.path)) || (/^src\/.+\.ras$/.test(f.path) && !f.path.endsWith(".d.ras")))
        .map((f) => ({ file: f.path, module: f.path.startsWith("objects/") ? f.path.slice(8, -4) : moduleName(f.path) }))
        .sort((a, b) => a.module.localeCompare(b.module));
}

/** The first body line of the function `name` of a text (its declaration when it has none). */
function bodyLineOf(text: string, name: string): number {
    const found = macrosOf(text).find((m) => m.name.toLowerCase() === name.toLowerCase());
    if (!found) return 0;
    const lines = text.split(/\r\n|\r|\n/);
    for (let line = found.line; line < lines.length; line++) if (/\{\s*(\/\/.*)?$/.test(lines[line])) return Math.min(line + 1, lines.length - 1);
    return found.line;
}

/** Opens (or adds) the handler of an event of an object of the sheet, or its macro. */
async function objectEvent(m: ObjectEvent) {
    if (!m.codeName) throw new Error(`シート ${m.sheet} のコード名が分かりません。`);
    const file = sheetFile(m.codeName);
    if (m.activex !== undefined) {
        // An ActiveX control: `@on("CommandButton1.Click")` in the sheet's file (spec §12.2.1).
        const events = m.activex && m.object ? sheetControlEvents(m.activex, m.object) : [];
        if (!events.length) {
            throw new Error(`${m.object || "このコントロール"} は Microsoft Forms 2.0 のコントロールではないため、イベントを選べません。ハンドラー（@on("${m.object || "<名前>"}.<イベント>")）は、VBE が作るイベント プロシージャの宣言どおりの引数で手で書いてください。`);
        }
        let event: EventDescriptor | undefined = events.find((e) => e.isDefault) ?? events[0];
        if (m.choose) {
            const sorted = [...events].sort((a, b) => Number(b.isDefault) - Number(a.isDefault));
            const i = await pick(`${m.object} の開くイベント`, sorted.map((e) => ({ label: e.name })));
            event = i === undefined ? undefined : sorted[i];
        }
        if (event) openHandler(hooks, file, m.object, event);
        return;
    }
    // A shape or a form control: its `@on("<name>.Click")` handler in the sheet's file, whose
    // generated macro (`Sheet1.ボタン1_click`) it runs (`OnAction`); or another macro.
    const label = m.object || "このコントロール";
    const ev = m.eventName ?? "クリック";
    const assign = (full: string) => {
        if (full.toLowerCase() !== m.macro.replace(/^.*!/, "").toLowerCase()) editWorkbook([{ op: "macro", sheet: m.sheet, object: m.object, shape: m.shape, macro: full }]);
    };
    const handler = () => {
        let text = hooks.text(file) ?? "";
        if (!findHandler(text, m.object, "Click")) {
            text += shapeHandlerText(text, m.object, bodyIndent(text, "    "), "\n");
            hooks.write(file, text);
        }
        assign(shapeMacro(m.codeName, m.object));
        hooks.openCode(file, findHandler(text, m.object, "Click")?.bodyLine ?? 0);
    };
    const generated = m.object !== "" && m.macro !== "" && isShapeMacro(m.macro, m.codeName, m.object);
    if (!m.choose && m.macro && !generated) {
        const bare = m.macro.replace(/^.*!/, "");
        const dot = bare.lastIndexOf(".");
        const [module, name] = dot < 0 ? [undefined, bare] : [bare.slice(0, dot), bare.slice(dot + 1)];
        const found = macroFiles().find((f) => (module === undefined || f.module.toLowerCase() === module.toLowerCase()) && macrosOf(hooks.text(f.file) ?? "").some((x) => x.name.toLowerCase() === name.toLowerCase()));
        if (!found) throw new Error(`${label} の${ev}イベントに登録された ${m.macro} は、プロジェクトの RAS のファイルにありません（ブックの VBA だけにある関数かもしれません）。`);
        hooks.openCode(found.file, bodyLineOf(hooks.text(found.file) ?? "", name));
        return;
    }
    if (!m.choose && m.object) {
        handler();
        return;
    }
    const fresh = newMacroName(hooks.text(file) ?? "", m.object, m.shape);
    let target = { file, module: m.codeName, name: fresh };
    let create = true;
    if (m.choose) {
        const choices: { target?: typeof target; create: boolean; handler?: boolean; label: string; description?: string; detail?: string }[] = m.object
            ? [
                  {
                      handler: true,
                      create: false,
                      label: `@on("${m.object}.Click")`,
                      description: file,
                      detail: findHandler(hooks.text(file) ?? "", m.object, "Click") ? `ハンドラーがあります${generated ? "（登録されています）" : ""}` : "ハンドラーを追加します",
                  },
              ]
            : [{ target, create: true, label: `＋ 新しい関数: ${m.codeName}.${fresh}`, description: file }];
        for (const f of macroFiles()) {
            for (const x of macrosOf(hooks.text(f.file) ?? "")) {
                const full = `${f.module}.${x.name}`;
                choices.push({ target: { file: f.file, module: f.module, name: x.name }, create: false, label: full, description: f.file, detail: m.macro.replace(/^.*!/, "").toLowerCase() === full.toLowerCase() ? "登録されています" : undefined });
            }
        }
        const i = await pick(`${label} の${ev}イベントに登録するハンドラー・関数`, choices);
        if (i === undefined) return;
        if (choices[i].handler) {
            handler();
            return;
        }
        target = choices[i].target ?? target;
        create = choices[i].create;
    }
    if (create) {
        const text = hooks.text(target.file) ?? "";
        const separator = !text.trim() ? "" : text.endsWith("\n") ? "\n" : "\n\n";
        hooks.write(target.file, `${text}${separator}export function ${target.name}() {\n    \n}\n`);
    }
    assign(`${target.module}.${target.name}`);
    hooks.openCode(target.file, bodyLineOf(hooks.text(target.file) ?? "", target.name));
}

/** Asks for one of the items (the Designer's and the sheets' menus of events and macros). */
function pick(title: string, items: { label: string; description?: string; detail?: string }[]): Promise<number | undefined> {
    const dialog = $<HTMLDialogElement>("picker");
    const list = $("picker-list");
    $("picker-title").textContent = title;
    list.textContent = "";
    return new Promise((resolve) => {
        let chosen: number | undefined;
        items.forEach((item, i) => {
            const b = el("button", "pick");
            b.type = "button";
            b.append(el("span", "pick-label", item.label));
            if (item.description) b.append(el("span", "pick-desc", item.description));
            if (item.detail) b.append(el("span", "pick-detail", item.detail));
            b.onclick = () => {
                chosen = i;
                dialog.close();
            };
            list.append(b);
        });
        dialog.onclose = () => {
            dialog.onclose = null;
            resolve(chosen);
        };
        dialog.showModal();
        (list.firstElementChild as HTMLElement | null)?.focus();
    });
}

/** The project's files as the Designer and the events of the sheets use them. */
const hooks: DesignerHooks = {
    text: (path) => files.find((f) => f.path === path)?.text,
    write(path, text) {
        const f = files.find((x) => x.path === path);
        if (f) f.text = text;
        else files.push({ path, text });
        states.delete(path);
        if (active === path && !$("editor").hidden) {
            active = "";
            openFile(path);
        }
        renderFiles();
        scheduleCompile(0);
    },
    openText(path, at) {
        formsAsText.add(path);
        active = "";
        openFile(path, at);
    },
    openCode(path, line) {
        if (active === path && editor.state.doc.toString() !== hooks.text(path)) states.delete(path);
        active = "";
        const text = hooks.text(path) ?? "";
        openFile(path, { line, col: text.split("\n")[line]?.length ?? 0 });
    },
    pick,
};

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
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
    clearBookHistory();
    setWorkbookPath(file.name);
    if (active === "ras.toml") {
        active = "";
        openFile("ras.toml");
    }
    renderBook();
    scheduleCompile(0);
    toast(`${file.name} を読み込みました。プログラムはこのブックで動きます（シート: ${sheetNames().join(", ")}）`);
}

/**
 * `ras pack`: the VBA written into a workbook, saved as .xlsm. The book is the one the book tab
 * shows: the one read into the page (or a new one), or after the run, if there was one.
 */
function packWorkbook(book: Workbook) {
    const input = { files: sourceFiles(), toml: toml() };
    let packed;
    try {
        packed = packBook(JSON.stringify(input), book.bytes, book.name);
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

const SPLIT = "ras-playground:split";

/** The splitter between the source and the output: dragged, or moved with the arrow keys. */
function setUpSplitter() {
    const panes = document.querySelector<HTMLElement>(".panes")!;
    const splitter = $("splitter");
    const clamp = (r: number) => {
        const w = panes.clientWidth || 1;
        // Each pane keeps at least 240 px (or a fifth of a narrow page).
        const min = Math.min(240 / w, 0.2);
        return Math.min(Math.max(r, min), 1 - min);
    };
    const apply = (r: number, keep: boolean) => {
        const v = clamp(r);
        panes.style.setProperty("--split", String(v));
        splitter.setAttribute("aria-valuenow", String(Math.round(v * 100)));
        if (keep) {
            try {
                localStorage.setItem(SPLIT, String(v));
            } catch {
                // Not kept without storage.
            }
        }
    };
    let saved = NaN;
    try {
        saved = parseFloat(localStorage.getItem(SPLIT) || "");
    } catch {
        // The default.
    }
    apply(Number.isFinite(saved) ? saved : 0.5, false);
    splitter.setAttribute("aria-valuemin", "0");
    splitter.setAttribute("aria-valuemax", "100");
    const current = () => parseFloat(panes.style.getPropertyValue("--split")) || 0.5;
    splitter.addEventListener("pointerdown", (e) => {
        if (e.button !== 0) return;
        e.preventDefault();
        splitter.setPointerCapture(e.pointerId);
        document.body.classList.add("splitting");
        const rect = panes.getBoundingClientRect();
        const move = (ev: PointerEvent) => apply((ev.clientX - rect.left) / rect.width, false);
        const up = () => {
            splitter.removeEventListener("pointermove", move);
            splitter.removeEventListener("pointerup", up);
            splitter.removeEventListener("pointercancel", up);
            document.body.classList.remove("splitting");
            apply(current(), true);
        };
        splitter.addEventListener("pointermove", move);
        splitter.addEventListener("pointerup", up);
        splitter.addEventListener("pointercancel", up);
    });
    splitter.addEventListener("dblclick", () => apply(0.5, true));
    splitter.addEventListener("keydown", (e) => {
        const step = e.shiftKey ? 0.1 : 0.02;
        if (e.key === "ArrowLeft") apply(current() - step, true);
        else if (e.key === "ArrowRight") apply(current() + step, true);
        else if (e.key === "Home") apply(0, true);
        else if (e.key === "End") apply(1, true);
        else return;
        e.preventDefault();
    });
    // A narrower window keeps both panes usable.
    window.addEventListener("resize", () => apply(current(), false));
}

const EXPLORER = "ras-playground:explorer-pane";

/** The explorer shown or hidden (kept for the next visit). */
function toggleExplorer(show = $("explorer").hidden) {
    $("explorer").hidden = !show;
    $("explorer-splitter").hidden = !show;
    $("explorer-toggle").setAttribute("aria-pressed", String(show));
    try {
        const saved = JSON.parse(localStorage.getItem(EXPLORER) || "{}");
        localStorage.setItem(EXPLORER, JSON.stringify({ ...saved, hidden: !show }));
    } catch {
        // Not kept without storage.
    }
}

/** The explorer's width (dragged on its edge) and whether it is shown. */
function setUpExplorerPane() {
    const pane = $("explorer");
    const splitter = $("explorer-splitter");
    let saved: { width?: number; hidden?: boolean } = {};
    try {
        saved = JSON.parse(localStorage.getItem(EXPLORER) || "{}") || {};
    } catch {
        // The default.
    }
    const clamp = (w: number) => Math.round(Math.min(Math.max(w, 140), Math.max(140, window.innerWidth * 0.5)));
    const apply = (w: number, keep: boolean) => {
        const v = clamp(w);
        pane.style.width = `${v}px`;
        if (!keep) return;
        try {
            const now = JSON.parse(localStorage.getItem(EXPLORER) || "{}") || {};
            localStorage.setItem(EXPLORER, JSON.stringify({ ...now, width: v }));
        } catch {
            // Not kept without storage.
        }
    };
    apply(typeof saved.width === "number" ? saved.width : 220, false);
    // Shown unless hidden before; on a narrow page (a phone), hidden until asked for.
    const show = saved.hidden === undefined ? window.innerWidth > 860 : !saved.hidden;
    pane.hidden = !show;
    splitter.hidden = !show;
    $("explorer-toggle").setAttribute("aria-pressed", String(show));
    $("explorer-toggle").onclick = () => toggleExplorer();
    splitter.addEventListener("pointerdown", (e) => {
        if (e.button !== 0) return;
        e.preventDefault();
        splitter.setPointerCapture(e.pointerId);
        document.body.classList.add("splitting");
        const left = pane.getBoundingClientRect().left;
        const move = (ev: PointerEvent) => apply(ev.clientX - left, false);
        const up = () => {
            splitter.removeEventListener("pointermove", move);
            splitter.removeEventListener("pointerup", up);
            splitter.removeEventListener("pointercancel", up);
            document.body.classList.remove("splitting");
            apply(pane.getBoundingClientRect().width, true);
        };
        splitter.addEventListener("pointermove", move);
        splitter.addEventListener("pointerup", up);
        splitter.addEventListener("pointercancel", up);
    });
    splitter.addEventListener("dblclick", () => apply(220, true));
    splitter.addEventListener("keydown", (e) => {
        const w = pane.getBoundingClientRect().width;
        const step = e.shiftKey ? 40 : 10;
        if (e.key === "ArrowLeft") apply(w - step, true);
        else if (e.key === "ArrowRight") apply(w + step, true);
        else return;
        e.preventDefault();
    });
}

/**
 * The page is cross-origin isolated through sw.js (for the dialogs and the user forms); reloads
 * once to be. A hard reload (Ctrl+Shift+R) bypasses the service worker, so the page has to reload
 * again then. What keeps the reload from looping (when the service worker does not make the page
 * isolated) is the time of the last one, in localStorage: the isolated page is in another
 * browsing context group (COOP), whose sessionStorage the page before it does not see.
 */
async function isolate() {
    if (crossOriginIsolated || !("serviceWorker" in navigator) || !window.isSecureContext) return;
    try {
        const reg = await navigator.serviceWorker.register("sw.js");
        if (navigator.serviceWorker.controller) return;
        const last = Number(localStorage.getItem("ras-isolate"));
        if (last && Math.abs(Date.now() - last) < 10_000) return;
        localStorage.setItem("ras-isolate", String(Date.now()));
        await navigator.serviceWorker.ready;
        if (reg.active) location.reload();
    } catch {
        // The dialogs then get the default answers.
    }
}

/** Shows the project just put in place (a sample, a new one). */
function projectOpened() {
    stopRun();
    history.replaceState(null, "", location.pathname + location.search);
    const a = active;
    active = "";
    openFile(a);
    $("output").innerHTML = '<p class="hint">▶ 実行 で関数を、テスト で test/*.test.ras を、Excel を使わずにブラウザの中で動かします。</p>';
    renderBook();
    runCompile();
    save();
    showTab("vba");
}

async function main() {
    void isolate();
    setUpSplitter();
    $("editor").innerHTML = '<div class="loading">コンパイラーを読み込んでいます…</div>';
    const response = fetch(versioned("ras_wasm_bg.wasm"));
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
        projectOpened();
    };
    $("new").onclick = () => {
        if (!confirm("新しいプロジェクトを作ると、いまのファイルとブックは置き換わります。作りますか？")) return;
        newProject();
        sample.value = "";
        projectOpened();
        toast("新しいプロジェクトを作りました（src/main.ras と空のブック Playground.xlsx）");
    };

    editor = new EditorView({ parent: $("editor") });
    explorer = new Explorer($("tree"), {
        files: () => files.map((f) => f.path),
        active: () => active,
        errors: errorCounts,
        open: (path) => openFile(path),
        create: createFile,
        moduleName: (folder) => {
            for (let i = 1; ; i++) {
                const stem = `module${i}`;
                if (moduleNameError(`${folder}/${stem}.ras`) === undefined) return stem;
            }
        },
        rename: (from, to) => {
            const f = files.find((x) => x.path === from);
            return f ? renameFile(f, to) : false;
        },
        remove: (path) => {
            const f = files.find((x) => x.path === path);
            if (f) removeFile(f);
        },
        fixed: (path) =>
            path === "ras.toml"
                ? { rename: "ras.toml はプロジェクトの設定のため、名前の変更や削除はできません", remove: "ras.toml はプロジェクトの設定のため、名前の変更や削除はできません" }
                : path.startsWith("forms/")
                  ? { rename: "フォームのファイルの名前は変えられません（フォーム名はファイル名と form.name で決まります）" }
                  : {},
        message: toast,
    });
    for (const i of document.querySelectorAll<HTMLElement>(".ras-icon[data-icon]")) i.innerHTML = rasIcons[i.dataset.icon!]?.(`ras-icon-${i.dataset.icon}`) ?? "";
    $("explorer-new-module").onclick = () => explorer.newModule();
    $("explorer-new-file").onclick = () => explorer.newFile();
    $("explorer-new-form").onclick = () => explorer.newFile("forms");
    $("explorer-collapse").onclick = () => explorer.collapseAll();
    setUpExplorerPane();
    designer = new FormDesigner($<HTMLIFrameElement>("designer"), hooks);
    formPanels = new FormPanels(() => undefined);
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
    for (const b of document.querySelectorAll<HTMLButtonElement>("#book-which button")) {
        b.onclick = () => {
            showBefore = b.dataset.which === "before";
            renderBook();
        };
    }
    window.addEventListener("message", (e) => {
        const frame = bookFrameOf(e.source);
        if (!frame) return;
        const m = e.data && (e.data as { rasView?: unknown }).rasView;
        if (m && typeof m === "object") bookMessage(m, frame);
    });
    $("book-download").onclick = () => packWorkbook(lastRun && !showBefore ? lastRun.after : workbook);
    document.addEventListener("keydown", (e) => {
        // Ctrl+B: the explorer shown or hidden, as the side bar of VS Code.
        if (e.key.toLowerCase() === "b" && (e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey) {
            e.preventDefault();
            toggleExplorer();
            return;
        }
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
