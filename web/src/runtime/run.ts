// Runs the debug build of a RAS program in the browser as `ras run` and `ras test` run it with
// Node: the launcher `ras-main.js` (or `ras-test-main.js`) on the stand-in for Excel, whose
// workbook is the WebAssembly `Book` of crates/ras-wasm.

import { Buffer } from "buffer";
import * as path from "path-browserify";
import { MemFs } from "./memfs";
import { NodeWorld, formatConsole } from "./node";

/** The `Book` class of the wasm module (crates/ras-wasm/src/lib.rs). */
export interface WasmBook {
    toBytes(): Uint8Array;
    saveBytes(path: string): Uint8Array;
    readonly savedPath: string;
    call(method: string, args: string): string;
    free(): void;
}
export interface WasmBookClass {
    fromBytes(data: Uint8Array): WasmBook;
    create(): WasmBook;
    isFunction(name: string): boolean;
    methods(): string;
    controls(data: Uint8Array, sheets: string): string;
}

export const ROOT = "/project";
export const DEBUG_DIR = `${ROOT}/build/debug`;

/** The `ras_book` addon as the run time requires it (crates/ras-book/src/node.rs). */
export function bookAddon(Book: WasmBookClass, fs: MemFs) {
    const methods = new Set<string>(JSON.parse(Book.methods()));
    const wrap = (w: WasmBook) =>
        new Proxy(
            {},
            {
                get(_t, name) {
                    if (typeof name !== "string") return undefined;
                    if (name === "toBytes") return () => Buffer.from(w.toBytes());
                    if (name === "save")
                        return (p: string) => {
                            const target = fs.abs(p);
                            const bytes = w.saveBytes(target);
                            const written = w.savedPath;
                            fs.mkdirSync(path.dirname(written), { recursive: true });
                            fs.writeFileSync(written, bytes);
                            return written;
                        };
                    if (methods.has(name))
                        return (...args: unknown[]) => {
                            const r = w.call(name, JSON.stringify(args, (_k, v) => (v === undefined ? null : v)));
                            return JSON.parse(r);
                        };
                    return undefined;
                },
                has(_t, name) {
                    return typeof name === "string" && (methods.has(name) || name === "toBytes" || name === "save");
                },
            },
        );
    return {
        Book: {
            open: (p: string) => wrap(Book.fromBytes(new Uint8Array(fs.readFileSync(p) as Buffer))),
            fromBytes: (data: Uint8Array) => wrap(Book.fromBytes(new Uint8Array(data))),
            create: () => wrap(Book.create()),
            isFunction: (name: string) => Book.isFunction(String(name)),
            // The controls of the sheets of a workbook file (their state, kept while the program runs).
            controls: (p: string, sheets: string[]) => JSON.parse(Book.controls(new Uint8Array(fs.readFileSync(p) as Buffer), JSON.stringify(sheets))),
        },
    };
}

export interface RunJob {
    /** `serve`: the run mode of the book tab (`ras-main.js --serve`, see `ServeSession`). */
    kind: "run" | "test" | "serve";
    /** The debug build: paths relative to build/debug. */
    js: { path: string; text: string }[];
    /** The project's files (sources, the workbook): paths relative to the project. */
    files: { path: string; data: string | Uint8Array }[];
    /** `run`: the entry point (`Module.proc`); `test`: the test file (`test/a.test.ras`). */
    target: string;
    /** `ras run`'s answers of the dialogs (RAS_ANSWERS). */
    answers?: unknown;
    /** `ras test`'s filter of test names. */
    filter?: string;
}

/** A dialog the program shows (`msgBox`, `inputBox`, `application.inputBox`). */
export type Dialog =
    | { kind: "msgBox"; prompt: string; title: string | null; buttons: number }
    | { kind: "inputBox"; prompt: string; title: string | null; default: string | null };

/**
 * Asks the person running the program: the button of a `msgBox` (`vbYes` = 6, ...), the text
 * typed into an `inputBox` (false: cancelled). Undefined: the answer `ras run` gives without one.
 */
export type Ask = (dialog: Dialog) => number | string | false | undefined;

/**
 * The host that draws the user forms of a run (spec §14.10.7), as the runtime's ras-forms.js
 * uses it: `send` a form's state (`show` / `update` / `hide` / `unload` / `error` / `blocked`),
 * `receive` what the user does next (waiting for it). The page is the host in the playground.
 */
export interface FormChannel {
    send(message: unknown): void;
    receive(): unknown;
    /** What the user has done meanwhile, without waiting (null: nothing). */
    poll?(): unknown;
}

const BUTTON_NAMES = ["", "vbOK", "vbCancel", "vbAbort", "vbRetry", "vbIgnore", "vbYes", "vbNo"];

/** Puts `ask` before the dialogs of `ras run` (main.js sets `hooks.msgBox` / `hooks.inputBox`). */
function askingHooks(hooks: Record<string, unknown>, ask: Ask, say: (text: string) => void) {
    for (const name of ["msgBox", "inputBox"] as const) {
        let given: ((...a: unknown[]) => unknown) | undefined = hooks[name] as never;
        Object.defineProperty(hooks, name, {
            configurable: true,
            enumerable: true,
            set(f) {
                given = f;
            },
            get() {
                return (...args: unknown[]) => {
                    const answer =
                        name === "msgBox"
                            ? ask({ kind: "msgBox", prompt: String(args[0]), buttons: Number(args[1]) || 0, title: args[2] === undefined ? null : String(args[2]) })
                            : ask({ kind: "inputBox", prompt: String(args[0]), title: args[1] === undefined ? null : String(args[1]), default: args[2] === undefined || args[2] === null ? null : String(args[2]) });
                    if (answer === undefined) return given ? given(...args) : undefined;
                    if (name === "msgBox") say(`[MsgBox] ${args[0]} → ${BUTTON_NAMES[answer as number] || answer}`);
                    else say(`[InputBox] ${args[0]} → ${answer === false ? "（キャンセル）" : answer}`);
                    return answer;
                };
            },
        });
    }
}

export interface RunResult {
    exitCode: number;
    /** What was printed, in order. */
    output: { stream: "stdout" | "stderr"; text: string }[];
    /** `ras run`'s report (ras-main.js), or null. */
    report: unknown;
    /** `ras test`'s events (ras-test-main.js), one object each. */
    events: unknown[];
    /** The workbooks the program saved: paths relative to the project. */
    saved: { path: string; data: Uint8Array }[];
}

/** Without cross-origin isolation there is no `SharedArrayBuffer`: `sleep` waits busily then. */
function allowWaiting() {
    const g = globalThis as Record<string, unknown>;
    if (typeof g.SharedArrayBuffer === "undefined") g.SharedArrayBuffer = ArrayBuffer;
    const original = Atomics.wait as (arr: Int32Array, index: number, value: number, timeout?: number) => string;
    Atomics.wait = ((arr: Int32Array, index: number, value: number, timeout?: number) => {
        try {
            return original.call(Atomics, arr, index, value, timeout);
        } catch {
            const end = Date.now() + Math.min(timeout ?? 0, 60_000);
            while (Date.now() < end) {
                // waiting
            }
            return "timed-out";
        }
    }) as typeof Atomics.wait;
}

let waiting = false;

/** The program's world: the project in a `MemFs`, `ras-main.js` (or `ras-test-main.js`) to run. */
function prepare(Book: WasmBookClass, job: RunJob, onOutput?: (stream: "stdout" | "stderr", text: string) => void, ask?: Ask, forms?: FormChannel) {
    if (!waiting) {
        allowWaiting();
        waiting = true;
    }
    const fs = new MemFs(ROOT);
    fs.mkdirSync("/tmp", { recursive: true });
    for (const f of job.files) fs.put(path.join(ROOT, f.path), f.data);
    for (const f of job.js) fs.put(path.join(DEBUG_DIR, f.path), f.text);
    const output: RunResult["output"] = [];
    const out = (stream: "stdout" | "stderr", text: string) => {
        const last = output[output.length - 1];
        if (last && last.stream === stream) last.text += text;
        else output.push({ stream, text });
        if (onOutput) onOutput(stream, text);
    };
    const reportFile = "/tmp/ras-run-report.json";
    const eventsFile = "/tmp/ras-test-events.jsonl";
    const env: Record<string, string> = { RAS_BOOK_ADDON: "ras_book", RAS_DIALOGS: "off" };
    let argv: string[];
    let script: string;
    if (job.kind === "run") {
        script = `${DEBUG_DIR}/ras-main.js`;
        argv = [script, job.target];
        env.RAS_RUN_REPORT = reportFile;
        if (job.answers) env.RAS_ANSWERS = JSON.stringify(job.answers);
    } else if (job.kind === "serve") {
        script = `${DEBUG_DIR}/ras-main.js`;
        argv = [script, "--serve", SERVE_DIR];
    } else {
        script = `${DEBUG_DIR}/ras-test-main.js`;
        argv = [script, job.target];
        env.RAS_TEST_EVENTS = eventsFile;
        env.RAS_TEST_OPTIONS = JSON.stringify({ filter: job.filter || null, saveWorkbooks: false });
    }
    const loaded = (id: string, exports: unknown) => {
        if (id !== `${DEBUG_DIR}/ras-runtime.js`) return;
        const hooks = (exports as { hooks: Record<string, unknown> }).hooks;
        if (ask) askingHooks(hooks, ask, (text) => out("stdout", `${text}\n`));
        if (forms) hooks.formHost = forms;
    };
    const world = new NodeWorld({ fs, out, env, argv, extra: { ras_book: bookAddon(Book, fs) }, loaded });
    return { fs, world, script, output, out, reportFile, eventsFile };
}

/** Calls `f` with `console` printing into `out` (as the program's standard output and error). */
function printing<T>(out: (stream: "stdout" | "stderr", text: string) => void, f: () => T): T {
    const c = console as unknown as Record<string, (...a: unknown[]) => void>;
    const saved: Record<string, (...a: unknown[]) => void> = {};
    for (const [name, stream] of [
        ["log", "stdout"],
        ["info", "stdout"],
        ["debug", "stdout"],
        ["warn", "stderr"],
        ["error", "stderr"],
    ] as const) {
        saved[name] = c[name];
        c[name] = (...a: unknown[]) => out(stream, `${formatConsole(a)}\n`);
    }
    try {
        return f();
    } finally {
        Object.assign(c, saved);
    }
}

export function runJob(Book: WasmBookClass, job: RunJob, onOutput?: (stream: "stdout" | "stderr", text: string) => void, ask?: Ask, forms?: FormChannel): RunResult {
    const { fs, world, script, output, out, reportFile, eventsFile } = prepare(Book, job, onOutput, ask, forms);
    const exitCode = printing(out, () => world.run(script));
    let report: unknown = null;
    if (fs.existsSync(reportFile)) {
        try {
            report = JSON.parse(fs.readFileSync(reportFile, "utf8") as string);
        } catch {
            report = null;
        }
    }
    const events: unknown[] = [];
    if (fs.existsSync(eventsFile)) {
        for (const line of (fs.readFileSync(eventsFile, "utf8") as string).split("\n")) {
            if (!line.trim()) continue;
            try {
                events.push(JSON.parse(line));
            } catch {
                // not an event
            }
        }
    }
    const savedBooks: RunResult["saved"] = [];
    for (const [rel, data] of fs.filesUnder(`${ROOT}/build/debug/workbooks`)) {
        savedBooks.push({ path: `build/debug/workbooks/${rel}`, data });
    }
    return { exitCode, output, report, events, saved: savedBooks };
}

// ---------------------------------------------------------------------------- the run mode

/** Where `ras-main.js --serve` writes the workbooks after each request. */
const SERVE_DIR = "/tmp/ras-serve";
/** What starts an answer of `ras-main.js --serve` on its standard output. */
const SERVE_MARK = "\u001eRAS ";

/** What the program answers a request with (`$rt.serve` of the run time). */
export interface ServeReply {
    id: number;
    ok: boolean;
    error?: { code: number; description: string; file: string | null; line: number };
    /** The errors the event handlers did not catch (the program goes on, as in Excel). */
    errors?: { handler: string; code: number; description: string; file: string | null; line: number }[];
    /** The sheet active and its selection, as the program left them. */
    where?: { sheet: string | null; cell: string | null; selection: string | null };
    /** The revision of the project's workbook (it changes with each change of its cells). */
    rev?: number;
    /** The revision of the state of the controls (their values, captions, items). */
    controlsRev?: number;
    cancel?: boolean;
    snapshotError?: string;
}

/** The project's workbook as the program has it after a request (when it changed). */
export interface ServeSnapshot {
    bytes: Uint8Array;
    /** The controls as the program has them (JSON, as `Book.controls` gives it), or null. */
    controls: string | null;
}

/**
 * The run mode of the book tab: `ras-main.js --serve`, kept running in this world. A request
 * (what is done on the sheet, as the page posts it) goes to its standard input; its answer is
 * what it writes on its standard output after the `SERVE_MARK`.
 */
export class ServeSession {
    private readonly world: NodeWorld;
    private readonly fs: MemFs;
    private readonly out: (stream: "stdout" | "stderr", text: string) => void;
    private readonly print: (stream: "stdout" | "stderr", text: string) => void;
    private readonly input = new Map<string, ((...a: unknown[]) => void)[]>();
    private reply: ServeReply | null = null;
    /** What was printed of a line not ended yet. */
    private pending = "";
    private next = 1;
    private rev = -1;
    private controlsRev = -1;
    /** The exit code when the program ended (by itself), or null. */
    exited: number | null = null;

    constructor(Book: WasmBookClass, job: RunJob, onOutput: (stream: "stdout" | "stderr", text: string) => void, ask?: Ask, forms?: FormChannel) {
        this.out = onOutput;
        // (The answers are not printed: they are taken from the lines written.)
        const print = (stream: "stdout" | "stderr", text: string) => (stream === "stdout" ? this.lines(text) : onOutput(stream, text));
        const p = prepare(Book, { ...job, kind: "serve" }, print, ask, forms);
        this.world = p.world;
        this.fs = p.fs;
        const input = this.input;
        const stdin = {
            setEncoding: () => stdin,
            on(event: string, f: (...a: unknown[]) => void) {
                input.set(event, [...(input.get(event) || []), f]);
                return stdin;
            },
            once(event: string, f: (...a: unknown[]) => void) {
                return stdin.on(event, f);
            },
            resume: () => stdin,
            pause: () => stdin,
        };
        this.world.process.stdin = stdin;
        this.print = print;
        const code = printing(print, () => this.world.start(p.script));
        if (code !== undefined) this.exited = this.world.finish(code);
    }

    private line(text: string) {
        const at = text.indexOf(SERVE_MARK);
        if (at < 0) {
            this.out("stdout", `${text}\n`);
            return;
        }
        if (at > 0) this.out("stdout", `${text.slice(0, at)}\n`);
        try {
            this.reply = JSON.parse(text.slice(at + SERVE_MARK.length)) as ServeReply;
        } catch {
            // (Not an answer after all.)
        }
    }

    /** Hands the program a request; gives its answer, with the workbook when it changed. */
    request(m: Record<string, unknown>): { reply: ServeReply; snapshot: ServeSnapshot | null } {
        const ended = (description: string): { reply: ServeReply; snapshot: null } => ({ reply: { id: 0, ok: false, error: { code: 0, description, file: null, line: 0 } }, snapshot: null });
        if (this.exited !== null) return ended(`プログラムが終了しました（終了コード ${this.exited}）。`);
        this.reply = null;
        const code = printing(this.print, () =>
            this.world.guard(() => {
                for (const f of this.input.get("data") || []) f(`${JSON.stringify({ ...m, id: this.next++ })}\n`);
            }),
        );
        if (code !== undefined) {
            this.exited = this.world.finish(code);
            if (this.reply === null) return ended(`プログラムが終了しました（終了コード ${this.exited}）。`);
        }
        const reply = this.reply ?? { id: 0, ok: false, error: { code: 0, description: "プログラムが答えませんでした。", file: null, line: 0 } };
        return { reply, snapshot: this.snapshot(reply) };
    }

    /** What the program printed on its standard output: its answers, and the rest (shown). */
    private lines(text: string) {
        this.pending += text;
        for (let i = this.pending.indexOf("\n"); i >= 0; i = this.pending.indexOf("\n")) {
            this.line(this.pending.slice(0, i));
            this.pending = this.pending.slice(i + 1);
        }
    }

    private snapshot(reply: ServeReply): ServeSnapshot | null {
        if (reply.rev === undefined) return null;
        const controlsRev = reply.controlsRev ?? this.controlsRev;
        if (reply.rev === this.rev && controlsRev === this.controlsRev) return null;
        try {
            const list = JSON.parse(this.fs.readFileSync(`${SERVE_DIR}/snapshot.json`, "utf8") as string) as { books: { file: string; project: boolean }[] };
            const own = list.books.find((b) => b.project);
            if (!own) return null;
            const bytes = new Uint8Array(this.fs.readFileSync(`${SERVE_DIR}/${own.file}`) as Buffer);
            const controlsFile = `${SERVE_DIR}/controls.json`;
            const controls = this.fs.existsSync(controlsFile) ? (this.fs.readFileSync(controlsFile, "utf8") as string) : null;
            this.rev = reply.rev;
            this.controlsRev = controlsRev;
            return { bytes, controls };
        } catch {
            return null;
        }
    }
}
