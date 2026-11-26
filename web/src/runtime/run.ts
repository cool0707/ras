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
        },
    };
}

export interface RunJob {
    kind: "run" | "test";
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

export function runJob(Book: WasmBookClass, job: RunJob, onOutput?: (stream: "stdout" | "stderr", text: string) => void, ask?: Ask, forms?: FormChannel): RunResult {
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
    let exitCode: number;
    try {
        exitCode = world.run(script);
    } finally {
        Object.assign(c, saved);
    }
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
